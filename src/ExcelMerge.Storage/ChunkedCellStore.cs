using System.Buffers;
using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using ExcelMerge.Domain;
using Microsoft.Win32.SafeHandles;

namespace ExcelMerge.Storage;

/// <summary>
/// An append-only, chunked row store with a disk-backed fixed-width offset index.
/// One writer is serialized while independent readers may run concurrently.
/// </summary>
public sealed class ChunkedCellStore : IDisposable, IAsyncDisposable
{
    private const int FormatVersion = 1;
    private const int IndexHeaderSize = 16;
    private const int IndexEntrySize = 24;
    private const int ChunkHeaderSize = 16;
    private const int RecordHeaderSize = 8;

    private static ReadOnlySpan<byte> IndexMagic => "EMRIDX01"u8;
    private static ReadOnlySpan<byte> ChunkMagic => "EMCELL01"u8;

    private readonly WorkspaceSettings _settings;
    private readonly SafeFileHandle _indexHandle;
    private readonly List<DataChunk> _chunks = new();
    private readonly SemaphoreSlim _appendGate = new(1, 1);
    private readonly RowCache _cache;
    private readonly object _stateSync = new();
    private readonly Action<ChunkedCellStore>? _onDisposed;
    private readonly Lazy<Task> _disposeTask;
    private readonly byte[] _indexEntryBuffer = new byte[IndexEntrySize];

    private TaskCompletionSource? _operationsDrained;
    private Exception? _fault;
    private long _rowCount;
    private int _activeOperations;
    private bool _disposeRequested;

    internal ChunkedCellStore(
        string directoryPath,
        WorkspaceSettings settings,
        Action<ChunkedCellStore>? onDisposed)
    {
        DirectoryPath = Path.GetFullPath(directoryPath);
        _settings = settings;
        _onDisposed = onDisposed;
        _cache = new RowCache(settings.RowCacheCapacity, settings.RowCacheByteLimit);

        EnsureDirectoryIsEmpty(DirectoryPath);
        IndexFilePath = Path.Combine(DirectoryPath, "rows.idx");

        SafeFileHandle? indexHandle = null;
        try
        {
            indexHandle = OpenNewFile(IndexFilePath);
            WriteIndexHeader(indexHandle);
            _indexHandle = indexHandle;
            _chunks.Add(CreateChunk(0));
        }
        catch
        {
            indexHandle?.Dispose();
            TryDeleteCreatedFiles(DirectoryPath);
            throw;
        }

        _disposeTask = new Lazy<Task>(DisposeCoreAsync, LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public string DirectoryPath { get; }

    public string IndexFilePath { get; }

    public long RowCount => Interlocked.Read(ref _rowCount);

    public int ChunkCount
    {
        get
        {
            lock (_stateSync)
            {
                return _chunks.Count;
            }
        }
    }

    public long DataLengthBytes
    {
        get
        {
            lock (_stateSync)
            {
                long length = 0;
                foreach (var chunk in _chunks)
                {
                    length = checked(length + chunk.Length);
                }

                return length;
            }
        }
    }

    public long IndexLengthBytes => checked(IndexHeaderSize + (RowCount * IndexEntrySize));

    public int CachedRowCount => _cache.Count;

    public long CachedRowSizeBytes => _cache.SizeBytes;

    public bool IsDisposed
    {
        get
        {
            lock (_stateSync)
            {
                return _disposeRequested;
            }
        }
    }

    public IReadOnlyList<string> GetDataFilePaths()
    {
        lock (_stateSync)
        {
            return _chunks.Select(static chunk => chunk.Path).ToArray();
        }
    }

    /// <summary>
    /// Appends one complete row. Cancellation or an I/O failure leaves the row
    /// count and index unchanged; partially written data is truncated.
    /// </summary>
    public async ValueTask<RowOffset> AppendRowAsync(
        RowRecord row,
        CancellationToken cancellationToken = default)
    {
        EnterOperation();

        var gateEntered = false;
        byte[]? recordBuffer = null;

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            await _appendGate.WaitAsync(cancellationToken).ConfigureAwait(false);
            gateEntered = true;

            var payloadLength = RowCodec.GetEncodedLength(
                row,
                _settings.MaximumRowSizeBytes,
                cancellationToken);
            var recordLength = checked(RecordHeaderSize + payloadLength);

            recordBuffer = ArrayPool<byte>.Shared.Rent(recordLength);
            var record = recordBuffer.AsMemory(0, recordLength);
            var payload = record.Span[RecordHeaderSize..];
            RowCodec.Encode(row, payload, cancellationToken);

            BinaryPrimitives.WriteInt32LittleEndian(record.Span, payloadLength);
            BinaryPrimitives.WriteUInt32LittleEndian(record.Span[sizeof(int)..], Crc32.Compute(payload));

            cancellationToken.ThrowIfCancellationRequested();
            var chunk = EnsureAppendChunk(recordLength);
            var originalChunkLength = chunk.Length;
            var rowOrdinal = RowCount;
            var indexOffset = GetIndexEntryPosition(rowOrdinal);
            var rowOffset = new RowOffset(
                rowOrdinal,
                chunk.Number,
                originalChunkLength,
                recordLength,
                row.CellCount,
                row.RowIndex);

            try
            {
                await RandomAccess.WriteAsync(
                    chunk.Handle,
                    record,
                    originalChunkLength,
                    cancellationToken).ConfigureAwait(false);

                WriteIndexEntry(_indexEntryBuffer, rowOffset);
                await RandomAccess.WriteAsync(
                    _indexHandle,
                    _indexEntryBuffer,
                    indexOffset,
                    cancellationToken).ConfigureAwait(false);

                lock (_stateSync)
                {
                    chunk.Length = checked(originalChunkLength + recordLength);
                }

                Interlocked.Exchange(ref _rowCount, rowOrdinal + 1);
                return rowOffset;
            }
            catch (Exception appendException)
            {
                try
                {
                    RandomAccess.SetLength(chunk.Handle, originalChunkLength);
                    RandomAccess.SetLength(_indexHandle, indexOffset);
                }
                catch (Exception rollbackException)
                {
                    MarkFaulted(rollbackException);
                    throw new IOException(
                        "The row append failed and the store could not roll back its files.",
                        new AggregateException(appendException, rollbackException));
                }

                ExceptionDispatchInfo.Capture(appendException).Throw();
                throw;
            }
        }
        finally
        {
            if (recordBuffer is not null)
            {
                ArrayPool<byte>.Shared.Return(recordBuffer, clearArray: true);
            }

            if (gateEntered)
            {
                _appendGate.Release();
            }

            ExitOperation();
        }
    }

    public async ValueTask<RowOffset> GetRowOffsetAsync(
        long rowOrdinal,
        CancellationToken cancellationToken = default)
    {
        ValidateRowOrdinal(rowOrdinal);
        EnterOperation();

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            return await ReadRowOffsetCoreAsync(rowOrdinal, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            ExitOperation();
        }
    }

    public async ValueTask<RowRecord> ReadRowAsync(
        long rowOrdinal,
        CancellationToken cancellationToken = default)
    {
        ValidateRowOrdinal(rowOrdinal);
        EnterOperation();

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (_cache.TryGet(rowOrdinal, out var cachedRow))
            {
                return cachedRow;
            }

            var rowOffset = await ReadRowOffsetCoreAsync(rowOrdinal, cancellationToken).ConfigureAwait(false);
            return await ReadRowCoreAsync(rowOffset, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            ExitOperation();
        }
    }

    public async ValueTask<RowRecord> ReadRowAsync(
        RowOffset rowOffset,
        CancellationToken cancellationToken = default)
    {
        EnterOperation();

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var indexedOffset = await ReadRowOffsetCoreAsync(
                rowOffset.RowOrdinal,
                cancellationToken).ConfigureAwait(false);

            if (indexedOffset != rowOffset)
            {
                throw new ArgumentException("The row offset does not match this store's index.", nameof(rowOffset));
            }

            if (_cache.TryGet(rowOffset.RowOrdinal, out var cachedRow))
            {
                return cachedRow;
            }

            return await ReadRowCoreAsync(rowOffset, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            ExitOperation();
        }
    }

    /// <summary>
    /// Reads a stable snapshot of the requested range in append order.
    /// Rows appended after enumeration starts are not included.
    /// </summary>
    public async IAsyncEnumerable<RowRecord> ReadRowsAsync(
        long startRowOrdinal = 0,
        long? count = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ThrowIfUnavailable();

        var snapshotCount = RowCount;
        if (startRowOrdinal < 0 || startRowOrdinal > snapshotCount)
        {
            throw new ArgumentOutOfRangeException(nameof(startRowOrdinal));
        }

        var rowsToRead = count ?? (snapshotCount - startRowOrdinal);
        if (rowsToRead < 0 || rowsToRead > snapshotCount - startRowOrdinal)
        {
            throw new ArgumentOutOfRangeException(nameof(count));
        }

        var end = checked(startRowOrdinal + rowsToRead);
        for (var rowOrdinal = startRowOrdinal; rowOrdinal < end; rowOrdinal++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return await ReadRowAsync(rowOrdinal, cancellationToken).ConfigureAwait(false);
        }
    }

    public bool TryGetCachedRow(long rowOrdinal, out RowRecord row)
    {
        ThrowIfUnavailable();
        return _cache.TryGet(rowOrdinal, out row);
    }

    public void ClearCache()
    {
        ThrowIfUnavailable();
        _cache.Clear();
    }

    /// <summary>
    /// Flushes row data first and the offset index last, making all completed
    /// appends durable to the extent supported by the underlying file system.
    /// </summary>
    public async ValueTask FlushAsync(CancellationToken cancellationToken = default)
    {
        EnterOperation();
        var gateEntered = false;

        try
        {
            await _appendGate.WaitAsync(cancellationToken).ConfigureAwait(false);
            gateEntered = true;

            DataChunk[] chunks;
            lock (_stateSync)
            {
                chunks = _chunks.ToArray();
            }

            foreach (var chunk in chunks)
            {
                cancellationToken.ThrowIfCancellationRequested();
                RandomAccess.FlushToDisk(chunk.Handle);
            }

            cancellationToken.ThrowIfCancellationRequested();
            RandomAccess.FlushToDisk(_indexHandle);
        }
        finally
        {
            if (gateEntered)
            {
                _appendGate.Release();
            }

            ExitOperation();
        }
    }

    public void Dispose() => _disposeTask.Value.GetAwaiter().GetResult();

    public ValueTask DisposeAsync() => new(_disposeTask.Value);

    private async Task<RowOffset> ReadRowOffsetCoreAsync(
        long rowOrdinal,
        CancellationToken cancellationToken)
    {
        if (rowOrdinal >= RowCount)
        {
            throw new ArgumentOutOfRangeException(nameof(rowOrdinal));
        }

        var buffer = ArrayPool<byte>.Shared.Rent(IndexEntrySize);
        try
        {
            var indexPosition = GetIndexEntryPosition(rowOrdinal);
            await ReadExactlyAsync(
                _indexHandle,
                buffer.AsMemory(0, IndexEntrySize),
                indexPosition,
                "The row index ended unexpectedly.",
                cancellationToken).ConfigureAwait(false);

            var source = buffer.AsSpan(0, IndexEntrySize);
            var chunkNumber = BinaryPrimitives.ReadInt32LittleEndian(source);
            var byteOffset = BinaryPrimitives.ReadInt64LittleEndian(source[sizeof(int)..]);
            var byteLength = BinaryPrimitives.ReadInt32LittleEndian(source[12..]);
            var cellCount = BinaryPrimitives.ReadInt32LittleEndian(source[16..]);
            var sourceRowIndex = BinaryPrimitives.ReadInt32LittleEndian(source[20..]);

            RowOffset rowOffset;
            try
            {
                rowOffset = new RowOffset(
                    rowOrdinal,
                    chunkNumber,
                    byteOffset,
                    byteLength,
                    cellCount,
                    sourceRowIndex);
            }
            catch (ArgumentOutOfRangeException exception)
            {
                throw new StorageCorruptionException("The row index contains an invalid offset.", exception);
            }

            if (rowOffset.ByteOffset < ChunkHeaderSize ||
                rowOffset.ByteLength < RecordHeaderSize ||
                rowOffset.ByteLength > (long)_settings.MaximumRowSizeBytes + RecordHeaderSize)
            {
                throw new StorageCorruptionException("The row index contains an invalid record range.");
            }

            var chunk = GetChunk(rowOffset.ChunkNumber);
            long endOffset;
            try
            {
                endOffset = rowOffset.EndByteOffset;
            }
            catch (OverflowException exception)
            {
                throw new StorageCorruptionException("The row offset overflows the chunk address space.", exception);
            }

            if (endOffset > RandomAccess.GetLength(chunk.Handle))
            {
                throw new StorageCorruptionException("The row record extends beyond its data chunk.");
            }

            return rowOffset;
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
        }
    }

    private async Task<RowRecord> ReadRowCoreAsync(
        RowOffset rowOffset,
        CancellationToken cancellationToken)
    {
        var chunk = GetChunk(rowOffset.ChunkNumber);
        var buffer = ArrayPool<byte>.Shared.Rent(rowOffset.ByteLength);

        try
        {
            var record = buffer.AsMemory(0, rowOffset.ByteLength);
            await ReadExactlyAsync(
                chunk.Handle,
                record,
                rowOffset.ByteOffset,
                "The row record ended unexpectedly.",
                cancellationToken).ConfigureAwait(false);

            var payloadLength = BinaryPrimitives.ReadInt32LittleEndian(record.Span);
            var expectedCrc = BinaryPrimitives.ReadUInt32LittleEndian(record.Span[sizeof(int)..]);
            if (payloadLength != rowOffset.ByteLength - RecordHeaderSize)
            {
                throw new StorageCorruptionException("The row record length does not match its index entry.");
            }

            var payload = record.Span[RecordHeaderSize..];
            if (Crc32.Compute(payload) != expectedCrc)
            {
                throw new StorageCorruptionException("The row record checksum is invalid.");
            }

            cancellationToken.ThrowIfCancellationRequested();
            var row = RowCodec.Decode(
                payload,
                rowOffset.CellCount,
                rowOffset.SourceRowIndex,
                cancellationToken);
            _cache.Add(rowOffset.RowOrdinal, row);
            return row;
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
        }
    }

    private DataChunk EnsureAppendChunk(int recordLength)
    {
        lock (_stateSync)
        {
            var current = _chunks[^1];
            var hasRows = current.Length > ChunkHeaderSize;
            var wouldExceedTarget = recordLength > _settings.ChunkSizeBytes - current.Length;

            if (hasRows && wouldExceedTarget)
            {
                current = CreateChunk(_chunks.Count);
                _chunks.Add(current);
            }

            return current;
        }
    }

    private DataChunk GetChunk(int chunkNumber)
    {
        lock (_stateSync)
        {
            if ((uint)chunkNumber >= (uint)_chunks.Count)
            {
                throw new StorageCorruptionException("The row index refers to a missing data chunk.");
            }

            return _chunks[chunkNumber];
        }
    }

    private DataChunk CreateChunk(int chunkNumber)
    {
        var path = Path.Combine(DirectoryPath, $"rows-{chunkNumber:D8}.bin");
        var handle = OpenNewFile(path);

        try
        {
            Span<byte> header = stackalloc byte[ChunkHeaderSize];
            ChunkMagic.CopyTo(header);
            BinaryPrimitives.WriteInt32LittleEndian(header[8..], FormatVersion);
            BinaryPrimitives.WriteInt32LittleEndian(header[12..], chunkNumber);
            RandomAccess.Write(handle, header, 0);
            return new DataChunk(chunkNumber, path, handle, ChunkHeaderSize);
        }
        catch
        {
            handle.Dispose();
            TryDeleteFile(path);
            throw;
        }
    }

    private static void WriteIndexHeader(SafeFileHandle handle)
    {
        Span<byte> header = stackalloc byte[IndexHeaderSize];
        IndexMagic.CopyTo(header);
        BinaryPrimitives.WriteInt32LittleEndian(header[8..], FormatVersion);
        BinaryPrimitives.WriteInt32LittleEndian(header[12..], IndexEntrySize);
        RandomAccess.Write(handle, header, 0);
    }

    private static void WriteIndexEntry(Span<byte> destination, RowOffset rowOffset)
    {
        BinaryPrimitives.WriteInt32LittleEndian(destination, rowOffset.ChunkNumber);
        BinaryPrimitives.WriteInt64LittleEndian(destination[4..], rowOffset.ByteOffset);
        BinaryPrimitives.WriteInt32LittleEndian(destination[12..], rowOffset.ByteLength);
        BinaryPrimitives.WriteInt32LittleEndian(destination[16..], rowOffset.CellCount);
        BinaryPrimitives.WriteInt32LittleEndian(destination[20..], rowOffset.SourceRowIndex);
    }

    private static async ValueTask ReadExactlyAsync(
        SafeFileHandle handle,
        Memory<byte> destination,
        long fileOffset,
        string endOfFileMessage,
        CancellationToken cancellationToken)
    {
        var totalRead = 0;
        while (totalRead < destination.Length)
        {
            var bytesRead = await RandomAccess.ReadAsync(
                handle,
                destination[totalRead..],
                checked(fileOffset + totalRead),
                cancellationToken).ConfigureAwait(false);

            if (bytesRead == 0)
            {
                throw new StorageCorruptionException(endOfFileMessage);
            }

            totalRead += bytesRead;
        }
    }

    private static long GetIndexEntryPosition(long rowOrdinal)
    {
        try
        {
            return checked(IndexHeaderSize + (rowOrdinal * IndexEntrySize));
        }
        catch (OverflowException exception)
        {
            throw new ArgumentOutOfRangeException(nameof(rowOrdinal), exception);
        }
    }

    private void ValidateRowOrdinal(long rowOrdinal)
    {
        if (rowOrdinal < 0 || rowOrdinal >= RowCount)
        {
            throw new ArgumentOutOfRangeException(nameof(rowOrdinal));
        }
    }

    private void EnterOperation()
    {
        lock (_stateSync)
        {
            ObjectDisposedException.ThrowIf(_disposeRequested, this);
            if (_fault is not null)
            {
                throw new InvalidOperationException("The row store is faulted and cannot be used.", _fault);
            }

            _activeOperations++;
        }
    }

    private void ExitOperation()
    {
        TaskCompletionSource? drained = null;
        lock (_stateSync)
        {
            _activeOperations--;
            if (_activeOperations == 0 && _disposeRequested)
            {
                drained = _operationsDrained;
            }
        }

        drained?.TrySetResult();
    }

    private void ThrowIfUnavailable()
    {
        lock (_stateSync)
        {
            ObjectDisposedException.ThrowIf(_disposeRequested, this);
            if (_fault is not null)
            {
                throw new InvalidOperationException("The row store is faulted and cannot be used.", _fault);
            }
        }
    }

    private void MarkFaulted(Exception exception)
    {
        lock (_stateSync)
        {
            _fault ??= exception;
        }
    }

    private async Task DisposeCoreAsync()
    {
        Task waitForOperations;
        lock (_stateSync)
        {
            _disposeRequested = true;
            waitForOperations = _activeOperations == 0
                ? Task.CompletedTask
                : (_operationsDrained ??= new TaskCompletionSource(
                    TaskCreationOptions.RunContinuationsAsynchronously)).Task;
        }

        await waitForOperations.ConfigureAwait(false);

        try
        {
            _cache.Clear();
            _indexHandle.Dispose();

            DataChunk[] chunks;
            lock (_stateSync)
            {
                chunks = _chunks.ToArray();
            }

            foreach (var chunk in chunks)
            {
                chunk.Handle.Dispose();
            }

            _appendGate.Dispose();
        }
        finally
        {
            _onDisposed?.Invoke(this);
        }
    }

    private static SafeFileHandle OpenNewFile(string path) => File.OpenHandle(
        path,
        FileMode.CreateNew,
        FileAccess.ReadWrite,
        FileShare.Read,
        FileOptions.Asynchronous | FileOptions.RandomAccess);

    private static void EnsureDirectoryIsEmpty(string path)
    {
        Directory.CreateDirectory(path);
        using var entries = Directory.EnumerateFileSystemEntries(path).GetEnumerator();
        if (entries.MoveNext())
        {
            throw new IOException($"The row store directory '{path}' is not empty.");
        }
    }

    private static void TryDeleteCreatedFiles(string directoryPath)
    {
        try
        {
            foreach (var path in Directory.EnumerateFiles(directoryPath, "rows-*"))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Workspace disposal owns final cleanup.
        }
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch
        {
            // Workspace disposal owns final cleanup.
        }
    }

    private sealed class DataChunk(
        int number,
        string path,
        SafeFileHandle handle,
        long length)
    {
        public int Number { get; } = number;

        public string Path { get; } = path;

        public SafeFileHandle Handle { get; } = handle;

        public long Length { get; set; } = length;
    }
}

using System.Buffers;
using System.Buffers.Binary;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace ExcelMerge.Storage;

public sealed record ChunkedTextStoreOptions
{
    public long ChunkSizeBytes { get; init; } = WorkspaceOptions.DefaultChunkSizeBytes;

    public int MaximumTextSizeBytes { get; init; } = WorkspaceOptions.DefaultMaximumRowSizeBytes;

    public int CacheCapacity { get; init; } = 1_024;

    public long CacheByteLimit { get; init; } = 16L * 1024 * 1024;

    public bool DeleteOnDispose { get; init; } = true;

    internal ChunkedTextStoreSettings Validate()
    {
        if (ChunkSizeBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(ChunkSizeBytes));
        if (MaximumTextSizeBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaximumTextSizeBytes));
        if (CacheCapacity < 0)
            throw new ArgumentOutOfRangeException(nameof(CacheCapacity));
        if (CacheByteLimit < 0)
            throw new ArgumentOutOfRangeException(nameof(CacheByteLimit));

        return new ChunkedTextStoreSettings(
            ChunkSizeBytes,
            MaximumTextSizeBytes,
            CacheCapacity,
            CacheByteLimit,
            DeleteOnDispose);
    }
}

/// <summary>
/// Append-only UTF-8 text storage backed by chunk files and a fixed-width offset index.
/// The decoded-value cache is bounded by both item count and approximate string size.
/// </summary>
public sealed class ChunkedTextStore : IDisposable, IAsyncDisposable
{
    private const int FormatVersion = 1;
    private const int IndexHeaderSize = 16;
    private const int IndexEntrySize = 16;
    private const int ChunkHeaderSize = 16;
    private static readonly UTF8Encoding Utf8 = new(false, true);
    private static ReadOnlySpan<byte> IndexMagic => "EMTIDX01"u8;
    private static ReadOnlySpan<byte> ChunkMagic => "EMTEXT01"u8;

    private readonly ChunkedTextStoreSettings _settings;
    private readonly SafeFileHandle _indexHandle;
    private readonly List<TextChunk> _chunks = [];
    private readonly Dictionary<long, CacheEntry> _cache = [];
    private readonly LinkedList<long> _cacheLru = [];
    private readonly object _sync = new();
    private readonly Action<ChunkedTextStore>? _onDisposed;
    private long _count;
    private long _cacheSizeBytes;
    private bool _disposed;

    public ChunkedTextStore(
        string directoryPath,
        ChunkedTextStoreOptions? options = null)
        : this(directoryPath, (options ?? new ChunkedTextStoreOptions()).Validate(), null)
    {
    }

    internal ChunkedTextStore(
        string directoryPath,
        ChunkedTextStoreSettings settings,
        Action<ChunkedTextStore>? onDisposed)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(directoryPath);
        DirectoryPath = Path.GetFullPath(directoryPath);
        _settings = settings;
        _onDisposed = onDisposed;
        EnsureDirectoryIsEmpty(DirectoryPath);
        IndexFilePath = Path.Combine(DirectoryPath, "strings.idx");

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
            TryDeleteDirectory(DirectoryPath);
            throw;
        }
    }

    public string DirectoryPath { get; }

    public string IndexFilePath { get; }

    public long Count
    {
        get
        {
            lock (_sync)
            {
                return _count;
            }
        }
    }

    public int ChunkCount
    {
        get
        {
            lock (_sync)
            {
                return _chunks.Count;
            }
        }
    }

    public int CachedTextCount
    {
        get
        {
            lock (_sync)
            {
                return _cache.Count;
            }
        }
    }

    public long CachedTextSizeBytes
    {
        get
        {
            lock (_sync)
            {
                return _cacheSizeBytes;
            }
        }
    }

    public long Append(string value, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(value);
        cancellationToken.ThrowIfCancellationRequested();

        var byteCount = Utf8.GetByteCount(value);
        if (byteCount > _settings.MaximumTextSizeBytes)
        {
            throw new ArgumentOutOfRangeException(
                nameof(value),
                $"Encoded text exceeds the {_settings.MaximumTextSizeBytes:N0}-byte limit.");
        }

        byte[]? rented = null;
        try
        {
            var bytes = byteCount == 0
                ? Span<byte>.Empty
                : (rented = ArrayPool<byte>.Shared.Rent(byteCount)).AsSpan(0, byteCount);
            if (byteCount != 0)
            {
                Utf8.GetBytes(value, bytes);
            }

            lock (_sync)
            {
                ThrowIfDisposed();
                cancellationToken.ThrowIfCancellationRequested();
                var chunk = EnsureAppendChunk(byteCount);
                var dataOffset = chunk.Length;
                var indexOffset = checked(IndexHeaderSize + (_count * IndexEntrySize));
                Span<byte> indexEntry = stackalloc byte[IndexEntrySize];
                BinaryPrimitives.WriteInt32LittleEndian(indexEntry, chunk.Number);
                BinaryPrimitives.WriteInt32LittleEndian(indexEntry[4..], byteCount);
                BinaryPrimitives.WriteInt64LittleEndian(indexEntry[8..], dataOffset);

                try
                {
                    if (byteCount != 0)
                    {
                        RandomAccess.Write(chunk.Handle, bytes, dataOffset);
                    }

                    RandomAccess.Write(_indexHandle, indexEntry, indexOffset);
                    chunk.Length = checked(dataOffset + byteCount);
                    return _count++;
                }
                catch
                {
                    RandomAccess.SetLength(chunk.Handle, dataOffset);
                    RandomAccess.SetLength(_indexHandle, indexOffset);
                    throw;
                }
            }
        }
        finally
        {
            if (rented is not null)
            {
                ArrayPool<byte>.Shared.Return(rented, clearArray: true);
            }
        }
    }

    public string Read(long index, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_sync)
        {
            ThrowIfDisposed();
            if ((ulong)index >= (ulong)_count)
                throw new ArgumentOutOfRangeException(nameof(index));

            if (_cache.TryGetValue(index, out var cached))
            {
                Touch(index, cached);
                return cached.Value;
            }

            Span<byte> indexEntry = stackalloc byte[IndexEntrySize];
            ReadExactly(
                _indexHandle,
                indexEntry,
                checked(IndexHeaderSize + (index * IndexEntrySize)),
                "The text index ended unexpectedly.");
            var chunkNumber = BinaryPrimitives.ReadInt32LittleEndian(indexEntry);
            var byteLength = BinaryPrimitives.ReadInt32LittleEndian(indexEntry[4..]);
            var byteOffset = BinaryPrimitives.ReadInt64LittleEndian(indexEntry[8..]);
            if (chunkNumber < 0 || chunkNumber >= _chunks.Count ||
                byteLength < 0 || byteLength > _settings.MaximumTextSizeBytes ||
                byteOffset < ChunkHeaderSize)
            {
                throw new StorageCorruptionException("The text index contains an invalid offset.");
            }

            var chunk = _chunks[chunkNumber];
            long endOffset;
            try
            {
                endOffset = checked(byteOffset + byteLength);
            }
            catch (OverflowException exception)
            {
                throw new StorageCorruptionException("The text index offset overflows.", exception);
            }

            if (endOffset > RandomAccess.GetLength(chunk.Handle))
                throw new StorageCorruptionException("The text record extends beyond its data chunk.");

            cancellationToken.ThrowIfCancellationRequested();
            byte[]? rented = null;
            try
            {
                var value = string.Empty;
                if (byteLength != 0)
                {
                    rented = ArrayPool<byte>.Shared.Rent(byteLength);
                    var bytes = rented.AsSpan(0, byteLength);
                    ReadExactly(chunk.Handle, bytes, byteOffset, "The text record ended unexpectedly.");
                    value = Utf8.GetString(bytes);
                }

                AddToCache(index, value);
                return value;
            }
            catch (DecoderFallbackException exception)
            {
                throw new StorageCorruptionException("The text record is not valid UTF-8.", exception);
            }
            finally
            {
                if (rented is not null)
                    ArrayPool<byte>.Shared.Return(rented, clearArray: true);
            }
        }
    }

    public void Flush()
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            foreach (var chunk in _chunks)
            {
                RandomAccess.FlushToDisk(chunk.Handle);
            }

            RandomAccess.FlushToDisk(_indexHandle);
        }
    }

    public void Dispose()
    {
        lock (_sync)
        {
            if (_disposed)
                return;
            _disposed = true;
            _cache.Clear();
            _cacheLru.Clear();
            _cacheSizeBytes = 0;
            foreach (var chunk in _chunks)
            {
                chunk.Handle.Dispose();
            }
            _indexHandle.Dispose();
        }

        _onDisposed?.Invoke(this);
        if (_settings.DeleteOnDispose)
        {
            TryDeleteDirectory(DirectoryPath);
        }
    }

    public ValueTask DisposeAsync()
    {
        Dispose();
        return ValueTask.CompletedTask;
    }

    private TextChunk EnsureAppendChunk(int byteLength)
    {
        var current = _chunks[^1];
        if (byteLength != 0 &&
            current.Length > ChunkHeaderSize &&
            checked(current.Length + byteLength) > _settings.ChunkSizeBytes)
        {
            current = CreateChunk(_chunks.Count);
            _chunks.Add(current);
        }

        return current;
    }

    private TextChunk CreateChunk(int number)
    {
        var path = Path.Combine(DirectoryPath, $"strings-{number:D6}.bin");
        var handle = OpenNewFile(path);
        try
        {
            Span<byte> header = stackalloc byte[ChunkHeaderSize];
            header.Clear();
            ChunkMagic.CopyTo(header);
            BinaryPrimitives.WriteInt32LittleEndian(header[8..], FormatVersion);
            BinaryPrimitives.WriteInt32LittleEndian(header[12..], number);
            RandomAccess.Write(handle, header, 0);
            return new TextChunk(number, path, handle, ChunkHeaderSize);
        }
        catch
        {
            handle.Dispose();
            throw;
        }
    }

    private static void WriteIndexHeader(SafeFileHandle handle)
    {
        Span<byte> header = stackalloc byte[IndexHeaderSize];
        header.Clear();
        IndexMagic.CopyTo(header);
        BinaryPrimitives.WriteInt32LittleEndian(header[8..], FormatVersion);
        BinaryPrimitives.WriteInt32LittleEndian(header[12..], IndexEntrySize);
        RandomAccess.Write(handle, header, 0);
    }

    private void AddToCache(long index, string value)
    {
        if (_settings.CacheCapacity == 0 || _settings.CacheByteLimit == 0)
            return;

        var size = EstimateSize(value);
        if (size > _settings.CacheByteLimit)
            return;

        var node = _cacheLru.AddFirst(index);
        _cache.Add(index, new CacheEntry(value, size, node));
        _cacheSizeBytes = checked(_cacheSizeBytes + size);
        while ((_cache.Count > _settings.CacheCapacity || _cacheSizeBytes > _settings.CacheByteLimit) &&
               _cacheLru.Last is { } last)
        {
            var entry = _cache[last.Value];
            _cache.Remove(last.Value);
            _cacheLru.RemoveLast();
            _cacheSizeBytes -= entry.SizeBytes;
        }
    }

    private void Touch(long index, CacheEntry entry)
    {
        _cacheLru.Remove(entry.Node);
        _cacheLru.AddFirst(entry.Node);
    }

    private static long EstimateSize(string value) => checked(24L + (value.Length * sizeof(char)));

    private static SafeFileHandle OpenNewFile(string path) => File.OpenHandle(
        path,
        FileMode.CreateNew,
        FileAccess.ReadWrite,
        FileShare.Read,
        FileOptions.RandomAccess);

    private static void ReadExactly(
        SafeFileHandle handle,
        Span<byte> destination,
        long fileOffset,
        string errorMessage)
    {
        var read = 0;
        while (read < destination.Length)
        {
            var count = RandomAccess.Read(handle, destination[read..], checked(fileOffset + read));
            if (count == 0)
                throw new StorageCorruptionException(errorMessage);
            read += count;
        }
    }

    private static void EnsureDirectoryIsEmpty(string path)
    {
        Directory.CreateDirectory(path);
        if (Directory.EnumerateFileSystemEntries(path).Any())
            throw new IOException($"Text store directory '{path}' is not empty.");
    }

    private void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
                Directory.Delete(path, recursive: true);
        }
        catch
        {
            // Workspace cleanup retries abandoned stores.
        }
    }

    private sealed class TextChunk(
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

    private sealed record CacheEntry(
        string Value,
        long SizeBytes,
        LinkedListNode<long> Node);
}

internal sealed record ChunkedTextStoreSettings(
    long ChunkSizeBytes,
    int MaximumTextSizeBytes,
    int CacheCapacity,
    long CacheByteLimit,
    bool DeleteOnDispose);

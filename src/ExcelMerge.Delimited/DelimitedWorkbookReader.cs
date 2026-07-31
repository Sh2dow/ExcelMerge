using System.Text;
using ExcelMerge.Domain;
using ExcelMerge.Storage;

namespace ExcelMerge.Delimited;

/// <summary>Streams RFC 4180 CSV or tab-delimited text into a workspace row store.</summary>
public sealed class DelimitedWorkbookReader : IDelimitedWorkbookReader
{
    public async ValueTask<WorkbookMetadata> ReadMetadataAsync(
        string filePath,
        DelimitedReaderOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        var scan = await ScanAsync(
            filePath,
            options ?? new DelimitedReaderOptions(),
            progress: null,
            consumeRow: null,
            cancellationToken).ConfigureAwait(false);
        return scan.Metadata;
    }

    public async ValueTask<DelimitedReaderResult> IndexAsync(
        string filePath,
        Workspace workspace,
        DelimitedReaderOptions? options = null,
        IProgress<DelimitedReaderProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        ArgumentNullException.ThrowIfNull(workspace);

        var store = workspace.CreateCellStore();
        try
        {
            var scan = await ScanAsync(
                filePath,
                options ?? new DelimitedReaderOptions(),
                progress,
                async (row, token) =>
                {
                    await store.AppendRowAsync(row, token).ConfigureAwait(false);
                },
                cancellationToken).ConfigureAwait(false);
            await store.FlushAsync(cancellationToken).ConfigureAwait(false);

            var worksheet = new DelimitedReaderWorksheet(scan.Metadata.Sheets[0], store);
            return new DelimitedReaderResult(
                scan.Metadata,
                worksheet,
                scan.Format,
                scan.TextEncoding,
                scan.HasByteOrderMark);
        }
        catch
        {
            await DisposeFailedStoreAsync(store).ConfigureAwait(false);
            throw;
        }
    }

    private static async ValueTask<ScanResult> ScanAsync(
        string filePath,
        DelimitedReaderOptions options,
        IProgress<DelimitedReaderProgress>? progress,
        Func<RowRecord, CancellationToken, ValueTask>? consumeRow,
        CancellationToken cancellationToken)
    {
        var fullPath = Path.GetFullPath(filePath);
        var settings = ReaderSettings.Create(fullPath, options);
        cancellationToken.ThrowIfCancellationRequested();

        await using var stream = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 128 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        var source = new SourceFile(fullPath, stream);
        progress?.Report(new DelimitedReaderProgress(
            DelimitedReaderStage.Detecting,
            ContentLength: source.Length));

        var detectedEncoding = DetectEncoding(stream, fullPath);
        stream.Position = detectedEncoding.PreambleLength;
        using var reader = new StreamReader(
            stream,
            detectedEncoding.Encoding,
            detectEncodingFromByteOrderMarks: false,
            bufferSize: 128 * 1024,
            leaveOpen: true);

        long rowsRead = 0;
        long cellsRead = 0;
        long nonEmptyCells = 0;
        var maximumColumnCount = 0;
        progress?.Report(new DelimitedReaderProgress(
            DelimitedReaderStage.Reading,
            ContentLength: source.Length));

        await DelimitedParser.ParseAsync(
            reader,
            settings.Delimiter,
            fullPath,
            settings.MaximumFieldCharacters,
            settings.MaximumColumns,
            ConsumeRecordAsync,
            cancellationToken).ConfigureAwait(false);

        source.ThrowIfChanged();
        var sheet = new SheetMetadata(
            "delimited:0",
            settings.WorksheetName,
            0,
            FirstRowIndex: rowsRead == 0 ? null : 0,
            LastRowIndex: rowsRead == 0 ? null : checked((int)rowsRead - 1),
            FirstColumnIndex: maximumColumnCount == 0 ? null : 0,
            LastColumnIndex: maximumColumnCount == 0 ? null : maximumColumnCount - 1,
            NonEmptyCellCount: nonEmptyCells);
        var metadata = new WorkbookMetadata(
            source.DisplayName,
            settings.Format == DelimitedFileFormat.Csv ? WorkbookFormat.Csv : WorkbookFormat.Tsv,
            [sheet],
            fullPath,
            source.Length,
            source.LastModifiedUtc);
        progress?.Report(new DelimitedReaderProgress(
            DelimitedReaderStage.Completed,
            rowsRead,
            cellsRead,
            source.Length));
        return new ScanResult(
            metadata,
            settings.Format,
            detectedEncoding.TextEncoding,
            detectedEncoding.HasByteOrderMark);

        async ValueTask ConsumeRecordAsync(string[] fields)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if ((settings.MaximumRows > 0 && rowsRead >= settings.MaximumRows) ||
                rowsRead >= int.MaxValue)
            {
                throw new DelimitedReaderException(
                    DelimitedReaderError.RowLimitExceeded,
                    "The input exceeds the configured row limit.",
                    fullPath,
                    rowsRead + 1);
            }

            CellRecord[]? cells = consumeRow is null ? null : new CellRecord[fields.Length];
            var rowIndex = checked((int)rowsRead);
            for (var columnIndex = 0; columnIndex < fields.Length; columnIndex++)
            {
                var field = fields[columnIndex];
                if (field.Length != 0)
                {
                    nonEmptyCells++;
                }

                if (cells is not null)
                {
                    cells[columnIndex] = new CellRecord(
                        new CellAddress(rowIndex, columnIndex),
                        CellValue.FromText(field, field));
                }
            }

            if (cells is not null)
            {
                await consumeRow!(new RowRecord(rowIndex, cells), cancellationToken)
                    .ConfigureAwait(false);
            }

            rowsRead++;
            cellsRead = checked(cellsRead + fields.Length);
            maximumColumnCount = Math.Max(maximumColumnCount, fields.Length);
            if (rowsRead % settings.ProgressIntervalRows == 0)
            {
                progress?.Report(new DelimitedReaderProgress(
                    DelimitedReaderStage.Reading,
                    rowsRead,
                    cellsRead,
                    source.Length));
            }
        }
    }

    private static DetectedEncoding DetectEncoding(FileStream stream, string sourcePath)
    {
        Span<byte> prefix = stackalloc byte[4];
        stream.Position = 0;
        var count = stream.Read(prefix);
        stream.Position = 0;

        if (count >= 4 &&
            ((prefix[0] == 0xFF && prefix[1] == 0xFE && prefix[2] == 0x00 && prefix[3] == 0x00) ||
             (prefix[0] == 0x00 && prefix[1] == 0x00 && prefix[2] == 0xFE && prefix[3] == 0xFF)))
        {
            throw UnsupportedEncoding(sourcePath);
        }

        if (count >= 3 && prefix[0] == 0xEF && prefix[1] == 0xBB && prefix[2] == 0xBF)
        {
            return new DetectedEncoding(
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
                DelimitedTextEncoding.Utf8,
                HasByteOrderMark: true,
                PreambleLength: 3);
        }

        if (count >= 2 && prefix[0] == 0xFF && prefix[1] == 0xFE)
        {
            return new DetectedEncoding(
                new UnicodeEncoding(bigEndian: false, byteOrderMark: false, throwOnInvalidBytes: true),
                DelimitedTextEncoding.Utf16LittleEndian,
                HasByteOrderMark: true,
                PreambleLength: 2);
        }

        if (count >= 2 && prefix[0] == 0xFE && prefix[1] == 0xFF)
        {
            return new DetectedEncoding(
                new UnicodeEncoding(bigEndian: true, byteOrderMark: false, throwOnInvalidBytes: true),
                DelimitedTextEncoding.Utf16BigEndian,
                HasByteOrderMark: true,
                PreambleLength: 2);
        }

        if (count >= 3 && prefix[0] == 0x2B && prefix[1] == 0x2F && prefix[2] == 0x76)
        {
            throw UnsupportedEncoding(sourcePath);
        }

        return new DetectedEncoding(
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true),
            DelimitedTextEncoding.Utf8,
            HasByteOrderMark: false,
            PreambleLength: 0);
    }

    private static DelimitedReaderException UnsupportedEncoding(string sourcePath) =>
        new(
            DelimitedReaderError.UnsupportedEncoding,
            "Only UTF-8 and byte-order-marked UTF-16 text are supported.",
            sourcePath);

    private static async ValueTask DisposeFailedStoreAsync(ChunkedCellStore store)
    {
        var directoryPath = store.DirectoryPath;
        try
        {
            await store.DisposeAsync().ConfigureAwait(false);
            if (Directory.Exists(directoryPath))
            {
                Directory.Delete(directoryPath, recursive: true);
            }
        }
        catch
        {
            // Preserve the indexing exception; workspace disposal owns final cleanup.
        }
    }

    private sealed record ReaderSettings(
        DelimitedFileFormat Format,
        char Delimiter,
        string WorksheetName,
        int ProgressIntervalRows,
        int MaximumFieldCharacters,
        int MaximumColumns,
        int MaximumRows)
    {
        public static ReaderSettings Create(string fullPath, DelimitedReaderOptions options)
        {
            ArgumentNullException.ThrowIfNull(options);
            if (options.ProgressIntervalRows <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(options.ProgressIntervalRows));
            }

            if (options.MaximumFieldCharacters < 0 ||
                options.MaximumColumns < 0 ||
                options.MaximumRows < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(options));
            }

            if (options.WorksheetName is not null && string.IsNullOrWhiteSpace(options.WorksheetName))
            {
                throw new ArgumentException("The worksheet name cannot be empty.", nameof(options));
            }

            var format = ResolveFormat(fullPath, options.Format);
            var worksheetName = options.WorksheetName ?? Path.GetFileNameWithoutExtension(fullPath);
            if (string.IsNullOrWhiteSpace(worksheetName))
            {
                worksheetName = "Data";
            }

            return new ReaderSettings(
                format,
                format == DelimitedFileFormat.Csv ? ',' : '\t',
                worksheetName,
                options.ProgressIntervalRows,
                options.MaximumFieldCharacters,
                options.MaximumColumns,
                options.MaximumRows);
        }

        private static DelimitedFileFormat ResolveFormat(
            string fullPath,
            DelimitedFileFormat requestedFormat)
        {
            if (requestedFormat is DelimitedFileFormat.Csv or DelimitedFileFormat.Tsv)
            {
                return requestedFormat;
            }

            if (requestedFormat != DelimitedFileFormat.Auto)
            {
                throw new ArgumentOutOfRangeException(nameof(requestedFormat));
            }

            if (DelimitedFormatDetector.TryDetect(fullPath, out var detected))
            {
                return detected;
            }

            throw new DelimitedReaderException(
                DelimitedReaderError.UnsupportedFileExtension,
                "Automatic format detection requires a .csv or .tsv extension.",
                fullPath);
        }
    }

    private sealed class SourceFile
    {
        public SourceFile(string fullPath, FileStream stream)
        {
            FullPath = fullPath;
            var fileInfo = new FileInfo(fullPath);
            DisplayName = fileInfo.Name;
            Length = stream.Length;
            LastWriteTimeUtc = fileInfo.LastWriteTimeUtc;
            LastModifiedUtc = new DateTimeOffset(LastWriteTimeUtc);
        }

        public string FullPath { get; }

        public string DisplayName { get; }

        public long Length { get; }

        public DateTime LastWriteTimeUtc { get; }

        public DateTimeOffset LastModifiedUtc { get; }

        public void ThrowIfChanged()
        {
            var fileInfo = new FileInfo(FullPath);
            fileInfo.Refresh();
            if (!fileInfo.Exists ||
                fileInfo.Length != Length ||
                fileInfo.LastWriteTimeUtc != LastWriteTimeUtc)
            {
                throw new DelimitedReaderException(
                    DelimitedReaderError.SourceChanged,
                    "The source changed while it was being read.",
                    FullPath);
            }
        }
    }

    private readonly record struct DetectedEncoding(
        Encoding Encoding,
        DelimitedTextEncoding TextEncoding,
        bool HasByteOrderMark,
        int PreambleLength);

    private readonly record struct ScanResult(
        WorkbookMetadata Metadata,
        DelimitedFileFormat Format,
        DelimitedTextEncoding TextEncoding,
        bool HasByteOrderMark);
}

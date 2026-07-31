using System.Runtime.CompilerServices;
using System.Globalization;
using System.Text;
using ExcelMerge.Domain;

namespace ExcelMerge.Delimited;

/// <summary>Streams complete worksheet rows to a same-directory transactional CSV or TSV file.</summary>
public sealed class DelimitedContentWriter : IDelimitedContentWriter
{
    private const string RecordSeparator = "\r\n";

    public ValueTask<DelimitedWriteResult> WriteAsync(
        string destinationPath,
        IEnumerable<RowRecord> rows,
        DelimitedWriterOptions? options = null,
        IProgress<DelimitedWriterProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(rows);
        return WriteAsync(
            destinationPath,
            AsAsyncEnumerable(rows),
            options,
            progress,
            cancellationToken);
    }

    public async ValueTask<DelimitedWriteResult> WriteAsync(
        string destinationPath,
        IAsyncEnumerable<RowRecord> rows,
        DelimitedWriterOptions? options = null,
        IProgress<DelimitedWriterProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        ArgumentNullException.ThrowIfNull(rows);
        options ??= new DelimitedWriterOptions();
        ValidateOptions(options);

        var fullPath = Path.GetFullPath(destinationPath);
        var format = ResolveFormat(fullPath, options.Format);
        var destinationDirectory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(destinationDirectory) || !Directory.Exists(destinationDirectory))
        {
            throw new DirectoryNotFoundException(
                $"The destination directory '{destinationDirectory}' does not exist.");
        }

        if (!options.Overwrite && File.Exists(fullPath))
        {
            throw new DelimitedWriterException(
                DelimitedWriterError.DestinationExists,
                "The destination already exists and overwrite is disabled.",
                fullPath);
        }

        cancellationToken.ThrowIfCancellationRequested();
        progress?.Report(new DelimitedWriterProgress(DelimitedWriterStage.Writing));

        var delimiter = format == DelimitedFileFormat.Csv ? ',' : '\t';
        var transactionPath = CreateTemporaryPath(destinationDirectory, Path.GetFileName(fullPath));
        var committed = false;
        var committing = false;
        long sourceRowsProcessed = 0;
        long recordsWritten = 0;
        long cellsWritten = 0;
        long nextRowIndex = 0;
        int? currentRowIndex = null;
        int? currentColumnIndex = null;

        try
        {
            await using (var output = new FileStream(
                transactionPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                bufferSize: 64 * 1024,
                FileOptions.Asynchronous | FileOptions.SequentialScan))
            {
                await using (var writer = new StreamWriter(
                    output,
                    new UTF8Encoding(
                        encoderShouldEmitUTF8Identifier: options.EmitUtf8ByteOrderMark,
                        throwOnInvalidBytes: true),
                    bufferSize: 64 * 1024,
                    leaveOpen: true))
                {
                    await foreach (var row in rows
                        .WithCancellation(cancellationToken)
                        .ConfigureAwait(false))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        currentRowIndex = row.RowIndex;
                        currentColumnIndex = null;
                        if (row.RowIndex < nextRowIndex)
                        {
                            throw new DelimitedWriterException(
                                DelimitedWriterError.InvalidRowOrder,
                                "Rows must be unique and ordered by RowIndex.",
                                fullPath,
                                row.RowIndex);
                        }

                        while (nextRowIndex < row.RowIndex)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            await writer.WriteAsync(
                                RecordSeparator.AsMemory(),
                                cancellationToken).ConfigureAwait(false);
                            nextRowIndex++;
                            recordsWritten++;
                        }

                        var writtenCells = await WriteRowAsync(
                            writer,
                            row,
                            delimiter,
                            fullPath,
                            columnIndex => currentColumnIndex = columnIndex,
                            cancellationToken).ConfigureAwait(false);
                        cellsWritten = checked(cellsWritten + writtenCells);
                        sourceRowsProcessed++;
                        recordsWritten++;
                        nextRowIndex = (long)row.RowIndex + 1;

                        if (sourceRowsProcessed % options.ProgressIntervalRows == 0)
                        {
                            progress?.Report(new DelimitedWriterProgress(
                                DelimitedWriterStage.Writing,
                                sourceRowsProcessed,
                                recordsWritten,
                                cellsWritten));
                        }
                    }

                    cancellationToken.ThrowIfCancellationRequested();
                    await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
                }

                cancellationToken.ThrowIfCancellationRequested();
                output.Flush(flushToDisk: true);
            }

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new DelimitedWriterProgress(
                DelimitedWriterStage.Committing,
                sourceRowsProcessed,
                recordsWritten,
                cellsWritten));
            committing = true;
            Commit(transactionPath, fullPath, options.Overwrite);
            committed = true;

            var contentLength = new FileInfo(fullPath).Length;
            progress?.Report(new DelimitedWriterProgress(
                DelimitedWriterStage.Completed,
                sourceRowsProcessed,
                recordsWritten,
                cellsWritten));
            return new DelimitedWriteResult(
                fullPath,
                format,
                sourceRowsProcessed,
                recordsWritten,
                cellsWritten,
                contentLength,
                options.EmitUtf8ByteOrderMark);
        }
        catch (DelimitedWriterException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (EncoderFallbackException exception)
        {
            throw new DelimitedWriterException(
                DelimitedWriterError.InvalidCellValue,
                "A cell contains text that cannot be encoded as valid UTF-8.",
                fullPath,
                currentRowIndex,
                currentColumnIndex,
                exception);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            throw new DelimitedWriterException(
                committing ? DelimitedWriterError.CommitFailed : DelimitedWriterError.IoFailure,
                committing
                    ? "The transactional output could not be committed atomically."
                    : "The delimited output could not be written transactionally.",
                fullPath,
                currentRowIndex,
                currentColumnIndex,
                exception);
        }
        finally
        {
            if (!committed)
            {
                TryDelete(transactionPath);
            }
        }
    }

    private static async ValueTask<long> WriteRowAsync(
        StreamWriter writer,
        RowRecord row,
        char delimiter,
        string destinationPath,
        Action<int> setCurrentColumn,
        CancellationToken cancellationToken)
    {
        var nextColumnIndex = 0;
        var previousColumnIndex = -1;
        for (var index = 0; index < row.Cells.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = row.Cells.Span[index];
            var columnIndex = cell.Address.ColumnIndex;
            setCurrentColumn(columnIndex);
            if (cell.Address.RowIndex != row.RowIndex)
            {
                throw new DelimitedWriterException(
                    DelimitedWriterError.InvalidCellAddress,
                    "A cell address does not belong to its row.",
                    destinationPath,
                    row.RowIndex,
                    columnIndex);
            }

            if (columnIndex <= previousColumnIndex)
            {
                throw new DelimitedWriterException(
                    DelimitedWriterError.InvalidCellOrder,
                    "Cells must be unique and ordered by column index.",
                    destinationPath,
                    row.RowIndex,
                    columnIndex);
            }

            while (nextColumnIndex < columnIndex)
            {
                if (nextColumnIndex > 0)
                {
                    await writer.WriteAsync(delimiter).ConfigureAwait(false);
                }

                nextColumnIndex++;
            }

            if (nextColumnIndex > 0)
            {
                await writer.WriteAsync(delimiter).ConfigureAwait(false);
            }

            await WriteFieldAsync(
                writer,
                GetRawText(cell.Value),
                delimiter,
                cancellationToken).ConfigureAwait(false);
            nextColumnIndex++;
            previousColumnIndex = columnIndex;
        }

        await writer.WriteAsync(
            RecordSeparator.AsMemory(),
            cancellationToken).ConfigureAwait(false);
        return row.Cells.Length;
    }

    private static string GetRawText(CellValue value)
    {
        if (value.IsFormula)
        {
            return "=" + (value.Formula ?? string.Empty);
        }

        var scalar = value.Scalar.GetValueOrDefault();
        return scalar.Kind switch
        {
            CellKind.Blank => string.Empty,
            CellKind.Text or CellKind.Error => scalar.TextValue ?? string.Empty,
            CellKind.Number => scalar.NumberValue.GetValueOrDefault().ToString("R", CultureInfo.InvariantCulture),
            CellKind.Boolean => scalar.BooleanValue.GetValueOrDefault() ? "TRUE" : "FALSE",
            CellKind.DateTime => scalar.DateTimeValue.GetValueOrDefault().ToString("O", CultureInfo.InvariantCulture),
            _ => string.Empty,
        };
    }

    private static async ValueTask WriteFieldAsync(
        StreamWriter writer,
        string value,
        char delimiter,
        CancellationToken cancellationToken)
    {
        var requiresQuotes = value.IndexOfAny([delimiter, '"', '\r', '\n']) >= 0;
        if (!requiresQuotes)
        {
            await writer.WriteAsync(value.AsMemory(), cancellationToken).ConfigureAwait(false);
            return;
        }

        await writer.WriteAsync('"').ConfigureAwait(false);
        var segmentStart = 0;
        for (var index = 0; index < value.Length; index++)
        {
            if (value[index] != '"')
            {
                continue;
            }

            if (index > segmentStart)
            {
                await writer.WriteAsync(
                    value.AsMemory(segmentStart, index - segmentStart),
                    cancellationToken).ConfigureAwait(false);
            }

            await writer.WriteAsync("\"\"".AsMemory(), cancellationToken).ConfigureAwait(false);
            segmentStart = index + 1;
        }

        if (segmentStart < value.Length)
        {
            await writer.WriteAsync(value.AsMemory(segmentStart), cancellationToken).ConfigureAwait(false);
        }

        await writer.WriteAsync('"').ConfigureAwait(false);
    }

    private static DelimitedFileFormat ResolveFormat(
        string destinationPath,
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

        if (DelimitedFormatDetector.TryDetect(destinationPath, out var detected))
        {
            return detected;
        }

        throw new DelimitedWriterException(
            DelimitedWriterError.UnsupportedFileExtension,
            "Automatic format detection requires a .csv or .tsv extension.",
            destinationPath);
    }

    private static void Commit(string transactionPath, string destinationPath, bool overwrite)
    {
        if (File.Exists(destinationPath))
        {
            if (!overwrite)
            {
                throw new DelimitedWriterException(
                    DelimitedWriterError.DestinationExists,
                    "The destination already exists and overwrite is disabled.",
                    destinationPath);
            }

            try
            {
                File.Replace(transactionPath, destinationPath, destinationBackupFileName: null);
            }
            catch (PlatformNotSupportedException)
            {
                File.Move(transactionPath, destinationPath, overwrite: true);
            }
        }
        else
        {
            File.Move(transactionPath, destinationPath);
        }
    }

    private static string CreateTemporaryPath(string directory, string destinationFileName)
    {
        for (var attempt = 0; attempt < 10; attempt++)
        {
            var candidate = Path.Combine(
                directory,
                $".{destinationFileName}.{Guid.NewGuid():N}.tmp");
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new IOException("A unique transactional output path could not be allocated.");
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
            // Preserve the write failure. A later cleanup can remove an abandoned transaction file.
        }
    }

    private static void ValidateOptions(DelimitedWriterOptions options)
    {
        if (options.Format is not (
            DelimitedFileFormat.Auto or
            DelimitedFileFormat.Csv or
            DelimitedFileFormat.Tsv))
        {
            throw new ArgumentOutOfRangeException(nameof(options.Format));
        }

        if (options.ProgressIntervalRows <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(options.ProgressIntervalRows),
                "The progress interval must be positive.");
        }
    }

    private static async IAsyncEnumerable<RowRecord> AsAsyncEnumerable(
        IEnumerable<RowRecord> rows,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        await Task.CompletedTask.ConfigureAwait(false);
        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return row;
        }
    }
}

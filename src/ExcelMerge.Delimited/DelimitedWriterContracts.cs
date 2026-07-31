using ExcelMerge.Domain;

namespace ExcelMerge.Delimited;

public interface IDelimitedContentWriter
{
    ValueTask<DelimitedWriteResult> WriteAsync(
        string destinationPath,
        IAsyncEnumerable<RowRecord> rows,
        DelimitedWriterOptions? options = null,
        IProgress<DelimitedWriterProgress>? progress = null,
        CancellationToken cancellationToken = default);

    ValueTask<DelimitedWriteResult> WriteAsync(
        string destinationPath,
        IEnumerable<RowRecord> rows,
        DelimitedWriterOptions? options = null,
        IProgress<DelimitedWriterProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public sealed record DelimitedWriterOptions
{
    /// <summary>
    /// Uses the .csv or .tsv extension when set to Auto. Unknown extensions require an explicit format.
    /// </summary>
    public DelimitedFileFormat Format { get; init; } = DelimitedFileFormat.Auto;

    public bool Overwrite { get; init; } = true;

    public bool EmitUtf8ByteOrderMark { get; init; }

    public int ProgressIntervalRows { get; init; } = 1024;
}

public enum DelimitedWriterStage
{
    Writing,
    Committing,
    Completed,
}

public readonly record struct DelimitedWriterProgress(
    DelimitedWriterStage Stage,
    long SourceRowsProcessed = 0,
    long RecordsWritten = 0,
    long CellsWritten = 0);

public sealed record DelimitedWriteResult(
    string DestinationPath,
    DelimitedFileFormat Format,
    long SourceRowsProcessed,
    long RecordsWritten,
    long CellsWritten,
    long ContentLength,
    bool HasByteOrderMark);

public enum DelimitedWriterError
{
    UnsupportedFileExtension,
    DestinationExists,
    InvalidRowOrder,
    InvalidCellOrder,
    InvalidCellAddress,
    InvalidCellValue,
    CommitFailed,
    IoFailure,
}

public sealed class DelimitedWriterException : IOException
{
    public DelimitedWriterException(
        DelimitedWriterError error,
        string message,
        string? destinationPath = null,
        int? rowIndex = null,
        int? columnIndex = null,
        Exception? innerException = null)
        : base(message, innerException)
    {
        Error = error;
        DestinationPath = destinationPath;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
    }

    public DelimitedWriterError Error { get; }

    public string? DestinationPath { get; }

    public int? RowIndex { get; }

    public int? ColumnIndex { get; }
}

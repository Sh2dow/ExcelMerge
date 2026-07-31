using ExcelMerge.Domain;
using ExcelMerge.Storage;

namespace ExcelMerge.Delimited;

public enum DelimitedFileFormat
{
    Auto,
    Csv,
    Tsv,
}

public enum DelimitedTextEncoding
{
    Utf8,
    Utf16LittleEndian,
    Utf16BigEndian,
}

/// <summary>Extension-based detection for unambiguous CSV and TSV file names.</summary>
public static class DelimitedFormatDetector
{
    public static bool TryDetect(string filePath, out DelimitedFileFormat format)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

        var extension = Path.GetExtension(filePath);
        if (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
        {
            format = DelimitedFileFormat.Csv;
            return true;
        }

        if (extension.Equals(".tsv", StringComparison.OrdinalIgnoreCase))
        {
            format = DelimitedFileFormat.Tsv;
            return true;
        }

        format = DelimitedFileFormat.Auto;
        return false;
    }
}

public interface IDelimitedWorkbookReader
{
    ValueTask<WorkbookMetadata> ReadMetadataAsync(
        string filePath,
        DelimitedReaderOptions? options = null,
        CancellationToken cancellationToken = default);

    ValueTask<DelimitedReaderResult> IndexAsync(
        string filePath,
        Workspace workspace,
        DelimitedReaderOptions? options = null,
        IProgress<DelimitedReaderProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

public sealed record DelimitedReaderOptions
{
    /// <summary>
    /// Uses the .csv or .tsv extension when set to Auto. Unknown extensions require an explicit format.
    /// </summary>
    public DelimitedFileFormat Format { get; init; } = DelimitedFileFormat.Auto;

    /// <summary>Overrides the single worksheet name. The file name without extension is used by default.</summary>
    public string? WorksheetName { get; init; }

    /// <summary>Number of indexed rows between progress reports.</summary>
    public int ProgressIntervalRows { get; init; } = 1024;

    /// <summary>Maximum characters in one decoded field. Zero disables this adapter limit.</summary>
    public int MaximumFieldCharacters { get; init; }

    /// <summary>Maximum fields in one record. Zero disables this adapter limit.</summary>
    public int MaximumColumns { get; init; }

    /// <summary>Maximum logical records. Zero permits every row index representable by the domain.</summary>
    public int MaximumRows { get; init; }
}

public enum DelimitedReaderStage
{
    Detecting,
    Reading,
    Completed,
}

public readonly record struct DelimitedReaderProgress(
    DelimitedReaderStage Stage,
    long RowsRead = 0,
    long CellsRead = 0,
    long ContentLength = 0);

public enum DelimitedReaderError
{
    UnsupportedFileExtension,
    UnsupportedEncoding,
    InvalidEncoding,
    UnexpectedQuote,
    UnexpectedCharacterAfterQuote,
    UnterminatedQuotedField,
    FieldLimitExceeded,
    ColumnLimitExceeded,
    RowLimitExceeded,
    SourceChanged,
}

/// <summary>A deterministic format, encoding, syntax, limit, or source-stability failure.</summary>
public sealed class DelimitedReaderException : IOException
{
    public DelimitedReaderException(
        DelimitedReaderError error,
        string message,
        string? sourcePath = null,
        long? recordNumber = null,
        int? fieldNumber = null,
        long? characterOffset = null,
        Exception? innerException = null)
        : base(message, innerException)
    {
        Error = error;
        SourcePath = sourcePath;
        RecordNumber = recordNumber;
        FieldNumber = fieldNumber;
        CharacterOffset = characterOffset;
    }

    public DelimitedReaderError Error { get; }

    public string? SourcePath { get; }

    /// <summary>One-based logical record number, when the failure is record-specific.</summary>
    public long? RecordNumber { get; }

    /// <summary>One-based field number, when the failure is field-specific.</summary>
    public int? FieldNumber { get; }

    /// <summary>One-based decoded character offset, when known.</summary>
    public long? CharacterOffset { get; }
}

public sealed class DelimitedReaderResult
{
    internal DelimitedReaderResult(
        WorkbookMetadata metadata,
        DelimitedReaderWorksheet worksheet,
        DelimitedFileFormat format,
        DelimitedTextEncoding textEncoding,
        bool hasByteOrderMark)
    {
        Metadata = metadata;
        Worksheet = worksheet;
        Worksheets = Array.AsReadOnly([worksheet]);
        Format = format;
        TextEncoding = textEncoding;
        HasByteOrderMark = hasByteOrderMark;
    }

    public WorkbookMetadata Metadata { get; }

    public DelimitedReaderWorksheet Worksheet { get; }

    public IReadOnlyList<DelimitedReaderWorksheet> Worksheets { get; }

    public DelimitedFileFormat Format { get; }

    public DelimitedTextEncoding TextEncoding { get; }

    public bool HasByteOrderMark { get; }
}

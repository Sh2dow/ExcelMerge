using System.Collections.ObjectModel;
using System.Globalization;
using ExcelMerge.Domain;
using ExcelMerge.Storage;

namespace ExcelMerge.OpenXml;

/// <summary>Reads workbook metadata and streams selected worksheets into disk-backed row stores.</summary>
public interface IOpenXmlWorkbookReader
{
    ValueTask<WorkbookMetadata> ReadMetadataAsync(
        string filePath,
        OpenXmlReaderOptions? options = null,
        CancellationToken cancellationToken = default);

    ValueTask<OpenXmlReaderResult> IndexAsync(
        string filePath,
        Workspace workspace,
        OpenXmlReaderOptions? options = null,
        IProgress<OpenXmlReaderProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

/// <summary>Controls worksheet selection, display conversion, progress cadence, and package limits.</summary>
public sealed record OpenXmlReaderOptions
{
    /// <summary>
    /// Relationship identifiers to index. When both selection collections are null, every worksheet is
    /// indexed. When either collection is supplied, a worksheet matching an identifier or name is indexed.
    /// </summary>
    public IReadOnlyCollection<string>? SheetIds { get; init; }

    /// <summary>Case-insensitive worksheet names to index.</summary>
    public IReadOnlyCollection<string>? SheetNames { get; init; }

    /// <summary>Whether to populate <see cref="CellValue.DisplayText"/>.</summary>
    public bool IncludeDisplayText { get; init; } = true;

    /// <summary>Culture used for practical number and date display formatting.</summary>
    public CultureInfo DisplayCulture { get; init; } = CultureInfo.InvariantCulture;

    /// <summary>Number of rows between indexing progress reports. Must be positive.</summary>
    public int ProgressIntervalRows { get; init; } = 1024;

    /// <summary>
    /// Maximum aggregate uncompressed ZIP entry size. Zero disables this preflight limit.
    /// </summary>
    public long MaximumUncompressedPackageBytes { get; init; }

    /// <summary>
    /// Maximum characters accepted by the Open XML SDK for one XML part. Zero uses the SDK's unlimited
    /// setting, which is appropriate for very large streamed worksheets.
    /// </summary>
    public long MaximumCharactersInPart { get; init; }
}

public enum OpenXmlReaderStage
{
    Preflight,
    DiscoveringWorkbook,
    ReadingSharedStrings,
    IndexingWorksheet,
    Completed,
}

/// <summary>Monotonic operation counters; total worksheet rows are not known without reading the sheet.</summary>
public readonly record struct OpenXmlReaderProgress(
    OpenXmlReaderStage Stage,
    int CompletedWorksheetCount,
    int TotalWorksheetCount,
    string? SheetId = null,
    string? SheetName = null,
    long RowsRead = 0,
    long CellsRead = 0,
    long SharedStringsRead = 0);

public enum OpenXmlReaderError
{
    UnsupportedFileExtension,
    LegacyBinaryWorkbook,
    EncryptedPackage,
    NotZipPackage,
    InvalidPackage,
    UnsupportedWorkbookType,
    MissingWorkbookPart,
    InvalidRelationship,
    UnsupportedSheetType,
    SheetNotFound,
    InvalidSharedStringTable,
    InvalidStyles,
    InvalidWorksheet,
    SourceChanged,
}

/// <summary>A fail-closed format, package, relationship, or worksheet rejection.</summary>
public sealed class OpenXmlReaderException : IOException
{
    public OpenXmlReaderException(
        OpenXmlReaderError error,
        string message,
        string? sourcePath = null,
        string? sheetId = null,
        Exception? innerException = null)
        : base(message, innerException)
    {
        Error = error;
        SourcePath = sourcePath;
        SheetId = sheetId;
    }

    public OpenXmlReaderError Error { get; }

    public string? SourcePath { get; }

    public string? SheetId { get; }
}

/// <summary>The indexed worksheets and their post-scan metadata.</summary>
public sealed class OpenXmlReaderResult
{
    internal OpenXmlReaderResult(
        WorkbookMetadata metadata,
        IReadOnlyList<OpenXmlReaderWorksheet> worksheets)
    {
        Metadata = metadata;
        Worksheets = worksheets;
    }

    public WorkbookMetadata Metadata { get; }

    public IReadOnlyList<OpenXmlReaderWorksheet> Worksheets { get; }
}

/// <summary>
/// A sparse worksheet snapshot backed by a <see cref="ChunkedCellStore"/>. The caller-provided
/// <see cref="Workspace"/> owns the store lifetime.
/// </summary>
public sealed class OpenXmlReaderWorksheet : IWorksheetSnapshot
{
    private readonly int[] _storedRowIndices;

    internal OpenXmlReaderWorksheet(
        SheetMetadata metadata,
        ChunkedCellStore cellStore,
        int[] storedRowIndices)
    {
        Metadata = metadata;
        CellStore = cellStore;
        _storedRowIndices = storedRowIndices;
    }

    public SheetMetadata Metadata { get; }

    public ChunkedCellStore CellStore { get; }

    public int StoredRowCount => _storedRowIndices.Length;

    public ReadOnlyMemory<int> StoredRowIndices => _storedRowIndices;

    public async ValueTask<RowRecord?> GetRowAsync(
        int rowIndex,
        CancellationToken cancellationToken = default)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
        cancellationToken.ThrowIfCancellationRequested();

        var ordinal = Array.BinarySearch(_storedRowIndices, rowIndex);
        if (ordinal < 0)
        {
            return null;
        }

        return await CellStore.ReadRowAsync(ordinal, cancellationToken).ConfigureAwait(false);
    }

    public async ValueTask<WorksheetRowBatch> ReadRowsAsync(
        int startRowIndex,
        int rowCount,
        CancellationToken cancellationToken = default)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(startRowIndex);
        ArgumentOutOfRangeException.ThrowIfNegative(rowCount);
        cancellationToken.ThrowIfCancellationRequested();

        var endRowIndex = (long)startRowIndex + rowCount;
        var startOrdinal = LowerBound(startRowIndex);
        var endOrdinal = LowerBound(endRowIndex);
        var rows = new RowRecord[endOrdinal - startOrdinal];

        for (var index = 0; index < rows.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            rows[index] = await CellStore.ReadRowAsync(
                startOrdinal + index,
                cancellationToken).ConfigureAwait(false);
        }

        return new WorksheetRowBatch(startRowIndex, rowCount, rows);
    }

    internal static IReadOnlyList<OpenXmlReaderWorksheet> AsReadOnly(
        OpenXmlReaderWorksheet[] worksheets) =>
        new ReadOnlyCollection<OpenXmlReaderWorksheet>(worksheets);

    private int LowerBound(long rowIndex)
    {
        var low = 0;
        var high = _storedRowIndices.Length;
        while (low < high)
        {
            var middle = low + ((high - low) >> 1);
            if (_storedRowIndices[middle] < rowIndex)
            {
                low = middle + 1;
            }
            else
            {
                high = middle;
            }
        }

        return low;
    }
}

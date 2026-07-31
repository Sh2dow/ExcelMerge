using System.Security.Cryptography;
using ExcelMerge.Delimited;
using ExcelMerge.Domain;
using ExcelMerge.OpenXml;
using ExcelMerge.Storage;

namespace ExcelMerge.Application;

internal sealed record IndexedApplicationWorkbook(
    string FilePath,
    WorkbookMetadata Metadata,
    IReadOnlyList<IWorksheetSnapshot> Worksheets,
    OpenXmlWorkbookSource? OpenXmlSource,
    ApplicationFileFingerprint Fingerprint)
{
    public WorkbookFormat Format => Metadata.Format;
}

internal readonly record struct ApplicationFileFingerprint(
    long ContentLength,
    long LastWriteTimeUtcTicks,
    string Sha256)
{
    public static async ValueTask<ApplicationFileFingerprint> CaptureAsync(
        string path,
        CancellationToken cancellationToken)
    {
        var fullPath = Path.GetFullPath(path);
        var before = new FileInfo(fullPath);
        before.Refresh();
        if (!before.Exists)
        {
            throw new FileNotFoundException("The source file does not exist.", fullPath);
        }

        await using var stream = new FileStream(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read | FileShare.Delete,
            bufferSize: 128 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        var hash = await SHA256.HashDataAsync(stream, cancellationToken).ConfigureAwait(false);
        var after = new FileInfo(fullPath);
        after.Refresh();
        if (!after.Exists ||
            before.Length != after.Length ||
            before.LastWriteTimeUtc != after.LastWriteTimeUtc)
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.InvalidRequest,
                "A source file changed while its session fingerprint was captured.",
                path: fullPath);
        }

        return new ApplicationFileFingerprint(
            after.Length,
            after.LastWriteTimeUtc.Ticks,
            Convert.ToHexString(hash));
    }

    public async ValueTask VerifyAsync(string path, CancellationToken cancellationToken)
    {
        var actual = await CaptureAsync(path, cancellationToken).ConfigureAwait(false);
        if (actual != this)
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.InvalidRequest,
                "A source file changed after the session was opened.",
                path: Path.GetFullPath(path));
        }
    }
}

internal sealed class ApplicationWorkbookLoader(
    int progressIntervalRows)
{
    public WorkbookFormat DetectFormat(string path, WorkbookSide side)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.InvalidRequest,
                "A source path cannot be empty.",
                side,
                path);
        }

        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.MissingSource,
                "A source file does not exist.",
                side,
                fullPath);
        }

        return Path.GetExtension(fullPath).ToLowerInvariant() switch
        {
            ".xlsx" => WorkbookFormat.Xlsx,
            ".csv" => WorkbookFormat.Csv,
            ".tsv" => WorkbookFormat.Tsv,
            ".xls" => throw new ExcelMergeApplicationException(
                ApplicationError.UnsupportedLegacyWorkbook,
                "Legacy .xls workbooks are not supported. Convert the file to .xlsx first.",
                side,
                fullPath),
            _ => throw new ExcelMergeApplicationException(
                ApplicationError.UnsupportedFormat,
                "Only .xlsx, .csv, and .tsv files are supported.",
                side,
                fullPath),
        };
    }

    public async ValueTask<IndexedApplicationWorkbook> LoadAsync(
        string path,
        WorkbookSide side,
        WorkbookFormat format,
        Workspace workspace,
        IProgress<ApplicationProgress>? progress,
        CancellationToken cancellationToken)
    {
        var fullPath = Path.GetFullPath(path);
        progress?.Report(new ApplicationProgress(ApplicationStage.ReadingMetadata, side));
        switch (format)
        {
            case WorkbookFormat.Xlsx:
                {
                    var reader = new OpenXmlWorkbookReader();
                    var options = new OpenXmlReaderOptions
                    {
                        ProgressIntervalRows = progressIntervalRows,
                    };
                    await reader.ReadMetadataAsync(fullPath, options, cancellationToken)
                        .ConfigureAwait(false);
                    var result = await reader.IndexAsync(
                        fullPath,
                        workspace,
                        options,
                        new InlineProgress<OpenXmlReaderProgress>(value => progress?.Report(
                            new ApplicationProgress(
                                ApplicationStage.Indexing,
                                side,
                                value.SheetId,
                                value.SheetName,
                                value.CompletedWorksheetCount,
                                value.TotalWorksheetCount,
                                value.RowsRead,
                                value.CellsRead))),
                        cancellationToken).ConfigureAwait(false);
                    var source = await OpenXmlWorkbookSource.CaptureAsync(fullPath, cancellationToken)
                        .ConfigureAwait(false);
                    return new IndexedApplicationWorkbook(
                        fullPath,
                        result.Metadata,
                        result.Worksheets.Cast<IWorksheetSnapshot>().ToArray(),
                        source,
                        new ApplicationFileFingerprint(
                            source.Fingerprint.ContentLength,
                            source.Fingerprint.LastWriteTimeUtcTicks,
                            source.Fingerprint.Sha256));
                }
            case WorkbookFormat.Csv:
            case WorkbookFormat.Tsv:
                {
                    var reader = new DelimitedWorkbookReader();
                    var options = new DelimitedReaderOptions
                    {
                        Format = format == WorkbookFormat.Csv
                            ? DelimitedFileFormat.Csv
                            : DelimitedFileFormat.Tsv,
                        WorksheetName = "Data",
                        ProgressIntervalRows = progressIntervalRows,
                    };
                    await reader.ReadMetadataAsync(fullPath, options, cancellationToken)
                        .ConfigureAwait(false);
                    var result = await reader.IndexAsync(
                        fullPath,
                        workspace,
                        options,
                        new InlineProgress<DelimitedReaderProgress>(value => progress?.Report(
                            new ApplicationProgress(
                                ApplicationStage.Indexing,
                                side,
                                resultSheetId(format),
                                Rows: value.RowsRead,
                                Cells: value.CellsRead))),
                        cancellationToken).ConfigureAwait(false);
                    var fingerprint = await ApplicationFileFingerprint.CaptureAsync(
                        fullPath,
                        cancellationToken).ConfigureAwait(false);
                    return new IndexedApplicationWorkbook(
                        fullPath,
                        result.Metadata,
                        result.Worksheets.Cast<IWorksheetSnapshot>().ToArray(),
                        OpenXmlSource: null,
                        fingerprint);
                }
            default:
                throw new ExcelMergeApplicationException(
                    ApplicationError.UnsupportedFormat,
                    $"Workbook format '{format}' is not supported.",
                    side,
                    fullPath);
        }

        static string resultSheetId(WorkbookFormat sourceFormat) =>
            sourceFormat == WorkbookFormat.Csv ? "csv:0" : "tsv:0";
    }
}

internal sealed class InlineProgress<T>(Action<T> callback) : IProgress<T>
{
    public void Report(T value) => callback(value);
}

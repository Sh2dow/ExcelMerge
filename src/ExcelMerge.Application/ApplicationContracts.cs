using ExcelMerge.Domain;
using ExcelMerge.Engine;
using ExcelMerge.Storage;

namespace ExcelMerge.Application;

public enum WorkbookSide
{
    Base,
    Local,
    Remote,
}

public enum ApplicationStage
{
    Validating,
    ReadingMetadata,
    Indexing,
    Comparing,
    Merging,
    Saving,
    Completed,
}

public readonly record struct ApplicationProgress(
    ApplicationStage Stage,
    WorkbookSide? Side = null,
    string? SheetId = null,
    string? SheetName = null,
    long Completed = 0,
    long? Total = null,
    long Rows = 0,
    long Cells = 0);

public sealed record ExcelMergeApplicationOptions
{
    public WorkspaceOptions Workspace { get; init; } = new();

    public WorksheetComparisonOptions Comparison { get; init; } = new();

    public int ProgressIntervalRows { get; init; } = 1024;

    public TimeSpan StaleWorkspaceMinimumAge { get; init; } = WorkspaceMaintenance.DefaultMinimumAge;
}

public sealed record WorksheetSelection(string? Id = null, string? Name = null)
{
    internal void Validate(string parameterName)
    {
        if (string.IsNullOrWhiteSpace(Id) == string.IsNullOrWhiteSpace(Name))
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.InvalidSheetSelection,
                "A worksheet selection must specify exactly one identifier or name.",
                parameterName: parameterName);
        }
    }
}

public sealed record CompareRequest(
    string LocalPath,
    string RemotePath,
    WorksheetSelection? LocalSheet = null,
    WorksheetSelection? RemoteSheet = null,
    WorksheetComparisonOptions? Options = null);

public sealed record MergeRequest(
    string BasePath,
    string LocalPath,
    string RemotePath,
    WorksheetComparisonOptions? Options = null);

public sealed record SaveRequest(
    string DestinationPath,
    bool Overwrite = true,
    bool ValidatePackage = true,
    long MinimumFreeSpaceReserveBytes = 16L * 1024 * 1024);

public sealed record ApplicationSaveResult(
    string DestinationPath,
    WorkbookFormat Format,
    long ContentLength);

public sealed record CompareSheetResult(
    string Id,
    IWorksheetSnapshot? LocalWorksheet,
    IWorksheetSnapshot? RemoteWorksheet,
    TwoWayDiffResult Difference);

public sealed record MergeSheetResult(
    string Id,
    IWorksheetSnapshot? BaseWorksheet,
    IWorksheetSnapshot? LocalWorksheet,
    IWorksheetSnapshot? RemoteWorksheet,
    ThreeWayMergeResult Merge);

public enum ApplicationError
{
    InvalidRequest,
    MissingSource,
    UnsupportedFormat,
    UnsupportedLegacyWorkbook,
    MixedFormats,
    InvalidSheetSelection,
    InvalidResolution,
    UnresolvedConflicts,
    UnsupportedOutput,
    SessionDisposed,
}

public sealed class ExcelMergeApplicationException : InvalidOperationException
{
    public ExcelMergeApplicationException(
        ApplicationError error,
        string message,
        WorkbookSide? side = null,
        string? path = null,
        string? parameterName = null,
        Exception? innerException = null)
        : base(message, innerException)
    {
        Error = error;
        Side = side;
        Path = path;
        ParameterName = parameterName;
    }

    public ApplicationError Error { get; }

    public WorkbookSide? Side { get; }

    public string? Path { get; }

    public string? ParameterName { get; }
}

using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

public enum MergeDecisionScope
{
    Cell,
    Row,
    RowMetadata,
    Sheet,
    SheetMetadata,
}

public enum AutomaticMergeKind
{
    UseLocal,
    UseRemote,
    UseEither,
    Delete,
}

/// <summary>A non-conflicting result choice. Source choices include value, style, and structure.</summary>
public readonly record struct AutomaticMergeDecision(
    MergeDecisionScope Scope,
    AutomaticMergeKind Kind,
    ConflictLocation Location,
    CellValue? ResultValue = null);

public enum MergeViewRowState
{
    Unchanged,
    Automatic,
    Conflict,
}

/// <summary>
/// A three-way row mapping. ConflictStartIndex and ConflictCount address the result's conflict array.
/// </summary>
public readonly record struct MergeViewRow(
    int? BaseRowIndex,
    int? LocalRowIndex,
    int? RemoteRowIndex,
    ChangeKind LocalKind,
    ChangeKind RemoteKind,
    MergeViewRowState State,
    int ConflictStartIndex,
    int ConflictCount);

/// <summary>BASE-relative changes and conflict data for one selected sheet.</summary>
public sealed record ThreeWayMergeResult(
    SheetChange LocalChange,
    SheetChange RemoteChange,
    ReadOnlyMemory<ConflictRecord> Conflicts,
    ReadOnlyMemory<CellResolution> CellResolutions,
    ReadOnlyMemory<RowResolution> RowResolutions,
    ReadOnlyMemory<AutomaticMergeDecision> AutomaticDecisions,
    ReadOnlyMemory<MergeViewRow> ViewRows,
    long NextConflictId)
{
    public bool HasConflicts => !Conflicts.IsEmpty;
}

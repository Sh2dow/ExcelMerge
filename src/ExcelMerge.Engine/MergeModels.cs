using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

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
    string SheetGroupId,
    SheetChange LocalChange,
    SheetChange RemoteChange,
    ReadOnlyMemory<ConflictRecord> Conflicts,
    ReadOnlyMemory<CellResolution> CellResolutions,
    ReadOnlyMemory<RowResolution> RowResolutions,
    ReadOnlyMemory<AutomaticMergeDecision> AutomaticDecisions,
    ReadOnlyMemory<MergeRowMapping> RowMappings,
    ReadOnlyMemory<MergeViewRow> ViewRows,
    long NextConflictId)
{
    public bool HasConflicts => !Conflicts.IsEmpty;
}

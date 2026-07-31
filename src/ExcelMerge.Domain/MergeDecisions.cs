namespace ExcelMerge.Domain;

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

/// <summary>A non-conflicting result choice produced by the merge engine.</summary>
public readonly record struct AutomaticMergeDecision(
    MergeDecisionScope Scope,
    AutomaticMergeKind Kind,
    ConflictLocation Location,
    CellValue? ResultValue = null);

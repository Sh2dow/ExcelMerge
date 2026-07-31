namespace ExcelMerge.Domain;

public enum ChangeKind
{
    Unchanged,
    Added,
    Removed,
    Modified,
}

/// <summary>
/// A side-specific cell delta against BASE. Added changes have no base address/value; removed
/// changes have no source address/value.
/// </summary>
public readonly record struct CellChange(
    SourceSide Source,
    ChangeKind Kind,
    CellAddress? BaseAddress,
    CellAddress? SourceAddress,
    CellValue? BaseValue,
    CellValue? SourceValue);

/// <summary>A side-specific aligned row delta against BASE.</summary>
public readonly record struct RowChange(
    SourceSide Source,
    ChangeKind Kind,
    int? BaseRowIndex,
    int? SourceRowIndex,
    ReadOnlyMemory<CellChange> CellChanges);

/// <summary>A side-specific worksheet delta against BASE.</summary>
public readonly record struct SheetChange(
    SourceSide Source,
    ChangeKind Kind,
    SheetMetadata? BaseSheet,
    SheetMetadata? SourceSheet,
    ReadOnlyMemory<RowChange> RowChanges);

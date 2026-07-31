namespace ExcelMerge.Domain;

public enum ConflictKind
{
    CellValue,
    CellDeleteEdit,
    RowAddAdd,
    RowDeleteEdit,
    RowMetadata,
    IncompatibleRowInsertion,
    SheetAddAdd,
    SheetDeleteEdit,
    AmbiguousSheetRename,
}

/// <summary>
/// Locates aligned BASE, LOCAL, and REMOTE rows. <see cref="ColumnIndex"/> is null for row- and
/// sheet-level conflicts.
/// </summary>
public readonly record struct ConflictLocation(
    string SheetId,
    int? BaseRowIndex = null,
    int? LocalRowIndex = null,
    int? RemoteRowIndex = null,
    int? ColumnIndex = null);

/// <summary>Values are populated for cell conflicts and omitted for structural conflicts.</summary>
public readonly record struct ConflictRecord(
    long Id,
    ConflictKind Kind,
    ConflictLocation Location,
    CellValue? BaseValue = null,
    CellValue? LocalValue = null,
    CellValue? RemoteValue = null);

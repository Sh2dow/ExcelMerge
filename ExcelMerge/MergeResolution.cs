#nullable enable

namespace ExcelMerge;

public enum MergeResolution
{
    Unresolved,
    Local,
    Remote,
    Both,
    Custom,
}

public readonly record struct MergeRowKey(string SheetName, int RowIndex);

public readonly record struct MergeCellKey(string SheetName, int RowIndex, int ColumnIndex);

public readonly record struct MergeCellResolution(MergeResolution Resolution, string? CustomValue = null);

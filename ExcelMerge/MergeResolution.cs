namespace ExcelMerge;

public enum MergeResolution
{
    Unresolved,
    Local,
    Remote,
    Both,
}

public readonly record struct MergeRowKey(string SheetName, int RowIndex);

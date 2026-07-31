namespace ExcelMerge.Domain;

/// <summary>Identifies an input workbook in a three-way merge.</summary>
public enum WorkbookSide
{
    Base,
    Local,
    Remote,
}

/// <summary>Identifies the side that produced a change relative to BASE.</summary>
public enum SourceSide
{
    Local,
    Remote,
}

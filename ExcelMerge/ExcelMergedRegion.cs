namespace ExcelMerge;

public readonly record struct ExcelMergedRegion(
    int FirstRow,
    int LastRow,
    int FirstColumn,
    int LastColumn)
{
    public bool Overlaps(ExcelMergedRegion other)
    {
        return FirstRow <= other.LastRow
            && LastRow >= other.FirstRow
            && FirstColumn <= other.LastColumn
            && LastColumn >= other.FirstColumn;
    }
}

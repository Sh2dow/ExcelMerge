namespace ExcelMerge.OpenXml;

internal sealed class OpenXmlRowIndexTransform
{
    private readonly int[] _insertionBoundaries;
    private readonly int[] _insertionPrefixCounts;
    private readonly int[] _deletedRows;

    public OpenXmlRowIndexTransform(CompiledRowTransform transform)
    {
        ArgumentNullException.ThrowIfNull(transform);
        var insertions = transform.Insertions.OrderBy(static pair => pair.Key).ToArray();
        _insertionBoundaries = insertions.Select(static pair => pair.Key).ToArray();
        _insertionPrefixCounts = new int[insertions.Length];
        var count = 0;
        for (var index = 0; index < insertions.Length; index++)
        {
            count = checked(count + insertions[index].Value.Count);
            _insertionPrefixCounts[index] = count;
        }

        _deletedRows = transform.DeletedLocalRows.Order().ToArray();
    }

    public int MapLocalRow(int rowIndex) => checked(
        rowIndex + CountInsertionsAtOrBefore(rowIndex) - CountDeletedBefore(rowIndex));

    public int? MapReferenceRow(int rowIndex) =>
        Array.BinarySearch(_deletedRows, rowIndex) >= 0
            ? null
            : MapLocalRow(rowIndex);

    public int MapInsertion(int boundary, int ordinal) => checked(
        boundary + CountInsertionsBefore(boundary) + ordinal - CountDeletedBefore(boundary));

    private int CountInsertionsBefore(int rowIndex) =>
        PrefixCount(LowerBound(_insertionBoundaries, rowIndex) - 1);

    private int CountInsertionsAtOrBefore(int rowIndex) =>
        PrefixCount(UpperBound(_insertionBoundaries, rowIndex) - 1);

    private int CountDeletedBefore(int rowIndex) => LowerBound(_deletedRows, rowIndex);

    private int PrefixCount(int index) => index < 0 ? 0 : _insertionPrefixCounts[index];

    private static int LowerBound(int[] values, int target)
    {
        var low = 0;
        var high = values.Length;
        while (low < high)
        {
            var middle = low + ((high - low) >> 1);
            if (values[middle] < target)
                low = middle + 1;
            else
                high = middle;
        }

        return low;
    }

    private static int UpperBound(int[] values, int target)
    {
        var low = 0;
        var high = values.Length;
        while (low < high)
        {
            var middle = low + ((high - low) >> 1);
            if (values[middle] <= target)
                low = middle + 1;
            else
                high = middle;
        }

        return low;
    }
}

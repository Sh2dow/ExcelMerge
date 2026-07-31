using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

/// <summary>A deterministic dual hash and semantic cell count for a row.</summary>
public readonly record struct RowSignature(
    ulong PrimaryHash,
    ulong SecondaryHash,
    int SignificantCellCount);

public readonly record struct RowKeySignature(
    RowSignature Signature,
    bool HasNonBlankValue);

/// <summary>Computes stable row hashes and provides the exact verification paired with them.</summary>
public sealed class RowSignatureCalculator
{
    private readonly ComparisonContext _context;

    public RowSignatureCalculator(WorksheetComparisonOptions? options = null)
    {
        _context = new ComparisonContext(new NormalizedComparisonOptions(
            options ?? new WorksheetComparisonOptions()));
    }

    public RowSignature Compute(RowRecord row, CancellationToken cancellationToken = default) =>
        _context.ComputeSignature(row, cancellationToken);

    public RowKeySignature ComputeKey(RowRecord row, CancellationToken cancellationToken = default) =>
        _context.ComputeKeySignature(row, cancellationToken);

    public bool VerifyEqual(
        RowRecord left,
        RowRecord right,
        CancellationToken cancellationToken = default) =>
        _context.RowsEqual(left, right, cancellationToken);

    public bool VerifyKeyEqual(
        RowRecord left,
        RowRecord right,
        CancellationToken cancellationToken = default) =>
        _context.KeysEqual(left, right, cancellationToken);
}

/// <summary>Exact semantic row equality using the configured typed-cell and metadata options.</summary>
public sealed class RowValueComparer : IEqualityComparer<RowRecord>
{
    private readonly ComparisonContext _context;

    public RowValueComparer(WorksheetComparisonOptions? options = null)
    {
        _context = new ComparisonContext(new NormalizedComparisonOptions(
            options ?? new WorksheetComparisonOptions()));
    }

    public bool Equals(RowRecord x, RowRecord y) =>
        _context.RowsEqual(x, y, CancellationToken.None);

    public int GetHashCode(RowRecord obj)
    {
        var signature = _context.ComputeSignature(obj, CancellationToken.None);
        return unchecked((int)(signature.PrimaryHash ^ (signature.PrimaryHash >> 32)));
    }
}

internal sealed class ComparisonContext
{
    public ComparisonContext(NormalizedComparisonOptions options)
    {
        Options = options;
    }

    public NormalizedComparisonOptions Options { get; }

    public RowSignature ComputeSignature(RowRecord row, CancellationToken cancellationToken)
    {
        var primary = new StableHashBuilder(StableHashLane.Primary);
        var secondary = new StableHashBuilder(StableHashLane.Secondary);
        primary.AddByte(0x52);
        secondary.AddByte(0x52);

        var count = 0;
        foreach (var cell in row.Cells.Span)
        {
            if ((count % Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            if (IsIgnoredSparseCell(cell))
                continue;

            primary.AddInt32(cell.Address.ColumnIndex);
            secondary.AddInt32(cell.Address.ColumnIndex);
            primary.AddUInt64(Options.CellComparer.GetStableHash(
                cell.Value,
                StableHashLane.Primary,
                cancellationToken));
            secondary.AddUInt64(Options.CellComparer.GetStableHash(
                cell.Value,
                StableHashLane.Secondary,
                cancellationToken));

            if (Options.Source.CompareCellStyles)
            {
                AddNullableInt(ref primary, cell.StyleIndex);
                AddNullableInt(ref secondary, cell.StyleIndex);
            }

            count++;
        }

        primary.AddInt32(count);
        secondary.AddInt32(count);
        if (Options.Source.CompareRowMetadata)
        {
            AddRowMetadata(ref primary, row);
            AddRowMetadata(ref secondary, row);
        }

        return new RowSignature(primary.Value, secondary.Value, count);
    }

    public RowKeySignature ComputeKeySignature(RowRecord row, CancellationToken cancellationToken)
    {
        var primary = new StableHashBuilder(StableHashLane.Primary);
        var secondary = new StableHashBuilder(StableHashLane.Secondary);
        primary.AddByte(0x4B);
        secondary.AddByte(0x4B);

        var hasNonBlank = false;
        foreach (var columnIndex in Options.KeyColumns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = FindCell(row, columnIndex);
            var value = cell?.Value ?? CellValue.Blank;
            hasNonBlank |= value.Kind != CellKind.Blank;

            primary.AddInt32(columnIndex);
            secondary.AddInt32(columnIndex);
            primary.AddUInt64(Options.CellComparer.GetStableHash(
                value,
                StableHashLane.Primary,
                cancellationToken));
            secondary.AddUInt64(Options.CellComparer.GetStableHash(
                value,
                StableHashLane.Secondary,
                cancellationToken));
        }

        var signature = new RowSignature(
            primary.Value,
            secondary.Value,
            Options.KeyColumns.Length);
        return new RowKeySignature(signature, hasNonBlank);
    }

    public bool RowsEqual(RowRecord left, RowRecord right, CancellationToken cancellationToken)
    {
        if (Options.Source.CompareRowMetadata && !RowMetadataEqual(left, right))
            return false;

        var leftCells = left.Cells.Span;
        var rightCells = right.Cells.Span;
        var leftIndex = 0;
        var rightIndex = 0;
        var comparisons = 0;

        while (true)
        {
            leftIndex = NextSignificantCell(leftCells, leftIndex);
            rightIndex = NextSignificantCell(rightCells, rightIndex);
            if (leftIndex == leftCells.Length || rightIndex == rightCells.Length)
                return leftIndex == leftCells.Length && rightIndex == rightCells.Length;

            if ((comparisons++ % Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();

            ref readonly var leftCell = ref leftCells[leftIndex];
            ref readonly var rightCell = ref rightCells[rightIndex];
            if (leftCell.Address.ColumnIndex != rightCell.Address.ColumnIndex ||
                !CellRecordsEqual(leftCell, rightCell, cancellationToken))
            {
                return false;
            }

            leftIndex++;
            rightIndex++;
        }
    }

    public bool KeysEqual(RowRecord left, RowRecord right, CancellationToken cancellationToken)
    {
        foreach (var columnIndex in Options.KeyColumns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var leftValue = FindCell(left, columnIndex)?.Value ?? CellValue.Blank;
            var rightValue = FindCell(right, columnIndex)?.Value ?? CellValue.Blank;
            if (!Options.CellComparer.Equals(leftValue, rightValue, cancellationToken))
                return false;
        }

        return true;
    }

    public bool CellRecordsEqual(
        in CellRecord left,
        in CellRecord right,
        CancellationToken cancellationToken)
    {
        return Options.CellComparer.Equals(left.Value, right.Value, cancellationToken) &&
            (!Options.Source.CompareCellStyles || left.StyleIndex == right.StyleIndex);
    }

    public bool RowMetadataEqual(in RowRecord left, in RowRecord right)
    {
        return NullableDoubleBitsEqual(left.Height, right.Height) &&
            left.IsHidden == right.IsHidden &&
            left.StyleIndex == right.StyleIndex;
    }

    public bool SheetMetadataEqual(SheetMetadata left, SheetMetadata right)
    {
        if (!Options.Source.CompareSheetMetadata)
            return true;

        return string.Equals(left.Name, right.Name, StringComparison.Ordinal) &&
            left.Position == right.Position &&
            left.Visibility == right.Visibility;
    }

    public bool IsIgnoredSparseCell(in CellRecord cell)
    {
        if (!Options.Source.TreatExplicitBlankCellsAsMissing || cell.Value.Kind != CellKind.Blank)
            return false;
        if (Options.Source.CompareCellStyles && cell.StyleIndex.HasValue)
            return false;
        return !Options.Source.CellValues.CompareDisplayText || cell.Value.DisplayText is null;
    }

    public CellRecord? FindCell(in RowRecord row, int columnIndex)
    {
        var cells = row.Cells.Span;
        var low = 0;
        var high = cells.Length;
        while (low < high)
        {
            var middle = low + ((high - low) >> 1);
            var middleColumn = cells[middle].Address.ColumnIndex;
            if (middleColumn < columnIndex)
                low = middle + 1;
            else
                high = middle;
        }

        if (low >= cells.Length || cells[low].Address.ColumnIndex != columnIndex)
            return null;

        var cell = cells[low];
        return IsIgnoredSparseCell(cell) ? null : cell;
    }

    private int NextSignificantCell(ReadOnlySpan<CellRecord> cells, int start)
    {
        while (start < cells.Length && IsIgnoredSparseCell(cells[start]))
            start++;
        return start;
    }

    private static bool NullableDoubleBitsEqual(double? left, double? right)
    {
        if (left.HasValue != right.HasValue)
            return false;
        return !left.HasValue ||
            BitConverter.DoubleToInt64Bits(left.Value) == BitConverter.DoubleToInt64Bits(right!.Value);
    }

    private static void AddRowMetadata(ref StableHashBuilder hash, in RowRecord row)
    {
        hash.AddByte(row.Height.HasValue ? (byte)1 : (byte)0);
        if (row.Height.HasValue)
            hash.AddInt64(BitConverter.DoubleToInt64Bits(row.Height.Value));
        hash.AddByte(row.IsHidden ? (byte)1 : (byte)0);
        AddNullableInt(ref hash, row.StyleIndex);
    }

    private static void AddNullableInt(ref StableHashBuilder hash, int? value)
    {
        hash.AddByte(value.HasValue ? (byte)1 : (byte)0);
        if (value.HasValue)
            hash.AddInt32(value.Value);
    }
}

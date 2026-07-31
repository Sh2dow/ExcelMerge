namespace ExcelMerge.Domain;

/// <summary>A bounded result for a zero-based worksheet row interval.</summary>
public readonly record struct WorksheetRowBatch(
    int StartRowIndex,
    int RowCount,
    ReadOnlyMemory<RowRecord> Rows);

/// <summary>
/// Random and range access to an immutable worksheet. Implementations may be memory-, file-, or
/// memory-map-backed and should avoid materializing the complete worksheet for a range request.
/// </summary>
public interface IWorksheetSnapshot
{
    SheetMetadata Metadata { get; }

    int StoredRowCount { get; }

    ValueTask<RowRecord?> GetRowAsync(
        int rowIndex,
        CancellationToken cancellationToken = default);

    /// <summary>Returns stored rows in the half-open interval [startRowIndex, startRowIndex + rowCount).</summary>
    ValueTask<WorksheetRowBatch> ReadRowsAsync(
        int startRowIndex,
        int rowCount,
        CancellationToken cancellationToken = default);
}

/// <summary>An in-memory worksheet snapshot for bounded sheets, tests, and adapter handoff batches.</summary>
/// <remarks><paramref name="Rows"/> must be ordered by <see cref="RowRecord.RowIndex"/>.</remarks>
public sealed record WorksheetSnapshot(
    SheetMetadata Metadata,
    ReadOnlyMemory<RowRecord> Rows) : IWorksheetSnapshot
{
    public int StoredRowCount => Rows.Length;

    public ValueTask<RowRecord?> GetRowAsync(
        int rowIndex,
        CancellationToken cancellationToken = default)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
        cancellationToken.ThrowIfCancellationRequested();

        var index = LowerBound(rowIndex);
        RowRecord? row = index < Rows.Length && Rows.Span[index].RowIndex == rowIndex
            ? Rows.Span[index]
            : null;
        return new ValueTask<RowRecord?>(row);
    }

    public ValueTask<WorksheetRowBatch> ReadRowsAsync(
        int startRowIndex,
        int rowCount,
        CancellationToken cancellationToken = default)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(startRowIndex);
        ArgumentOutOfRangeException.ThrowIfNegative(rowCount);
        cancellationToken.ThrowIfCancellationRequested();

        var start = LowerBound(startRowIndex);
        var end = LowerBound((long)startRowIndex + rowCount);
        var batch = new WorksheetRowBatch(startRowIndex, rowCount, Rows.Slice(start, end - start));
        return new ValueTask<WorksheetRowBatch>(batch);
    }

    private int LowerBound(long rowIndex)
    {
        var rows = Rows.Span;
        var low = 0;
        var high = rows.Length;
        while (low < high)
        {
            var middle = low + ((high - low) >> 1);
            if (rows[middle].RowIndex < rowIndex)
                low = middle + 1;
            else
                high = middle;
        }

        return low;
    }
}

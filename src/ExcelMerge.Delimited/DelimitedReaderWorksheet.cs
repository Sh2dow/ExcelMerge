using ExcelMerge.Domain;
using ExcelMerge.Storage;

namespace ExcelMerge.Delimited;

/// <summary>A contiguous, workspace-backed worksheet for one delimited text file.</summary>
public sealed class DelimitedReaderWorksheet : IWorksheetSnapshot
{
    internal DelimitedReaderWorksheet(SheetMetadata metadata, ChunkedCellStore cellStore)
    {
        Metadata = metadata;
        CellStore = cellStore;
    }

    public SheetMetadata Metadata { get; }

    internal ChunkedCellStore CellStore { get; }

    public int StoredRowCount => checked((int)CellStore.RowCount);

    public async ValueTask<RowRecord?> GetRowAsync(
        int rowIndex,
        CancellationToken cancellationToken = default)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
        cancellationToken.ThrowIfCancellationRequested();
        if (rowIndex >= CellStore.RowCount)
        {
            return null;
        }

        return await CellStore.ReadRowAsync(rowIndex, cancellationToken).ConfigureAwait(false);
    }

    public async ValueTask<WorksheetRowBatch> ReadRowsAsync(
        int startRowIndex,
        int rowCount,
        CancellationToken cancellationToken = default)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(startRowIndex);
        ArgumentOutOfRangeException.ThrowIfNegative(rowCount);
        cancellationToken.ThrowIfCancellationRequested();
        var available = Math.Max(0, Math.Min(rowCount, StoredRowCount - startRowIndex));
        var rows = new RowRecord[available];
        for (var index = 0; index < rows.Length; index++)
        {
            rows[index] = await CellStore.ReadRowAsync(
                startRowIndex + index,
                cancellationToken).ConfigureAwait(false);
        }

        return new WorksheetRowBatch(startRowIndex, rowCount, rows);
    }
}

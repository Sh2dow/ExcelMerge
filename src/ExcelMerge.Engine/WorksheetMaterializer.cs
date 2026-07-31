using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

internal sealed record MaterializedWorksheet(
    SheetMetadata Metadata,
    RowRecord[] Rows);

internal static class WorksheetMaterializer
{
    public static async ValueTask<MaterializedWorksheet> ReadAsync(
        IWorksheetSnapshot snapshot,
        NormalizedComparisonOptions options,
        CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        cancellationToken.ThrowIfCancellationRequested();

        var storedRowCount = snapshot.StoredRowCount;
        if (storedRowCount < 0)
        {
            throw new InvalidWorksheetSnapshotException(
                $"Sheet '{snapshot.Metadata.Id}' reports a negative stored row count.");
        }

        if (storedRowCount > options.Source.MaxMaterializedRows)
        {
            throw new WorksheetLimitExceededException(
                nameof(options.Source.MaxMaterializedRows),
                options.Source.MaxMaterializedRows,
                storedRowCount);
        }

        ValidateMetadata(snapshot.Metadata);
        if (snapshot.Metadata.NonEmptyCellCount > options.Source.MaxMaterializedCells)
        {
            throw new WorksheetLimitExceededException(
                nameof(options.Source.MaxMaterializedCells),
                options.Source.MaxMaterializedCells,
                snapshot.Metadata.NonEmptyCellCount.Value);
        }

        if (storedRowCount == 0)
            return new MaterializedWorksheet(snapshot.Metadata, Array.Empty<RowRecord>());

        var firstRow = snapshot.Metadata.FirstRowIndex ?? 0;
        var knownLastRow = snapshot.Metadata.LastRowIndex;
        if (knownLastRow.HasValue)
        {
            var knownSpan = (long)knownLastRow.Value - firstRow + 1;
            if (knownSpan > options.Source.MaxScannedRowSpan)
            {
                throw new WorksheetLimitExceededException(
                    nameof(options.Source.MaxScannedRowSpan),
                    options.Source.MaxScannedRowSpan,
                    knownSpan);
            }
        }

        var rows = new List<RowRecord>(storedRowCount);
        long cellCount = 0;
        long scannedSpan = 0;
        var cursor = firstRow;
        var previousRowIndex = -1;

        while (rows.Count < storedRowCount)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var remainingScan = options.Source.MaxScannedRowSpan - scannedSpan;
            if (remainingScan <= 0)
            {
                throw new WorksheetLimitExceededException(
                    nameof(options.Source.MaxScannedRowSpan),
                    options.Source.MaxScannedRowSpan,
                    checked(scannedSpan + 1));
            }

            long availablePhysicalRows = knownLastRow.HasValue
                ? (long)knownLastRow.Value - cursor + 1
                : int.MaxValue - (long)cursor + 1;
            if (availablePhysicalRows <= 0)
                break;

            var requestCount = (int)Math.Min(
                options.Source.ReadBatchRowSpan,
                Math.Min(remainingScan, availablePhysicalRows));
            var batch = await snapshot.ReadRowsAsync(cursor, requestCount, cancellationToken)
                .ConfigureAwait(false);
            ValidateBatchEnvelope(snapshot.Metadata.Id, batch, cursor, requestCount);

            var batchRows = batch.Rows.Span;
            for (var rowOffset = 0; rowOffset < batchRows.Length; rowOffset++)
            {
                if ((rowOffset % options.Source.CancellationCheckInterval) == 0)
                    cancellationToken.ThrowIfCancellationRequested();

                var row = batchRows[rowOffset];
                if (row.RowIndex < cursor || row.RowIndex >= (long)cursor + requestCount)
                {
                    throw new InvalidWorksheetSnapshotException(
                        $"Sheet '{snapshot.Metadata.Id}' returned row {row.RowIndex} outside " +
                        $"the requested interval [{cursor}, {(long)cursor + requestCount}).");
                }

                if (row.RowIndex <= previousRowIndex)
                {
                    throw new InvalidWorksheetSnapshotException(
                        $"Rows in sheet '{snapshot.Metadata.Id}' are not strictly increasing.");
                }

                ValidateRow(snapshot.Metadata.Id, row, options, cancellationToken);
                previousRowIndex = row.RowIndex;
                rows.Add(row);
                if (rows.Count > storedRowCount)
                {
                    throw new InvalidWorksheetSnapshotException(
                        $"Sheet '{snapshot.Metadata.Id}' returned more rows than StoredRowCount.");
                }

                cellCount = checked(cellCount + row.CellCount);
                if (cellCount > options.Source.MaxMaterializedCells)
                {
                    throw new WorksheetLimitExceededException(
                        nameof(options.Source.MaxMaterializedCells),
                        options.Source.MaxMaterializedCells,
                        cellCount);
                }
            }

            scannedSpan += requestCount;
            if (cursor > int.MaxValue - requestCount)
                break;
            cursor += requestCount;
        }

        if (rows.Count != storedRowCount)
        {
            throw new InvalidWorksheetSnapshotException(
                $"Sheet '{snapshot.Metadata.Id}' reports {storedRowCount} stored rows but returned " +
                $"{rows.Count} within its declared or bounded row range.");
        }

        return new MaterializedWorksheet(snapshot.Metadata, rows.ToArray());
    }

    private static void ValidateMetadata(SheetMetadata metadata)
    {
        if (metadata.Id is null || metadata.Name is null)
        {
            throw new InvalidWorksheetSnapshotException(
                "Worksheet metadata must have a non-null identifier and name.");
        }

        if (metadata.Position < 0 || metadata.NonEmptyCellCount < 0)
        {
            throw new InvalidWorksheetSnapshotException(
                $"Sheet '{metadata.Id}' contains a negative position or cell count.");
        }

        if (metadata.FirstRowIndex < 0 || metadata.LastRowIndex < 0 ||
            metadata.FirstColumnIndex < 0 || metadata.LastColumnIndex < 0)
        {
            throw new InvalidWorksheetSnapshotException(
                $"Sheet '{metadata.Id}' contains a negative used-range index.");
        }

        if (metadata.FirstRowIndex.HasValue && metadata.LastRowIndex.HasValue &&
            metadata.FirstRowIndex.Value > metadata.LastRowIndex.Value)
        {
            throw new InvalidWorksheetSnapshotException(
                $"Sheet '{metadata.Id}' has an inverted row range.");
        }

        if (metadata.FirstColumnIndex.HasValue && metadata.LastColumnIndex.HasValue &&
            metadata.FirstColumnIndex.Value > metadata.LastColumnIndex.Value)
        {
            throw new InvalidWorksheetSnapshotException(
                $"Sheet '{metadata.Id}' has an inverted column range.");
        }
    }

    private static void ValidateBatchEnvelope(
        string sheetId,
        WorksheetRowBatch batch,
        int requestedStart,
        int requestedCount)
    {
        if (batch.StartRowIndex != requestedStart || batch.RowCount != requestedCount)
        {
            throw new InvalidWorksheetSnapshotException(
                $"Sheet '{sheetId}' returned batch [{batch.StartRowIndex}, " +
                $"{(long)batch.StartRowIndex + batch.RowCount}) for requested interval " +
                $"[{requestedStart}, {(long)requestedStart + requestedCount}).");
        }

        if (batch.Rows.Length > requestedCount)
        {
            throw new InvalidWorksheetSnapshotException(
                $"Sheet '{sheetId}' returned more stored rows than fit in the requested interval.");
        }
    }

    private static void ValidateRow(
        string sheetId,
        in RowRecord row,
        NormalizedComparisonOptions options,
        CancellationToken cancellationToken)
    {
        var previousColumnIndex = -1;
        var cells = row.Cells.Span;
        for (var cellIndex = 0; cellIndex < cells.Length; cellIndex++)
        {
            if ((cellIndex % options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();

            ref readonly var cell = ref cells[cellIndex];
            if (cell.Address.RowIndex != row.RowIndex)
            {
                throw new InvalidWorksheetSnapshotException(
                    $"Cell {cell.Address} in sheet '{sheetId}' belongs to row {row.RowIndex}.");
            }

            if (cell.Address.ColumnIndex <= previousColumnIndex)
            {
                throw new InvalidWorksheetSnapshotException(
                    $"Cells in row {row.RowIndex} of sheet '{sheetId}' are not strictly ordered.");
            }

            if (cell.StyleIndex < 0)
            {
                throw new InvalidWorksheetSnapshotException(
                    $"Cell {cell.Address} in sheet '{sheetId}' has a negative style index.");
            }

            previousColumnIndex = cell.Address.ColumnIndex;
        }
    }
}

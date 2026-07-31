using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

public sealed class TwoWayDiffEngine
{
    public ValueTask<TwoWayDiffResult> CompareAsync(
        IWorksheetSnapshot? baseSheet,
        IWorksheetSnapshot? sourceSheet,
        WorksheetComparisonOptions? options = null,
        CancellationToken cancellationToken = default) =>
        CompareAsync(
            baseSheet,
            sourceSheet,
            SourceSide.Local,
            options,
            cancellationToken);

    public async ValueTask<TwoWayDiffResult> CompareAsync(
        IWorksheetSnapshot? baseSheet,
        IWorksheetSnapshot? sourceSheet,
        SourceSide sourceSide,
        WorksheetComparisonOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (baseSheet is null && sourceSheet is null)
            throw new ArgumentException("At least one worksheet snapshot is required.");
        if (!Enum.IsDefined(sourceSide))
            throw new ArgumentOutOfRangeException(nameof(sourceSide));

        var normalized = new NormalizedComparisonOptions(options ?? new WorksheetComparisonOptions());
        var context = new ComparisonContext(normalized);
        MaterializedWorksheet? materializedBase = null;
        MaterializedWorksheet? materializedSource = null;

        if (baseSheet is not null)
        {
            materializedBase = await WorksheetMaterializer.ReadAsync(
                baseSheet,
                normalized,
                cancellationToken).ConfigureAwait(false);
        }

        if (sourceSheet is not null)
        {
            materializedSource = await WorksheetMaterializer.ReadAsync(
                sourceSheet,
                normalized,
                cancellationToken).ConfigureAwait(false);
        }

        return DiffBuilder.Build(
            materializedBase,
            materializedSource,
            sourceSide,
            context,
            cancellationToken).Result;
    }
}

internal sealed record BuiltDiff(
    TwoWayDiffResult Result,
    PreparedRows BaseRows,
    PreparedRows SourceRows,
    AlignmentStep[] Alignment);

internal static class DiffBuilder
{
    public static BuiltDiff Build(
        MaterializedWorksheet? baseSheet,
        MaterializedWorksheet? sourceSheet,
        SourceSide sourceSide,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        var baseRows = PreparedRows.Create(
            baseSheet?.Rows ?? Array.Empty<RowRecord>(),
            context,
            cancellationToken);
        var sourceRows = PreparedRows.Create(
            sourceSheet?.Rows ?? Array.Empty<RowRecord>(),
            context,
            cancellationToken);
        return Build(
            baseSheet,
            sourceSheet,
            sourceSide,
            context,
            baseRows,
            sourceRows,
            cancellationToken);
    }

    public static BuiltDiff Build(
        MaterializedWorksheet? baseSheet,
        MaterializedWorksheet? sourceSheet,
        SourceSide sourceSide,
        ComparisonContext context,
        PreparedRows baseRows,
        PreparedRows sourceRows,
        CancellationToken cancellationToken)
    {
        var alignment = RowAligner.Align(baseRows, sourceRows, context, cancellationToken);
        var rowChanges = new List<RowChange>();
        var viewRows = new List<DiffViewRow>(alignment.Length);
        var addedRows = 0;
        var removedRows = 0;
        var modifiedRows = 0;
        long changedCells = 0;

        for (var index = 0; index < alignment.Length; index++)
        {
            if ((index % context.Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            var step = alignment[index];
            var kind = RowAligner.GetChangeKind(
                step,
                baseRows,
                sourceRows,
                context,
                cancellationToken);
            CellChange[] cellChanges;
            if (kind == ChangeKind.Unchanged)
            {
                cellChanges = Array.Empty<CellChange>();
            }
            else
            {
                cellChanges = BuildCellChanges(
                    step,
                    baseRows,
                    sourceRows,
                    sourceSide,
                    context,
                    cancellationToken);
                rowChanges.Add(new RowChange(
                    sourceSide,
                    kind,
                    step.BaseOrdinal >= 0 ? baseRows.Rows[step.BaseOrdinal].RowIndex : null,
                    step.SourceOrdinal >= 0 ? sourceRows.Rows[step.SourceOrdinal].RowIndex : null,
                    cellChanges));
                changedCells += cellChanges.Length;

                switch (kind)
                {
                    case ChangeKind.Added:
                        addedRows++;
                        break;
                    case ChangeKind.Removed:
                        removedRows++;
                        break;
                    case ChangeKind.Modified:
                        modifiedRows++;
                        break;
                }
            }

            if (kind != ChangeKind.Unchanged || context.Options.Source.IncludeUnchangedViewRows)
            {
                viewRows.Add(new DiffViewRow(
                    step.BaseOrdinal >= 0 ? baseRows.Rows[step.BaseOrdinal].RowIndex : null,
                    step.SourceOrdinal >= 0 ? sourceRows.Rows[step.SourceOrdinal].RowIndex : null,
                    kind,
                    BuildRuns(cellChanges)));
            }
        }

        var sheetKind = GetSheetChangeKind(
            baseSheet,
            sourceSheet,
            rowChanges.Count != 0,
            context);
        var sheetChange = new SheetChange(
            sourceSide,
            sheetKind,
            baseSheet?.Metadata,
            sourceSheet?.Metadata,
            rowChanges.ToArray());
        var changedRowCount = addedRows + removedRows + modifiedRows;
        var statistics = new DiffStatistics(
            viewRows.Count,
            changedRowCount,
            addedRows,
            removedRows,
            modifiedRows,
            changedCells);
        var result = new TwoWayDiffResult(sheetChange, viewRows.ToArray(), statistics);
        return new BuiltDiff(result, baseRows, sourceRows, alignment);
    }

    private static ChangeKind GetSheetChangeKind(
        MaterializedWorksheet? baseSheet,
        MaterializedWorksheet? sourceSheet,
        bool hasRowChanges,
        ComparisonContext context)
    {
        if (baseSheet is null)
            return ChangeKind.Added;
        if (sourceSheet is null)
            return ChangeKind.Removed;
        return hasRowChanges || !context.SheetMetadataEqual(baseSheet.Metadata, sourceSheet.Metadata)
            ? ChangeKind.Modified
            : ChangeKind.Unchanged;
    }

    private static CellChange[] BuildCellChanges(
        AlignmentStep step,
        PreparedRows baseRows,
        PreparedRows sourceRows,
        SourceSide sourceSide,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        if (step.BaseOrdinal < 0)
        {
            return BuildWholeRowChanges(
                sourceRows.Rows[step.SourceOrdinal],
                sourceSide,
                added: true,
                context,
                cancellationToken);
        }

        if (step.SourceOrdinal < 0)
        {
            return BuildWholeRowChanges(
                baseRows.Rows[step.BaseOrdinal],
                sourceSide,
                added: false,
                context,
                cancellationToken);
        }

        var baseCells = baseRows.Rows[step.BaseOrdinal].Cells.Span;
        var sourceCells = sourceRows.Rows[step.SourceOrdinal].Cells.Span;
        var baseIndex = 0;
        var sourceIndex = 0;
        var changes = new List<CellChange>();
        var work = 0;

        while (baseIndex < baseCells.Length || sourceIndex < sourceCells.Length)
        {
            if ((work++ % context.Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            baseIndex = NextSignificant(baseCells, baseIndex, context);
            sourceIndex = NextSignificant(sourceCells, sourceIndex, context);
            if (baseIndex == baseCells.Length && sourceIndex == sourceCells.Length)
                break;

            if (sourceIndex == sourceCells.Length ||
                (baseIndex < baseCells.Length &&
                    baseCells[baseIndex].Address.ColumnIndex < sourceCells[sourceIndex].Address.ColumnIndex))
            {
                ref readonly var removed = ref baseCells[baseIndex++];
                changes.Add(new CellChange(
                    sourceSide,
                    ChangeKind.Removed,
                    removed.Address,
                    null,
                    removed.Value,
                    null));
                continue;
            }

            if (baseIndex == baseCells.Length ||
                sourceCells[sourceIndex].Address.ColumnIndex < baseCells[baseIndex].Address.ColumnIndex)
            {
                ref readonly var added = ref sourceCells[sourceIndex++];
                changes.Add(new CellChange(
                    sourceSide,
                    ChangeKind.Added,
                    null,
                    added.Address,
                    null,
                    added.Value));
                continue;
            }

            ref readonly var baseCell = ref baseCells[baseIndex++];
            ref readonly var sourceCell = ref sourceCells[sourceIndex++];
            if (!context.CellRecordsEqual(baseCell, sourceCell, cancellationToken))
            {
                changes.Add(new CellChange(
                    sourceSide,
                    ChangeKind.Modified,
                    baseCell.Address,
                    sourceCell.Address,
                    baseCell.Value,
                    sourceCell.Value));
            }
        }

        return changes.ToArray();
    }

    private static CellChange[] BuildWholeRowChanges(
        in RowRecord row,
        SourceSide sourceSide,
        bool added,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        var changes = new List<CellChange>(row.CellCount);
        var cells = row.Cells.Span;
        for (var index = 0; index < cells.Length; index++)
        {
            if ((index % context.Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            ref readonly var cell = ref cells[index];
            if (context.IsIgnoredSparseCell(cell))
                continue;

            changes.Add(added
                ? new CellChange(
                    sourceSide,
                    ChangeKind.Added,
                    null,
                    cell.Address,
                    null,
                    cell.Value)
                : new CellChange(
                    sourceSide,
                    ChangeKind.Removed,
                    cell.Address,
                    null,
                    cell.Value,
                    null));
        }

        return changes.ToArray();
    }

    private static ColumnChangeRun[] BuildRuns(CellChange[] cellChanges)
    {
        if (cellChanges.Length == 0)
            return Array.Empty<ColumnChangeRun>();

        var runs = new List<ColumnChangeRun>();
        var startColumn = GetColumn(cellChanges[0]);
        var previousColumn = startColumn;
        var kind = cellChanges[0].Kind;
        for (var index = 1; index < cellChanges.Length; index++)
        {
            var column = GetColumn(cellChanges[index]);
            if (column != previousColumn + 1 || cellChanges[index].Kind != kind)
            {
                runs.Add(new ColumnChangeRun(startColumn, previousColumn - startColumn + 1, kind));
                startColumn = column;
                kind = cellChanges[index].Kind;
            }

            previousColumn = column;
        }

        runs.Add(new ColumnChangeRun(startColumn, previousColumn - startColumn + 1, kind));
        return runs.ToArray();
    }

    private static int GetColumn(in CellChange change) =>
        change.SourceAddress?.ColumnIndex ?? change.BaseAddress!.Value.ColumnIndex;

    private static int NextSignificant(
        ReadOnlySpan<CellRecord> cells,
        int index,
        ComparisonContext context)
    {
        while (index < cells.Length && context.IsIgnoredSparseCell(cells[index]))
            index++;
        return index;
    }
}

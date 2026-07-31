using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

public sealed class ThreeWayMergeEngine
{
    public async ValueTask<ThreeWayMergeResult> MergeAsync(
        IWorksheetSnapshot? baseSheet,
        IWorksheetSnapshot? localSheet,
        IWorksheetSnapshot? remoteSheet,
        ThreeWayMergeOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (baseSheet is null && localSheet is null && remoteSheet is null)
            throw new ArgumentException("At least one worksheet snapshot is required.");

        options ??= new ThreeWayMergeOptions();
        ArgumentNullException.ThrowIfNull(options.Comparison);
        ArgumentOutOfRangeException.ThrowIfNegative(options.FirstConflictId);
        if (options.FirstConflictId == long.MaxValue)
        {
            throw new ArgumentOutOfRangeException(
                nameof(ThreeWayMergeOptions.FirstConflictId),
                options.FirstConflictId,
                "The first conflict identifier must leave room for allocation.");
        }
        var normalized = new NormalizedComparisonOptions(options.Comparison);
        var context = new ComparisonContext(normalized);

        var materializedBase = await ReadOptionalAsync(
            baseSheet,
            normalized,
            cancellationToken).ConfigureAwait(false);
        var materializedLocal = await ReadOptionalAsync(
            localSheet,
            normalized,
            cancellationToken).ConfigureAwait(false);
        var materializedRemote = await ReadOptionalAsync(
            remoteSheet,
            normalized,
            cancellationToken).ConfigureAwait(false);

        var preparedBase = PreparedRows.Create(
            materializedBase?.Rows ?? Array.Empty<RowRecord>(),
            context,
            cancellationToken);
        var preparedLocal = PreparedRows.Create(
            materializedLocal?.Rows ?? Array.Empty<RowRecord>(),
            context,
            cancellationToken);
        var preparedRemote = PreparedRows.Create(
            materializedRemote?.Rows ?? Array.Empty<RowRecord>(),
            context,
            cancellationToken);

        var localDiff = BuildSideDiff(
            materializedBase,
            materializedLocal,
            SourceSide.Local,
            preparedBase,
            preparedLocal,
            context,
            cancellationToken);
        var remoteDiff = BuildSideDiff(
            materializedBase,
            materializedRemote,
            SourceSide.Remote,
            preparedBase,
            preparedRemote,
            context,
            cancellationToken);
        var accumulator = new MergeAccumulator(options.FirstConflictId);
        var sheetId = materializedBase?.Metadata.Id ??
            materializedLocal?.Metadata.Id ??
            materializedRemote!.Metadata.Id;

        if (materializedBase is null || materializedLocal is null || materializedRemote is null)
        {
            MergeStructuralSheetCase(
                materializedBase,
                materializedLocal,
                materializedRemote,
                preparedBase,
                preparedLocal,
                preparedRemote,
                localDiff,
                remoteDiff,
                sheetId,
                context,
                accumulator,
                cancellationToken);
        }
        else
        {
            MergeExistingSheets(
                materializedBase,
                materializedLocal,
                materializedRemote,
                preparedBase,
                preparedLocal,
                preparedRemote,
                localDiff,
                remoteDiff,
                sheetId,
                context,
                accumulator,
                cancellationToken);
        }

        return accumulator.CreateResult(
            localDiff.Result.SheetChange,
            remoteDiff.Result.SheetChange);
    }

    private static async ValueTask<MaterializedWorksheet?> ReadOptionalAsync(
        IWorksheetSnapshot? snapshot,
        NormalizedComparisonOptions options,
        CancellationToken cancellationToken)
    {
        return snapshot is null
            ? null
            : await WorksheetMaterializer.ReadAsync(snapshot, options, cancellationToken)
                .ConfigureAwait(false);
    }

    private static BuiltDiff BuildSideDiff(
        MaterializedWorksheet? baseSheet,
        MaterializedWorksheet? sourceSheet,
        SourceSide sourceSide,
        PreparedRows baseRows,
        PreparedRows sourceRows,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        if (baseSheet is null && sourceSheet is null)
        {
            var emptyChange = new SheetChange(
                sourceSide,
                ChangeKind.Unchanged,
                null,
                null,
                ReadOnlyMemory<RowChange>.Empty);
            var emptyResult = new TwoWayDiffResult(
                emptyChange,
                ReadOnlyMemory<DiffViewRow>.Empty,
                default);
            return new BuiltDiff(emptyResult, baseRows, sourceRows, Array.Empty<AlignmentStep>());
        }

        return DiffBuilder.Build(
            baseSheet,
            sourceSheet,
            sourceSide,
            context,
            baseRows,
            sourceRows,
            cancellationToken);
    }

    private static void MergeExistingSheets(
        MaterializedWorksheet baseSheet,
        MaterializedWorksheet localSheet,
        MaterializedWorksheet remoteSheet,
        PreparedRows baseRows,
        PreparedRows localRows,
        PreparedRows remoteRows,
        BuiltDiff localDiff,
        BuiltDiff remoteDiff,
        string sheetId,
        ComparisonContext context,
        MergeAccumulator accumulator,
        CancellationToken cancellationToken)
    {
        MergeSheetMetadata(
            baseSheet.Metadata,
            localSheet.Metadata,
            remoteSheet.Metadata,
            sheetId,
            context,
            accumulator);

        var localMap = SideAlignmentMap.Create(
            baseRows.Rows.Length,
            localDiff.Alignment,
            context.Options.Source.CancellationCheckInterval,
            cancellationToken);
        var remoteMap = SideAlignmentMap.Create(
            baseRows.Rows.Length,
            remoteDiff.Alignment,
            context.Options.Source.CancellationCheckInterval,
            cancellationToken);
        for (var baseOrdinal = 0; baseOrdinal <= baseRows.Rows.Length; baseOrdinal++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            MergeInsertionSlot(
                localMap.GetInsertions(baseOrdinal),
                remoteMap.GetInsertions(baseOrdinal),
                localRows,
                remoteRows,
                sheetId,
                context,
                accumulator,
                cancellationToken);

            if (baseOrdinal == baseRows.Rows.Length)
                break;

            MergeBaseRow(
                baseOrdinal,
                localMap.BaseToSource[baseOrdinal],
                remoteMap.BaseToSource[baseOrdinal],
                baseRows,
                localRows,
                remoteRows,
                sheetId,
                context,
                accumulator,
                cancellationToken);
        }
    }

    private static void MergeSheetMetadata(
        SheetMetadata baseSheet,
        SheetMetadata localSheet,
        SheetMetadata remoteSheet,
        string sheetId,
        ComparisonContext context,
        MergeAccumulator accumulator)
    {
        var localChanged = !context.SheetMetadataEqual(baseSheet, localSheet);
        var remoteChanged = !context.SheetMetadataEqual(baseSheet, remoteSheet);
        if (!localChanged && !remoteChanged)
            return;

        var location = new ConflictLocation(sheetId);
        if (localChanged && !remoteChanged)
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.SheetMetadata,
                AutomaticMergeKind.UseLocal,
                location));
        }
        else if (!localChanged)
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.SheetMetadata,
                AutomaticMergeKind.UseRemote,
                location));
        }
        else if (context.SheetMetadataEqual(localSheet, remoteSheet))
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.SheetMetadata,
                AutomaticMergeKind.UseEither,
                location));
        }
        else
        {
            accumulator.AddRowConflict(ConflictKind.AmbiguousSheetRename, location);
        }
    }

    private static void MergeInsertionSlot(
        IReadOnlyList<int> localIndices,
        IReadOnlyList<int> remoteIndices,
        PreparedRows localRows,
        PreparedRows remoteRows,
        string sheetId,
        ComparisonContext context,
        MergeAccumulator accumulator,
        CancellationToken cancellationToken)
    {
        if (localIndices.Count == 0 && remoteIndices.Count == 0)
            return;

        var localBlock = PreparedRows.Select(
            localRows,
            localIndices,
            context.Options.Source.CancellationCheckInterval,
            cancellationToken);
        var remoteBlock = PreparedRows.Select(
            remoteRows,
            remoteIndices,
            context.Options.Source.CancellationCheckInterval,
            cancellationToken);
        var alignment = RowAligner.Align(localBlock, remoteBlock, context, cancellationToken);
        var conflictKind = localIndices.Count == 1 && remoteIndices.Count == 1
            ? ConflictKind.RowAddAdd
            : ConflictKind.IncompatibleRowInsertion;

        var alignmentIndex = 0;
        while (alignmentIndex < alignment.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var step = alignment[alignmentIndex];
            if (IsExactInsertionStep(
                step,
                localBlock,
                remoteBlock,
                context,
                cancellationToken))
            {
                AddInsertionViewRow(
                    step,
                    forceConflict: false,
                    conflictKind,
                    localIndices,
                    remoteIndices,
                    localRows,
                    remoteRows,
                    sheetId,
                    accumulator);
                alignmentIndex++;
                continue;
            }

            var gapStart = alignmentIndex;
            var hasLocalRows = false;
            var hasRemoteRows = false;
            while (alignmentIndex < alignment.Length && !IsExactInsertionStep(
                alignment[alignmentIndex],
                localBlock,
                remoteBlock,
                context,
                cancellationToken))
            {
                hasLocalRows |= alignment[alignmentIndex].BaseOrdinal >= 0;
                hasRemoteRows |= alignment[alignmentIndex].SourceOrdinal >= 0;
                alignmentIndex++;
            }

            var forceConflict = hasLocalRows && hasRemoteRows;
            for (var gapIndex = gapStart; gapIndex < alignmentIndex; gapIndex++)
            {
                if (((gapIndex - gapStart) % context.Options.Source.CancellationCheckInterval) == 0)
                    cancellationToken.ThrowIfCancellationRequested();
                AddInsertionViewRow(
                    alignment[gapIndex],
                    forceConflict,
                    conflictKind,
                    localIndices,
                    remoteIndices,
                    localRows,
                    remoteRows,
                    sheetId,
                    accumulator);
            }
        }
    }

    private static bool IsExactInsertionStep(
        AlignmentStep step,
        PreparedRows localBlock,
        PreparedRows remoteBlock,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        return step.BaseOrdinal >= 0 &&
            step.SourceOrdinal >= 0 &&
            RowAligner.RowsExactlyEqual(
                localBlock,
                step.BaseOrdinal,
                remoteBlock,
                step.SourceOrdinal,
                context,
                cancellationToken);
    }

    private static void AddInsertionViewRow(
        AlignmentStep step,
        bool forceConflict,
        ConflictKind conflictKind,
        IReadOnlyList<int> localIndices,
        IReadOnlyList<int> remoteIndices,
        PreparedRows localRows,
        PreparedRows remoteRows,
        string sheetId,
        MergeAccumulator accumulator)
    {
        var localOrdinal = step.BaseOrdinal >= 0 ? localIndices[step.BaseOrdinal] : -1;
        var remoteOrdinal = step.SourceOrdinal >= 0 ? remoteIndices[step.SourceOrdinal] : -1;
        var localRowIndex = localOrdinal >= 0 ? localRows.Rows[localOrdinal].RowIndex : (int?)null;
        var remoteRowIndex = remoteOrdinal >= 0
            ? remoteRows.Rows[remoteOrdinal].RowIndex
            : (int?)null;
        var location = new ConflictLocation(
            sheetId,
            LocalRowIndex: localRowIndex,
            RemoteRowIndex: remoteRowIndex);
        var conflictStart = accumulator.ConflictCount;

        if (forceConflict)
        {
            accumulator.AddRowConflict(conflictKind, location);
        }
        else if (localOrdinal < 0)
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.Row,
                AutomaticMergeKind.UseRemote,
                location));
        }
        else if (remoteOrdinal < 0)
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.Row,
                AutomaticMergeKind.UseLocal,
                location));
        }
        else
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.Row,
                AutomaticMergeKind.UseEither,
                location));
        }

        var conflictCount = accumulator.ConflictCount - conflictStart;
        accumulator.AddViewRow(new MergeViewRow(
            null,
            localRowIndex,
            remoteRowIndex,
            localOrdinal >= 0 ? ChangeKind.Added : ChangeKind.Unchanged,
            remoteOrdinal >= 0 ? ChangeKind.Added : ChangeKind.Unchanged,
            conflictCount == 0 ? MergeViewRowState.Automatic : MergeViewRowState.Conflict,
            conflictCount == 0 ? -1 : conflictStart,
            conflictCount));
    }

    private static void MergeBaseRow(
        int baseOrdinal,
        int localOrdinal,
        int remoteOrdinal,
        PreparedRows baseRows,
        PreparedRows localRows,
        PreparedRows remoteRows,
        string sheetId,
        ComparisonContext context,
        MergeAccumulator accumulator,
        CancellationToken cancellationToken)
    {
        var baseRow = baseRows.Rows[baseOrdinal];
        var localRowIndex = localOrdinal >= 0 ? localRows.Rows[localOrdinal].RowIndex : (int?)null;
        var remoteRowIndex = remoteOrdinal >= 0
            ? remoteRows.Rows[remoteOrdinal].RowIndex
            : (int?)null;
        var localKind = localOrdinal < 0
            ? ChangeKind.Removed
            : RowAligner.RowsExactlyEqual(
                baseRows,
                baseOrdinal,
                localRows,
                localOrdinal,
                context,
                cancellationToken)
                ? ChangeKind.Unchanged
                : ChangeKind.Modified;
        var remoteKind = remoteOrdinal < 0
            ? ChangeKind.Removed
            : RowAligner.RowsExactlyEqual(
                baseRows,
                baseOrdinal,
                remoteRows,
                remoteOrdinal,
                context,
                cancellationToken)
                ? ChangeKind.Unchanged
                : ChangeKind.Modified;
        var location = new ConflictLocation(
            sheetId,
            baseRow.RowIndex,
            localRowIndex,
            remoteRowIndex);
        var conflictStart = accumulator.ConflictCount;
        var decisionStart = accumulator.DecisionCount;

        if (localOrdinal < 0 && remoteOrdinal < 0)
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.Row,
                AutomaticMergeKind.Delete,
                location));
        }
        else if (localOrdinal < 0 || remoteOrdinal < 0)
        {
            var retainedKind = localOrdinal < 0 ? remoteKind : localKind;
            if (retainedKind == ChangeKind.Unchanged)
            {
                accumulator.AddDecision(new AutomaticMergeDecision(
                    MergeDecisionScope.Row,
                    AutomaticMergeKind.Delete,
                    location));
            }
            else
            {
                accumulator.AddRowConflict(ConflictKind.RowDeleteEdit, location);
            }
        }
        else
        {
            MergeRowMetadata(
                baseRow,
                localRows.Rows[localOrdinal],
                remoteRows.Rows[remoteOrdinal],
                location,
                context,
                accumulator);
            MergeCells(
                baseRow,
                localRows.Rows[localOrdinal],
                remoteRows.Rows[remoteOrdinal],
                location,
                context,
                accumulator,
                cancellationToken);
        }

        var conflictCount = accumulator.ConflictCount - conflictStart;
        var hasDecision = accumulator.DecisionCount != decisionStart;
        var state = conflictCount != 0
            ? MergeViewRowState.Conflict
            : hasDecision || localKind != ChangeKind.Unchanged || remoteKind != ChangeKind.Unchanged
                ? MergeViewRowState.Automatic
                : MergeViewRowState.Unchanged;
        if (state != MergeViewRowState.Unchanged || context.Options.Source.IncludeUnchangedViewRows)
        {
            accumulator.AddViewRow(new MergeViewRow(
                baseRow.RowIndex,
                localRowIndex,
                remoteRowIndex,
                localKind,
                remoteKind,
                state,
                conflictCount == 0 ? -1 : conflictStart,
                conflictCount));
        }
    }

    private static void MergeRowMetadata(
        in RowRecord baseRow,
        in RowRecord localRow,
        in RowRecord remoteRow,
        ConflictLocation location,
        ComparisonContext context,
        MergeAccumulator accumulator)
    {
        if (!context.Options.Source.CompareRowMetadata)
            return;

        var localChanged = !context.RowMetadataEqual(baseRow, localRow);
        var remoteChanged = !context.RowMetadataEqual(baseRow, remoteRow);
        if (!localChanged && !remoteChanged)
            return;

        if (localChanged && !remoteChanged)
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.RowMetadata,
                AutomaticMergeKind.UseLocal,
                location));
        }
        else if (!localChanged)
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.RowMetadata,
                AutomaticMergeKind.UseRemote,
                location));
        }
        else if (context.RowMetadataEqual(localRow, remoteRow))
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.RowMetadata,
                AutomaticMergeKind.UseEither,
                location));
        }
        else
        {
            accumulator.AddRowConflict(ConflictKind.IncompatibleRowInsertion, location);
        }
    }

    private static void MergeCells(
        in RowRecord baseRow,
        in RowRecord localRow,
        in RowRecord remoteRow,
        ConflictLocation rowLocation,
        ComparisonContext context,
        MergeAccumulator accumulator,
        CancellationToken cancellationToken)
    {
        var baseCells = baseRow.Cells.Span;
        var localCells = localRow.Cells.Span;
        var remoteCells = remoteRow.Cells.Span;
        var baseIndex = 0;
        var localIndex = 0;
        var remoteIndex = 0;
        var work = 0;

        while (baseIndex < baseCells.Length ||
            localIndex < localCells.Length ||
            remoteIndex < remoteCells.Length)
        {
            if ((work++ % context.Options.Source.CancellationCheckInterval) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            baseIndex = NextSignificant(baseCells, baseIndex, context);
            localIndex = NextSignificant(localCells, localIndex, context);
            remoteIndex = NextSignificant(remoteCells, remoteIndex, context);
            if (baseIndex == baseCells.Length &&
                localIndex == localCells.Length &&
                remoteIndex == remoteCells.Length)
            {
                break;
            }

            var columnIndex = Math.Min(
                GetColumnOrMax(baseCells, baseIndex),
                Math.Min(
                    GetColumnOrMax(localCells, localIndex),
                    GetColumnOrMax(remoteCells, remoteIndex)));
            CellRecord? baseCell = TakeCell(baseCells, ref baseIndex, columnIndex);
            CellRecord? localCell = TakeCell(localCells, ref localIndex, columnIndex);
            CellRecord? remoteCell = TakeCell(remoteCells, ref remoteIndex, columnIndex);
            var localChanged = !CellsEqual(baseCell, localCell, context, cancellationToken);
            var remoteChanged = !CellsEqual(baseCell, remoteCell, context, cancellationToken);
            if (!localChanged && !remoteChanged)
                continue;

            var location = rowLocation with { ColumnIndex = columnIndex };
            if (localChanged && !remoteChanged)
            {
                accumulator.AddDecision(CreateCellDecision(
                    AutomaticMergeKind.UseLocal,
                    localCell,
                    location));
            }
            else if (!localChanged)
            {
                accumulator.AddDecision(CreateCellDecision(
                    AutomaticMergeKind.UseRemote,
                    remoteCell,
                    location));
            }
            else if (CellsEqual(localCell, remoteCell, context, cancellationToken))
            {
                accumulator.AddDecision(CreateCellDecision(
                    AutomaticMergeKind.UseEither,
                    localCell,
                    location));
            }
            else
            {
                var kind = baseCell.HasValue && (!localCell.HasValue || !remoteCell.HasValue)
                    ? ConflictKind.CellDeleteEdit
                    : ConflictKind.CellValue;
                accumulator.AddCellConflict(
                    kind,
                    location,
                    baseCell?.Value,
                    localCell?.Value,
                    remoteCell?.Value);
            }
        }
    }

    private static AutomaticMergeDecision CreateCellDecision(
        AutomaticMergeKind sourceKind,
        CellRecord? selectedCell,
        ConflictLocation location)
    {
        return selectedCell.HasValue
            ? new AutomaticMergeDecision(
                MergeDecisionScope.Cell,
                sourceKind,
                location,
                selectedCell.Value.Value)
            : new AutomaticMergeDecision(
                MergeDecisionScope.Cell,
                AutomaticMergeKind.Delete,
                location);
    }

    private static bool CellsEqual(
        CellRecord? left,
        CellRecord? right,
        ComparisonContext context,
        CancellationToken cancellationToken)
    {
        if (left.HasValue != right.HasValue)
            return false;
        return !left.HasValue || context.CellRecordsEqual(
            left.GetValueOrDefault(),
            right.GetValueOrDefault(),
            cancellationToken);
    }

    private static CellRecord? TakeCell(
        ReadOnlySpan<CellRecord> cells,
        ref int index,
        int columnIndex)
    {
        if (index >= cells.Length || cells[index].Address.ColumnIndex != columnIndex)
            return null;
        return cells[index++];
    }

    private static int GetColumnOrMax(ReadOnlySpan<CellRecord> cells, int index) =>
        index < cells.Length ? cells[index].Address.ColumnIndex : int.MaxValue;

    private static int NextSignificant(
        ReadOnlySpan<CellRecord> cells,
        int index,
        ComparisonContext context)
    {
        while (index < cells.Length && context.IsIgnoredSparseCell(cells[index]))
            index++;
        return index;
    }

    private static void MergeStructuralSheetCase(
        MaterializedWorksheet? baseSheet,
        MaterializedWorksheet? localSheet,
        MaterializedWorksheet? remoteSheet,
        PreparedRows baseRows,
        PreparedRows localRows,
        PreparedRows remoteRows,
        BuiltDiff localDiff,
        BuiltDiff remoteDiff,
        string sheetId,
        ComparisonContext context,
        MergeAccumulator accumulator,
        CancellationToken cancellationToken)
    {
        if (baseSheet is null)
        {
            MergeAddedSheetCase(
                localSheet,
                remoteSheet,
                localRows,
                remoteRows,
                sheetId,
                context,
                accumulator,
                cancellationToken);
            return;
        }

        if (localSheet is null && remoteSheet is null)
        {
            var location = new ConflictLocation(sheetId);
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.Sheet,
                AutomaticMergeKind.Delete,
                location));
            foreach (var row in baseRows.Rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                accumulator.AddViewRow(new MergeViewRow(
                    row.RowIndex,
                    null,
                    null,
                    ChangeKind.Removed,
                    ChangeKind.Removed,
                    MergeViewRowState.Automatic,
                    -1,
                    0));
            }

            return;
        }

        var retainedDiff = localSheet is null ? remoteDiff : localDiff;
        var unchanged = retainedDiff.Result.SheetChange.Kind == ChangeKind.Unchanged;
        var conflictIndex = -1;
        if (unchanged)
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.Sheet,
                AutomaticMergeKind.Delete,
                new ConflictLocation(sheetId)));
        }
        else
        {
            conflictIndex = accumulator.AddRowConflict(
                ConflictKind.SheetDeleteEdit,
                new ConflictLocation(sheetId));
        }

        AddDeleteEditViews(
            localSheet is null,
            retainedDiff,
            conflictIndex,
            context,
            accumulator,
            cancellationToken);
    }

    private static void MergeAddedSheetCase(
        MaterializedWorksheet? localSheet,
        MaterializedWorksheet? remoteSheet,
        PreparedRows localRows,
        PreparedRows remoteRows,
        string sheetId,
        ComparisonContext context,
        MergeAccumulator accumulator,
        CancellationToken cancellationToken)
    {
        var location = new ConflictLocation(sheetId);
        if (localSheet is null || remoteSheet is null)
        {
            var useLocal = localSheet is not null;
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.Sheet,
                useLocal ? AutomaticMergeKind.UseLocal : AutomaticMergeKind.UseRemote,
                location));
            var rows = useLocal ? localRows.Rows : remoteRows.Rows;
            foreach (var row in rows)
            {
                cancellationToken.ThrowIfCancellationRequested();
                accumulator.AddViewRow(new MergeViewRow(
                    null,
                    useLocal ? row.RowIndex : null,
                    useLocal ? null : row.RowIndex,
                    useLocal ? ChangeKind.Added : ChangeKind.Unchanged,
                    useLocal ? ChangeKind.Unchanged : ChangeKind.Added,
                    MergeViewRowState.Automatic,
                    -1,
                    0));
            }

            return;
        }

        var alignment = RowAligner.Align(localRows, remoteRows, context, cancellationToken);
        var equivalent = context.SheetMetadataEqual(localSheet.Metadata, remoteSheet.Metadata);
        foreach (var step in alignment)
        {
            if (step.BaseOrdinal < 0 || step.SourceOrdinal < 0 ||
                !RowAligner.RowsExactlyEqual(
                    localRows,
                    step.BaseOrdinal,
                    remoteRows,
                    step.SourceOrdinal,
                    context,
                    cancellationToken))
            {
                equivalent = false;
                break;
            }
        }

        var conflictIndex = -1;
        if (equivalent)
        {
            accumulator.AddDecision(new AutomaticMergeDecision(
                MergeDecisionScope.Sheet,
                AutomaticMergeKind.UseEither,
                location));
        }
        else
        {
            conflictIndex = accumulator.AddRowConflict(ConflictKind.SheetAddAdd, location);
        }

        foreach (var step in alignment)
        {
            cancellationToken.ThrowIfCancellationRequested();
            accumulator.AddViewRow(new MergeViewRow(
                null,
                step.BaseOrdinal >= 0 ? localRows.Rows[step.BaseOrdinal].RowIndex : null,
                step.SourceOrdinal >= 0 ? remoteRows.Rows[step.SourceOrdinal].RowIndex : null,
                step.BaseOrdinal >= 0 ? ChangeKind.Added : ChangeKind.Unchanged,
                step.SourceOrdinal >= 0 ? ChangeKind.Added : ChangeKind.Unchanged,
                equivalent ? MergeViewRowState.Automatic : MergeViewRowState.Conflict,
                conflictIndex,
                conflictIndex < 0 ? 0 : 1));
        }
    }

    private static void AddDeleteEditViews(
        bool localDeleted,
        BuiltDiff retainedDiff,
        int conflictIndex,
        ComparisonContext context,
        MergeAccumulator accumulator,
        CancellationToken cancellationToken)
    {
        foreach (var step in retainedDiff.Alignment)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var baseRowIndex = step.BaseOrdinal >= 0
                ? retainedDiff.BaseRows.Rows[step.BaseOrdinal].RowIndex
                : (int?)null;
            var retainedRowIndex = step.SourceOrdinal >= 0
                ? retainedDiff.SourceRows.Rows[step.SourceOrdinal].RowIndex
                : (int?)null;
            var retainedKind = RowAligner.GetChangeKind(
                step,
                retainedDiff.BaseRows,
                retainedDiff.SourceRows,
                context,
                cancellationToken);
            var deletedKind = step.BaseOrdinal >= 0 ? ChangeKind.Removed : ChangeKind.Unchanged;
            accumulator.AddViewRow(new MergeViewRow(
                baseRowIndex,
                localDeleted ? null : retainedRowIndex,
                localDeleted ? retainedRowIndex : null,
                localDeleted ? deletedKind : retainedKind,
                localDeleted ? retainedKind : deletedKind,
                conflictIndex < 0 ? MergeViewRowState.Automatic : MergeViewRowState.Conflict,
                conflictIndex,
                conflictIndex < 0 ? 0 : 1));
        }
    }

    private sealed class SideAlignmentMap
    {
        private static readonly IReadOnlyList<int> Empty = Array.Empty<int>();
        private readonly Dictionary<int, List<int>> _insertions;

        private SideAlignmentMap(int[] baseToSource, Dictionary<int, List<int>> insertions)
        {
            BaseToSource = baseToSource;
            _insertions = insertions;
        }

        public int[] BaseToSource { get; }

        public IReadOnlyList<int> GetInsertions(int slot) =>
            _insertions.TryGetValue(slot, out var rows) ? rows : Empty;

        public static SideAlignmentMap Create(
            int baseRowCount,
            AlignmentStep[] alignment,
            int cancellationCheckInterval,
            CancellationToken cancellationToken)
        {
            var map = new int[baseRowCount];
            Array.Fill(map, -1);
            var insertions = new Dictionary<int, List<int>>();
            var slot = 0;
            for (var index = 0; index < alignment.Length; index++)
            {
                if ((index % cancellationCheckInterval) == 0)
                    cancellationToken.ThrowIfCancellationRequested();
                var step = alignment[index];
                if (step.BaseOrdinal < 0)
                {
                    if (!insertions.TryGetValue(slot, out var rows))
                    {
                        rows = [];
                        insertions.Add(slot, rows);
                    }

                    rows.Add(step.SourceOrdinal);
                }
                else
                {
                    map[step.BaseOrdinal] = step.SourceOrdinal;
                    slot = step.BaseOrdinal + 1;
                }
            }

            return new SideAlignmentMap(map, insertions);
        }
    }

    private sealed class MergeAccumulator
    {
        private readonly List<ConflictRecord> _conflicts = [];
        private readonly List<CellResolution> _cellResolutions = [];
        private readonly List<RowResolution> _rowResolutions = [];
        private readonly List<AutomaticMergeDecision> _decisions = [];
        private readonly List<MergeViewRow> _viewRows = [];
        private long _nextConflictId;

        public MergeAccumulator(long firstConflictId)
        {
            _nextConflictId = firstConflictId;
        }

        public int ConflictCount => _conflicts.Count;

        public int DecisionCount => _decisions.Count;

        public void AddDecision(AutomaticMergeDecision decision) => _decisions.Add(decision);

        public void AddViewRow(MergeViewRow row) => _viewRows.Add(row);

        public int AddCellConflict(
            ConflictKind kind,
            ConflictLocation location,
            CellValue? baseValue,
            CellValue? localValue,
            CellValue? remoteValue)
        {
            var id = AllocateConflictId();
            var index = _conflicts.Count;
            _conflicts.Add(new ConflictRecord(
                id,
                kind,
                location,
                baseValue,
                localValue,
                remoteValue));
            _cellResolutions.Add(CellResolution.Unresolved(id));
            return index;
        }

        public int AddRowConflict(ConflictKind kind, ConflictLocation location)
        {
            var id = AllocateConflictId();
            var index = _conflicts.Count;
            _conflicts.Add(new ConflictRecord(id, kind, location));
            _rowResolutions.Add(RowResolution.Unresolved(id));
            return index;
        }

        public ThreeWayMergeResult CreateResult(SheetChange localChange, SheetChange remoteChange) =>
            new(
                localChange,
                remoteChange,
                _conflicts.ToArray(),
                _cellResolutions.ToArray(),
                _rowResolutions.ToArray(),
                _decisions.ToArray(),
                _viewRows.ToArray(),
                _nextConflictId);

        private long AllocateConflictId()
        {
            if (_nextConflictId == long.MaxValue)
                throw new InvalidOperationException("The conflict identifier range is exhausted.");
            return _nextConflictId++;
        }
    }
}

using System.Runtime.CompilerServices;
using ExcelMerge.Domain;

namespace ExcelMerge.Application;

internal static class DelimitedMergeAssembler
{
    public static async IAsyncEnumerable<RowRecord> AssembleAsync(
        MergePlan plan,
        MergeSheetResult sheetResult,
        [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(sheetResult);
        var sheetPlan = plan.Worksheets.ToArray().Single(sheet => string.Equals(
            sheet.Id,
            sheetResult.Id,
            StringComparison.Ordinal));
        var sheetChoice = GetSheetChoice(plan, sheetPlan);
        if (sheetChoice == RowChoice.None)
        {
            yield break;
        }

        var resultRowIndex = 0;
        foreach (var mapping in sheetPlan.Rows.ToArray())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (sheetChoice is RowChoice.Local or RowChoice.Remote)
            {
                var sourceIndex = sheetChoice == RowChoice.Local
                    ? mapping.LocalRowIndex
                    : mapping.RemoteRowIndex;
                var snapshot = sheetChoice == RowChoice.Local
                    ? sheetResult.LocalWorksheet
                    : sheetResult.RemoteWorksheet;
                if (sourceIndex.HasValue && snapshot is not null)
                {
                    var sourceRow = await ReadRequiredRowAsync(
                        snapshot,
                        sourceIndex.Value,
                        cancellationToken).ConfigureAwait(false);
                    yield return Reindex(sourceRow, resultRowIndex++);
                }

                continue;
            }

            var selection = GetRowChoice(plan, sheetPlan, mapping);
            switch (selection.Kind)
            {
                case RowChoice.None:
                    break;
                case RowChoice.Local:
                case RowChoice.Remote:
                    {
                        var snapshot = selection.Kind == RowChoice.Local
                            ? sheetResult.LocalWorksheet
                            : sheetResult.RemoteWorksheet;
                        var sourceIndex = selection.Kind == RowChoice.Local
                            ? mapping.LocalRowIndex
                            : mapping.RemoteRowIndex;
                        if (snapshot is not null && sourceIndex.HasValue)
                        {
                            var sourceRow = await ReadRequiredRowAsync(
                                snapshot,
                                sourceIndex.Value,
                                cancellationToken).ConfigureAwait(false);
                            var resultRow = selection.IsCompleteRowChoice
                                ? Reindex(sourceRow, resultRowIndex)
                                : ApplyCellChoices(
                                    plan,
                                    sheetPlan,
                                    mapping,
                                    sourceRow,
                                    resultRowIndex);
                            yield return resultRow;
                            resultRowIndex++;
                        }

                        break;
                    }
                case RowChoice.Both:
                    if (mapping.LocalRowIndex.HasValue && sheetResult.LocalWorksheet is not null)
                    {
                        yield return Reindex(
                            await ReadRequiredRowAsync(
                                sheetResult.LocalWorksheet,
                                mapping.LocalRowIndex.Value,
                                cancellationToken).ConfigureAwait(false),
                            resultRowIndex++);
                    }

                    if (mapping.RemoteRowIndex.HasValue && sheetResult.RemoteWorksheet is not null)
                    {
                        yield return Reindex(
                            await ReadRequiredRowAsync(
                                sheetResult.RemoteWorksheet,
                                mapping.RemoteRowIndex.Value,
                                cancellationToken).ConfigureAwait(false),
                            resultRowIndex++);
                    }

                    break;
                case RowChoice.Custom:
                    yield return Reindex(selection.CustomRow!.Value, resultRowIndex++);
                    break;
                default:
                    throw new ExcelMergeApplicationException(
                        ApplicationError.UnsupportedOutput,
                        "The delimited result contains an unsupported row choice.");
            }
        }
    }

    private static RowSelection GetRowChoice(
        MergePlan plan,
        SheetMergePlan sheet,
        in MergeRowMapping mapping)
    {
        var location = new ConflictLocation(
            sheet.Id,
            mapping.BaseRowIndex,
            mapping.LocalRowIndex,
            mapping.RemoteRowIndex);
        foreach (var rowOverride in plan.RowOverrides.Span)
        {
            if (rowOverride.Covers(location))
            {
                return FromResolution(rowOverride.Kind, rowOverride.CustomRow, complete: true);
            }
        }

        foreach (var conflict in plan.Conflicts.Span)
        {
            if (!SameRow(conflict.Location, location) ||
                conflict.Kind is ConflictKind.CellValue or ConflictKind.CellDeleteEdit or
                    ConflictKind.RowMetadata or ConflictKind.SheetAddAdd or
                    ConflictKind.SheetDeleteEdit or ConflictKind.AmbiguousSheetRename)
            {
                continue;
            }

            var resolution = plan.RowResolutions.ToArray().Single(item => item.ConflictId == conflict.Id);
            return FromResolution(resolution.Kind, resolution.CustomRow, complete: true);
        }

        foreach (var decision in sheet.AutomaticDecisions.Span)
        {
            if (decision.Scope != MergeDecisionScope.Row || !SameRow(decision.Location, location))
            {
                continue;
            }

            return decision.Kind switch
            {
                AutomaticMergeKind.UseLocal or AutomaticMergeKind.UseEither =>
                    new RowSelection(RowChoice.Local, IsCompleteRowChoice: true),
                AutomaticMergeKind.UseRemote =>
                    new RowSelection(RowChoice.Remote, IsCompleteRowChoice: true),
                AutomaticMergeKind.Delete =>
                    new RowSelection(RowChoice.None, IsCompleteRowChoice: true),
                _ => throw new ExcelMergeApplicationException(
                    ApplicationError.UnsupportedOutput,
                    "The delimited result contains an undefined automatic row choice."),
            };
        }

        return mapping.LocalRowIndex.HasValue
            ? new RowSelection(RowChoice.Local)
            : mapping.RemoteRowIndex.HasValue
                ? new RowSelection(RowChoice.Remote)
                : new RowSelection(RowChoice.None);
    }

    private static RowChoice GetSheetChoice(MergePlan plan, SheetMergePlan sheet)
    {
        var choice = RowChoice.Merged;
        foreach (var decision in sheet.AutomaticDecisions.Span)
        {
            if (decision.Scope != MergeDecisionScope.Sheet)
            {
                continue;
            }

            choice = decision.Kind switch
            {
                AutomaticMergeKind.UseLocal => sheet.LocalSheet is null ? RowChoice.None : RowChoice.Local,
                AutomaticMergeKind.UseRemote => sheet.RemoteSheet is null ? RowChoice.None : RowChoice.Remote,
                AutomaticMergeKind.UseEither => sheet.LocalSheet is not null
                    ? RowChoice.Local
                    : sheet.RemoteSheet is not null ? RowChoice.Remote : RowChoice.None,
                AutomaticMergeKind.Delete => RowChoice.None,
                _ => choice,
            };
        }

        foreach (var conflict in plan.Conflicts.Span)
        {
            if (!string.Equals(conflict.Location.SheetId, sheet.Id, StringComparison.Ordinal) ||
                conflict.Kind is not (ConflictKind.SheetAddAdd or ConflictKind.SheetDeleteEdit))
            {
                continue;
            }

            var resolution = plan.RowResolutions.ToArray().Single(item => item.ConflictId == conflict.Id);
            choice = resolution.Kind switch
            {
                ResolutionKind.Local => sheet.LocalSheet is null ? RowChoice.None : RowChoice.Local,
                ResolutionKind.Remote => sheet.RemoteSheet is null ? RowChoice.None : RowChoice.Remote,
                _ => throw new ExcelMergeApplicationException(
                    ApplicationError.UnsupportedOutput,
                    "Delimited sheet conflicts support only LOCAL or REMOTE output."),
            };
        }

        return choice;
    }

    private static RowRecord ApplyCellChoices(
        MergePlan plan,
        SheetMergePlan sheet,
        in MergeRowMapping mapping,
        in RowRecord sourceRow,
        int resultRowIndex)
    {
        var cells = sourceRow.Cells.ToArray().ToDictionary(
            static cell => cell.Address.ColumnIndex,
            static cell => cell);
        foreach (var decision in sheet.AutomaticDecisions.Span)
        {
            if (decision.Scope != MergeDecisionScope.Cell ||
                !MatchesMapping(decision.Location, sheet.Id, mapping) ||
                decision.Location.ColumnIndex is not { } columnIndex ||
                decision.Kind is AutomaticMergeKind.UseLocal or AutomaticMergeKind.UseEither)
            {
                continue;
            }

            SetCell(cells, columnIndex, decision.ResultValue);
        }

        foreach (var conflict in plan.Conflicts.Span)
        {
            if (conflict.Kind is not (ConflictKind.CellValue or ConflictKind.CellDeleteEdit) ||
                !MatchesMapping(conflict.Location, sheet.Id, mapping) ||
                conflict.Location.ColumnIndex is not { } columnIndex)
            {
                continue;
            }

            var resolution = plan.CellResolutions.ToArray().Single(item => item.ConflictId == conflict.Id);
            var value = resolution.Kind switch
            {
                ResolutionKind.Local => conflict.LocalValue,
                ResolutionKind.Remote => conflict.RemoteValue,
                ResolutionKind.Custom => resolution.CustomValue,
                _ => throw new ExcelMergeApplicationException(
                    ApplicationError.UnresolvedConflicts,
                    $"Cell conflict {conflict.Id} is unresolved."),
            };
            SetCell(cells, columnIndex, value);
        }

        return new RowRecord(
            resultRowIndex,
            cells.OrderBy(static pair => pair.Key)
                .Select(pair => new CellRecord(
                    new CellAddress(resultRowIndex, pair.Key),
                    pair.Value.Value,
                    pair.Value.StyleIndex))
                .ToArray(),
            sourceRow.Height,
            sourceRow.IsHidden,
            sourceRow.StyleIndex);

        static void SetCell(
            Dictionary<int, CellRecord> cells,
            int columnIndex,
            CellValue? value)
        {
            if (!value.HasValue || value.Value.Kind == CellKind.Blank)
            {
                cells.Remove(columnIndex);
            }
            else
            {
                cells.TryGetValue(columnIndex, out var existing);
                cells[columnIndex] = new CellRecord(
                    new CellAddress(0, columnIndex),
                    value.Value,
                    existing.StyleIndex);
            }
        }
    }

    private static RowSelection FromResolution(
        ResolutionKind kind,
        RowRecord? customRow,
        bool complete) => kind switch
        {
            ResolutionKind.Local => new RowSelection(RowChoice.Local, complete),
            ResolutionKind.Remote => new RowSelection(RowChoice.Remote, complete),
            ResolutionKind.Both => new RowSelection(RowChoice.Both, complete),
            ResolutionKind.Custom => new RowSelection(RowChoice.Custom, complete, customRow),
            _ => throw new ExcelMergeApplicationException(
                ApplicationError.UnresolvedConflicts,
                "A structural conflict is unresolved."),
        };

    private static async ValueTask<RowRecord> ReadRequiredRowAsync(
        IWorksheetSnapshot worksheet,
        int rowIndex,
        CancellationToken cancellationToken) =>
        await worksheet.GetRowAsync(rowIndex, cancellationToken).ConfigureAwait(false) ??
        throw new ExcelMergeApplicationException(
            ApplicationError.UnsupportedOutput,
            $"Result row {rowIndex} is missing from worksheet '{worksheet.Metadata.Name}'.");

    private static RowRecord Reindex(in RowRecord row, int resultRowIndex) =>
        new(
            resultRowIndex,
            row.Cells.ToArray().Select(cell => new CellRecord(
                new CellAddress(resultRowIndex, cell.Address.ColumnIndex),
                cell.Value,
                cell.StyleIndex)).ToArray(),
            row.Height,
            row.IsHidden,
            row.StyleIndex);

    private static bool MatchesMapping(
        ConflictLocation location,
        string sheetId,
        in MergeRowMapping mapping) =>
        string.Equals(location.SheetId, sheetId, StringComparison.Ordinal) &&
        location.BaseRowIndex == mapping.BaseRowIndex &&
        location.LocalRowIndex == mapping.LocalRowIndex &&
        location.RemoteRowIndex == mapping.RemoteRowIndex;

    private static bool SameRow(ConflictLocation left, ConflictLocation right) =>
        string.Equals(left.SheetId, right.SheetId, StringComparison.Ordinal) &&
        left.BaseRowIndex == right.BaseRowIndex &&
        left.LocalRowIndex == right.LocalRowIndex &&
        left.RemoteRowIndex == right.RemoteRowIndex;

    private enum RowChoice
    {
        Merged,
        None,
        Local,
        Remote,
        Both,
        Custom,
    }

    private readonly record struct RowSelection(
        RowChoice Kind,
        bool IsCompleteRowChoice = false,
        RowRecord? CustomRow = null);
}

using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

/// <summary>Combines selected-sheet results into the immutable plan consumed by a merge writer.</summary>
public static class MergePlanFactory
{
    public static MergePlan Create(
        WorkbookMetadata baseWorkbook,
        WorkbookMetadata localWorkbook,
        WorkbookMetadata remoteWorkbook,
        IEnumerable<ThreeWayMergeResult> sheetResults,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(baseWorkbook);
        ArgumentNullException.ThrowIfNull(localWorkbook);
        ArgumentNullException.ThrowIfNull(remoteWorkbook);
        ArgumentNullException.ThrowIfNull(sheetResults);

        var changes = new List<SheetChange>();
        var conflicts = new List<ConflictRecord>();
        var cellResolutions = new List<CellResolution>();
        var rowResolutions = new List<RowResolution>();
        var conflictIds = new HashSet<long>();

        foreach (var result in sheetResults)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(result);
            changes.Add(result.LocalChange);
            changes.Add(result.RemoteChange);

            foreach (var conflict in result.Conflicts.Span)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (conflict.Id < 0 || !conflictIds.Add(conflict.Id))
                {
                    throw new ArgumentException(
                        $"Conflict identifier {conflict.Id} is invalid or duplicated.",
                        nameof(sheetResults));
                }

                conflicts.Add(conflict);
            }

            cellResolutions.AddRange(result.CellResolutions.Span);
            rowResolutions.AddRange(result.RowResolutions.Span);
        }

        ValidateResolutions(
            conflicts,
            cellResolutions,
            rowResolutions,
            nameof(sheetResults),
            cancellationToken);
        return new MergePlan(
            baseWorkbook,
            localWorkbook,
            remoteWorkbook,
            changes.ToArray(),
            conflicts.ToArray(),
            cellResolutions.ToArray(),
            rowResolutions.ToArray());
    }

    private static void ValidateResolutions(
        List<ConflictRecord> conflicts,
        List<CellResolution> cellResolutions,
        List<RowResolution> rowResolutions,
        string parameterName,
        CancellationToken cancellationToken)
    {
        var expectedCellIds = new HashSet<long>();
        var expectedRowIds = new HashSet<long>();
        foreach (var conflict in conflicts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (conflict.Kind is ConflictKind.CellValue or ConflictKind.CellDeleteEdit)
                expectedCellIds.Add(conflict.Id);
            else
                expectedRowIds.Add(conflict.Id);
        }

        var actualCellIds = new HashSet<long>();
        foreach (var resolution in cellResolutions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!actualCellIds.Add(resolution.ConflictId))
                ThrowResolutionError(resolution.ConflictId, parameterName);
        }

        var actualRowIds = new HashSet<long>();
        foreach (var resolution in rowResolutions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!actualRowIds.Add(resolution.ConflictId))
                ThrowResolutionError(resolution.ConflictId, parameterName);
        }

        if (!expectedCellIds.SetEquals(actualCellIds) || !expectedRowIds.SetEquals(actualRowIds))
        {
            throw new ArgumentException(
                "Every conflict must have exactly one resolution of the matching scope.",
                parameterName);
        }
    }

    private static void ThrowResolutionError(
        long conflictId,
        string parameterName)
    {
        throw new ArgumentException(
            $"Conflict identifier {conflictId} has more than one resolution.",
            parameterName);
    }
}

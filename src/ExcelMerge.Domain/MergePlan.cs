namespace ExcelMerge.Domain;

/// <summary>
/// An immutable, side-explicit plan consumed by a merge writer. Collections are expected to remain
/// valid for the lifetime of the plan.
/// </summary>
public sealed record MergePlan(
    WorkbookMetadata BaseWorkbook,
    WorkbookMetadata LocalWorkbook,
    WorkbookMetadata RemoteWorkbook,
    ReadOnlyMemory<SheetMergePlan> Worksheets,
    ReadOnlyMemory<ConflictRecord> Conflicts,
    ReadOnlyMemory<CellResolution> CellResolutions,
    ReadOnlyMemory<RowResolution> RowResolutions)
{
    public ReadOnlyMemory<RowMergeOverride> RowOverrides { get; init; } =
        ReadOnlyMemory<RowMergeOverride>.Empty;

    public ReadOnlyMemory<SheetChange> SheetChanges
    {
        get
        {
            var changes = new SheetChange[checked(Worksheets.Length * 2)];
            for (var index = 0; index < Worksheets.Length; index++)
            {
                changes[index * 2] = Worksheets.Span[index].LocalChange;
                changes[(index * 2) + 1] = Worksheets.Span[index].RemoteChange;
            }

            return changes;
        }
    }

    public bool HasUnresolvedResolutions
    {
        get
        {
            foreach (var resolution in CellResolutions.Span)
            {
                if (!resolution.IsResolved && !IsCoveredByRowOverride(resolution.ConflictId))
                    return true;
            }

            foreach (var resolution in RowResolutions.Span)
            {
                if (!resolution.IsResolved && !IsCoveredByRowOverride(resolution.ConflictId))
                    return true;
            }

            return false;
        }
    }

    private bool IsCoveredByRowOverride(long conflictId)
    {
        ConflictLocation? location = null;
        foreach (var conflict in Conflicts.Span)
        {
            if (conflict.Id == conflictId)
            {
                location = conflict.Location;
                break;
            }
        }

        if (!location.HasValue)
        {
            return false;
        }

        foreach (var rowOverride in RowOverrides.Span)
        {
            if (rowOverride.Covers(location.Value))
            {
                return true;
            }
        }

        return false;
    }
}

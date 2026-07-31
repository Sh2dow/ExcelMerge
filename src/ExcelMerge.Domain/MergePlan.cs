namespace ExcelMerge.Domain;

/// <summary>
/// An immutable, side-explicit plan consumed by a merge writer. Collections are expected to remain
/// valid for the lifetime of the plan.
/// </summary>
public sealed record MergePlan(
    WorkbookMetadata BaseWorkbook,
    WorkbookMetadata LocalWorkbook,
    WorkbookMetadata RemoteWorkbook,
    ReadOnlyMemory<SheetChange> SheetChanges,
    ReadOnlyMemory<ConflictRecord> Conflicts,
    ReadOnlyMemory<CellResolution> CellResolutions,
    ReadOnlyMemory<RowResolution> RowResolutions)
{
    public bool HasUnresolvedResolutions
    {
        get
        {
            foreach (var resolution in CellResolutions.Span)
            {
                if (!resolution.IsResolved)
                    return true;
            }

            foreach (var resolution in RowResolutions.Span)
            {
                if (!resolution.IsResolved)
                    return true;
            }

            return false;
        }
    }
}

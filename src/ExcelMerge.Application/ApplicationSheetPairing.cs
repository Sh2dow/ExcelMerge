using ExcelMerge.Domain;

namespace ExcelMerge.Application;

internal readonly record struct CompareWorksheetPair(
    IWorksheetSnapshot? Local,
    IWorksheetSnapshot? Remote);

internal readonly record struct MergeWorksheetGroup(
    IWorksheetSnapshot? Base,
    IWorksheetSnapshot? Local,
    IWorksheetSnapshot? Remote);

internal static class ApplicationSheetPairing
{
    public static IReadOnlyList<CompareWorksheetPair> PairCompare(
        IndexedApplicationWorkbook local,
        IndexedApplicationWorkbook remote,
        WorksheetSelection? localSelection,
        WorksheetSelection? remoteSelection)
    {
        if (localSelection is not null || remoteSelection is not null)
        {
            var selectedLocal = ResolveSelection(
                local.Worksheets,
                localSelection,
                WorkbookSide.Local);
            var selectedRemote = ResolveSelection(
                remote.Worksheets,
                remoteSelection,
                WorkbookSide.Remote);
            return [new CompareWorksheetPair(selectedLocal, selectedRemote)];
        }

        var unmatchedRemote = new HashSet<IWorksheetSnapshot>(remote.Worksheets);
        var result = new List<CompareWorksheetPair>();
        foreach (var localSheet in local.Worksheets)
        {
            var remoteSheet = remote.Worksheets.FirstOrDefault(candidate =>
                unmatchedRemote.Contains(candidate) &&
                string.Equals(
                    candidate.Metadata.Name,
                    localSheet.Metadata.Name,
                    StringComparison.OrdinalIgnoreCase));
            if (remoteSheet is not null)
            {
                unmatchedRemote.Remove(remoteSheet);
            }

            result.Add(new CompareWorksheetPair(localSheet, remoteSheet));
        }

        result.AddRange(remote.Worksheets
            .Where(unmatchedRemote.Contains)
            .Select(static sheet => new CompareWorksheetPair(null, sheet)));
        return result;
    }

    public static IReadOnlyList<MergeWorksheetGroup> PairMerge(
        IndexedApplicationWorkbook @base,
        IndexedApplicationWorkbook local,
        IndexedApplicationWorkbook remote)
    {
        var groups = @base.Worksheets
            .Select(static sheet => new MutableGroup { Base = sheet })
            .ToList();
        var unmatchedLocal = AttachByName(groups, local.Worksheets, static group => group.Base, localSide: true);
        var unmatchedRemote = AttachByName(groups, remote.Worksheets, static group => group.Base, localSide: false);

        var baseWithoutLocal = groups.Where(static group => group.Local is null).ToArray();
        AttachByOrdinal(baseWithoutLocal, unmatchedLocal, localSide: true);
        unmatchedLocal.RemoveAll(sheet => groups.Any(group => ReferenceEquals(group.Local, sheet)));

        var baseWithoutRemote = groups.Where(static group => group.Remote is null).ToArray();
        AttachByOrdinal(baseWithoutRemote, unmatchedRemote, localSide: false);
        unmatchedRemote.RemoveAll(sheet => groups.Any(group => ReferenceEquals(group.Remote, sheet)));

        foreach (var localSheet in unmatchedLocal.ToArray())
        {
            var matchingRemote = unmatchedRemote.FirstOrDefault(remoteSheet => string.Equals(
                remoteSheet.Metadata.Name,
                localSheet.Metadata.Name,
                StringComparison.OrdinalIgnoreCase));
            groups.Add(new MutableGroup
            {
                Local = localSheet,
                Remote = matchingRemote,
            });
            unmatchedLocal.Remove(localSheet);
            if (matchingRemote is not null)
            {
                unmatchedRemote.Remove(matchingRemote);
            }
        }

        groups.AddRange(unmatchedRemote.Select(static sheet => new MutableGroup { Remote = sheet }));
        return groups
            .OrderBy(static group => group.Local?.Metadata.Position ?? int.MaxValue)
            .ThenBy(static group => group.Base?.Metadata.Position ?? int.MaxValue)
            .ThenBy(static group => group.Remote?.Metadata.Position ?? int.MaxValue)
            .Select(static group => new MergeWorksheetGroup(group.Base, group.Local, group.Remote))
            .ToArray();
    }

    private static List<IWorksheetSnapshot> AttachByName(
        List<MutableGroup> groups,
        IReadOnlyList<IWorksheetSnapshot> candidates,
        Func<MutableGroup, IWorksheetSnapshot?> anchor,
        bool localSide)
    {
        var unmatched = new List<IWorksheetSnapshot>();
        foreach (var candidate in candidates)
        {
            var group = groups.FirstOrDefault(item =>
                GetSide(item, localSide) is null &&
                anchor(item) is { } anchorSheet &&
                string.Equals(
                    anchorSheet.Metadata.Name,
                    candidate.Metadata.Name,
                    StringComparison.OrdinalIgnoreCase));
            if (group is null)
            {
                unmatched.Add(candidate);
            }
            else
            {
                SetSide(group, candidate, localSide);
            }
        }

        return unmatched;
    }

    private static void AttachByOrdinal(
        IReadOnlyList<MutableGroup> groups,
        IReadOnlyList<IWorksheetSnapshot> candidates,
        bool localSide)
    {
        var count = Math.Min(groups.Count, candidates.Count);
        for (var index = 0; index < count; index++)
        {
            SetSide(groups[index], candidates[index], localSide);
        }
    }

    private static IWorksheetSnapshot? ResolveSelection(
        IReadOnlyList<IWorksheetSnapshot> worksheets,
        WorksheetSelection? selection,
        WorkbookSide side)
    {
        if (selection is null)
        {
            return worksheets.FirstOrDefault() ?? throw new ExcelMergeApplicationException(
                ApplicationError.InvalidSheetSelection,
                "The workbook does not contain a selectable worksheet.",
                side);
        }

        selection.Validate(side.ToString());
        var result = selection.Id is not null
            ? worksheets.FirstOrDefault(sheet => string.Equals(
                sheet.Metadata.Id,
                selection.Id,
                StringComparison.Ordinal))
            : worksheets.FirstOrDefault(sheet => string.Equals(
                sheet.Metadata.Name,
                selection.Name,
                StringComparison.OrdinalIgnoreCase));
        return result ?? throw new ExcelMergeApplicationException(
            ApplicationError.InvalidSheetSelection,
            $"Worksheet '{selection.Id ?? selection.Name}' was not found.",
            side);
    }

    private static IWorksheetSnapshot? GetSide(MutableGroup group, bool localSide) =>
        localSide ? group.Local : group.Remote;

    private static void SetSide(
        MutableGroup group,
        IWorksheetSnapshot worksheet,
        bool localSide)
    {
        if (localSide)
            group.Local = worksheet;
        else
            group.Remote = worksheet;
    }

    private sealed class MutableGroup
    {
        public IWorksheetSnapshot? Base { get; init; }

        public IWorksheetSnapshot? Local { get; set; }

        public IWorksheetSnapshot? Remote { get; set; }
    }
}

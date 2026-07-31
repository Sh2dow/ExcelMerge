using ExcelMerge.Domain;

namespace ExcelMerge.OpenXml;

internal sealed record CompiledOpenXmlMerge(
    IReadOnlyList<CompiledWorksheetPatch> Worksheets,
    bool RequiresRecalculation);

internal enum CompiledWorksheetAction
{
    KeepLocal,
    Delete,
    ImportRemote,
}

internal enum CompiledRowSourceKind
{
    Remote,
    Custom,
}

internal sealed record CompiledRowSource(
    CompiledRowSourceKind Kind,
    int? RemoteRowIndex = null,
    RowRecord? CustomRow = null);

internal sealed record CompiledRowTransform(
    IReadOnlySet<int> DeletedLocalRows,
    IReadOnlyDictionary<int, CompiledRowSource> Replacements,
    IReadOnlyDictionary<int, IReadOnlyList<CompiledRowSource>> Insertions,
    IReadOnlyDictionary<int, int> RemoteMetadataRows)
{
    public bool HasChanges =>
        DeletedLocalRows.Count != 0 ||
        Replacements.Count != 0 ||
        Insertions.Count != 0 ||
        RemoteMetadataRows.Count != 0;

    public bool RequiresRemoteRows =>
        RemoteMetadataRows.Count != 0 ||
        Replacements.Values.Any(static source => source.Kind == CompiledRowSourceKind.Remote) ||
        Insertions.Values.SelectMany(static sources => sources)
            .Any(static source => source.Kind == CompiledRowSourceKind.Remote);
}

internal enum CompiledCellSourceKind
{
    Remote,
    Custom,
}

internal readonly record struct CompiledCellPatch(
    CellValue? Value,
    CompiledCellSourceKind Source);

internal sealed record CompiledWorksheetPatch(
    string SheetGroupId,
    string? LocalRelationshipId,
    string? LocalName,
    string? RemoteRelationshipId,
    string? RemoteName,
    SheetVisibility RemoteVisibility,
    string? ResultName,
    SheetVisibility? ResultVisibility,
    CompiledWorksheetAction Action,
    CompiledRowTransform RowTransform,
    ReadOnlyMemory<MergeRowMapping> RowMappings,
    IReadOnlyDictionary<int, IReadOnlyDictionary<int, CompiledCellPatch>> Rows);

internal static class OpenXmlMergePlanCompiler
{
    public static CompiledOpenXmlMerge Compile(MergePlan plan, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ValidateWorkbook(plan.BaseWorkbook, nameof(plan.BaseWorkbook));
        ValidateWorkbook(plan.LocalWorkbook, nameof(plan.LocalWorkbook));
        ValidateWorkbook(plan.RemoteWorkbook, nameof(plan.RemoteWorkbook));

        if (plan.HasUnresolvedResolutions)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnresolvedConflicts,
                "The merge plan contains unresolved conflicts.");
        }

        var conflicts = BuildConflictMap(plan.Conflicts, cancellationToken);
        var cellResolutions = BuildCellResolutionMap(plan.CellResolutions, cancellationToken);
        var rowResolutions = BuildRowResolutionMap(plan.RowResolutions, cancellationToken);
        ValidateResolutionCoverage(conflicts, cellResolutions, rowResolutions);
        var allRowOverrides = ValidateRowOverrides(plan, cancellationToken);

        var worksheetIds = new HashSet<string>(StringComparer.Ordinal);
        var worksheets = new List<CompiledWorksheetPatch>();
        var requiresRecalculation = false;
        foreach (var worksheet in plan.Worksheets.Span)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!worksheetIds.Add(worksheet.Id))
            {
                throw InvalidPlan($"Worksheet group '{worksheet.Id}' is duplicated.");
            }

            var patches = new Dictionary<int, Dictionary<int, CompiledCellPatch>>();
            var rowTransform = EmptyRowTransform();
            var rowOverrides = allRowOverrides
                .Where(rowOverride => string.Equals(
                    rowOverride.SheetId,
                    worksheet.Id,
                    StringComparison.Ordinal))
                .ToArray();
            var action = DetermineSheetAction(
                worksheet,
                conflicts,
                rowResolutions,
                cancellationToken);
            var (resultName, resultVisibility) = DetermineSheetMetadata(
                worksheet,
                conflicts,
                rowResolutions,
                action,
                cancellationToken);
            if (action == CompiledWorksheetAction.KeepLocal)
            {
                CompileAutomaticDecisions(
                    worksheet,
                    rowOverrides,
                    patches,
                    cancellationToken);
                rowTransform = CompileRowTransform(
                    worksheet,
                    conflicts,
                    rowResolutions,
                    rowOverrides,
                    cancellationToken);
                CompileConflictResolutions(
                    worksheet,
                    conflicts,
                    cellResolutions,
                    rowResolutions,
                    rowOverrides,
                    patches,
                    cancellationToken);
            }

            if (action == CompiledWorksheetAction.KeepLocal &&
                patches.Count == 0 &&
                !rowTransform.HasChanges &&
                resultName is null &&
                !resultVisibility.HasValue)
            {
                continue;
            }

            if (action == CompiledWorksheetAction.KeepLocal && worksheet.LocalSheet is null)
            {
                throw Unsupported(
                    worksheet.Id,
                    "A result worksheet cannot be created before remote worksheet import is enabled.");
            }

            if (action == CompiledWorksheetAction.ImportRemote && worksheet.RemoteSheet is null)
            {
                throw InvalidPlan($"Worksheet group '{worksheet.Id}' selects a missing REMOTE sheet.");
            }

            var frozenRows = patches.ToDictionary(
                static pair => pair.Key,
                static pair => (IReadOnlyDictionary<int, CompiledCellPatch>)pair.Value);
            worksheets.Add(new CompiledWorksheetPatch(
                worksheet.Id,
                worksheet.LocalSheet?.Id,
                worksheet.LocalSheet?.Name,
                worksheet.RemoteSheet?.Id,
                worksheet.RemoteSheet?.Name,
                worksheet.RemoteSheet?.Visibility ?? SheetVisibility.Visible,
                resultName,
                resultVisibility,
                action,
                rowTransform,
                worksheet.Rows,
                frozenRows));
            requiresRecalculation |= action != CompiledWorksheetAction.KeepLocal ||
                patches.Count != 0 ||
                rowTransform.HasChanges ||
                resultName is not null;
        }

        return new CompiledOpenXmlMerge(worksheets, requiresRecalculation);
    }

    private static (string? Name, SheetVisibility? Visibility) DetermineSheetMetadata(
        SheetMergePlan worksheet,
        IReadOnlyDictionary<long, ConflictRecord> conflicts,
        IReadOnlyDictionary<long, RowResolution> rowResolutions,
        CompiledWorksheetAction action,
        CancellationToken cancellationToken)
    {
        if (action != CompiledWorksheetAction.KeepLocal || worksheet.LocalSheet is null)
        {
            return default;
        }

        var useRemote = false;
        foreach (var decision in worksheet.AutomaticDecisions.Span)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (decision.Scope == MergeDecisionScope.SheetMetadata &&
                decision.Kind == AutomaticMergeKind.UseRemote)
            {
                useRemote = true;
            }
        }

        foreach (var conflict in conflicts.Values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (conflict.Kind != ConflictKind.AmbiguousSheetRename ||
                !string.Equals(conflict.Location.SheetId, worksheet.Id, StringComparison.Ordinal))
            {
                continue;
            }

            var resolution = rowResolutions[conflict.Id];
            if (resolution.Kind == ResolutionKind.Remote)
            {
                useRemote = true;
            }
            else if (resolution.Kind != ResolutionKind.Local)
            {
                throw Unsupported(
                    worksheet.Id,
                    $"Worksheet metadata conflict {conflict.Id} requires resolution '{resolution.Kind}'.");
            }
        }

        if (!useRemote)
        {
            return default;
        }

        var remote = worksheet.RemoteSheet ?? throw InvalidPlan(
            $"Worksheet group '{worksheet.Id}' selects missing REMOTE metadata.");
        return (
            string.Equals(remote.Name, worksheet.LocalSheet.Name, StringComparison.Ordinal)
                ? null
                : remote.Name,
            remote.Visibility == worksheet.LocalSheet.Visibility
                ? null
                : remote.Visibility);
    }

    private static CompiledRowTransform CompileRowTransform(
        SheetMergePlan worksheet,
        IReadOnlyDictionary<long, ConflictRecord> conflicts,
        IReadOnlyDictionary<long, RowResolution> rowResolutions,
        IReadOnlyList<RowMergeOverride> rowOverrides,
        CancellationToken cancellationToken)
    {
        var deletedRows = new HashSet<int>();
        var replacements = new Dictionary<int, CompiledRowSource>();
        var insertions = new Dictionary<int, List<CompiledRowSource>>();
        var metadataRows = new Dictionary<int, int>();

        foreach (var decision in worksheet.AutomaticDecisions.Span)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsOverridden(decision.Location, rowOverrides))
            {
                continue;
            }

            if (decision.Scope == MergeDecisionScope.RowMetadata)
            {
                if (decision.Kind == AutomaticMergeKind.UseRemote)
                {
                    AddRemoteMetadataSelection(worksheet, decision.Location, metadataRows);
                }
                else if (decision.Kind is not (AutomaticMergeKind.UseLocal or AutomaticMergeKind.UseEither))
                {
                    throw InvalidPlan("The plan contains an invalid row metadata decision.");
                }

                continue;
            }

            if (decision.Scope != MergeDecisionScope.Row)
            {
                continue;
            }

            switch (decision.Kind)
            {
                case AutomaticMergeKind.UseLocal:
                case AutomaticMergeKind.UseEither:
                    break;
                case AutomaticMergeKind.Delete:
                    if (decision.Location.LocalRowIndex is { } deletedRow)
                    {
                        AddDeletion(deletedRows, replacements, deletedRow, worksheet.Id);
                    }

                    break;
                case AutomaticMergeKind.UseRemote:
                    AddRemoteSelection(
                        worksheet,
                        decision.Location,
                        deletedRows,
                        replacements,
                        insertions);
                    break;
                default:
                    throw InvalidPlan("The plan contains an undefined row decision.");
            }
        }

        foreach (var conflict in conflicts.Values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(conflict.Location.SheetId, worksheet.Id, StringComparison.Ordinal) ||
                conflict.Kind is ConflictKind.CellValue or ConflictKind.CellDeleteEdit or
                    ConflictKind.SheetAddAdd or ConflictKind.SheetDeleteEdit or ConflictKind.AmbiguousSheetRename)
            {
                continue;
            }

            if (IsOverridden(conflict.Location, rowOverrides))
            {
                continue;
            }

            var resolution = rowResolutions[conflict.Id];
            if (conflict.Kind == ConflictKind.RowMetadata)
            {
                if (resolution.Kind == ResolutionKind.Remote)
                {
                    AddRemoteMetadataSelection(worksheet, conflict.Location, metadataRows);
                }
                else if (resolution.Kind != ResolutionKind.Local)
                {
                    throw Unsupported(
                        worksheet.Id,
                        $"Row metadata conflict {conflict.Id} requires resolution '{resolution.Kind}'.");
                }

                continue;
            }

            switch (resolution.Kind)
            {
                case ResolutionKind.Local:
                    break;
                case ResolutionKind.Remote:
                    AddRemoteSelection(
                        worksheet,
                        conflict.Location,
                        deletedRows,
                        replacements,
                        insertions);
                    break;
                case ResolutionKind.Both:
                    AddBothSelection(worksheet, conflict.Location, insertions);
                    break;
                case ResolutionKind.Custom:
                    AddCustomSelection(
                        worksheet,
                        conflict.Location,
                        resolution.CustomRow ?? throw InvalidPlan(
                            $"Custom row resolution {conflict.Id} has no row."),
                        replacements,
                        insertions);
                    break;
                default:
                    throw InvalidPlan($"Structural conflict {conflict.Id} is unresolved.");
            }
        }

        foreach (var rowOverride in rowOverrides)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var location = new ConflictLocation(
                rowOverride.SheetId,
                rowOverride.BaseRowIndex,
                rowOverride.LocalRowIndex,
                rowOverride.RemoteRowIndex);
            switch (rowOverride.Kind)
            {
                case ResolutionKind.Local:
                    break;
                case ResolutionKind.Remote:
                    AddRemoteSelection(
                        worksheet,
                        location,
                        deletedRows,
                        replacements,
                        insertions);
                    break;
                case ResolutionKind.Both:
                    AddBothSelection(worksheet, location, insertions);
                    break;
                case ResolutionKind.Custom:
                    AddCustomSelection(
                        worksheet,
                        location,
                        rowOverride.CustomRow ?? throw InvalidPlan(
                            "A custom row override has no row."),
                        replacements,
                        insertions);
                    break;
                default:
                    throw InvalidPlan("A row override contains an invalid resolution.");
            }
        }

        return new CompiledRowTransform(
            deletedRows,
            replacements,
            insertions.ToDictionary(
                static pair => pair.Key,
                static pair => (IReadOnlyList<CompiledRowSource>)pair.Value),
            metadataRows);
    }

    private static void AddRemoteMetadataSelection(
        SheetMergePlan worksheet,
        ConflictLocation location,
        Dictionary<int, int> metadataRows)
    {
        if (location.LocalRowIndex is not { } localRowIndex ||
            location.RemoteRowIndex is not { } remoteRowIndex ||
            !metadataRows.TryAdd(localRowIndex, remoteRowIndex))
        {
            throw InvalidPlan(
                $"Worksheet group '{worksheet.Id}' has an invalid REMOTE row metadata instruction.");
        }
    }

    private static void AddRemoteSelection(
        SheetMergePlan worksheet,
        ConflictLocation location,
        HashSet<int> deletedRows,
        Dictionary<int, CompiledRowSource> replacements,
        Dictionary<int, List<CompiledRowSource>> insertions)
    {
        if (location.RemoteRowIndex is not { } remoteRowIndex)
        {
            if (location.LocalRowIndex is { } localRowToDelete)
            {
                AddDeletion(deletedRows, replacements, localRowToDelete, worksheet.Id);
                return;
            }

            throw InvalidPlan("A REMOTE row selection has no row on either side.");
        }

        var source = new CompiledRowSource(
            CompiledRowSourceKind.Remote,
            RemoteRowIndex: remoteRowIndex);
        if (location.LocalRowIndex is { } localRowIndex)
        {
            AddReplacement(deletedRows, replacements, localRowIndex, source, worksheet.Id);
        }
        else
        {
            AddInsertion(
                insertions,
                FindInsertionBoundary(worksheet, location),
                source);
        }
    }

    private static void AddBothSelection(
        SheetMergePlan worksheet,
        ConflictLocation location,
        Dictionary<int, List<CompiledRowSource>> insertions)
    {
        if (location.RemoteRowIndex is not { } remoteRowIndex)
        {
            return;
        }

        var boundary = location.LocalRowIndex is { } localRowIndex
            ? checked(localRowIndex + 1)
            : FindInsertionBoundary(worksheet, location);
        AddInsertion(
            insertions,
            boundary,
            new CompiledRowSource(
                CompiledRowSourceKind.Remote,
                RemoteRowIndex: remoteRowIndex));
    }

    private static void AddCustomSelection(
        SheetMergePlan worksheet,
        ConflictLocation location,
        RowRecord customRow,
        Dictionary<int, CompiledRowSource> replacements,
        Dictionary<int, List<CompiledRowSource>> insertions)
    {
        var source = new CompiledRowSource(
            CompiledRowSourceKind.Custom,
            CustomRow: customRow);
        if (location.LocalRowIndex is { } localRowIndex)
        {
            if (!replacements.TryAdd(localRowIndex, source))
            {
                throw InvalidPlan(
                    $"Worksheet group '{worksheet.Id}' has duplicate replacement for LOCAL row {localRowIndex}.");
            }
        }
        else
        {
            AddInsertion(insertions, FindInsertionBoundary(worksheet, location), source);
        }
    }

    private static int FindInsertionBoundary(
        SheetMergePlan worksheet,
        ConflictLocation location)
    {
        var mappings = worksheet.Rows.Span;
        var mappingIndex = -1;
        for (var index = 0; index < mappings.Length; index++)
        {
            var mapping = mappings[index];
            if (mapping.BaseRowIndex == location.BaseRowIndex &&
                mapping.LocalRowIndex == location.LocalRowIndex &&
                mapping.RemoteRowIndex == location.RemoteRowIndex)
            {
                mappingIndex = index;
                break;
            }
        }

        if (mappingIndex < 0)
        {
            throw InvalidPlan($"Worksheet group '{worksheet.Id}' has an incomplete row mapping.");
        }

        for (var index = mappingIndex + 1; index < mappings.Length; index++)
        {
            if (mappings[index].LocalRowIndex is { } nextLocalRow)
            {
                return nextLocalRow;
            }
        }

        for (var index = mappingIndex - 1; index >= 0; index--)
        {
            if (mappings[index].LocalRowIndex is { } previousLocalRow)
            {
                return checked(previousLocalRow + 1);
            }
        }

        return 0;
    }

    private static void AddDeletion(
        HashSet<int> deletedRows,
        IReadOnlyDictionary<int, CompiledRowSource> replacements,
        int rowIndex,
        string sheetId)
    {
        if (replacements.ContainsKey(rowIndex) || !deletedRows.Add(rowIndex))
        {
            throw InvalidPlan(
                $"Worksheet group '{sheetId}' has conflicting instructions for LOCAL row {rowIndex}.");
        }
    }

    private static void AddReplacement(
        IReadOnlySet<int> deletedRows,
        Dictionary<int, CompiledRowSource> replacements,
        int rowIndex,
        CompiledRowSource source,
        string sheetId)
    {
        if (deletedRows.Contains(rowIndex) || !replacements.TryAdd(rowIndex, source))
        {
            throw InvalidPlan(
                $"Worksheet group '{sheetId}' has conflicting instructions for LOCAL row {rowIndex}.");
        }
    }

    private static void AddInsertion(
        Dictionary<int, List<CompiledRowSource>> insertions,
        int boundary,
        CompiledRowSource source)
    {
        if (!insertions.TryGetValue(boundary, out var rows))
        {
            rows = [];
            insertions.Add(boundary, rows);
        }

        rows.Add(source);
    }

    private static CompiledRowTransform EmptyRowTransform() =>
        new(
            new HashSet<int>(),
            new Dictionary<int, CompiledRowSource>(),
            new Dictionary<int, IReadOnlyList<CompiledRowSource>>(),
            new Dictionary<int, int>());

    private static CompiledWorksheetAction DetermineSheetAction(
        SheetMergePlan worksheet,
        IReadOnlyDictionary<long, ConflictRecord> conflicts,
        IReadOnlyDictionary<long, RowResolution> rowResolutions,
        CancellationToken cancellationToken)
    {
        var action = worksheet.LocalSheet is null
            ? CompiledWorksheetAction.Delete
            : CompiledWorksheetAction.KeepLocal;

        foreach (var decision in worksheet.AutomaticDecisions.Span)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (decision.Scope != MergeDecisionScope.Sheet)
            {
                continue;
            }

            action = decision.Kind switch
            {
                AutomaticMergeKind.UseLocal => worksheet.LocalSheet is null
                    ? CompiledWorksheetAction.Delete
                    : CompiledWorksheetAction.KeepLocal,
                AutomaticMergeKind.UseRemote => worksheet.RemoteSheet is null
                    ? CompiledWorksheetAction.Delete
                    : CompiledWorksheetAction.ImportRemote,
                AutomaticMergeKind.UseEither => worksheet.LocalSheet is not null
                    ? CompiledWorksheetAction.KeepLocal
                    : worksheet.RemoteSheet is not null
                        ? CompiledWorksheetAction.ImportRemote
                        : CompiledWorksheetAction.Delete,
                AutomaticMergeKind.Delete => CompiledWorksheetAction.Delete,
                _ => throw InvalidPlan("The plan contains an undefined sheet decision."),
            };
        }

        foreach (var conflict in conflicts.Values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(conflict.Location.SheetId, worksheet.Id, StringComparison.Ordinal) ||
                conflict.Kind is not (ConflictKind.SheetAddAdd or ConflictKind.SheetDeleteEdit))
            {
                continue;
            }

            if (!rowResolutions.TryGetValue(conflict.Id, out var resolution))
            {
                throw InvalidPlan($"Sheet conflict {conflict.Id} has no row resolution.");
            }

            action = resolution.Kind switch
            {
                ResolutionKind.Local => worksheet.LocalSheet is null
                    ? CompiledWorksheetAction.Delete
                    : CompiledWorksheetAction.KeepLocal,
                ResolutionKind.Remote => worksheet.RemoteSheet is null
                    ? CompiledWorksheetAction.Delete
                    : CompiledWorksheetAction.ImportRemote,
                _ => throw Unsupported(
                    worksheet.Id,
                    $"Sheet conflict {conflict.Id} requires resolution '{resolution.Kind}'."),
            };
        }

        return action;
    }

    private static void CompileAutomaticDecisions(
        SheetMergePlan worksheet,
        IReadOnlyList<RowMergeOverride> rowOverrides,
        Dictionary<int, Dictionary<int, CompiledCellPatch>> patches,
        CancellationToken cancellationToken)
    {
        foreach (var decision in worksheet.AutomaticDecisions.Span)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ValidateDecision(decision, worksheet.Id);
            if (IsOverridden(decision.Location, rowOverrides))
            {
                continue;
            }

            switch (decision.Scope)
            {
                case MergeDecisionScope.Cell:
                    if (decision.Kind is AutomaticMergeKind.UseLocal or AutomaticMergeKind.UseEither)
                    {
                        continue;
                    }

                    AddCellPatch(
                        patches,
                        worksheet,
                        decision.Location,
                        decision.Kind == AutomaticMergeKind.Delete ? null : decision.ResultValue,
                        CompiledCellSourceKind.Remote,
                        "automatic cell decision");
                    break;

                case MergeDecisionScope.Row:
                    continue;

                case MergeDecisionScope.RowMetadata:
                    continue;

                case MergeDecisionScope.Sheet:
                    continue;

                case MergeDecisionScope.SheetMetadata:
                    continue;

                default:
                    throw InvalidPlan("The plan contains an undefined automatic decision scope.");
            }
        }
    }

    private static void CompileConflictResolutions(
        SheetMergePlan worksheet,
        IReadOnlyDictionary<long, ConflictRecord> conflicts,
        IReadOnlyDictionary<long, CellResolution> cellResolutions,
        IReadOnlyDictionary<long, RowResolution> rowResolutions,
        IReadOnlyList<RowMergeOverride> rowOverrides,
        Dictionary<int, Dictionary<int, CompiledCellPatch>> patches,
        CancellationToken cancellationToken)
    {
        foreach (var conflict in conflicts.Values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.Equals(conflict.Location.SheetId, worksheet.Id, StringComparison.Ordinal))
            {
                continue;
            }
            if (IsOverridden(conflict.Location, rowOverrides))
            {
                continue;
            }

            if (conflict.Kind is ConflictKind.CellValue or ConflictKind.CellDeleteEdit)
            {
                var resolution = cellResolutions[conflict.Id];
                switch (resolution.Kind)
                {
                    case ResolutionKind.Local:
                        break;
                    case ResolutionKind.Remote:
                        AddCellPatch(
                            patches,
                            worksheet,
                            conflict.Location,
                            conflict.RemoteValue,
                            CompiledCellSourceKind.Remote,
                            "REMOTE cell resolution");
                        break;
                    case ResolutionKind.Custom:
                        AddCellPatch(
                            patches,
                            worksheet,
                            conflict.Location,
                            resolution.CustomValue,
                            CompiledCellSourceKind.Custom,
                            "custom cell resolution");
                        break;
                    default:
                        throw InvalidPlan(
                            $"Cell conflict {conflict.Id} has invalid resolution '{resolution.Kind}'.");
                }

                continue;
            }

            var rowResolution = rowResolutions[conflict.Id];
            if (conflict.Kind is ConflictKind.SheetAddAdd or
                ConflictKind.SheetDeleteEdit or
                ConflictKind.AmbiguousSheetRename ||
                rowResolution.Kind == ResolutionKind.Local)
            {
                continue;
            }

            // Row-level structural resolutions are compiled separately into the row transform.
        }
    }

    private static void AddCellPatch(
        Dictionary<int, Dictionary<int, CompiledCellPatch>> patches,
        SheetMergePlan worksheet,
        ConflictLocation location,
        CellValue? value,
        CompiledCellSourceKind source,
        string description)
    {
        if (!string.Equals(location.SheetId, worksheet.Id, StringComparison.Ordinal) ||
            location.LocalRowIndex is not { } rowIndex ||
            location.ColumnIndex is not { } columnIndex)
        {
            throw Unsupported(
                worksheet.Id,
                $"The {description} does not identify a cell in the LOCAL worksheet.");
        }

        var hasLocalRow = false;
        foreach (var row in worksheet.Rows.Span)
        {
            if (row.LocalRowIndex == rowIndex)
            {
                hasLocalRow = true;
                break;
            }
        }

        if (!hasLocalRow)
        {
            throw InvalidPlan(
                $"The {description} targets LOCAL row {rowIndex}, which is absent from the row mapping.");
        }

        if (!patches.TryGetValue(rowIndex, out var rowPatches))
        {
            rowPatches = new Dictionary<int, CompiledCellPatch>();
            patches.Add(rowIndex, rowPatches);
        }

        if (!rowPatches.TryAdd(columnIndex, new CompiledCellPatch(value, source)))
        {
            throw InvalidPlan(
                $"More than one result instruction targets row {rowIndex}, column {columnIndex}.");
        }
    }

    private static Dictionary<long, ConflictRecord> BuildConflictMap(
        ReadOnlyMemory<ConflictRecord> source,
        CancellationToken cancellationToken)
    {
        var result = new Dictionary<long, ConflictRecord>();
        foreach (var conflict in source.Span)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (conflict.Id < 0 || !result.TryAdd(conflict.Id, conflict))
            {
                throw InvalidPlan($"Conflict identifier {conflict.Id} is invalid or duplicated.");
            }
        }

        return result;
    }

    private static RowMergeOverride[] ValidateRowOverrides(
        MergePlan plan,
        CancellationToken cancellationToken)
    {
        var overrides = plan.RowOverrides.ToArray();
        var seen = new HashSet<(string SheetId, int? Base, int? Local, int? Remote)>();
        foreach (var rowOverride in overrides)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(rowOverride.SheetId) ||
                !Enum.IsDefined(rowOverride.Kind) ||
                rowOverride.Kind == ResolutionKind.Unresolved ||
                !seen.Add((
                    rowOverride.SheetId,
                    rowOverride.BaseRowIndex,
                    rowOverride.LocalRowIndex,
                    rowOverride.RemoteRowIndex)))
            {
                throw InvalidPlan("The plan contains an invalid or duplicated row override.");
            }

            SheetMergePlan? worksheet = null;
            foreach (var candidate in plan.Worksheets.Span)
            {
                if (string.Equals(candidate.Id, rowOverride.SheetId, StringComparison.Ordinal))
                {
                    worksheet = candidate;
                    break;
                }
            }

            if (worksheet is null)
            {
                throw InvalidPlan(
                    $"Row override worksheet group '{rowOverride.SheetId}' is missing.");
            }

            var hasMapping = false;
            foreach (var mapping in worksheet.Rows.Span)
            {
                if (mapping.BaseRowIndex == rowOverride.BaseRowIndex &&
                    mapping.LocalRowIndex == rowOverride.LocalRowIndex &&
                    mapping.RemoteRowIndex == rowOverride.RemoteRowIndex)
                {
                    hasMapping = true;
                    break;
                }
            }

            if (!hasMapping)
            {
                throw InvalidPlan(
                    $"Row override in worksheet group '{rowOverride.SheetId}' has no matching row mapping.");
            }
        }

        return overrides;
    }

    private static bool IsOverridden(
        ConflictLocation location,
        IReadOnlyList<RowMergeOverride> rowOverrides)
    {
        foreach (var rowOverride in rowOverrides)
        {
            if (rowOverride.Covers(location))
            {
                return true;
            }
        }

        return false;
    }

    private static Dictionary<long, CellResolution> BuildCellResolutionMap(
        ReadOnlyMemory<CellResolution> source,
        CancellationToken cancellationToken)
    {
        var result = new Dictionary<long, CellResolution>();
        foreach (var resolution in source.Span)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!Enum.IsDefined(resolution.Kind) || !result.TryAdd(resolution.ConflictId, resolution))
            {
                throw InvalidPlan(
                    $"Cell resolution for conflict {resolution.ConflictId} is invalid or duplicated.");
            }
        }

        return result;
    }

    private static Dictionary<long, RowResolution> BuildRowResolutionMap(
        ReadOnlyMemory<RowResolution> source,
        CancellationToken cancellationToken)
    {
        var result = new Dictionary<long, RowResolution>();
        foreach (var resolution in source.Span)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!Enum.IsDefined(resolution.Kind) || !result.TryAdd(resolution.ConflictId, resolution))
            {
                throw InvalidPlan(
                    $"Row resolution for conflict {resolution.ConflictId} is invalid or duplicated.");
            }
        }

        return result;
    }

    private static void ValidateResolutionCoverage(
        IReadOnlyDictionary<long, ConflictRecord> conflicts,
        IReadOnlyDictionary<long, CellResolution> cellResolutions,
        IReadOnlyDictionary<long, RowResolution> rowResolutions)
    {
        foreach (var conflict in conflicts.Values)
        {
            var cellScope = conflict.Kind is ConflictKind.CellValue or ConflictKind.CellDeleteEdit;
            if (cellScope != cellResolutions.ContainsKey(conflict.Id) ||
                cellScope == rowResolutions.ContainsKey(conflict.Id))
            {
                throw InvalidPlan(
                    $"Conflict {conflict.Id} does not have exactly one matching resolution.");
            }
        }

        if (cellResolutions.Keys.Any(id => !conflicts.ContainsKey(id)) ||
            rowResolutions.Keys.Any(id => !conflicts.ContainsKey(id)))
        {
            throw InvalidPlan("The plan contains a resolution for an unknown conflict.");
        }
    }

    private static void ValidateDecision(AutomaticMergeDecision decision, string sheetId)
    {
        if (!Enum.IsDefined(decision.Scope) ||
            !Enum.IsDefined(decision.Kind) ||
            !string.Equals(decision.Location.SheetId, sheetId, StringComparison.Ordinal))
        {
            throw InvalidPlan("The plan contains an invalid automatic merge decision.");
        }

        if (decision.Scope == MergeDecisionScope.Cell &&
            decision.Kind is not AutomaticMergeKind.Delete &&
            !decision.ResultValue.HasValue)
        {
            throw InvalidPlan("A non-delete cell decision requires a result value.");
        }
    }

    private static void ValidateWorkbook(WorkbookMetadata workbook, string parameterName)
    {
        ArgumentNullException.ThrowIfNull(workbook, parameterName);
        if (workbook.Format != WorkbookFormat.Xlsx)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPlan,
                "The Open XML writer requires .xlsx workbook metadata.");
        }
    }

    private static OpenXmlWriterException InvalidPlan(string message) =>
        new(OpenXmlWriterError.InvalidPlan, message);

    private static OpenXmlWriterException Unsupported(string sheetId, string message) =>
        new(
            OpenXmlWriterError.UnsupportedOperation,
            $"Worksheet group '{sheetId}' is not supported: {message}");
}

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelMerge.OpenXml;

internal sealed record OpenXmlDefinedNameState(
    DefinedName DefinedName,
    string? ScopeRelationshipId);

internal sealed record OpenXmlSheetReferenceTransform(
    string RelationshipId,
    string SourceName,
    string? ResultName,
    OpenXmlRowIndexTransform? RowTransform,
    bool IsDeleted,
    string RenamePlaceholder)
{
    public bool HasRowTransform => RowTransform is not null || IsDeleted;

    public bool HasRename =>
        ResultName is not null &&
        !string.Equals(SourceName, ResultName, StringComparison.Ordinal);

    public int? MapReferenceRow(int rowIndex) =>
        IsDeleted ? null : RowTransform?.MapReferenceRow(rowIndex) ?? rowIndex;
}

internal static class OpenXmlWorkbookReferenceTransformer
{
    public static IReadOnlyList<OpenXmlDefinedNameState> CaptureDefinedNames(
        WorkbookPart workbookPart)
    {
        ArgumentNullException.ThrowIfNull(workbookPart);
        var sheets = GetSheets(workbookPart);
        var result = new List<OpenXmlDefinedNameState>();
        foreach (var definedName in workbookPart.Workbook.DefinedNames?
            .Elements<DefinedName>() ?? [])
        {
            string? scopeRelationshipId = null;
            if (definedName.LocalSheetId?.Value is { } localSheetId)
            {
                if (localSheetId >= sheets.Count)
                {
                    throw InvalidPackage(
                        $"Defined name '{definedName.Name?.Value}' has an invalid local sheet scope.");
                }

                scopeRelationshipId = sheets[checked((int)localSheetId)].Id?.Value ?? throw InvalidPackage(
                    "A locally scoped defined name refers to a sheet without a relationship.");
            }

            result.Add(new OpenXmlDefinedNameState(definedName, scopeRelationshipId));
        }

        return result;
    }

    public static IReadOnlyList<OpenXmlSheetReferenceTransform> BuildTransforms(
        CompiledOpenXmlMerge compiled)
    {
        ArgumentNullException.ThrowIfNull(compiled);
        var result = new List<OpenXmlSheetReferenceTransform>();
        foreach (var patch in compiled.Worksheets)
        {
            if (patch.LocalRelationshipId is null || patch.LocalName is null)
            {
                continue;
            }

            var resultName = patch.Action switch
            {
                CompiledWorksheetAction.KeepLocal => patch.ResultName ?? patch.LocalName,
                CompiledWorksheetAction.ImportRemote => patch.RemoteName ?? patch.SheetGroupId,
                CompiledWorksheetAction.Delete => null,
                _ => throw new OpenXmlWriterException(
                    OpenXmlWriterError.InvalidPlan,
                    "A worksheet patch has an undefined action."),
            };
            var rowTransform = patch.Action == CompiledWorksheetAction.KeepLocal &&
                patch.RowTransform.HasChanges
                    ? new OpenXmlRowIndexTransform(patch.RowTransform)
                    : null;
            var isDeleted = patch.Action == CompiledWorksheetAction.Delete;
            if (rowTransform is null && !isDeleted &&
                string.Equals(resultName, patch.LocalName, StringComparison.Ordinal))
            {
                continue;
            }

            result.Add(new OpenXmlSheetReferenceTransform(
                patch.LocalRelationshipId,
                patch.LocalName,
                resultName,
                rowTransform,
                isDeleted,
                $"_em_{Guid.NewGuid():N}"));
        }

        return result;
    }

    public static bool Apply(
        WorkbookPart localWorkbookPart,
        WorkbookPart? remoteWorkbookPart,
        CompiledOpenXmlMerge compiled,
        IReadOnlyList<OpenXmlDefinedNameState> originalDefinedNames,
        IReadOnlyList<OpenXmlSheetReferenceTransform> transforms,
        string scratchDirectory,
        List<string> cleanupPaths,
        CancellationToken cancellationToken)
    {
        var changed = ReconcileDefinedNames(
            localWorkbookPart,
            originalDefinedNames,
            transforms,
            cancellationToken);
        changed |= ImportDefinedNames(
            localWorkbookPart,
            remoteWorkbookPart,
            compiled,
            cancellationToken);

        if (transforms.Count == 0)
        {
            return changed;
        }

        foreach (var sheet in GetSheets(localWorkbookPart))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var relationshipId = sheet.Id?.Value ?? throw InvalidPackage(
                "A result worksheet has no relationship identifier.");
            if (!localWorkbookPart.TryGetPartById(relationshipId, out var part) ||
                part is not WorksheetPart worksheetPart)
            {
                throw InvalidPackage(
                    $"Result worksheet relationship '{relationshipId}' has no worksheet part.");
            }

            if (CanTransformWorksheetFormula(relationshipId, transforms))
            {
                var scratchPath = CreateScratchPath(scratchDirectory, "formula");
                cleanupPaths.Add(scratchPath);
                try
                {
                    changed |= TransformWorksheetFormulas(
                        worksheetPart,
                        relationshipId,
                        transforms,
                        scratchPath,
                        cancellationToken);
                }
                finally
                {
                    DeleteScratchPath(cleanupPaths, scratchPath);
                }
            }

            changed |= TransformTableFormulas(
                worksheetPart,
                relationshipId,
                transforms,
                cancellationToken);
        }

        return changed;
    }

    private static bool CanTransformWorksheetFormula(
        string relationshipId,
        IReadOnlyList<OpenXmlSheetReferenceTransform> transforms) =>
        transforms.Any(transform =>
            transform.HasRename ||
            transform.HasRowTransform && !string.Equals(
                transform.RelationshipId,
                relationshipId,
                StringComparison.Ordinal));

    private static bool ReconcileDefinedNames(
        WorkbookPart workbookPart,
        IReadOnlyList<OpenXmlDefinedNameState> originalDefinedNames,
        IReadOnlyList<OpenXmlSheetReferenceTransform> transforms,
        CancellationToken cancellationToken)
    {
        var sheets = GetSheets(workbookPart);
        var positions = sheets
            .Select((sheet, index) => (RelationshipId: sheet.Id?.Value, Index: index))
            .Where(static item => item.RelationshipId is not null)
            .ToDictionary(
                static item => item.RelationshipId!,
                static item => item.Index,
                StringComparer.Ordinal);
        var changed = false;
        foreach (var state in originalDefinedNames)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var definedName = state.DefinedName;
            if (state.ScopeRelationshipId is not null)
            {
                if (!positions.TryGetValue(state.ScopeRelationshipId, out var resultPosition))
                {
                    definedName.Remove();
                    changed = true;
                    continue;
                }

                if (definedName.LocalSheetId?.Value != (uint)resultPosition)
                {
                    definedName.LocalSheetId = checked((uint)resultPosition);
                    changed = true;
                }
            }

            var originalText = definedName.Text ?? string.Empty;
            var transformedText = TransformFormulaText(
                originalText,
                state.ScopeRelationshipId,
                transforms,
                rewriteOwnRows: true,
                definedName: true);
            if (!string.Equals(originalText, transformedText, StringComparison.Ordinal))
            {
                definedName.Text = transformedText;
                changed = true;
            }
        }

        if (workbookPart.Workbook.DefinedNames is { HasChildren: false } emptyDefinedNames)
        {
            emptyDefinedNames.Remove();
            changed = true;
        }

        return changed;
    }

    private static bool ImportDefinedNames(
        WorkbookPart localWorkbookPart,
        WorkbookPart? remoteWorkbookPart,
        CompiledOpenXmlMerge compiled,
        CancellationToken cancellationToken)
    {
        if (remoteWorkbookPart is null)
        {
            return false;
        }

        var remoteSheets = GetSheets(remoteWorkbookPart);
        var localSheets = GetSheets(localWorkbookPart);
        var changed = false;
        foreach (var patch in compiled.Worksheets.Where(static patch =>
            patch.Action == CompiledWorksheetAction.ImportRemote))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var remotePosition = remoteSheets.FindIndex(sheet => string.Equals(
                sheet.Id?.Value,
                patch.RemoteRelationshipId,
                StringComparison.Ordinal));
            var resultName = patch.RemoteName ?? patch.SheetGroupId;
            var localPosition = localSheets.FindIndex(sheet => string.Equals(
                sheet.Name?.Value,
                resultName,
                StringComparison.OrdinalIgnoreCase));
            if (remotePosition < 0 || localPosition < 0)
            {
                throw InvalidPackage(
                    $"Imported worksheet '{resultName}' could not be resolved for defined-name import.");
            }

            var sourceNames = remoteWorkbookPart.Workbook.DefinedNames?
                .Elements<DefinedName>()
                .Where(name => name.LocalSheetId?.Value == (uint)remotePosition)
                .ToArray() ?? [];
            if (sourceNames.Length == 0)
            {
                continue;
            }

            var container = localWorkbookPart.Workbook.DefinedNames;
            if (container is null)
            {
                container = new DefinedNames();
                localWorkbookPart.Workbook.DefinedNames = container;
            }

            foreach (var sourceName in sourceNames)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var name = sourceName.Name?.Value;
                if (string.IsNullOrWhiteSpace(name) || container.Elements<DefinedName>().Any(existing =>
                    existing.LocalSheetId?.Value == (uint)localPosition &&
                    string.Equals(existing.Name?.Value, name, StringComparison.OrdinalIgnoreCase)))
                {
                    throw new OpenXmlWriterException(
                        OpenXmlWriterError.UnsupportedOperation,
                        $"Imported worksheet '{resultName}' has a duplicate or invalid defined name '{name}'.");
                }

                if (OpenXmlFormulaRowRewriter.ContainsForeignSheetReference(
                    sourceName.Text ?? string.Empty,
                    patch.RemoteName ?? patch.SheetGroupId))
                {
                    throw new OpenXmlWriterException(
                        OpenXmlWriterError.UnsupportedOperation,
                        $"Imported defined name '{name}' depends on another REMOTE worksheet.");
                }

                var clone = (DefinedName)sourceName.CloneNode(deep: true);
                clone.LocalSheetId = checked((uint)localPosition);
                container.Append(clone);
                changed = true;
            }
        }

        return changed;
    }

    private static bool TransformWorksheetFormulas(
        WorksheetPart worksheetPart,
        string relationshipId,
        IReadOnlyList<OpenXmlSheetReferenceTransform> transforms,
        string scratchPath,
        CancellationToken cancellationToken)
    {
        var changed = false;
        using (var reader = DocumentFormat.OpenXml.OpenXmlReader.Create(
            worksheetPart,
            readMiscNodes: true))
        using (var stream = new FileStream(
            scratchPath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 128 * 1024,
            FileOptions.SequentialScan))
        using (var writer = DocumentFormat.OpenXml.OpenXmlWriter.Create(stream, Encoding.UTF8))
        {
            writer.WriteStartDocument();
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.IsStartElement && IsFormulaElement(reader.ElementType))
                {
                    var formula = reader.LoadCurrentElement() as OpenXmlLeafTextElement ??
                        throw InvalidPackage("A worksheet formula could not be loaded.");
                    var original = formula.Text ?? string.Empty;
                    var transformed = TransformFormulaText(
                        original,
                        relationshipId,
                        transforms,
                        rewriteOwnRows: false,
                        definedName: false);
                    if (!string.Equals(original, transformed, StringComparison.Ordinal))
                    {
                        formula.Text = transformed;
                        changed = true;
                    }

                    writer.WriteElement(formula);
                }
                else if (reader.IsStartElement &&
                    typeof(OpenXmlLeafTextElement).IsAssignableFrom(reader.ElementType))
                {
                    var leaf = reader.LoadCurrentElement() ??
                        throw InvalidPackage($"Leaf element '{reader.LocalName}' could not be loaded.");
                    writer.WriteElement(leaf);
                }
                else if (reader.IsStartElement)
                {
                    writer.WriteStartElement(reader);
                }
                else if (reader.IsEndElement)
                {
                    writer.WriteEndElement();
                }
                else if (reader.IsMiscNode)
                {
                    writer.WriteString(reader.GetText());
                }
            }
        }

        if (changed)
        {
            using var input = new FileStream(
                scratchPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read,
                bufferSize: 128 * 1024,
                FileOptions.SequentialScan);
            worksheetPart.FeedData(input);
        }

        return changed;
    }

    private static bool TransformTableFormulas(
        WorksheetPart worksheetPart,
        string relationshipId,
        IReadOnlyList<OpenXmlSheetReferenceTransform> transforms,
        CancellationToken cancellationToken)
    {
        var changed = false;
        foreach (var tablePart in worksheetPart.TableDefinitionParts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var tableChanged = false;
            foreach (var formula in tablePart.Table
                .Descendants<OpenXmlLeafTextElement>()
                .Where(static element => element is CalculatedColumnFormula or TotalsRowFormula))
            {
                var original = formula.Text ?? string.Empty;
                var transformed = TransformFormulaText(
                    original,
                    relationshipId,
                    transforms,
                    rewriteOwnRows: true,
                    definedName: false);
                if (!string.Equals(original, transformed, StringComparison.Ordinal))
                {
                    formula.Text = transformed;
                    tableChanged = true;
                }
            }

            if (tableChanged)
            {
                tablePart.Table.Save();
                changed = true;
            }
        }

        return changed;
    }

    private static void DeleteScratchPath(List<string> cleanupPaths, string path)
    {
        cleanupPaths.Remove(path);
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            cleanupPaths.Add(path);
        }
    }

    private static string TransformFormulaText(
        string formula,
        string? ownerRelationshipId,
        IReadOnlyList<OpenXmlSheetReferenceTransform> transforms,
        bool rewriteOwnRows,
        bool definedName)
    {
        var result = formula;
        foreach (var transform in transforms)
        {
            if (!transform.HasRowTransform ||
                string.Equals(transform.RelationshipId, ownerRelationshipId, StringComparison.Ordinal) &&
                !rewriteOwnRows)
            {
                continue;
            }

            var rewriteUnqualified = string.Equals(
                transform.RelationshipId,
                ownerRelationshipId,
                StringComparison.Ordinal);
            result = definedName
                ? OpenXmlFormulaRowRewriter.RewriteDefinedName(
                    result,
                    transform.MapReferenceRow,
                    transform.SourceName,
                    rewriteUnqualified)
                : OpenXmlFormulaRowRewriter.Rewrite(
                    result,
                    transform.MapReferenceRow,
                    transform.SourceName,
                    rewriteUnqualified);
        }

        foreach (var transform in transforms.Where(static transform => transform.HasRename))
        {
            result = OpenXmlFormulaRowRewriter.RenameSheetReferences(
                result,
                transform.SourceName,
                transform.RenamePlaceholder);
        }

        foreach (var transform in transforms.Where(static transform => transform.HasRename))
        {
            result = OpenXmlFormulaRowRewriter.RenameSheetReferences(
                result,
                transform.RenamePlaceholder,
                transform.ResultName!);
        }

        return result;
    }

    private static bool IsFormulaElement(Type elementType) =>
        elementType == typeof(CellFormula) ||
        elementType == typeof(Formula) ||
        elementType == typeof(Formula1) ||
        elementType == typeof(Formula2);

    private static List<Sheet> GetSheets(WorkbookPart workbookPart) =>
        workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList() ??
        throw InvalidPackage("The workbook does not contain a sheet collection.");

    private static string CreateScratchPath(string directory, string prefix)
    {
        for (var attempt = 0; attempt < 10; attempt++)
        {
            var path = Path.Combine(directory, $".{prefix}.{Guid.NewGuid():N}.tmp");
            if (!File.Exists(path) && !Directory.Exists(path))
            {
                return path;
            }
        }

        throw new IOException("A unique workbook-reference scratch path could not be allocated.");
    }

    private static OpenXmlWriterException InvalidPackage(string message) =>
        new(OpenXmlWriterError.InvalidPackage, message);
}

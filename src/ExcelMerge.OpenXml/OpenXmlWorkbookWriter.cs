using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelMerge.Domain;
using System.IO.Compression;

namespace ExcelMerge.OpenXml;

/// <summary>Writes a validated merge result through a same-directory transactional package.</summary>
public sealed class OpenXmlWorkbookWriter : IOpenXmlWorkbookWriter
{
    public async ValueTask<OpenXmlWriteResult> WriteAsync(
        OpenXmlWriteRequest request,
        OpenXmlWriterOptions? options = null,
        IProgress<OpenXmlWriterProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        options ??= new OpenXmlWriterOptions();
        ValidateOptions(options);
        cancellationToken.ThrowIfCancellationRequested();

        var compiled = OpenXmlMergePlanCompiler.Compile(request.Plan, cancellationToken);
        var destinationPath = request.DestinationPath;
        var destinationDirectory = Path.GetDirectoryName(destinationPath);
        if (string.IsNullOrEmpty(destinationDirectory) || !Directory.Exists(destinationDirectory))
        {
            throw new DirectoryNotFoundException(
                $"The destination directory '{destinationDirectory}' does not exist.");
        }

        if (!options.Overwrite && File.Exists(destinationPath))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.DestinationExists,
                "The destination already exists and overwrite is disabled.",
                destinationPath);
        }

        progress?.Report(new OpenXmlWriterProgress(
            OpenXmlWriterStage.Preflight,
            TotalWorksheetCount: compiled.Worksheets.Count));

        var cleanupPaths = new List<string>();
        FileStream? baseStream = null;
        FileStream? localStream = null;
        FileStream? remoteStream = null;
        string? transactionPath = null;
        var committed = false;
        try
        {
            baseStream = await OpenVerifiedAsync(request.BaseSource, cancellationToken).ConfigureAwait(false);
            localStream = await OpenVerifiedAsync(request.LocalSource, cancellationToken).ConfigureAwait(false);
            remoteStream = await OpenVerifiedAsync(request.RemoteSource, cancellationToken).ConfigureAwait(false);
            OpenXmlSourcePackagePreflight.Validate(request.BaseSource.FilePath, cancellationToken);
            OpenXmlSourcePackagePreflight.Validate(request.LocalSource.FilePath, cancellationToken);
            OpenXmlSourcePackagePreflight.Validate(request.RemoteSource.FilePath, cancellationToken);
            EnsureTemporarySpace(
                destinationPath,
                request.LocalSource,
                request.RemoteSource,
                options);

            transactionPath = CreateTemporaryPath(destinationDirectory, Path.GetFileName(destinationPath));
            cleanupPaths.Add(transactionPath);
            progress?.Report(new OpenXmlWriterProgress(
                OpenXmlWriterStage.CopyingLocal,
                TotalWorksheetCount: compiled.Worksheets.Count));
            await CopyLocalAsync(localStream, transactionPath, cancellationToken).ConfigureAwait(false);

            progress?.Report(new OpenXmlWriterProgress(
                OpenXmlWriterStage.ApplyingPlan,
                TotalWorksheetCount: compiled.Worksheets.Count));
            ApplyPlan(
                transactionPath,
                compiled,
                request.RemoteSource.FilePath,
                destinationDirectory,
                cleanupPaths,
                progress,
                cancellationToken);

            if (options.ValidatePackage)
            {
                progress?.Report(new OpenXmlWriterProgress(
                    OpenXmlWriterStage.Validating,
                    compiled.Worksheets.Count,
                    compiled.Worksheets.Count));
                OpenXmlPackageValidator.Validate(
                    transactionPath,
                    options.MaximumValidationErrors,
                    cancellationToken);
            }

            await VerifyOpenSourceAsync(request.BaseSource, baseStream, cancellationToken).ConfigureAwait(false);
            await VerifyOpenSourceAsync(request.LocalSource, localStream, cancellationToken).ConfigureAwait(false);
            await VerifyOpenSourceAsync(request.RemoteSource, remoteStream, cancellationToken).ConfigureAwait(false);
            await FlushToDiskAsync(transactionPath, cancellationToken).ConfigureAwait(false);

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report(new OpenXmlWriterProgress(
                OpenXmlWriterStage.Committing,
                compiled.Worksheets.Count,
                compiled.Worksheets.Count));
            Commit(transactionPath, destinationPath);
            committed = true;
            cleanupPaths.Remove(transactionPath);

            var resultSource = await OpenXmlWorkbookSource.CaptureAsync(
                destinationPath,
                CancellationToken.None).ConfigureAwait(false);
            progress?.Report(new OpenXmlWriterProgress(
                OpenXmlWriterStage.Completed,
                compiled.Worksheets.Count,
                compiled.Worksheets.Count));
            return new OpenXmlWriteResult(
                destinationPath,
                resultSource.Fingerprint.ContentLength,
                resultSource.Fingerprint);
        }
        catch (OpenXmlWriterException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is
            IOException or
            UnauthorizedAccessException or
            OpenXmlPackageException or
            InvalidDataException or
            FormatException)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.IoFailure,
                "The workbook result could not be written transactionally.",
                destinationPath,
                innerException: exception);
        }
        finally
        {
            await DisposeAsync(remoteStream).ConfigureAwait(false);
            await DisposeAsync(localStream).ConfigureAwait(false);
            await DisposeAsync(baseStream).ConfigureAwait(false);
            if (!committed)
            {
                DeletePaths(cleanupPaths);
            }
            else
            {
                DeletePaths(cleanupPaths);
            }
        }
    }

    private static async ValueTask<FileStream> OpenVerifiedAsync(
        OpenXmlWorkbookSource source,
        CancellationToken cancellationToken)
    {
        var stream = OpenXmlWorkbookSource.OpenRead(source.FilePath);
        try
        {
            await VerifyOpenSourceAsync(source, stream, cancellationToken).ConfigureAwait(false);
            return stream;
        }
        catch
        {
            await stream.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    private static async ValueTask VerifyOpenSourceAsync(
        OpenXmlWorkbookSource source,
        FileStream stream,
        CancellationToken cancellationToken)
    {
        var actual = await OpenXmlWorkbookSource.ComputeFingerprintAsync(
            source.FilePath,
            stream,
            cancellationToken).ConfigureAwait(false);
        if (actual != source.Fingerprint)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.SourceChanged,
                "A workbook source changed after the merge session was created.",
                source.FilePath);
        }
    }

    private static async Task CopyLocalAsync(
        FileStream localStream,
        string transactionPath,
        CancellationToken cancellationToken)
    {
        localStream.Position = 0;
        await using var output = new FileStream(
            transactionPath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 128 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        await localStream.CopyToAsync(output, cancellationToken).ConfigureAwait(false);
        await output.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private static void ApplyPlan(
        string transactionPath,
        CompiledOpenXmlMerge compiled,
        string remotePath,
        string destinationDirectory,
        List<string> cleanupPaths,
        IProgress<OpenXmlWriterProgress>? progress,
        CancellationToken cancellationToken)
    {
        using var document = SpreadsheetDocument.Open(
            transactionPath,
            isEditable: true,
            new OpenSettings { AutoSave = false });
        if (document.DigitalSignatureOriginPart is not null)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "Digitally signed workbooks cannot be modified.",
                transactionPath);
        }

        var workbookPart = document.WorkbookPart ?? throw new OpenXmlWriterException(
            OpenXmlWriterError.InvalidPackage,
            "The LOCAL package does not contain a workbook part.",
            transactionPath);
        var originalDefinedNames = OpenXmlWorkbookReferenceTransformer.CaptureDefinedNames(workbookPart);
        var referenceTransforms = OpenXmlWorkbookReferenceTransformer.BuildTransforms(compiled);
        ValidateFinalWorksheetNames(workbookPart, compiled);
        var uses1904DateSystem = workbookPart.Workbook.WorkbookProperties?.Date1904?.Value ?? false;
        SpreadsheetDocument? remoteDocument = null;
        WorkbookPart? remoteWorkbookPart = null;
        OpenXmlWriterSharedStrings? remoteSharedStrings = null;
        OpenXmlStyleMapper? styleMapper = null;
        if (compiled.Worksheets.Any(static worksheet =>
            worksheet.Action == CompiledWorksheetAction.ImportRemote ||
            worksheet.RowTransform.RequiresRemoteRows))
        {
            remoteDocument = SpreadsheetDocument.Open(
                remotePath,
                isEditable: false,
                new OpenSettings { AutoSave = false });
            remoteWorkbookPart = remoteDocument.WorkbookPart ?? throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                "The REMOTE package does not contain a workbook part.",
                remotePath);
            styleMapper = new OpenXmlStyleMapper(workbookPart, remoteWorkbookPart);
            var sharedStringStorePath = CreateTemporaryPath(
                destinationDirectory,
                "remote-shared-strings");
            cleanupPaths.Add(sharedStringStorePath);
            remoteSharedStrings = OpenXmlWriterSharedStrings.Load(
                remoteWorkbookPart.SharedStringTablePart,
                sharedStringStorePath,
                cancellationToken);
        }

        var referencesChanged = false;
        try
        {
            for (var index = 0; index < compiled.Worksheets.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var worksheet = compiled.Worksheets[index];
                switch (worksheet.Action)
                {
                    case CompiledWorksheetAction.KeepLocal when
                        worksheet.Rows.Count != 0 || worksheet.RowTransform.HasChanges:
                        {
                            var scratchPath = CreateTemporaryPath(
                                destinationDirectory,
                                $"sheet-{index:D4}.xml");
                            cleanupPaths.Add(scratchPath);
                            try
                            {
                                OpenXmlWorksheetPatcher.Apply(
                                    workbookPart,
                                    remoteWorkbookPart,
                                    remoteSharedStrings,
                                    styleMapper,
                                    worksheet,
                                    scratchPath,
                                    uses1904DateSystem,
                                    cancellationToken);
                            }
                            finally
                            {
                                DeleteTemporaryPath(cleanupPaths, scratchPath);
                            }
                            break;
                        }
                    case CompiledWorksheetAction.Delete:
                        DeleteWorksheet(workbookPart, worksheet.LocalRelationshipId);
                        break;
                    case CompiledWorksheetAction.ImportRemote:
                        ImportWorksheet(
                            workbookPart,
                            remoteWorkbookPart!,
                            remoteSharedStrings!,
                            styleMapper!,
                            worksheet,
                            destinationDirectory,
                            cleanupPaths,
                            cancellationToken);
                        break;
                }

                if (worksheet.Action == CompiledWorksheetAction.KeepLocal &&
                    (worksheet.ResultName is not null || worksheet.ResultVisibility.HasValue))
                {
                    UpdateWorksheetMetadata(workbookPart, worksheet);
                }

                progress?.Report(new OpenXmlWriterProgress(
                    OpenXmlWriterStage.ApplyingPlan,
                    index + 1,
                    compiled.Worksheets.Count,
                    worksheet.SheetGroupId));
            }

            referencesChanged = OpenXmlWorkbookReferenceTransformer.Apply(
                workbookPart,
                remoteWorkbookPart,
                compiled,
                originalDefinedNames,
                referenceTransforms,
                destinationDirectory,
                cleanupPaths,
                cancellationToken);
            if (compiled.Worksheets.Count != 0 || referencesChanged)
            {
                workbookPart.Workbook.Save();
            }
        }
        finally
        {
            remoteSharedStrings?.Dispose();
            remoteDocument?.Dispose();
        }

        if (compiled.RequiresRecalculation || referencesChanged)
        {
            if (workbookPart.CalculationChainPart is { } calculationChainPart)
            {
                workbookPart.DeletePart(calculationChainPart);
            }

            var calculation = workbookPart.Workbook.CalculationProperties;
            if (calculation is null)
            {
                calculation = new CalculationProperties();
                workbookPart.Workbook.Append(calculation);
            }

            calculation.CalculationMode = CalculateModeValues.Auto;
            calculation.FullCalculationOnLoad = true;
            calculation.ForceFullCalculation = true;
            workbookPart.Workbook.Save();
        }

        NormalizeFontElementOrder(workbookPart);
    }

    private static void NormalizeFontElementOrder(WorkbookPart workbookPart)
    {
        var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
        if (stylesheet is null)
        {
            return;
        }

        var changed = false;
        foreach (var font in stylesheet.Descendants<Font>())
        {
            var children = font.ChildElements.ToArray();
            var ordered = children.OrderBy(FontElementOrder).ToArray();
            if (children.Select((child, index) => ReferenceEquals(child, ordered[index])).All(static value => value))
            {
                continue;
            }

            foreach (var child in children)
            {
                child.Remove();
            }

            foreach (var child in ordered)
            {
                font.Append(child);
            }

            changed = true;
        }

        if (changed)
        {
            stylesheet.Save();
        }
    }

    private static int FontElementOrder(OpenXmlElement element) => element switch
    {
        Bold => 0,
        Italic => 1,
        Strike => 2,
        Condense => 3,
        Extend => 4,
        Outline => 5,
        Shadow => 6,
        Underline => 7,
        VerticalTextAlignment => 8,
        FontSize => 9,
        Color => 10,
        FontName => 11,
        FontFamilyNumbering => 12,
        FontCharSet => 13,
        FontScheme => 14,
        _ => 15,
    };

    private static void UpdateWorksheetMetadata(
        WorkbookPart workbookPart,
        CompiledWorksheetPatch patch)
    {
        var relationshipId = patch.LocalRelationshipId ?? throw new OpenXmlWriterException(
            OpenXmlWriterError.InvalidPlan,
            $"Worksheet group '{patch.SheetGroupId}' has no LOCAL relationship.");
        var sheet = workbookPart.Workbook.Sheets?
            .Elements<Sheet>()
            .FirstOrDefault(candidate => string.Equals(
                candidate.Id?.Value,
                relationshipId,
                StringComparison.Ordinal)) ?? throw new OpenXmlWriterException(
                    OpenXmlWriterError.InvalidPlan,
                    $"LOCAL worksheet relationship '{relationshipId}' is missing.");
        if (patch.ResultName is not null)
        {
            sheet.Name = patch.ResultName;
        }

        if (patch.ResultVisibility is { } visibility)
        {
            sheet.State = visibility switch
            {
                SheetVisibility.Visible => SheetStateValues.Visible,
                SheetVisibility.Hidden => SheetStateValues.Hidden,
                SheetVisibility.VeryHidden => SheetStateValues.VeryHidden,
                _ => throw new OpenXmlWriterException(
                    OpenXmlWriterError.InvalidPlan,
                    "The result worksheet has an undefined visibility."),
            };
        }
    }

    private static void DeleteWorksheet(WorkbookPart workbookPart, string? relationshipId)
    {
        if (relationshipId is null)
        {
            return;
        }

        var sheets = workbookPart.Workbook.Sheets ?? throw new OpenXmlWriterException(
            OpenXmlWriterError.InvalidPackage,
            "The LOCAL workbook does not contain a sheet collection.");
        var sheet = sheets.Elements<Sheet>().FirstOrDefault(
            candidate => string.Equals(candidate.Id?.Value, relationshipId, StringComparison.Ordinal));
        if (sheet is null)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPlan,
                $"LOCAL worksheet relationship '{relationshipId}' is missing.");
        }

        if (!workbookPart.TryGetPartById(relationshipId, out var part))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                $"LOCAL worksheet relationship '{relationshipId}' has no part.");
        }

        sheet.Remove();
        workbookPart.DeletePart(part);
    }

    private static void ImportWorksheet(
        WorkbookPart localWorkbookPart,
        WorkbookPart remoteWorkbookPart,
        OpenXmlWriterSharedStrings remoteSharedStrings,
        OpenXmlStyleMapper styleMapper,
        CompiledWorksheetPatch patch,
        string destinationDirectory,
        List<string> cleanupPaths,
        CancellationToken cancellationToken)
    {
        if (patch.RemoteRelationshipId is null)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPlan,
                $"REMOTE worksheet group '{patch.SheetGroupId}' has no relationship.");
        }

        if (!remoteWorkbookPart.TryGetPartById(patch.RemoteRelationshipId, out var remotePart) ||
            remotePart is not WorksheetPart remoteWorksheetPart)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                $"REMOTE worksheet relationship '{patch.RemoteRelationshipId}' has no worksheet part.");
        }

        ValidateImportedWorksheet(
            localWorkbookPart,
            remoteWorksheetPart,
            patch);
        DeleteWorksheet(localWorkbookPart, patch.LocalRelationshipId);
        cancellationToken.ThrowIfCancellationRequested();
        var importedPart = localWorkbookPart.AddPart(remoteWorksheetPart);
        var scratchPath = CreateTemporaryPath(
            destinationDirectory,
            $"import-{patch.SheetGroupId}.xml");
        cleanupPaths.Add(scratchPath);
        try
        {
            OpenXmlImportedWorksheetTransformer.Transform(
                importedPart,
                remoteSharedStrings,
                styleMapper,
                scratchPath,
                cancellationToken);
        }
        finally
        {
            DeleteTemporaryPath(cleanupPaths, scratchPath);
        }
        RemapImportedTableIds(localWorkbookPart, importedPart);
        var relationshipId = localWorkbookPart.GetIdOfPart(importedPart);
        var sheets = localWorkbookPart.Workbook.Sheets ?? localWorkbookPart.Workbook.AppendChild(new Sheets());
        var maximumSheetId = sheets.Elements<Sheet>()
            .Select(static sheet => sheet.SheetId?.Value ?? 0U)
            .DefaultIfEmpty(0U)
            .Max();
        if (maximumSheetId == uint.MaxValue)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "A unique worksheet identifier could not be allocated.");
        }

        var nextSheetId = maximumSheetId + 1U;
        var sheet = new Sheet
        {
            Name = patch.RemoteName ?? patch.SheetGroupId,
            SheetId = nextSheetId,
            Id = relationshipId,
        };
        if (patch.RemoteVisibility == SheetVisibility.Hidden)
        {
            sheet.State = SheetStateValues.Hidden;
        }
        else if (patch.RemoteVisibility == SheetVisibility.VeryHidden)
        {
            sheet.State = SheetStateValues.VeryHidden;
        }

        sheets.Append(sheet);
    }

    private static void ValidateFinalWorksheetNames(
        WorkbookPart workbookPart,
        CompiledOpenXmlMerge compiled)
    {
        var result = workbookPart.Workbook.Sheets?
            .Elements<Sheet>()
            .Select(sheet => new WorksheetNameState(
                sheet.Id?.Value ?? throw new OpenXmlWriterException(
                    OpenXmlWriterError.InvalidPackage,
                    "A LOCAL worksheet has no relationship identifier."),
                sheet.Name?.Value ?? string.Empty))
            .ToList() ?? throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                "The LOCAL workbook does not contain a sheet collection.");

        foreach (var patch in compiled.Worksheets)
        {
            var localIndex = patch.LocalRelationshipId is null
                ? -1
                : result.FindIndex(sheet => string.Equals(
                    sheet.RelationshipId,
                    patch.LocalRelationshipId,
                    StringComparison.Ordinal));
            switch (patch.Action)
            {
                case CompiledWorksheetAction.KeepLocal:
                    if (localIndex < 0)
                    {
                        throw new OpenXmlWriterException(
                            OpenXmlWriterError.InvalidPlan,
                            $"LOCAL worksheet relationship '{patch.LocalRelationshipId}' is missing.");
                    }

                    if (patch.ResultName is not null)
                    {
                        result[localIndex] = result[localIndex] with { Name = patch.ResultName };
                    }

                    break;
                case CompiledWorksheetAction.Delete:
                    if (localIndex >= 0)
                    {
                        result.RemoveAt(localIndex);
                    }

                    break;
                case CompiledWorksheetAction.ImportRemote:
                    if (localIndex >= 0)
                    {
                        result.RemoveAt(localIndex);
                    }

                    result.Add(new WorksheetNameState(
                        $"import:{patch.SheetGroupId}",
                        patch.RemoteName ?? patch.SheetGroupId));
                    break;
                default:
                    throw new OpenXmlWriterException(
                        OpenXmlWriterError.InvalidPlan,
                        "A worksheet patch has an undefined action.");
            }
        }

        if (result.Count == 0)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "A workbook result must contain at least one worksheet.");
        }

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in result)
        {
            if (!IsValidWorksheetName(sheet.Name) || !names.Add(sheet.Name))
            {
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.InvalidPlan,
                    $"Result worksheet name '{sheet.Name}' is invalid or duplicated.");
            }
        }
    }

    private static bool IsValidWorksheetName(string name) =>
        name.Length is > 0 and <= 31 &&
        name[0] != '\'' &&
        name[^1] != '\'' &&
        name.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) < 0;

    private static void ValidateImportedWorksheet(
        WorkbookPart localWorkbookPart,
        WorksheetPart remoteWorksheetPart,
        CompiledWorksheetPatch patch)
    {
        if (remoteWorksheetPart.PivotTableParts.Any() ||
            remoteWorksheetPart.QueryTableParts.Any() ||
            remoteWorksheetPart.SingleCellTablePart is not null ||
            remoteWorksheetPart.WorksheetThreadedCommentsParts.Any())
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                $"REMOTE worksheet group '{patch.SheetGroupId}' depends on unsupported pivot, query, single-cell table, or threaded-comment parts.");
        }

        ValidateImportedPartGraph(remoteWorksheetPart, patch.SheetGroupId);
        var hyperlinkRelationshipIds = remoteWorksheetPart.HyperlinkRelationships
            .Select(static relationship => relationship.Id)
            .ToHashSet(StringComparer.Ordinal);
        if (remoteWorksheetPart.ExternalRelationships.Any(relationship =>
            !hyperlinkRelationshipIds.Contains(relationship.Id)))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                $"REMOTE worksheet group '{patch.SheetGroupId}' contains an unsupported external relationship.");
        }

        WorksheetPart? replacedWorksheet = null;
        if (patch.LocalRelationshipId is not null &&
            localWorkbookPart.TryGetPartById(patch.LocalRelationshipId, out var replacedPart))
        {
            replacedWorksheet = replacedPart as WorksheetPart;
        }

        var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var worksheetPart in localWorkbookPart.WorksheetParts)
        {
            if (ReferenceEquals(worksheetPart, replacedWorksheet))
            {
                continue;
            }

            foreach (var tablePart in worksheetPart.TableDefinitionParts)
            {
                AddTableNames(tablePart.Table, existingNames, "LOCAL");
            }
        }

        foreach (var definedName in localWorkbookPart.Workbook.DefinedNames?
            .Elements<DefinedName>() ?? [])
        {
            if (!string.IsNullOrWhiteSpace(definedName.Name?.Value))
            {
                existingNames.Add(definedName.Name!.Value!);
            }
        }

        foreach (var tablePart in remoteWorksheetPart.TableDefinitionParts)
        {
            AddTableNames(tablePart.Table, existingNames, "REMOTE import");
        }
    }

    private static void ValidateImportedPartGraph(OpenXmlPart root, string sheetGroupId)
    {
        var visited = new HashSet<OpenXmlPart>();
        var active = new HashSet<OpenXmlPart>();
        Visit(root);
        return;

        void Visit(OpenXmlPart part)
        {
            if (!active.Add(part))
            {
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.UnsupportedOperation,
                    $"REMOTE worksheet group '{sheetGroupId}' contains a cyclic relationship graph.");
            }

            if (!visited.Add(part))
            {
                active.Remove(part);
                return;
            }

            if (part is PivotTablePart or PivotTableCacheDefinitionPart or QueryTablePart or SingleCellTablePart)
            {
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.UnsupportedOperation,
                    $"REMOTE worksheet group '{sheetGroupId}' contains an unsupported workbook-scoped dependency.");
            }

            foreach (var child in part.Parts)
            {
                Visit(child.OpenXmlPart);
            }

            active.Remove(part);
        }
    }

    private static void AddTableNames(
        Table table,
        HashSet<string> workbookNames,
        string source)
    {
        var tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(table.Name?.Value))
        {
            tableNames.Add(table.Name!.Value!);
        }

        if (!string.IsNullOrWhiteSpace(table.DisplayName?.Value))
        {
            tableNames.Add(table.DisplayName!.Value!);
        }

        if (tableNames.Count == 0)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                $"A {source} table has no name.");
        }

        foreach (var name in tableNames)
        {
            if (!workbookNames.Add(name))
            {
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.UnsupportedOperation,
                    $"A {source} table name '{name}' collides with another workbook name.");
            }
        }
    }

    private static void RemapImportedTableIds(
        WorkbookPart localWorkbookPart,
        WorksheetPart importedWorksheetPart)
    {
        var usedIds = localWorkbookPart.WorksheetParts
            .Where(worksheetPart => !ReferenceEquals(worksheetPart, importedWorksheetPart))
            .SelectMany(static worksheetPart => worksheetPart.TableDefinitionParts)
            .Select(static tablePart => tablePart.Table.Id?.Value ?? 0U)
            .Where(static id => id != 0U)
            .ToHashSet();
        var nextId = usedIds.DefaultIfEmpty(0U).Max();
        foreach (var tablePart in importedWorksheetPart.TableDefinitionParts)
        {
            var requestedId = tablePart.Table.Id?.Value ?? 0U;
            if (requestedId == 0U || !usedIds.Add(requestedId))
            {
                do
                {
                    if (nextId == uint.MaxValue)
                    {
                        throw new OpenXmlWriterException(
                            OpenXmlWriterError.UnsupportedOperation,
                            "A unique table identifier could not be allocated.");
                    }

                    nextId++;
                }
                while (!usedIds.Add(nextId));

                tablePart.Table.Id = nextId;
                tablePart.Table.Save();
            }
            else
            {
                nextId = Math.Max(nextId, requestedId);
            }
        }
    }

    private sealed record WorksheetNameState(string RelationshipId, string Name);

    private static async Task FlushToDiskAsync(
        string transactionPath,
        CancellationToken cancellationToken)
    {
        await using var stream = new FileStream(
            transactionPath,
            FileMode.Open,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 1,
            FileOptions.Asynchronous | FileOptions.WriteThrough);
        cancellationToken.ThrowIfCancellationRequested();
        stream.Flush(flushToDisk: true);
    }

    private static void Commit(string transactionPath, string destinationPath)
    {
        if (File.Exists(destinationPath))
        {
            try
            {
                File.Replace(transactionPath, destinationPath, destinationBackupFileName: null);
            }
            catch (PlatformNotSupportedException)
            {
                File.Move(transactionPath, destinationPath, overwrite: true);
            }
        }
        else
        {
            File.Move(transactionPath, destinationPath);
        }
    }

    private static void EnsureTemporarySpace(
        string destinationPath,
        OpenXmlWorkbookSource localSource,
        OpenXmlWorkbookSource remoteSource,
        OpenXmlWriterOptions options)
    {
        var root = Path.GetPathRoot(destinationPath);
        if (string.IsNullOrEmpty(root))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InsufficientTemporarySpace,
                "The destination volume could not be determined.",
                destinationPath);
        }

        long required;
        try
        {
            var expandedLocalLength = GetExpandedPackageLength(localSource.FilePath);
            var expandedRemoteLength = GetExpandedPackageLength(remoteSource.FilePath);
            var maximumOutputLength = checked(
                localSource.Fingerprint.ContentLength + remoteSource.Fingerprint.ContentLength);
            required = checked(
                (maximumOutputLength * 2) +
                expandedLocalLength +
                expandedRemoteLength +
                options.MinimumFreeSpaceReserveBytes);
        }
        catch (OverflowException exception)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InsufficientTemporarySpace,
                "The temporary-space estimate overflowed.",
                destinationPath,
                innerException: exception);
        }

        var available = new DriveInfo(root).AvailableFreeSpace;
        if (available < required)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InsufficientTemporarySpace,
                $"The destination volume has {available:N0} bytes free; at least {required:N0} bytes are required.",
                destinationPath);
        }
    }

    private static long GetExpandedPackageLength(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        long length = 0;
        foreach (var entry in archive.Entries)
        {
            length = checked(length + entry.Length);
        }

        return length;
    }

    private static string CreateTemporaryPath(string directory, string fileName)
    {
        for (var attempt = 0; attempt < 10; attempt++)
        {
            var candidate = Path.Combine(
                directory,
                $".{fileName}.{Guid.NewGuid():N}.tmp");
            if (!File.Exists(candidate) && !Directory.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new IOException("A unique transactional output path could not be allocated.");
    }

    private static void ValidateOptions(OpenXmlWriterOptions options)
    {
        if (options.MaximumValidationErrors <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(options.MaximumValidationErrors),
                "The maximum validation error count must be positive.");
        }

        ArgumentOutOfRangeException.ThrowIfNegative(options.MinimumFreeSpaceReserveBytes);
    }

    private static async ValueTask DisposeAsync(FileStream? stream)
    {
        if (stream is not null)
        {
            await stream.DisposeAsync().ConfigureAwait(false);
        }
    }

    private static void DeletePaths(List<string> paths)
    {
        for (var index = paths.Count - 1; index >= 0; index--)
        {
            try
            {
                if (File.Exists(paths[index]))
                {
                    File.Delete(paths[index]);
                }
                else if (Directory.Exists(paths[index]))
                {
                    Directory.Delete(paths[index], recursive: true);
                }
            }
            catch
            {
                // Preserve the operation result. A later startup cleanup handles abandoned files.
            }
        }
    }

    private static void DeleteTemporaryPath(List<string> paths, string path)
    {
        paths.Remove(path);
        try
        {
            if (File.Exists(path))
                File.Delete(path);
            else if (Directory.Exists(path))
                Directory.Delete(path, recursive: true);
        }
        catch
        {
            paths.Add(path);
        }
    }
}

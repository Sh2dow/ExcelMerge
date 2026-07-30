#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace ExcelMerge;

public sealed class WorkbookMergeException : Exception
{
    public WorkbookMergeException(string message) : base(message)
    {
    }
}

public sealed class WorkbookMergeService
{
    private static readonly XNamespace SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace OfficeRelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace ContentTypeNamespace =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    public void Save(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook,
        IReadOnlyDictionary<MergeRowKey, MergeResolution> resolutions,
        string outputPath)
    {
        SaveCore(baseWorkbook, localWorkbook, remoteWorkbook, resolutions, null, outputPath, null, null);
    }

    public void Save(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook,
        IReadOnlyDictionary<MergeCellKey, MergeCellResolution> resolutions,
        string outputPath)
    {
        SaveCore(baseWorkbook, localWorkbook, remoteWorkbook, null, resolutions, outputPath, null, null);
    }

    public void Save(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook,
        IReadOnlyDictionary<MergeRowKey, MergeResolution> rowResolutions,
        IReadOnlyDictionary<MergeCellKey, MergeCellResolution> cellResolutions,
        string outputPath)
    {
        SaveCore(baseWorkbook, localWorkbook, remoteWorkbook, rowResolutions, cellResolutions, outputPath, null, null);
    }

    public void Save(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook,
        IReadOnlyDictionary<MergeRowKey, MergeResolution> rowResolutions,
        IReadOnlyDictionary<MergeCellKey, MergeCellResolution> cellResolutions,
        string localSourcePath,
        string remoteSourcePath,
        string outputPath)
    {
        SaveCore(
            baseWorkbook,
            localWorkbook,
            remoteWorkbook,
            rowResolutions,
            cellResolutions,
            outputPath,
            localSourcePath,
            remoteSourcePath);
    }

    private static void SaveCore(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook,
        IReadOnlyDictionary<MergeRowKey, MergeResolution>? rowResolutions,
        IReadOnlyDictionary<MergeCellKey, MergeCellResolution>? cellResolutions,
        string outputPath,
        string? localSourcePath,
        string? remoteSourcePath)
    {
        var extension = Path.GetExtension(outputPath);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            throw new WorkbookMergeException("RESULT must use the .xlsx or .xls extension.");

        var unresolved = new HashSet<MergeCellKey>();
        var results = new List<MergedSheetResult>();
        ValidateSheetRenameAmbiguity(baseWorkbook, localWorkbook, remoteWorkbook);
        foreach (var group in CreateSheetGroups(baseWorkbook, localWorkbook, remoteWorkbook))
        {
            if (ShouldOmitDeletedSheet(group))
                continue;
            if (group.Base != null && (group.Local == null || group.Remote == null))
                throw new WorkbookMergeException($"Sheet deletion combined with modifications is not supported safely: '{group.ResultName}'.");
            results.Add(MergeSheet(group, rowResolutions, cellResolutions, unresolved));
        }

        if (unresolved.Count > 0)
        {
            var locations = string.Join(", ", unresolved
                .OrderBy(key => key.SheetName)
                .ThenBy(key => key.RowIndex)
                .ThenBy(key => key.ColumnIndex)
                .Take(10)
                .Select(key => $"{key.SheetName}!{ColumnName(key.ColumnIndex)}{key.RowIndex + 1}"));
            throw new WorkbookMergeException($"Resolve all conflicts before saving RESULT: {locations}");
        }

        ValidateOutputLimits(results, extension.Equals(".xls", StringComparison.OrdinalIgnoreCase));

        var localKind = localSourcePath == null ? ExcelFileKind.Unknown : GetExcelFileKind(localSourcePath);
        var remoteKind = remoteSourcePath == null ? ExcelFileKind.Unknown : GetExcelFileKind(remoteSourcePath);
        if (localSourcePath != null
            && remoteSourcePath != null
            && localKind != ExcelFileKind.Unknown
            && remoteKind != ExcelFileKind.Unknown)
        {
            WritePreservingWorkbook(
                localWorkbook,
                remoteWorkbook,
                results,
                localKind,
                remoteKind,
                localSourcePath,
                remoteSourcePath,
                outputPath);
            return;
        }

        using IWorkbook outputWorkbook = extension.Equals(".xls", StringComparison.OrdinalIgnoreCase)
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var result in results)
        {
            var sheet = outputWorkbook.CreateSheet(CreateUniqueSheetName(result.Name, usedNames));
            foreach (var rowValue in result.Rows.OrderBy(pair => pair.Key))
            {
                var outputRow = sheet.CreateRow(rowValue.Key);
                foreach (var cellValue in rowValue.Value.OrderBy(pair => pair.Key))
                {
                    if (!string.IsNullOrEmpty(cellValue.Value.Value))
                        outputRow.CreateCell(cellValue.Key).SetCellValue(cellValue.Value.Value);
                }
            }

            foreach (var region in result.MergedRegions)
            {
                sheet.AddMergedRegion(new CellRangeAddress(
                    region.FirstRow,
                    region.LastRow,
                    region.FirstColumn,
                    region.LastColumn));
            }
        }

        if (outputWorkbook.NumberOfSheets == 0)
            outputWorkbook.CreateSheet("Result");

        using var output = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        outputWorkbook.Write(output);
    }

    private static ExcelFileKind GetExcelFileKind(string path)
    {
        if (!File.Exists(path))
            return ExcelFileKind.Unknown;
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            Span<byte> signature = stackalloc byte[8];
            if (stream.Read(signature) < signature.Length)
                return ExcelFileKind.Unknown;
            if (signature[0] == 0x50 && signature[1] == 0x4B)
            {
                stream.Position = 0;
                using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
                return archive.GetEntry("[Content_Types].xml") != null
                    && archive.GetEntry("xl/workbook.xml") != null
                    ? ExcelFileKind.Xlsx
                    : ExcelFileKind.Unknown;
            }

            ReadOnlySpan<byte> oleSignature = stackalloc byte[]
                { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            return signature.SequenceEqual(oleSignature) ? ExcelFileKind.Xls : ExcelFileKind.Unknown;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or UnauthorizedAccessException)
        {
            return ExcelFileKind.Unknown;
        }
    }

    private static void WritePreservingWorkbook(
        ExcelWorkbook localModel,
        ExcelWorkbook remoteModel,
        IReadOnlyList<MergedSheetResult> results,
        ExcelFileKind localKind,
        ExcelFileKind remoteKind,
        string localSourcePath,
        string remoteSourcePath,
        string outputPath)
    {
        if (!File.Exists(localSourcePath))
            throw new WorkbookMergeException($"LOCAL source file no longer exists: {localSourcePath}");
        if (!File.Exists(remoteSourcePath))
            throw new WorkbookMergeException($"REMOTE source file no longer exists: {remoteSourcePath}");
        var outputKind = Path.GetExtension(outputPath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            ? ExcelFileKind.Xlsx
            : ExcelFileKind.Xls;
        if (localKind != outputKind)
            throw new WorkbookMergeException("RESULT must use the same Excel extension as LOCAL to preserve workbook formatting.");

        if (CanPatchXlsxPackage(localModel, remoteModel, results, localKind, remoteKind))
        {
            WritePatchedXlsxPackage(
                localModel,
                remoteModel,
                results,
                localSourcePath,
                remoteSourcePath,
                outputPath);
            return;
        }

        var outputFullPath = Path.GetFullPath(outputPath);
        var outputDirectory = Path.GetDirectoryName(outputFullPath)
            ?? throw new WorkbookMergeException("RESULT must have a valid parent directory.");
        var writePath = Path.Combine(
            outputDirectory,
            $".excelmerge-preserved-{Guid.NewGuid():N}{Path.GetExtension(outputPath)}");

        try
        {
            using (var outputWorkbook = WorkbookFactory.Create(localSourcePath))
            using (var remoteWorkbook = WorkbookFactory.Create(remoteSourcePath))
            {
                var remoteStyles = new Dictionary<int, ICellStyle>();
                foreach (var result in results)
                {
                    var outputSheet = GetPhysicalSheet(outputWorkbook, localModel, result.Name, results.Count == 1);
                    var remoteSheet = GetPhysicalSheet(remoteWorkbook, remoteModel, result.Name, results.Count == 1);
                    if (outputSheet == null)
                    {
                        if (remoteSheet == null)
                            throw new WorkbookMergeException($"Cannot find source sheet '{result.Name}' in LOCAL or REMOTE.");
                        var usedNames = Enumerable.Range(0, outputWorkbook.NumberOfSheets)
                            .Select(index => outputWorkbook.GetSheetName(index))
                            .ToHashSet(StringComparer.OrdinalIgnoreCase);
                        outputSheet = outputWorkbook.CreateSheet(CreateUniqueSheetName(result.Name, usedNames));
                        CopySheet(remoteSheet, outputSheet, outputWorkbook, remoteStyles);
                        outputWorkbook.SetSheetVisibility(
                            outputWorkbook.GetSheetIndex(outputSheet),
                            remoteWorkbook.GetSheetVisibility(remoteWorkbook.GetSheetIndex(remoteSheet)));
                    }

                    ApplyDuplicatedRows(
                        outputSheet,
                        remoteSheet,
                        result.DuplicatedRows,
                        outputWorkbook,
                        remoteStyles);
                    ApplyCellPatches(outputSheet, remoteSheet, result, outputWorkbook, remoteStyles);
                    SynchronizeMergedRegions(outputSheet, result.MergedRegions);
                }

                var retainedLocalSheets = results
                    .Where(result => localModel.Sheets.ContainsKey(result.Name))
                    .Select(result => result.Name)
                    .ToHashSet(StringComparer.Ordinal);
                foreach (var name in localModel.Sheets.Keys.Where(name => !retainedLocalSheets.Contains(name)).ToList())
                {
                    var index = outputWorkbook.GetSheetIndex(name);
                    if (index >= 0)
                        outputWorkbook.RemoveSheetAt(index);
                }

                if (outputWorkbook.NumberOfSheets == 0)
                    outputWorkbook.CreateSheet("Result");

                using var output = new FileStream(writePath, FileMode.Create, FileAccess.Write, FileShare.None);
                outputWorkbook.Write(output);
            }
            if (outputKind == ExcelFileKind.Xlsx)
                RestoreLocalThemeParts(localSourcePath, writePath);

            File.Move(writePath, outputFullPath, true);
        }
        finally
        {
            if (File.Exists(writePath))
                File.Delete(writePath);
        }
    }

    private static void RestoreLocalThemeParts(string localSourcePath, string outputPath)
    {
        using var localPackage = ZipFile.OpenRead(localSourcePath);
        var localThemes = localPackage.Entries
            .Where(entry => entry.FullName.StartsWith("xl/theme/", StringComparison.Ordinal))
            .ToList();
        if (localThemes.Count == 0)
            return;

        using var outputPackage = ZipFile.Open(outputPath, ZipArchiveMode.Update);
        foreach (var source in localThemes)
        {
            var existing = outputPackage.GetEntry(source.FullName);
            existing?.Delete();
            var replacement = outputPackage.CreateEntry(source.FullName, CompressionLevel.Optimal);
            replacement.LastWriteTime = source.LastWriteTime;
            replacement.ExternalAttributes = source.ExternalAttributes;
            using var input = source.Open();
            using var output = replacement.Open();
            input.CopyTo(output);
        }
    }

    private static bool CanPatchXlsxPackage(
        ExcelWorkbook localModel,
        ExcelWorkbook remoteModel,
        IReadOnlyList<MergedSheetResult> results,
        ExcelFileKind localKind,
        ExcelFileKind remoteKind)
    {
        return localKind == ExcelFileKind.Xlsx
            && remoteKind == ExcelFileKind.Xlsx
            && results.Count == localModel.Sheets.Count
            && results.All(result => localModel.Sheets.ContainsKey(result.Name))
            && results.All(result => remoteModel.Sheets.ContainsKey(result.Name)
                || results.Count == 1 && remoteModel.Sheets.Count == 1)
            && results.All(result => result.DuplicatedRows.Count == 0);
    }

    private static void WritePatchedXlsxPackage(
        ExcelWorkbook localModel,
        ExcelWorkbook remoteModel,
        IReadOnlyList<MergedSheetResult> results,
        string localSourcePath,
        string remoteSourcePath,
        string outputPath)
    {
        var localEntries = ReadPackage(localSourcePath);
        var remoteEntries = ReadRemotePackage(remoteSourcePath, remoteModel, results);
        var localEntryMap = localEntries.ToDictionary(entry => entry.Name, StringComparer.Ordinal);
        var localSheets = GetWorksheetParts(localEntryMap);
        var remoteSheets = GetWorksheetParts(remoteEntries);
        var remoteSharedStrings = GetSharedStrings(remoteEntries);
        var stylesCompatible = localEntryMap.TryGetValue("xl/styles.xml", out var localStyles)
            && remoteEntries.TryGetValue("xl/styles.xml", out var remoteStyles)
            && localStyles.Content.SequenceEqual(remoteStyles.Content);
        var replacements = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var requiresFullRecalculation = false;

        foreach (var result in results)
        {
            if (!localSheets.TryGetValue(result.Name, out var localPart))
                throw new WorkbookMergeException($"Cannot find LOCAL worksheet part for '{result.Name}'.");
            var remotePart = GetWorksheetPart(remoteSheets, remoteModel, result.Name, results.Count == 1);
            if (remotePart == null || !remoteEntries.TryGetValue(remotePart, out var remoteEntry))
                throw new WorkbookMergeException($"Cannot find REMOTE worksheet part for '{result.Name}'.");

            var localDocument = ParseXml(localEntryMap[localPart].Content);
            var remoteDocument = ParseXml(remoteEntry.Content);
            var cellsChanged = ApplyOpenXmlCellPatches(
                localDocument,
                remoteDocument,
                result,
                remoteSharedStrings,
                stylesCompatible,
                out _);
            var mergedRegionsChanged = SynchronizeOpenXmlMergedRegions(localDocument, result.MergedRegions);
            requiresFullRecalculation |= cellsChanged;
            if (cellsChanged || mergedRegionsChanged)
                replacements[localPart] = SerializeXml(localDocument);
        }

        if (requiresFullRecalculation)
            PrepareFullRecalculation(localEntries, localEntryMap, replacements);

        var outputFullPath = Path.GetFullPath(outputPath);
        var outputDirectory = Path.GetDirectoryName(outputFullPath)
            ?? throw new WorkbookMergeException("RESULT must have a valid parent directory.");
        var writePath = Path.Combine(outputDirectory, $".excelmerge-patched-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var destination = ZipFile.Open(writePath, ZipArchiveMode.Create))
            {
                foreach (var sourceEntry in localEntries)
                {
                    var entry = destination.CreateEntry(sourceEntry.Name, CompressionLevel.Optimal);
                    entry.LastWriteTime = sourceEntry.LastWriteTime;
                    entry.ExternalAttributes = sourceEntry.ExternalAttributes;
                    using var output = entry.Open();
                    output.Write(replacements.GetValueOrDefault(sourceEntry.Name, sourceEntry.Content));
                }
            }
            if (File.Exists(outputFullPath))
                File.Replace(writePath, outputFullPath, null);
            else
                File.Move(writePath, outputFullPath);
        }
        finally
        {
            if (File.Exists(writePath))
                File.Delete(writePath);
        }
    }

    private static List<XlsxPackageEntry> ReadPackage(string path)
    {
        var entries = new List<XlsxPackageEntry>();
        using var archive = ZipFile.OpenRead(path);
        foreach (var entry in archive.Entries)
            entries.Add(ReadPackageEntry(entry));
        return entries;
    }

    private static Dictionary<string, XlsxPackageEntry> ReadRemotePackage(
        string path,
        ExcelWorkbook remoteModel,
        IReadOnlyList<MergedSheetResult> results)
    {
        using var archive = ZipFile.OpenRead(path);
        var entries = new Dictionary<string, XlsxPackageEntry>(StringComparer.Ordinal);

        void AddEntry(string name, bool required)
        {
            if (entries.ContainsKey(name))
                return;
            var entry = archive.GetEntry(name);
            if (entry == null)
            {
                if (required)
                    throw new WorkbookMergeException($"Required workbook part is missing: {name}.");
                return;
            }
            entries.Add(name, ReadPackageEntry(entry));
        }

        AddEntry("xl/workbook.xml", true);
        AddEntry("xl/_rels/workbook.xml.rels", true);
        AddEntry("xl/styles.xml", false);
        AddEntry("xl/sharedStrings.xml", false);
        var worksheetParts = GetWorksheetParts(entries);
        foreach (var result in results)
        {
            var part = GetWorksheetPart(worksheetParts, remoteModel, result.Name, results.Count == 1);
            if (part != null)
                AddEntry(part, true);
        }
        return entries;
    }

    private static XlsxPackageEntry ReadPackageEntry(ZipArchiveEntry entry)
    {
        if (entry.Length > int.MaxValue)
            throw new WorkbookMergeException($"Workbook part is too large to process: {entry.FullName}.");
        var content = new byte[(int)entry.Length];
        using var input = entry.Open();
        input.ReadExactly(content);
        return new XlsxPackageEntry(
            entry.FullName,
            entry.LastWriteTime,
            entry.ExternalAttributes,
            content);
    }

    private static Dictionary<string, string> GetWorksheetParts(
        IReadOnlyDictionary<string, XlsxPackageEntry> entries)
    {
        var workbook = ParseXml(entries["xl/workbook.xml"].Content);
        var relationships = ParseXml(entries["xl/_rels/workbook.xml.rels"].Content)
            .Root!
            .Elements(PackageRelationshipNamespace + "Relationship")
            .Where(relationship => ((string?)relationship.Attribute("Type"))?.EndsWith(
                    "/worksheet",
                    StringComparison.Ordinal) == true
                && !string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.Ordinal))
            .ToDictionary(
                relationship => (string)relationship.Attribute("Id")!,
                relationship => ResolvePackageTarget((string)relationship.Attribute("Target")!),
                StringComparer.Ordinal);
        return workbook.Descendants(SpreadsheetNamespace + "sheet").ToDictionary(
            sheet => (string)sheet.Attribute("name")!,
            sheet => relationships[(string)sheet.Attribute(OfficeRelationshipNamespace + "id")!],
            StringComparer.Ordinal);
    }

    private static string? GetWorksheetPart(
        IReadOnlyDictionary<string, string> parts,
        ExcelWorkbook model,
        string resultName,
        bool allowSoleSheetFallback)
    {
        if (parts.TryGetValue(resultName, out var part))
            return part;
        if (allowSoleSheetFallback && model.Sheets.Count == 1)
            return parts.Values.Single();
        return null;
    }

    private static string ResolvePackageTarget(string target)
    {
        var workbookUri = new Uri("http://excelmerge.local/xl/workbook.xml", UriKind.Absolute);
        var resolved = new Uri(workbookUri, target.Replace('\\', '/'));
        return Uri.UnescapeDataString(resolved.AbsolutePath.TrimStart('/'));
    }

    private static IReadOnlyList<XElement> GetSharedStrings(
        IReadOnlyDictionary<string, XlsxPackageEntry> entries)
    {
        return entries.TryGetValue("xl/sharedStrings.xml", out var entry)
            ? ParseXml(entry.Content).Root!.Elements(SpreadsheetNamespace + "si").ToList()
            : Array.Empty<XElement>();
    }

    private static bool ApplyOpenXmlCellPatches(
        XDocument localDocument,
        XDocument remoteDocument,
        MergedSheetResult result,
        IReadOnlyList<XElement> remoteSharedStrings,
        bool stylesCompatible,
        out bool formulasChanged)
    {
        var changed = false;
        formulasChanged = false;
        var localIndex = OpenXmlWorksheetIndex.Create(localDocument);
        var remoteCells = remoteDocument.Descendants(SpreadsheetNamespace + "c")
            .ToDictionary(cell => (string)cell.Attribute("r")!, StringComparer.Ordinal);
        foreach (var (rowIndex, cells) in result.Rows)
        {
            foreach (var (columnIndex, cellResult) in cells)
            {
                if (cellResult.Source == CellSource.Local)
                    continue;
                changed = true;

                var outputAddress = ColumnName(columnIndex) + (rowIndex + 1);
                var (outputCell, created) = GetOrCreateOpenXmlCell(localIndex, outputAddress, rowIndex, columnIndex);
                if (cellResult.Source == CellSource.Custom)
                {
                    ValidateStandaloneFormula(outputCell.Element(SpreadsheetNamespace + "f"), outputAddress);
                    if (outputCell.Attributes().Any(attribute => attribute.Name.LocalName is "cm" or "vm"))
                    {
                        throw new WorkbookMergeException(
                            $"Cell metadata references cannot be merged safely yet: {outputAddress}.");
                    }
                    formulasChanged |= outputCell.Element(SpreadsheetNamespace + "f") != null;
                    SetOpenXmlInlineString(outputCell, outputAddress, cellResult.Value);
                    continue;
                }

                var remoteAddress = ColumnName(columnIndex) + (cellResult.SourceRowIndex + 1);
                remoteCells.TryGetValue(remoteAddress, out var remoteCell);
                formulasChanged |= CopyOpenXmlCellPayload(
                    remoteCell,
                    outputCell,
                    outputAddress,
                    created,
                    stylesCompatible,
                    remoteSharedStrings);
            }
        }
        return changed;
    }

    private static (XElement Cell, bool Created) GetOrCreateOpenXmlCell(
        OpenXmlWorksheetIndex index,
        string address,
        int rowIndex,
        int columnIndex)
    {
        if (index.Cells.TryGetValue(address, out var existing))
            return (existing, false);

        var rowNumber = rowIndex + 1;
        if (!index.Rows.TryGetValue(rowNumber, out var rowIndexEntry))
        {
            var row = new XElement(SpreadsheetNamespace + "row", new XAttribute("r", rowNumber));
            var followingRows = rowNumber < int.MaxValue
                ? index.RowNumbers.GetViewBetween(rowNumber + 1, int.MaxValue)
                : null;
            if (followingRows == null || followingRows.Count == 0)
                index.SheetData.Add(row);
            else
                index.Rows[followingRows.Min].Row.AddBeforeSelf(row);
            rowIndexEntry = new OpenXmlRowIndex(row);
            index.Rows.Add(rowNumber, rowIndexEntry);
            index.RowNumbers.Add(rowNumber);
        }

        var cell = new XElement(SpreadsheetNamespace + "c", new XAttribute("r", address));
        var followingColumns = columnIndex < int.MaxValue
            ? rowIndexEntry.ColumnNumbers.GetViewBetween(columnIndex + 1, int.MaxValue)
            : null;
        if (followingColumns == null || followingColumns.Count == 0)
            rowIndexEntry.Row.Add(cell);
        else
            rowIndexEntry.Cells[followingColumns.Min].AddBeforeSelf(cell);
        index.Cells.Add(address, cell);
        rowIndexEntry.Cells.TryAdd(columnIndex, cell);
        rowIndexEntry.ColumnNumbers.Add(columnIndex);
        return (cell, true);
    }

    private static bool CopyOpenXmlCellPayload(
        XElement? source,
        XElement destination,
        string destinationAddress,
        bool destinationWasCreated,
        bool stylesCompatible,
        IReadOnlyList<XElement> sharedStrings)
    {
        var sourceFormula = source?.Element(SpreadsheetNamespace + "f");
        var destinationFormula = destination.Element(SpreadsheetNamespace + "f");
        ValidateStandaloneFormula(sourceFormula, destinationAddress);
        ValidateStandaloneFormula(destinationFormula, destinationAddress);
        if (source?.Attributes().Any(attribute => attribute.Name.LocalName is "cm" or "vm") == true
            || destination.Attributes().Any(attribute => attribute.Name.LocalName is "cm" or "vm"))
        {
            throw new WorkbookMergeException(
                $"Cell metadata references cannot be merged safely yet: {destinationAddress}.");
        }

        var formulaChanged = sourceFormula != null || destinationFormula != null;
        var destinationStyle = (string?)destination.Attribute("s");
        destination.RemoveAttributes();
        destination.SetAttributeValue("r", destinationAddress);
        if (!destinationWasCreated && destinationStyle != null)
            destination.SetAttributeValue("s", destinationStyle);
        else if (stylesCompatible && source?.Attribute("s") is { } sourceStyle)
            destination.SetAttributeValue("s", sourceStyle.Value);
        else if (destinationWasCreated
            && source?.Attribute("s") is { Value: not "0" }
            && !stylesCompatible)
        {
            throw new WorkbookMergeException(
                $"A REMOTE-only styled cell cannot be mapped safely into LOCAL yet: {destinationAddress}.");
        }

        destination.RemoveNodes();
        if (source == null)
            return formulaChanged;

        var sourceType = (string?)source.Attribute("t");
        foreach (var attribute in source.Attributes().Where(attribute => attribute.Name.LocalName is not ("r" or "s" or "t")))
            destination.SetAttributeValue(attribute.Name, attribute.Value);
        if (sourceType == "s")
        {
            var value = source.Element(SpreadsheetNamespace + "v")?.Value;
            if (value == null || !int.TryParse(value, out var index) || index < 0 || index >= sharedStrings.Count)
                throw new WorkbookMergeException($"REMOTE shared string index is invalid at {destinationAddress}.");
            destination.SetAttributeValue("t", "inlineStr");
            destination.Add(new XElement(
                SpreadsheetNamespace + "is",
                sharedStrings[index].Elements().Select(element => new XElement(element))));
            return formulaChanged;
        }

        if (sourceType != null)
            destination.SetAttributeValue("t", sourceType);
        destination.Add(source.Elements()
            .Where(element => element.Name != SpreadsheetNamespace + "v" || sourceFormula == null)
            .Select(element => new XElement(element)));
        return formulaChanged;
    }

    private static void ValidateStandaloneFormula(XElement? formula, string address)
    {
        var formulaType = (string?)formula?.Attribute("t");
        if (formulaType != null && !formulaType.Equals("normal", StringComparison.Ordinal))
        {
            throw new WorkbookMergeException(
                $"Shared, array, and data-table formulas cannot be patched safely yet: {address}.");
        }
    }

    private static void SetOpenXmlInlineString(XElement cell, string address, string value)
    {
        var style = (string?)cell.Attribute("s");
        cell.RemoveAttributes();
        cell.SetAttributeValue("r", address);
        if (style != null)
            cell.SetAttributeValue("s", style);
        cell.SetAttributeValue("t", "inlineStr");
        cell.ReplaceNodes(new XElement(
            SpreadsheetNamespace + "is",
            new XElement(
                SpreadsheetNamespace + "t",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                value)));
    }

    private static void PrepareFullRecalculation(
        ICollection<XlsxPackageEntry> localEntries,
        IReadOnlyDictionary<string, XlsxPackageEntry> localEntryMap,
        IDictionary<string, byte[]> replacements)
    {
        var workbook = ParseXml(localEntryMap["xl/workbook.xml"].Content);
        var root = workbook.Root ?? throw new WorkbookMergeException("Workbook XML has no root element.");
        var calculation = root.Element(SpreadsheetNamespace + "calcPr");
        if (calculation == null)
        {
            calculation = new XElement(SpreadsheetNamespace + "calcPr");
            var laterNames = new HashSet<XName>
            {
                SpreadsheetNamespace + "oleSize",
                SpreadsheetNamespace + "customWorkbookViews",
                SpreadsheetNamespace + "pivotCaches",
                SpreadsheetNamespace + "smartTagPr",
                SpreadsheetNamespace + "smartTagTypes",
                SpreadsheetNamespace + "webPublishing",
                SpreadsheetNamespace + "fileRecoveryPr",
                SpreadsheetNamespace + "webPublishObjects",
                SpreadsheetNamespace + "extLst",
            };
            var following = root.Elements().FirstOrDefault(element => laterNames.Contains(element.Name));
            if (following == null)
                root.Add(calculation);
            else
                following.AddBeforeSelf(calculation);
        }
        calculation.SetAttributeValue("calcMode", "auto");
        calculation.SetAttributeValue("fullCalcOnLoad", 1);
        calculation.SetAttributeValue("forceFullCalc", 1);
        replacements["xl/workbook.xml"] = SerializeXml(workbook);

        var relationshipsPath = "xl/_rels/workbook.xml.rels";
        var relationships = ParseXml(localEntryMap[relationshipsPath].Content);
        var calculationRelationships = relationships.Root!
            .Elements(PackageRelationshipNamespace + "Relationship")
            .Where(relationship => ((string?)relationship.Attribute("Type"))?.EndsWith(
                "/calcChain",
                StringComparison.Ordinal) == true)
            .ToList();
        var calculationParts = calculationRelationships
            .Select(relationship => ResolvePackageTarget((string)relationship.Attribute("Target")!))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var relationship in calculationRelationships)
            relationship.Remove();
        if (calculationRelationships.Count > 0)
            replacements[relationshipsPath] = SerializeXml(relationships);

        if (calculationParts.Count == 0)
            return;
        foreach (var entry in localEntries.Where(entry => calculationParts.Contains(entry.Name)).ToList())
            localEntries.Remove(entry);

        const string contentTypesPath = "[Content_Types].xml";
        var contentTypes = ParseXml(localEntryMap[contentTypesPath].Content);
        contentTypes.Root!
            .Elements(ContentTypeNamespace + "Override")
            .Where(overrideElement => calculationParts.Contains(
                ((string)overrideElement.Attribute("PartName")!).TrimStart('/')))
            .Remove();
        replacements[contentTypesPath] = SerializeXml(contentTypes);
    }

    private static bool SynchronizeOpenXmlMergedRegions(
        XDocument document,
        IReadOnlyList<ExcelMergedRegion> expectedRegions)
    {
        var mergeCells = document.Descendants(SpreadsheetNamespace + "mergeCells").FirstOrDefault();
        var existing = mergeCells?.Elements(SpreadsheetNamespace + "mergeCell")
            .Select(element => (string)element.Attribute("ref")!)
            .ToHashSet(StringComparer.Ordinal) ?? new HashSet<string>(StringComparer.Ordinal);
        var expected = expectedRegions.Select(FormatRegion).ToList();
        if (existing.SetEquals(expected))
            return false;
        if (expected.Count == 0)
        {
            mergeCells?.Remove();
            return true;
        }

        if (mergeCells == null)
        {
            mergeCells = new XElement(SpreadsheetNamespace + "mergeCells");
            var root = document.Root ?? throw new WorkbookMergeException("Worksheet XML has no root element.");
            var laterElementNames = new HashSet<XName>
            {
                SpreadsheetNamespace + "phoneticPr",
                SpreadsheetNamespace + "conditionalFormatting",
                SpreadsheetNamespace + "dataValidations",
                SpreadsheetNamespace + "hyperlinks",
                SpreadsheetNamespace + "printOptions",
                SpreadsheetNamespace + "pageMargins",
                SpreadsheetNamespace + "pageSetup",
                SpreadsheetNamespace + "headerFooter",
                SpreadsheetNamespace + "rowBreaks",
                SpreadsheetNamespace + "colBreaks",
                SpreadsheetNamespace + "customProperties",
                SpreadsheetNamespace + "cellWatches",
                SpreadsheetNamespace + "ignoredErrors",
                SpreadsheetNamespace + "smartTags",
                SpreadsheetNamespace + "drawing",
                SpreadsheetNamespace + "legacyDrawing",
                SpreadsheetNamespace + "legacyDrawingHF",
                SpreadsheetNamespace + "picture",
                SpreadsheetNamespace + "oleObjects",
                SpreadsheetNamespace + "controls",
                SpreadsheetNamespace + "webPublishItems",
                SpreadsheetNamespace + "tableParts",
                SpreadsheetNamespace + "extLst",
            };
            var following = root.Elements().FirstOrDefault(element => laterElementNames.Contains(element.Name));
            if (following == null)
                root.Add(mergeCells);
            else
                following.AddBeforeSelf(mergeCells);
        }
        mergeCells.RemoveNodes();
        mergeCells.SetAttributeValue("count", expected.Count);
        mergeCells.Add(expected.Select(region =>
            new XElement(SpreadsheetNamespace + "mergeCell", new XAttribute("ref", region))));
        return true;
    }

    private static string FormatRegion(ExcelMergedRegion region) =>
        $"{ColumnName(region.FirstColumn)}{region.FirstRow + 1}:{ColumnName(region.LastColumn)}{region.LastRow + 1}";

    private static int GetColumnIndex(string address)
    {
        var result = 0;
        for (var index = 0; index < address.Length && char.IsLetter(address[index]); index++)
            result = result * 26 + char.ToUpperInvariant(address[index]) - 'A' + 1;
        return result - 1;
    }

    private static XDocument ParseXml(byte[] content)
    {
        using var input = new MemoryStream(content);
        return XDocument.Load(input, LoadOptions.PreserveWhitespace);
    }

    private static byte[] SerializeXml(XDocument document)
    {
        using var output = new MemoryStream();
        using (var writer = new StreamWriter(output, new UTF8Encoding(false), leaveOpen: true))
            document.Save(writer, SaveOptions.DisableFormatting);
        return output.ToArray();
    }

    private static ISheet? GetPhysicalSheet(
        IWorkbook workbook,
        ExcelWorkbook model,
        string resultName,
        bool allowSoleSheetFallback)
    {
        if (model.Sheets.ContainsKey(resultName))
            return workbook.GetSheet(resultName);
        if (allowSoleSheetFallback && model.Sheets.Count == 1)
            return workbook.GetSheet(model.Sheets.Keys.Single()) ?? workbook.GetSheetAt(0);
        return null;
    }

    private static void ApplyDuplicatedRows(
        ISheet outputSheet,
        ISheet? remoteSheet,
        IReadOnlyList<int> duplicatedRows,
        IWorkbook outputWorkbook,
        IDictionary<int, ICellStyle> remoteStyles)
    {
        var shift = 0;
        foreach (var sourceRowIndex in duplicatedRows)
        {
            var outputRowIndex = sourceRowIndex + shift + 1;
            if (outputRowIndex <= outputSheet.LastRowNum)
                outputSheet.ShiftRows(outputRowIndex, outputSheet.LastRowNum, 1, true, false);

            var outputRow = outputSheet.GetRow(outputRowIndex) ?? outputSheet.CreateRow(outputRowIndex);
            var remoteRow = remoteSheet?.GetRow(sourceRowIndex);
            if (remoteRow != null && CopyRow(remoteRow, outputRow, outputWorkbook, remoteStyles))
                outputSheet.ForceFormulaRecalculation = true;
            shift++;
        }
    }

    private static void ApplyCellPatches(
        ISheet outputSheet,
        ISheet? remoteSheet,
        MergedSheetResult result,
        IWorkbook outputWorkbook,
        IDictionary<int, ICellStyle> remoteStyles)
    {
        foreach (var (rowIndex, cells) in result.Rows)
        {
            var outputRow = outputSheet.GetRow(rowIndex);
            if (outputRow == null && cells.Values.Any(cell => cell.Source != CellSource.Local))
            {
                outputRow = outputSheet.CreateRow(rowIndex);
                var remoteSourceRow = cells.Values
                    .Where(cell => cell.Source == CellSource.Remote)
                    .Select(cell => cell.SourceRowIndex)
                    .FirstOrDefault(-1);
                var remoteRow = remoteSourceRow >= 0 ? remoteSheet?.GetRow(remoteSourceRow) : null;
                if (remoteRow != null)
                    CopyRowFormatting(remoteRow, outputRow, outputWorkbook, remoteStyles);
            }
            if (outputRow == null)
                continue;

            foreach (var (columnIndex, cell) in cells)
            {
                if (cell.Source == CellSource.Local)
                    continue;

                var outputCell = outputRow.GetCell(columnIndex);
                if (cell.Source == CellSource.Custom)
                {
                    outputCell ??= outputRow.CreateCell(columnIndex);
                    outputCell.SetCellType(CellType.String);
                    outputCell.SetCellValue(cell.Value);
                    continue;
                }

                var remoteCell = remoteSheet?.GetRow(cell.SourceRowIndex)?.GetCell(columnIndex);
                if (remoteCell == null)
                {
                    outputCell?.SetCellType(CellType.Blank);
                    continue;
                }

                var created = outputCell == null;
                outputCell ??= outputRow.CreateCell(columnIndex);
                CopyCellValue(remoteCell, outputCell);
                if (created)
                    CopyCellStyle(remoteCell, outputCell, outputWorkbook, remoteStyles);
                if (remoteCell.CellType == CellType.Formula)
                    outputSheet.ForceFormulaRecalculation = true;
            }
        }
    }

    private static void CopySheet(
        ISheet source,
        ISheet destination,
        IWorkbook destinationWorkbook,
        IDictionary<int, ICellStyle> styleMap)
    {
        destination.DefaultColumnWidth = source.DefaultColumnWidth;
        destination.DefaultRowHeight = source.DefaultRowHeight;
        destination.DisplayGridlines = source.DisplayGridlines;
        destination.IsPrintGridlines = source.IsPrintGridlines;
        destination.FitToPage = source.FitToPage;
        destination.HorizontallyCenter = source.HorizontallyCenter;
        destination.VerticallyCenter = source.VerticallyCenter;
        destination.Autobreaks = source.Autobreaks;

        var maximumColumn = -1;
        var copiedFormula = false;
        for (var rowIndex = 0; rowIndex <= source.LastRowNum; rowIndex++)
        {
            var sourceRow = source.GetRow(rowIndex);
            if (sourceRow == null)
                continue;
            var destinationRow = destination.CreateRow(rowIndex);
            copiedFormula |= CopyRow(sourceRow, destinationRow, destinationWorkbook, styleMap);
            maximumColumn = Math.Max(maximumColumn, sourceRow.LastCellNum - 1);
        }

        for (var column = 0; column <= maximumColumn; column++)
        {
            destination.SetColumnWidth(column, source.GetColumnWidth(column));
            destination.SetColumnHidden(column, source.IsColumnHidden(column));
            var sourceStyle = source.GetColumnStyle(column);
            if (sourceStyle != null && sourceStyle.Index != 0)
                destination.SetDefaultColumnStyle(column, MapStyle(sourceStyle, destinationWorkbook, styleMap));
        }

        for (var index = 0; index < source.NumMergedRegions; index++)
        {
            var region = source.GetMergedRegion(index);
            destination.AddMergedRegion(new CellRangeAddress(
                region.FirstRow,
                region.LastRow,
                region.FirstColumn,
                region.LastColumn));
        }

        var pane = source.PaneInformation;
        if (pane != null && pane.IsFreezePane())
        {
            destination.CreateFreezePane(
                pane.VerticalSplitPosition,
                pane.HorizontalSplitPosition,
                pane.VerticalSplitLeftColumn,
                pane.HorizontalSplitTopRow);
        }
        if (copiedFormula)
            destination.ForceFormulaRecalculation = true;
    }

    private static bool CopyRow(
        IRow source,
        IRow destination,
        IWorkbook destinationWorkbook,
        IDictionary<int, ICellStyle> styleMap)
    {
        CopyRowFormatting(source, destination, destinationWorkbook, styleMap);
        var copiedFormula = false;
        for (var column = source.FirstCellNum; column >= 0 && column < source.LastCellNum; column++)
        {
            var sourceCell = source.GetCell(column);
            if (sourceCell == null)
                continue;
            var destinationCell = destination.CreateCell(column);
            CopyCellValue(sourceCell, destinationCell);
            CopyCellStyle(sourceCell, destinationCell, destinationWorkbook, styleMap);
            copiedFormula |= sourceCell.CellType == CellType.Formula;
        }
        return copiedFormula;
    }

    private static void CopyRowFormatting(
        IRow source,
        IRow destination,
        IWorkbook destinationWorkbook,
        IDictionary<int, ICellStyle> styleMap)
    {
        destination.Height = source.Height;
        destination.ZeroHeight = source.ZeroHeight;
        if (source.IsFormatted && source.RowStyle != null)
            destination.RowStyle = MapStyle(source.RowStyle, destinationWorkbook, styleMap);
    }

    private static void CopyCellValue(ICell source, ICell destination)
    {
        if (destination.CellType == CellType.Formula && source.CellType != CellType.Formula)
            destination.SetCellType(CellType.Blank);
        switch (source.CellType)
        {
            case CellType.Blank:
                destination.SetCellType(CellType.Blank);
                break;
            case CellType.Boolean:
                destination.SetCellValue(source.BooleanCellValue);
                break;
            case CellType.Error:
                destination.SetCellErrorValue(source.ErrorCellValue);
                break;
            case CellType.Formula:
                destination.SetCellFormula(source.CellFormula);
                break;
            case CellType.Numeric:
                destination.SetCellValue(source.NumericCellValue);
                break;
            case CellType.String:
                destination.SetCellValue(source.StringCellValue);
                break;
            default:
                destination.SetCellType(CellType.Blank);
                break;
        }
    }

    private static void CopyCellStyle(
        ICell source,
        ICell destination,
        IWorkbook destinationWorkbook,
        IDictionary<int, ICellStyle> styleMap)
    {
        if (source.CellStyle == null || source.CellStyle.Index == 0)
            return;
        destination.CellStyle = MapStyle(source.CellStyle, destinationWorkbook, styleMap);
    }

    private static ICellStyle MapStyle(
        ICellStyle source,
        IWorkbook destinationWorkbook,
        IDictionary<int, ICellStyle> styleMap)
    {
        if (styleMap.TryGetValue(source.Index, out var mapped))
            return mapped;
        mapped = destinationWorkbook.CreateCellStyle();
        mapped.CloneStyleFrom(source);
        styleMap[source.Index] = mapped;
        return mapped;
    }

    private static void SynchronizeMergedRegions(
        ISheet sheet,
        IReadOnlyList<ExcelMergedRegion> expectedRegions)
    {
        var existing = Enumerable.Range(0, sheet.NumMergedRegions)
            .Select(index => sheet.GetMergedRegion(index))
            .Select(region => new ExcelMergedRegion(
                region.FirstRow,
                region.LastRow,
                region.FirstColumn,
                region.LastColumn))
            .ToHashSet();
        if (existing.SetEquals(expectedRegions))
            return;

        for (var index = sheet.NumMergedRegions - 1; index >= 0; index--)
            sheet.RemoveMergedRegion(index);
        foreach (var region in expectedRegions)
        {
            sheet.AddMergedRegion(new CellRangeAddress(
                region.FirstRow,
                region.LastRow,
                region.FirstColumn,
                region.LastColumn));
        }
    }

    private static MergedSheetResult MergeSheet(
        SheetGroup group,
        IReadOnlyDictionary<MergeRowKey, MergeResolution>? rowResolutions,
        IReadOnlyDictionary<MergeCellKey, MergeCellResolution>? cellResolutions,
        ISet<MergeCellKey> unresolved)
    {
        var baseSnapshot = SheetSnapshot.Create(group.Base);
        var localSnapshot = SheetSnapshot.Create(group.Local);
        var remoteSnapshot = SheetSnapshot.Create(group.Remote);
        var rowIndices = baseSnapshot.Rows.Keys
            .Concat(localSnapshot.Rows.Keys)
            .Concat(remoteSnapshot.Rows.Keys)
            .Distinct()
            .OrderBy(index => index)
            .ToList();
        var rows = new SortedDictionary<int, Dictionary<int, MergedCellResult>>();
        var duplicatedRows = new List<int>();
        var shift = 0;

        foreach (var rowIndex in rowIndices)
        {
            var columns = baseSnapshot.GetColumns(rowIndex)
                .Concat(localSnapshot.GetColumns(rowIndex))
                .Concat(remoteSnapshot.GetColumns(rowIndex))
                .Distinct()
                .OrderBy(index => index)
                .ToList();
            var rowKey = new MergeRowKey(group.ResultName, rowIndex);
            var rowResolution = rowResolutions != null && rowResolutions.TryGetValue(rowKey, out var selected)
                ? selected
                : MergeResolution.Unresolved;
            var cellStates = new List<CellMergeState>(columns.Count);
            var conflictCount = 0;
            foreach (var column in columns)
            {
                var baseValue = baseSnapshot.GetValue(rowIndex, column);
                var localValue = localSnapshot.GetValue(rowIndex, column);
                var remoteValue = remoteSnapshot.GetValue(rowIndex, column);
                var conflict = IsConflict(baseValue, localValue, remoteValue);
                var resolution = conflict
                    ? rowResolution != MergeResolution.Unresolved
                        ? new MergeCellResolution(rowResolution)
                        : GetCellResolution(
                            cellResolutions,
                            new MergeCellKey(group.ResultName, rowIndex, column))
                    : default;
                if (conflict)
                    conflictCount++;
                cellStates.Add(new CellMergeState(
                    column,
                    baseValue,
                    localValue,
                    remoteValue,
                    conflict,
                    resolution));
            }
            var keepBoth = cellStates.Any(state =>
                state.IsConflict && state.Resolution.Resolution == MergeResolution.Both);

            if (!keepBoth)
            {
                foreach (var state in cellStates.Where(state =>
                    state.IsConflict && state.Resolution.Resolution == MergeResolution.Unresolved))
                {
                    unresolved.Add(new MergeCellKey(group.ResultName, rowIndex, state.Column));
                }
            }

            var outputRowIndex = rowIndex + shift;
            if (keepBoth)
            {
                rows[outputRowIndex] = localSnapshot.CopyRow(rowIndex, CellSource.Local);
                rows[outputRowIndex + 1] = remoteSnapshot.CopyRow(rowIndex, CellSource.Remote);
                duplicatedRows.Add(rowIndex);
                shift++;
                continue;
            }

            if (conflictCount > 0 && rowResolution == MergeResolution.Local)
            {
                rows[outputRowIndex] = CopySelectedRow(localSnapshot, rowIndex, columns, CellSource.Local);
                continue;
            }
            if (conflictCount > 0 && rowResolution == MergeResolution.Remote)
            {
                rows[outputRowIndex] = CopySelectedRow(remoteSnapshot, rowIndex, columns, CellSource.Remote);
                continue;
            }

            var mergedRow = new Dictionary<int, MergedCellResult>();
            foreach (var state in cellStates)
            {
                mergedRow[state.Column] = MergeValue(
                    state.BaseValue,
                    state.LocalValue,
                    state.RemoteValue,
                    state.Resolution,
                    rowIndex);
            }
            rows[outputRowIndex] = mergedRow;
        }

        var regions = MergeRegions(baseSnapshot.MergedRegions, localSnapshot.MergedRegions, remoteSnapshot.MergedRegions);
        var orderedDuplicatedRows = duplicatedRows.OrderBy(row => row).ToArray();
        if (regions.Any(region => region.FirstRow != region.LastRow
            && ContainsValueInRange(orderedDuplicatedRows, region.FirstRow, region.LastRow)))
            throw new WorkbookMergeException($"KEEP BOTH cannot safely duplicate a row inside a vertical merged range in sheet '{group.ResultName}'.");
        var transformedRegions = TransformRegions(regions, orderedDuplicatedRows);
        ValidateRegions(transformedRegions, group.ResultName);
        return new MergedSheetResult(group.ResultName, rows, transformedRegions, duplicatedRows);
    }

    private static Dictionary<int, MergedCellResult> CopySelectedRow(
        SheetSnapshot snapshot,
        int rowIndex,
        IEnumerable<int> columns,
        CellSource source)
    {
        return columns.ToDictionary(
            column => column,
            column => new MergedCellResult(snapshot.GetValue(rowIndex, column), source, rowIndex));
    }

    private static MergedCellResult MergeValue(
        string baseValue,
        string localValue,
        string remoteValue,
        MergeCellResolution resolution,
        int sourceRowIndex = -1)
    {
        if (localValue == remoteValue)
            return new MergedCellResult(localValue, CellSource.Local, sourceRowIndex);
        if (localValue == baseValue)
            return new MergedCellResult(remoteValue, CellSource.Remote, sourceRowIndex);
        if (remoteValue == baseValue)
            return new MergedCellResult(localValue, CellSource.Local, sourceRowIndex);
        return resolution.Resolution switch
        {
            MergeResolution.Remote => new MergedCellResult(remoteValue, CellSource.Remote, sourceRowIndex),
            MergeResolution.Custom => new MergedCellResult(
                resolution.CustomValue ?? string.Empty,
                CellSource.Custom,
                sourceRowIndex),
            _ => new MergedCellResult(localValue, CellSource.Local, sourceRowIndex),
        };
    }

    private static MergeCellResolution GetCellResolution(
        IReadOnlyDictionary<MergeCellKey, MergeCellResolution>? resolutions,
        MergeCellKey key)
    {
        return resolutions != null && resolutions.TryGetValue(key, out var resolution)
            ? resolution
            : new MergeCellResolution(MergeResolution.Unresolved);
    }

    private static bool IsConflict(string baseValue, string localValue, string remoteValue)
    {
        return localValue != remoteValue
            && localValue != baseValue
            && remoteValue != baseValue;
    }

    private static string ColumnName(int index)
    {
        var name = string.Empty;
        do
        {
            name = (char)('A' + index % 26) + name;
            index = index / 26 - 1;
        }
        while (index >= 0);
        return name;
    }

    private static HashSet<ExcelMergedRegion> MergeRegions(
        HashSet<ExcelMergedRegion> baseRegions,
        HashSet<ExcelMergedRegion> localRegions,
        HashSet<ExcelMergedRegion> remoteRegions)
    {
        var result = new HashSet<ExcelMergedRegion>(baseRegions);
        result.ExceptWith(baseRegions.Except(localRegions));
        result.ExceptWith(baseRegions.Except(remoteRegions));
        result.UnionWith(localRegions.Except(baseRegions));
        result.UnionWith(remoteRegions.Except(baseRegions));
        return result;
    }

    private static IReadOnlyList<ExcelMergedRegion> TransformRegions(
        IEnumerable<ExcelMergedRegion> regions,
        IReadOnlyList<int> duplicatedRows)
    {
        var result = new HashSet<ExcelMergedRegion>();
        foreach (var region in regions)
        {
            var shiftBefore = LowerBound(duplicatedRows, region.FirstRow);
            if (region.FirstRow == region.LastRow
                && shiftBefore < duplicatedRows.Count
                && duplicatedRows[shiftBefore] == region.FirstRow)
            {
                var row = region.FirstRow + shiftBefore;
                result.Add(region with { FirstRow = row, LastRow = row });
                result.Add(region with { FirstRow = row + 1, LastRow = row + 1 });
                continue;
            }

            result.Add(region with
            {
                FirstRow = region.FirstRow + shiftBefore,
                LastRow = region.LastRow + UpperBound(duplicatedRows, region.LastRow),
            });
        }
        return result.OrderBy(region => region.FirstRow).ThenBy(region => region.FirstColumn).ToList();
    }

    private static bool ContainsValueInRange(IReadOnlyList<int> values, int first, int last)
    {
        var index = LowerBound(values, first);
        return index < values.Count && values[index] <= last;
    }

    private static int LowerBound(IReadOnlyList<int> values, int target)
    {
        var low = 0;
        var high = values.Count;
        while (low < high)
        {
            var middle = low + (high - low) / 2;
            if (values[middle] < target)
                low = middle + 1;
            else
                high = middle;
        }
        return low;
    }

    private static int UpperBound(IReadOnlyList<int> values, int target)
    {
        var low = 0;
        var high = values.Count;
        while (low < high)
        {
            var middle = low + (high - low) / 2;
            if (values[middle] <= target)
                low = middle + 1;
            else
                high = middle;
        }
        return low;
    }

    private static void ValidateRegions(IReadOnlyList<ExcelMergedRegion> regions, string sheetName)
    {
        var active = new List<ExcelMergedRegion>();
        foreach (var region in regions.OrderBy(region => region.FirstRow).ThenBy(region => region.FirstColumn))
        {
            for (var index = active.Count - 1; index >= 0; index--)
            {
                var candidate = active[index];
                if (candidate.LastRow < region.FirstRow)
                {
                    active[index] = active[^1];
                    active.RemoveAt(active.Count - 1);
                }
                else if (candidate != region && candidate.Overlaps(region))
                    throw new WorkbookMergeException($"Merged-cell conflict in sheet '{sheetName}'.");
            }
            active.Add(region);
        }
    }

    private static bool ShouldOmitDeletedSheet(SheetGroup group)
    {
        if (group.Local == null && group.Remote == null)
            return true;
        if (group.Base == null)
            return false;
        if (group.Local == null && SheetSnapshot.AreEqual(group.Base, group.Remote))
            return true;
        return group.Remote == null && SheetSnapshot.AreEqual(group.Base, group.Local);
    }

    private static void ValidateOutputLimits(IEnumerable<MergedSheetResult> results, bool legacyXls)
    {
        var maximumAllowedRow = legacyXls ? 65535 : 1048575;
        var maximumAllowedColumn = legacyXls ? 255 : 16383;
        foreach (var result in results)
        {
            var maximumRow = result.Rows.Count == 0 ? 0 : result.Rows.Keys.Max();
            var maximumColumn = result.Rows.Values.SelectMany(row => row.Keys).DefaultIfEmpty(0).Max();
            if (maximumRow > maximumAllowedRow || maximumColumn > maximumAllowedColumn
                || result.MergedRegions.Any(region => region.LastRow > maximumAllowedRow || region.LastColumn > maximumAllowedColumn))
                throw new WorkbookMergeException($"Sheet '{result.Name}' exceeds the selected Excel format's row or column limits.");
        }
    }

    private static void ValidateSheetRenameAmbiguity(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook)
    {
        if (baseWorkbook.Sheets.Count <= 1 && localWorkbook.Sheets.Count <= 1 && remoteWorkbook.Sheets.Count <= 1)
            return;

        var removedFromBoth = baseWorkbook.Sheets.Keys
            .Where(name => !localWorkbook.Sheets.ContainsKey(name) && !remoteWorkbook.Sheets.ContainsKey(name))
            .ToList();
        var addedToBoth = localWorkbook.Sheets.Keys
            .Where(name => remoteWorkbook.Sheets.ContainsKey(name) && !baseWorkbook.Sheets.ContainsKey(name))
            .ToList();
        if (removedFromBoth.Count > 0 && addedToBoth.Count > 0)
            throw new WorkbookMergeException("Sheet renames cannot be matched safely yet. Keep sheet names stable before saving RESULT.");
    }

    private static IEnumerable<SheetGroup> CreateSheetGroups(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook)
    {
        if (baseWorkbook.Sheets.Count == 1 && localWorkbook.Sheets.Count == 1 && remoteWorkbook.Sheets.Count == 1)
        {
            var local = localWorkbook.Sheets.Single();
            yield return new SheetGroup(
                local.Key,
                baseWorkbook.Sheets.Values.Single(),
                local.Value,
                remoteWorkbook.Sheets.Values.Single());
            yield break;
        }

        var names = localWorkbook.Sheets.Keys
            .Concat(remoteWorkbook.Sheets.Keys)
            .Concat(baseWorkbook.Sheets.Keys)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        foreach (var name in names)
        {
            baseWorkbook.Sheets.TryGetValue(name, out var baseSheet);
            localWorkbook.Sheets.TryGetValue(name, out var localSheet);
            remoteWorkbook.Sheets.TryGetValue(name, out var remoteSheet);
            yield return new SheetGroup(name, baseSheet, localSheet, remoteSheet);
        }
    }

    private static string CreateUniqueSheetName(string name, ISet<string> usedNames)
    {
        var safeName = WorkbookUtil.CreateSafeSheetName(name);
        if (string.IsNullOrWhiteSpace(safeName))
            safeName = "Sheet";
        if (safeName.Length > 31)
            safeName = safeName[..31];

        var candidate = safeName;
        var suffix = 2;
        while (!usedNames.Add(candidate))
        {
            var suffixText = $" ({suffix++})";
            candidate = safeName[..Math.Min(safeName.Length, 31 - suffixText.Length)] + suffixText;
        }
        return candidate;
    }

    private sealed record SheetGroup(
        string ResultName,
        ExcelSheet? Base,
        ExcelSheet? Local,
        ExcelSheet? Remote);

    private sealed class OpenXmlWorksheetIndex
    {
        public XElement SheetData { get; }
        public Dictionary<string, XElement> Cells { get; } = new(StringComparer.Ordinal);
        public Dictionary<int, OpenXmlRowIndex> Rows { get; } = new();
        public SortedSet<int> RowNumbers { get; } = new();

        private OpenXmlWorksheetIndex(XElement sheetData)
        {
            SheetData = sheetData;
        }

        public static OpenXmlWorksheetIndex Create(XDocument document)
        {
            var sheetData = document.Descendants(SpreadsheetNamespace + "sheetData").Single();
            var index = new OpenXmlWorksheetIndex(sheetData);
            foreach (var row in sheetData.Elements(SpreadsheetNamespace + "row"))
            {
                var cells = row.Elements(SpreadsheetNamespace + "c").ToList();
                foreach (var cell in cells)
                {
                    var address = (string?)cell.Attribute("r");
                    if (address != null)
                        index.Cells.TryAdd(address, cell);
                }

                var rowNumber = (int?)row.Attribute("r");
                if (!rowNumber.HasValue || index.Rows.ContainsKey(rowNumber.Value))
                    continue;
                var rowIndex = new OpenXmlRowIndex(row);
                foreach (var cell in cells)
                {
                    var address = (string?)cell.Attribute("r");
                    if (address == null)
                        continue;
                    var columnIndex = GetColumnIndex(address);
                    rowIndex.Cells.TryAdd(columnIndex, cell);
                    rowIndex.ColumnNumbers.Add(columnIndex);
                }
                index.Rows.Add(rowNumber.Value, rowIndex);
                index.RowNumbers.Add(rowNumber.Value);
            }
            return index;
        }
    }

    private sealed class OpenXmlRowIndex
    {
        public XElement Row { get; }
        public Dictionary<int, XElement> Cells { get; } = new();
        public SortedSet<int> ColumnNumbers { get; } = new();

        public OpenXmlRowIndex(XElement row)
        {
            Row = row;
        }
    }

    private enum CellSource
    {
        Local,
        Remote,
        Custom,
    }

    private enum ExcelFileKind
    {
        Unknown,
        Xlsx,
        Xls,
    }

    private readonly record struct MergedCellResult(string Value, CellSource Source, int SourceRowIndex);

    private readonly record struct CellMergeState(
        int Column,
        string BaseValue,
        string LocalValue,
        string RemoteValue,
        bool IsConflict,
        MergeCellResolution Resolution);

    private sealed record XlsxPackageEntry(
        string Name,
        DateTimeOffset LastWriteTime,
        int ExternalAttributes,
        byte[] Content);

    private sealed record MergedSheetResult(
        string Name,
        SortedDictionary<int, Dictionary<int, MergedCellResult>> Rows,
        IReadOnlyList<ExcelMergedRegion> MergedRegions,
        IReadOnlyList<int> DuplicatedRows);

    private sealed class SheetSnapshot
    {
        public Dictionary<int, Dictionary<int, string>> Rows { get; } = new();
        public HashSet<ExcelMergedRegion> MergedRegions { get; } = new();

        public static SheetSnapshot Create(ExcelSheet? sheet)
        {
            var snapshot = new SheetSnapshot();
            if (sheet == null)
                return snapshot;
            foreach (var row in sheet.Rows.Values)
            {
                var values = new Dictionary<int, string>();
                var hasValue = false;
                foreach (var cell in row.Cells)
                {
                    values[cell.OriginalColumnIndex] = cell.Value;
                    hasValue |= !string.IsNullOrEmpty(cell.Value);
                }
                if (hasValue)
                    snapshot.Rows[row.Cells.FirstOrDefault()?.OriginalRowIndex ?? row.Index] = values;
            }
            snapshot.MergedRegions.UnionWith(sheet.MergedRegions);
            return snapshot;
        }

        public string GetValue(int row, int column)
        {
            return Rows.TryGetValue(row, out var values) && values.TryGetValue(column, out var value)
                ? value
                : string.Empty;
        }

        public IEnumerable<int> GetColumns(int row)
        {
            return Rows.TryGetValue(row, out var values) ? values.Keys : Enumerable.Empty<int>();
        }

        public Dictionary<int, MergedCellResult> CopyRow(int row, CellSource source)
        {
            return Rows.TryGetValue(row, out var values)
                ? values.ToDictionary(
                    pair => pair.Key,
                    pair => new MergedCellResult(pair.Value, source, row))
                : new Dictionary<int, MergedCellResult>();
        }

        public static bool AreEqual(ExcelSheet? first, ExcelSheet? second)
        {
            var left = Create(first);
            var right = Create(second);
            if (!left.MergedRegions.SetEquals(right.MergedRegions)
                || !left.Rows.Keys.OrderBy(value => value).SequenceEqual(right.Rows.Keys.OrderBy(value => value)))
                return false;
            return left.Rows.All(row => right.Rows.TryGetValue(row.Key, out var values)
                && row.Value.OrderBy(pair => pair.Key).SequenceEqual(values.OrderBy(pair => pair.Key)));
        }
    }
}

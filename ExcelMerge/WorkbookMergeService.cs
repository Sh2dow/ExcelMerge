#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
    public void Save(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook,
        IReadOnlyDictionary<MergeRowKey, MergeResolution> resolutions,
        string outputPath)
    {
        var extension = Path.GetExtension(outputPath);
        if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            throw new WorkbookMergeException("RESULT must use the .xlsx or .xls extension.");

        var unresolved = new HashSet<MergeRowKey>();
        var results = new List<MergedSheetResult>();
        ValidateSheetRenameAmbiguity(baseWorkbook, localWorkbook, remoteWorkbook);
        foreach (var group in CreateSheetGroups(baseWorkbook, localWorkbook, remoteWorkbook))
        {
            if (ShouldOmitDeletedSheet(group))
                continue;
            if (group.Base != null && (group.Local == null || group.Remote == null))
                throw new WorkbookMergeException($"Sheet deletion combined with modifications is not supported safely: '{group.ResultName}'.");
            results.Add(MergeSheet(group, resolutions, unresolved));
        }

        if (unresolved.Count > 0)
        {
            var locations = string.Join(", ", unresolved
                .OrderBy(key => key.SheetName)
                .ThenBy(key => key.RowIndex)
                .Take(10)
                .Select(key => $"{key.SheetName}!{key.RowIndex + 1}"));
            throw new WorkbookMergeException($"Resolve all conflicts before saving RESULT: {locations}");
        }

        ValidateOutputLimits(results, extension.Equals(".xls", StringComparison.OrdinalIgnoreCase));

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
                    if (!string.IsNullOrEmpty(cellValue.Value))
                        outputRow.CreateCell(cellValue.Key).SetCellValue(cellValue.Value);
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

    private static MergedSheetResult MergeSheet(
        SheetGroup group,
        IReadOnlyDictionary<MergeRowKey, MergeResolution> resolutions,
        ISet<MergeRowKey> unresolved)
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
        var rows = new SortedDictionary<int, Dictionary<int, string>>();
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
            var hasConflict = columns.Any(column => IsConflict(
                baseSnapshot.GetValue(rowIndex, column),
                localSnapshot.GetValue(rowIndex, column),
                remoteSnapshot.GetValue(rowIndex, column)));
            var key = new MergeRowKey(group.ResultName, rowIndex);
            var resolution = resolutions.TryGetValue(key, out var selected)
                ? selected
                : MergeResolution.Unresolved;
            if (hasConflict && resolution == MergeResolution.Unresolved)
                unresolved.Add(key);

            var outputRowIndex = rowIndex + shift;
            if (hasConflict && resolution == MergeResolution.Both)
            {
                rows[outputRowIndex] = localSnapshot.CopyRow(rowIndex);
                rows[outputRowIndex + 1] = remoteSnapshot.CopyRow(rowIndex);
                duplicatedRows.Add(rowIndex);
                shift++;
                continue;
            }

            var mergedRow = new Dictionary<int, string>();
            foreach (var column in columns)
            {
                var baseValue = baseSnapshot.GetValue(rowIndex, column);
                var localValue = localSnapshot.GetValue(rowIndex, column);
                var remoteValue = remoteSnapshot.GetValue(rowIndex, column);
                mergedRow[column] = MergeValue(baseValue, localValue, remoteValue, resolution);
            }
            rows[outputRowIndex] = mergedRow;
        }

        var regions = MergeRegions(baseSnapshot.MergedRegions, localSnapshot.MergedRegions, remoteSnapshot.MergedRegions);
        if (regions.Any(region => region.FirstRow != region.LastRow
            && duplicatedRows.Any(row => row >= region.FirstRow && row <= region.LastRow)))
            throw new WorkbookMergeException($"KEEP BOTH cannot safely duplicate a row inside a vertical merged range in sheet '{group.ResultName}'.");
        var transformedRegions = TransformRegions(regions, duplicatedRows);
        ValidateRegions(transformedRegions, group.ResultName);
        return new MergedSheetResult(group.ResultName, rows, transformedRegions);
    }

    private static string MergeValue(
        string baseValue,
        string localValue,
        string remoteValue,
        MergeResolution resolution)
    {
        if (localValue == remoteValue)
            return localValue;
        if (localValue == baseValue)
            return remoteValue;
        if (remoteValue == baseValue)
            return localValue;
        return resolution == MergeResolution.Remote ? remoteValue : localValue;
    }

    private static bool IsConflict(string baseValue, string localValue, string remoteValue)
    {
        return localValue != remoteValue
            && localValue != baseValue
            && remoteValue != baseValue;
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
        IReadOnlyCollection<int> duplicatedRows)
    {
        var duplicates = duplicatedRows.OrderBy(row => row).ToList();
        var result = new HashSet<ExcelMergedRegion>();
        foreach (var region in regions)
        {
            var shiftBefore = duplicates.Count(row => row < region.FirstRow);
            if (region.FirstRow == region.LastRow && duplicates.Contains(region.FirstRow))
            {
                var row = region.FirstRow + shiftBefore;
                result.Add(region with { FirstRow = row, LastRow = row });
                result.Add(region with { FirstRow = row + 1, LastRow = row + 1 });
                continue;
            }

            result.Add(region with
            {
                FirstRow = region.FirstRow + shiftBefore,
                LastRow = region.LastRow + duplicates.Count(row => row <= region.LastRow),
            });
        }
        return result.OrderBy(region => region.FirstRow).ThenBy(region => region.FirstColumn).ToList();
    }

    private static void ValidateRegions(IReadOnlyList<ExcelMergedRegion> regions, string sheetName)
    {
        for (var first = 0; first < regions.Count; first++)
        {
            for (var second = first + 1; second < regions.Count; second++)
            {
                if (regions[first] != regions[second] && regions[first].Overlaps(regions[second]))
                    throw new WorkbookMergeException($"Merged-cell conflict in sheet '{sheetName}'.");
            }
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

    private sealed record MergedSheetResult(
        string Name,
        SortedDictionary<int, Dictionary<int, string>> Rows,
        IReadOnlyList<ExcelMergedRegion> MergedRegions);

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
                foreach (var cell in row.Cells)
                    values[cell.OriginalColumnIndex] = cell.Value;
                if (values.Values.Any(value => !string.IsNullOrEmpty(value)))
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

        public Dictionary<int, string> CopyRow(int row)
        {
            return Rows.TryGetValue(row, out var values)
                ? new Dictionary<int, string>(values)
                : new Dictionary<int, string>();
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

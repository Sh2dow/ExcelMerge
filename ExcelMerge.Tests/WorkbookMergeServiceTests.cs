using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class WorkbookMergeServiceTests
{
    [TestMethod]
    public void SaveUsesSelectedLocalValue()
    {
        var path = TempPath();
        try
        {
            new WorkbookMergeService().Save(
                Workbook("10"),
                Workbook("11"),
                Workbook("12"),
                new Dictionary<MergeRowKey, MergeResolution>
                {
                    [new MergeRowKey("Sheet1", 0)] = MergeResolution.Local,
                },
                path);

            var result = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());
            Assert.AreEqual("11", result.Sheets["Sheet1"].Rows[0].Cells[0].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void SaveRowResolutionUsesTheEntireLocalRow()
    {
        var path = TempPath();
        try
        {
            new WorkbookMergeService().Save(
                Workbook("10", "base"),
                Workbook("11", "base"),
                Workbook("12", "remote change"),
                new Dictionary<MergeRowKey, MergeResolution>
                {
                    [new MergeRowKey("Sheet1", 0)] = MergeResolution.Local,
                },
                new Dictionary<MergeCellKey, MergeCellResolution>(),
                path);

            var cells = ExcelWorkbook.Create(path, new ExcelSheetReadConfig()).Sheets["Sheet1"].Rows[0].Cells;
            Assert.AreEqual("11", cells[0].Value);
            Assert.AreEqual("base", cells[1].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void SaveUsesIndependentCellResolutions()
    {
        var path = TempPath();
        try
        {
            new WorkbookMergeService().Save(
                Workbook("10", "20"),
                Workbook("11", "21"),
                Workbook("12", "22"),
                new Dictionary<MergeCellKey, MergeCellResolution>
                {
                    [new MergeCellKey("Sheet1", 0, 0)] = new(MergeResolution.Local),
                    [new MergeCellKey("Sheet1", 0, 1)] = new(MergeResolution.Remote),
                },
                path);

            var cells = ExcelWorkbook.Create(path, new ExcelSheetReadConfig()).Sheets["Sheet1"].Rows[0].Cells;
            Assert.AreEqual("11", cells[0].Value);
            Assert.AreEqual("22", cells[1].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void SaveUsesCustomConflictValue()
    {
        var path = TempPath();
        try
        {
            new WorkbookMergeService().Save(
                Workbook("10"),
                Workbook("11"),
                Workbook("12"),
                new Dictionary<MergeCellKey, MergeCellResolution>
                {
                    [new MergeCellKey("Sheet1", 0, 0)] = new(MergeResolution.Custom, "user value"),
                },
                path);

            var result = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());
            Assert.AreEqual("user value", result.Sheets["Sheet1"].Rows[0].Cells[0].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void SaveKeepBothDuplicatesTheConflictRow()
    {
        var path = TempPath();
        try
        {
            new WorkbookMergeService().Save(
                Workbook("10"),
                Workbook("11"),
                Workbook("12"),
                new Dictionary<MergeRowKey, MergeResolution>
                {
                    [new MergeRowKey("Sheet1", 0)] = MergeResolution.Both,
                },
                path);

            var result = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());
            Assert.AreEqual("11", result.Sheets["Sheet1"].Rows[0].Cells[0].Value);
            Assert.AreEqual("12", result.Sheets["Sheet1"].Rows[1].Cells[0].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void SavePreservesMergedCellRanges()
    {
        var path = TempPath();
        try
        {
            var workbook = Workbook("Title");
            workbook.Sheets["Sheet1"].MergedRegions.Add(new ExcelMergedRegion(0, 0, 0, 1));

            new WorkbookMergeService().Save(
                workbook,
                workbook,
                workbook,
                new Dictionary<MergeRowKey, MergeResolution>(),
                path);

            var result = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());
            CollectionAssert.Contains(
                result.Sheets["Sheet1"].MergedRegions,
                new ExcelMergedRegion(0, 0, 0, 1));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void SaveIgnoresDifferentTrailingBlankRows()
    {
        var path = TempPath();
        try
        {
            var baseWorkbook = Workbook("10");
            var localWorkbook = Workbook("11");
            var remoteWorkbook = Workbook("12");
            baseWorkbook.Sheets["Sheet1"].Rows.Add(2, new ExcelRow(2, new[] { new ExcelCell(string.Empty, 0, 2) }));
            localWorkbook.Sheets["Sheet1"].Rows.Add(2, new ExcelRow(2, new[] { new ExcelCell(string.Empty, 0, 2) }));
            remoteWorkbook.Sheets["Sheet1"].Rows.Add(1, new ExcelRow(1, new[] { new ExcelCell("remote addition", 0, 1) }));

            new WorkbookMergeService().Save(
                baseWorkbook,
                localWorkbook,
                remoteWorkbook,
                new Dictionary<MergeRowKey, MergeResolution>
                {
                    [new MergeRowKey("Sheet1", 0)] = MergeResolution.Both,
                },
                path);

            var result = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());
            Assert.AreEqual("11", result.Sheets["Sheet1"].Rows[0].Cells[0].Value);
            Assert.AreEqual("12", result.Sheets["Sheet1"].Rows[1].Cells[0].Value);
            Assert.AreEqual("remote addition", result.Sheets["Sheet1"].Rows[2].Cells[0].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void SavePreservesLocalFormattingWhileApplyingRemoteFormula()
    {
        var basePath = TempPath();
        var localPath = TempPath();
        var remotePath = TempPath();
        var outputPath = TempPath();
        try
        {
            WriteFormulaWorkbook(basePath, "1+1", IndexedColors.Grey25Percent.Index);
            WriteFormulaWorkbook(localPath, "1+2", IndexedColors.LightCornflowerBlue.Index);
            WriteFormulaWorkbook(remotePath, "1+3", IndexedColors.LightGreen.Index);
            var config = new ExcelSheetReadConfig();

            new WorkbookMergeService().Save(
                ExcelWorkbook.Create(basePath, config),
                ExcelWorkbook.Create(localPath, config),
                ExcelWorkbook.Create(remotePath, config),
                new Dictionary<MergeRowKey, MergeResolution>(),
                new Dictionary<MergeCellKey, MergeCellResolution>
                {
                    [new MergeCellKey("Sheet1", 0, 0)] = new(MergeResolution.Remote),
                },
                localPath,
                remotePath,
                outputPath);

            using var result = WorkbookFactory.Create(outputPath);
            var sheet = result.GetSheet("Sheet1");
            var cell = sheet.GetRow(0).GetCell(0);
            Assert.AreEqual(CellType.Formula, cell.CellType);
            Assert.AreEqual("1+3", cell.CellFormula);
            Assert.AreEqual(IndexedColors.LightCornflowerBlue.Index, cell.CellStyle.FillForegroundColor);
            Assert.AreEqual(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
            Assert.AreEqual(BorderStyle.Thick, cell.CellStyle.BorderBottom);
            Assert.IsTrue(result.GetFontAt(cell.CellStyle.FontIndex).IsBold);
            Assert.AreEqual(IndexedColors.DarkBlue.Index, result.GetFontAt(cell.CellStyle.FontIndex).Color);
            Assert.AreEqual(32f, sheet.GetRow(0).HeightInPoints);
            Assert.AreEqual(6400, sheet.GetColumnWidth(0));
            Assert.IsTrue(sheet.PaneInformation.IsFreezePane());
            Assert.AreEqual(1, sheet.NumMergedRegions);
            Assert.AreEqual("A2:B2", sheet.GetMergedRegion(0).FormatAsString());
        }
        finally
        {
            Delete(basePath, localPath, remotePath, outputPath);
        }
    }

    [TestMethod]
    public void SavePreservesFormattingWithExtensionlessGitInputs()
    {
        var basePath = TempPath(string.Empty);
        var localPath = TempPath(string.Empty);
        var remotePath = TempPath(string.Empty);
        var outputPath = TempPath();
        try
        {
            WriteFormulaWorkbook(basePath, "1+1", IndexedColors.Grey25Percent.Index);
            WriteFormulaWorkbook(localPath, "1+2", IndexedColors.LightCornflowerBlue.Index);
            WriteFormulaWorkbook(remotePath, "1+3", IndexedColors.LightGreen.Index);
            var config = new ExcelSheetReadConfig();

            new WorkbookMergeService().Save(
                ExcelWorkbook.Create(basePath, config),
                ExcelWorkbook.Create(localPath, config),
                ExcelWorkbook.Create(remotePath, config),
                new Dictionary<MergeRowKey, MergeResolution>(),
                new Dictionary<MergeCellKey, MergeCellResolution>
                {
                    [new MergeCellKey("Sheet1", 0, 0)] = new(MergeResolution.Remote),
                },
                localPath,
                remotePath,
                outputPath);

            using var result = WorkbookFactory.Create(outputPath);
            var cell = result.GetSheet("Sheet1").GetRow(0).GetCell(0);
            Assert.AreEqual("1+3", cell.CellFormula);
            Assert.AreEqual(IndexedColors.LightCornflowerBlue.Index, cell.CellStyle.FillForegroundColor);
            Assert.AreEqual(BorderStyle.Thick, cell.CellStyle.BorderBottom);
            Assert.AreEqual(6400, result.GetSheet("Sheet1").GetColumnWidth(0));
        }
        finally
        {
            Delete(basePath, localPath, remotePath, outputPath);
        }
    }

    [TestMethod]
    public void SaveCopiesFormattingForRemoteOnlySheet()
    {
        var basePath = TempPath();
        var localPath = TempPath();
        var remotePath = TempPath();
        var outputPath = TempPath();
        try
        {
            WriteWorkbook(basePath, includeRemoteSheet: false);
            WriteWorkbook(localPath, includeRemoteSheet: false);
            WriteWorkbook(remotePath, includeRemoteSheet: true);
            var config = new ExcelSheetReadConfig();

            new WorkbookMergeService().Save(
                ExcelWorkbook.Create(basePath, config),
                ExcelWorkbook.Create(localPath, config),
                ExcelWorkbook.Create(remotePath, config),
                new Dictionary<MergeRowKey, MergeResolution>(),
                new Dictionary<MergeCellKey, MergeCellResolution>(),
                localPath,
                remotePath,
                outputPath);

            using var result = WorkbookFactory.Create(outputPath);
            Assert.AreEqual(2, result.NumberOfSheets);
            var sheet = result.GetSheet("RemoteOnly");
            var cell = sheet.GetRow(0).GetCell(0);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            Assert.AreEqual(42.5, cell.NumericCellValue);
            Assert.AreEqual(IndexedColors.LightOrange.Index, cell.CellStyle.FillForegroundColor);
            Assert.AreEqual(FillPattern.SolidForeground, cell.CellStyle.FillPattern);
            Assert.IsTrue(result.GetFontAt(cell.CellStyle.FontIndex).IsBold);
            Assert.AreEqual(28f, sheet.GetRow(0).HeightInPoints);
            Assert.AreEqual(7200, sheet.GetColumnWidth(0));
            Assert.AreEqual(1, sheet.NumMergedRegions);
            Assert.AreEqual("A2:B2", sheet.GetMergedRegion(0).FormatAsString());
            Assert.AreEqual(SheetVisibility.VeryHidden, result.GetSheetVisibility(result.GetSheetIndex(sheet)));
        }
        finally
        {
            Delete(basePath, localPath, remotePath, outputPath);
        }
    }

    [TestMethod]
    public void SaveReplacesLocalFormulaWithRemoteLiteralAndKeepsStyle()
    {
        var basePath = TempPath();
        var localPath = TempPath();
        var remotePath = TempPath();
        var outputPath = TempPath();
        try
        {
            WriteNumericWorkbook(basePath, 1, IndexedColors.Grey25Percent.Index);
            WriteFormulaWorkbook(localPath, "1+1", IndexedColors.LightCornflowerBlue.Index);
            WriteNumericWorkbook(remotePath, 3, IndexedColors.LightGreen.Index);
            var config = new ExcelSheetReadConfig();

            new WorkbookMergeService().Save(
                ExcelWorkbook.Create(basePath, config),
                ExcelWorkbook.Create(localPath, config),
                ExcelWorkbook.Create(remotePath, config),
                new Dictionary<MergeRowKey, MergeResolution>(),
                new Dictionary<MergeCellKey, MergeCellResolution>
                {
                    [new MergeCellKey("Sheet1", 0, 0)] = new(MergeResolution.Remote),
                },
                localPath,
                remotePath,
                outputPath);

            using var result = WorkbookFactory.Create(outputPath);
            var cell = result.GetSheet("Sheet1").GetRow(0).GetCell(0);
            Assert.AreEqual(CellType.Numeric, cell.CellType);
            Assert.AreEqual(3d, cell.NumericCellValue);
            Assert.AreEqual(IndexedColors.LightCornflowerBlue.Index, cell.CellStyle.FillForegroundColor);
        }
        finally
        {
            Delete(basePath, localPath, remotePath, outputPath);
        }
    }

    [TestMethod]
    public void SaveKeepBothPreservesLocalAndRemoteRowStyles()
    {
        var basePath = TempPath();
        var localPath = TempPath();
        var remotePath = TempPath();
        var outputPath = TempPath();
        try
        {
            WriteFormulaWorkbook(basePath, "1+1", IndexedColors.Grey25Percent.Index);
            WriteFormulaWorkbook(localPath, "1+2", IndexedColors.LightCornflowerBlue.Index);
            WriteFormulaWorkbook(remotePath, "1+3", IndexedColors.LightGreen.Index);
            var config = new ExcelSheetReadConfig();
            var localTheme = Encoding.UTF8.GetBytes(
                "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Preserve me\"/>");
            AddPackageEntry(localPath, "xl/theme/theme1.xml", localTheme);

            new WorkbookMergeService().Save(
                ExcelWorkbook.Create(basePath, config),
                ExcelWorkbook.Create(localPath, config),
                ExcelWorkbook.Create(remotePath, config),
                new Dictionary<MergeRowKey, MergeResolution>
                {
                    [new MergeRowKey("Sheet1", 0)] = MergeResolution.Both,
                },
                new Dictionary<MergeCellKey, MergeCellResolution>(),
                localPath,
                remotePath,
                outputPath);

            using var result = WorkbookFactory.Create(outputPath);
            var sheet = result.GetSheet("Sheet1");
            Assert.AreEqual("1+2", sheet.GetRow(0).GetCell(0).CellFormula);
            Assert.AreEqual("1+3", sheet.GetRow(1).GetCell(0).CellFormula);
            Assert.AreEqual(
                IndexedColors.LightCornflowerBlue.Index,
                sheet.GetRow(0).GetCell(0).CellStyle.FillForegroundColor);
            Assert.AreEqual(
                IndexedColors.LightGreen.Index,
                sheet.GetRow(1).GetCell(0).CellStyle.FillForegroundColor);
            Assert.AreEqual(32f, sheet.GetRow(1).HeightInPoints);
            Assert.AreEqual("A3:B3", sheet.GetMergedRegion(0).FormatAsString());
            CollectionAssert.AreEqual(localTheme, ReadPackageEntry(outputPath, "xl/theme/theme1.xml"));
        }
        finally
        {
            Delete(basePath, localPath, remotePath, outputPath);
        }
    }

    [TestMethod]
    public void SavePreservesLocalFormattingInLegacyXls()
    {
        var basePath = TempPath(".xls");
        var localPath = TempPath(".xls");
        var remotePath = TempPath(".xls");
        var outputPath = TempPath(".xls");
        try
        {
            WriteLegacyWorkbook(basePath, "base", IndexedColors.Grey25Percent.Index);
            WriteLegacyWorkbook(localPath, "local", IndexedColors.LightCornflowerBlue.Index);
            WriteLegacyWorkbook(remotePath, "remote", IndexedColors.LightGreen.Index);
            var config = new ExcelSheetReadConfig();

            new WorkbookMergeService().Save(
                ExcelWorkbook.Create(basePath, config),
                ExcelWorkbook.Create(localPath, config),
                ExcelWorkbook.Create(remotePath, config),
                new Dictionary<MergeRowKey, MergeResolution>(),
                new Dictionary<MergeCellKey, MergeCellResolution>
                {
                    [new MergeCellKey("Sheet1", 0, 0)] = new(MergeResolution.Remote),
                },
                localPath,
                remotePath,
                outputPath);

            using var result = WorkbookFactory.Create(outputPath);
            var cell = result.GetSheet("Sheet1").GetRow(0).GetCell(0);
            Assert.AreEqual("remote", cell.StringCellValue);
            Assert.AreEqual(IndexedColors.LightCornflowerBlue.Index, cell.CellStyle.FillForegroundColor);
            Assert.AreEqual(BorderStyle.Thick, cell.CellStyle.BorderBottom);
        }
        finally
        {
            Delete(basePath, localPath, remotePath, outputPath);
        }
    }

    [TestMethod]
    public void SavePreservesUntouchedInlineStringsInPatchedXlsx()
    {
        var basePath = TempPath();
        var localPath = TempPath();
        var remotePath = TempPath();
        var outputPath = TempPath();
        try
        {
            WriteInlineStringWorkbook(basePath, "base");
            WriteInlineStringWorkbook(localPath, "base");
            WriteInlineStringWorkbook(remotePath, "remote");
            var localStyles = ReadPackageEntry(localPath, "xl/styles.xml");
            var config = new ExcelSheetReadConfig();

            new WorkbookMergeService().Save(
                ExcelWorkbook.Create(basePath, config),
                ExcelWorkbook.Create(localPath, config),
                ExcelWorkbook.Create(remotePath, config),
                new Dictionary<MergeRowKey, MergeResolution>(),
                new Dictionary<MergeCellKey, MergeCellResolution>(),
                localPath,
                remotePath,
                outputPath);

            using var result = WorkbookFactory.Create(outputPath);
            var sheet = result.GetSheet("Sheet1");
            Assert.AreEqual("remote", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.AreEqual("untouched inline text", sheet.GetRow(1).GetCell(0).StringCellValue);
            CollectionAssert.AreEqual(localStyles, ReadPackageEntry(outputPath, "xl/styles.xml"));
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XDocument worksheet;
            using (var input = new MemoryStream(ReadPackageEntry(outputPath, "xl/worksheets/sheet1.xml")))
                worksheet = XDocument.Load(input);
            var untouched = worksheet.Descendants(spreadsheet + "c")
                .Single(cell => (string)cell.Attribute("r")! == "A2");
            Assert.AreEqual("inlineStr", (string?)untouched.Attribute("t"));
            Assert.AreEqual("untouched inline text", untouched.Descendants(spreadsheet + "t").Single().Value);
        }
        finally
        {
            Delete(basePath, localPath, remotePath, outputPath);
        }
    }

    [TestMethod]
    public void SaveInsertsRemoteCellsAndRowsInWorksheetOrder()
    {
        var basePath = TempPath();
        var localPath = TempPath();
        var remotePath = TempPath();
        var outputPath = TempPath();
        try
        {
            WriteSparsePatchWorkbook(basePath, includeRemoteValues: false);
            WriteSparsePatchWorkbook(localPath, includeRemoteValues: false);
            WriteSparsePatchWorkbook(remotePath, includeRemoteValues: true);
            var config = new ExcelSheetReadConfig();

            new WorkbookMergeService().Save(
                ExcelWorkbook.Create(basePath, config),
                ExcelWorkbook.Create(localPath, config),
                ExcelWorkbook.Create(remotePath, config),
                new Dictionary<MergeRowKey, MergeResolution>(),
                new Dictionary<MergeCellKey, MergeCellResolution>(),
                localPath,
                remotePath,
                outputPath);

            using (var result = WorkbookFactory.Create(outputPath))
            {
                var sheet = result.GetSheet("Sheet1");
                Assert.AreEqual("remote B", sheet.GetRow(0).GetCell(1).StringCellValue);
                Assert.AreEqual("remote C", sheet.GetRow(0).GetCell(2).StringCellValue);
                Assert.AreEqual("remote E", sheet.GetRow(2).GetCell(4).StringCellValue);
            }

            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XDocument worksheet;
            using (var input = new MemoryStream(ReadPackageEntry(outputPath, "xl/worksheets/sheet1.xml")))
                worksheet = XDocument.Load(input);
            var rows = worksheet.Descendants(spreadsheet + "sheetData")
                .Single()
                .Elements(spreadsheet + "row")
                .ToList();
            CollectionAssert.AreEqual(new[] { "1", "3" }, rows.Select(row => (string)row.Attribute("r")!).ToArray());
            CollectionAssert.AreEqual(
                new[] { "A1", "B1", "C1", "D1" },
                rows[0].Elements(spreadsheet + "c").Select(cell => (string)cell.Attribute("r")!).ToArray());
            CollectionAssert.AreEqual(
                new[] { "E3" },
                rows[1].Elements(spreadsheet + "c").Select(cell => (string)cell.Attribute("r")!).ToArray());
        }
        finally
        {
            Delete(basePath, localPath, remotePath, outputPath);
        }
    }

    private static ExcelWorkbook Workbook(params string[] values)
    {
        var workbook = new ExcelWorkbook();
        var sheet = new ExcelSheet();
        sheet.Rows.Add(0, new ExcelRow(0, values.Select((value, column) => new ExcelCell(value, column, 0))));
        workbook.Sheets.Add("Sheet1", sheet);
        return workbook;
    }

    private static void WriteFormulaWorkbook(string path, string formula, short fillColor)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");
        var row = sheet.CreateRow(0);
        row.HeightInPoints = 32;
        var cell = row.CreateCell(0);
        cell.SetCellFormula(formula);
        var style = workbook.CreateCellStyle();
        style.FillForegroundColor = fillColor;
        style.FillPattern = FillPattern.SolidForeground;
        style.BorderBottom = BorderStyle.Thick;
        style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
        var font = workbook.CreateFont();
        font.IsBold = true;
        font.Color = IndexedColors.DarkBlue.Index;
        style.SetFont(font);
        cell.CellStyle = style;
        var titleRow = sheet.CreateRow(1);
        titleRow.CreateCell(0).SetCellValue("Merged title");
        sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
        sheet.SetColumnWidth(0, 6400);
        sheet.CreateFreezePane(1, 1);
        workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateFormulaCell(cell);
        using var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        workbook.Write(output);
    }

    private static void WriteNumericWorkbook(string path, double value, short fillColor)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");
        var row = sheet.CreateRow(0);
        row.HeightInPoints = 32;
        var cell = row.CreateCell(0);
        cell.SetCellValue(value);
        var style = workbook.CreateCellStyle();
        style.FillForegroundColor = fillColor;
        style.FillPattern = FillPattern.SolidForeground;
        style.BorderBottom = BorderStyle.Thick;
        style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
        var font = workbook.CreateFont();
        font.IsBold = true;
        font.Color = IndexedColors.DarkBlue.Index;
        style.SetFont(font);
        cell.CellStyle = style;
        sheet.CreateRow(1).CreateCell(0).SetCellValue("Merged title");
        sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
        sheet.SetColumnWidth(0, 6400);
        sheet.CreateFreezePane(1, 1);
        using var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        workbook.Write(output);
    }

    private static void WriteWorkbook(string path, bool includeRemoteSheet)
    {
        using var workbook = new XSSFWorkbook();
        workbook.CreateSheet("Existing").CreateRow(0).CreateCell(0).SetCellValue("unchanged");
        if (includeRemoteSheet)
        {
            var sheet = workbook.CreateSheet("RemoteOnly");
            workbook.SetSheetVisibility(workbook.GetSheetIndex(sheet), SheetVisibility.VeryHidden);
            var row = sheet.CreateRow(0);
            row.HeightInPoints = 28;
            var cell = row.CreateCell(0);
            cell.SetCellValue(42.5);
            var style = workbook.CreateCellStyle();
            style.FillForegroundColor = IndexedColors.LightOrange.Index;
            style.FillPattern = FillPattern.SolidForeground;
            var font = workbook.CreateFont();
            font.IsBold = true;
            style.SetFont(font);
            cell.CellStyle = style;
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Remote title");
            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 0, 1));
            sheet.SetColumnWidth(0, 7200);
        }

        using var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        workbook.Write(output);
    }

    private static void WriteLegacyWorkbook(string path, string value, short fillColor)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");
        var cell = sheet.CreateRow(0).CreateCell(0);
        cell.SetCellValue(value);
        var style = workbook.CreateCellStyle();
        style.FillForegroundColor = fillColor;
        style.FillPattern = FillPattern.SolidForeground;
        style.BorderBottom = BorderStyle.Thick;
        cell.CellStyle = style;
        using var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        workbook.Write(output);
    }

    private static void WriteInlineStringWorkbook(string path, string firstValue)
    {
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Sheet1");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(firstValue);
            sheet.CreateRow(1).CreateCell(0).SetCellValue("untouched inline text");
            using var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            workbook.Write(output);
        }

        var entries = new List<(string Name, DateTimeOffset Time, int Attributes, byte[] Content)>();
        using (var archive = ZipFile.OpenRead(path))
        {
            foreach (var entry in archive.Entries)
            {
                using var input = entry.Open();
                using var content = new MemoryStream();
                input.CopyTo(content);
                entries.Add((entry.FullName, entry.LastWriteTime, entry.ExternalAttributes, content.ToArray()));
            }
        }

        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var sheetEntry = entries.Single(entry => entry.Name == "xl/worksheets/sheet1.xml");
        XDocument document;
        using (var input = new MemoryStream(sheetEntry.Content))
            document = XDocument.Load(input);
        foreach (var cell in document.Descendants(spreadsheet + "c"))
        {
            var address = (string)cell.Attribute("r")!;
            var value = address == "A1" ? firstValue : "untouched inline text";
            cell.SetAttributeValue("t", "inlineStr");
            cell.ReplaceNodes(new XElement(
                spreadsheet + "is",
                new XElement(spreadsheet + "t", value)));
        }
        using (var content = new MemoryStream())
        {
            document.Save(content);
            var index = entries.FindIndex(entry => entry.Name == sheetEntry.Name);
            entries[index] = (sheetEntry.Name, sheetEntry.Time, sheetEntry.Attributes, content.ToArray());
        }

        var temporaryPath = path + ".tmp";
        using (var archive = ZipFile.Open(temporaryPath, ZipArchiveMode.Create))
        {
            foreach (var source in entries)
            {
                var entry = archive.CreateEntry(source.Name, CompressionLevel.Optimal);
                entry.LastWriteTime = source.Time;
                entry.ExternalAttributes = source.Attributes;
                using var output = entry.Open();
                output.Write(source.Content);
            }
        }
        File.Move(temporaryPath, path, true);
    }

    private static void WriteSparsePatchWorkbook(string path, bool includeRemoteValues)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");
        var firstRow = sheet.CreateRow(0);
        firstRow.CreateCell(0).SetCellValue("common");
        if (includeRemoteValues)
        {
            firstRow.CreateCell(1).SetCellValue("remote B");
            firstRow.CreateCell(2).SetCellValue("remote C");
            sheet.CreateRow(2).CreateCell(4).SetCellValue("remote E");
        }
        firstRow.CreateCell(3).SetCellValue("tail");
        using var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        workbook.Write(output);
    }

    private static byte[] ReadPackageEntry(string path, string entryName)
    {
        using var archive = ZipFile.OpenRead(path);
        using var input = archive.GetEntry(entryName)!.Open();
        using var content = new MemoryStream();
        input.CopyTo(content);
        return content.ToArray();
    }

    private static void AddPackageEntry(string path, string entryName, byte[] content)
    {
        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        archive.GetEntry(entryName)?.Delete();
        using var output = archive.CreateEntry(entryName, CompressionLevel.Optimal).Open();
        output.Write(content);
    }

    private static void Delete(params string[] paths)
    {
        foreach (var path in paths)
            File.Delete(path);
    }

    private static string TempPath(string extension = ".xlsx") =>
        Path.Combine(Path.GetTempPath(), $"excelmerge-result-{Guid.NewGuid():N}{extension}");
}

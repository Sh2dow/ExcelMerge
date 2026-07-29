using Microsoft.VisualStudio.TestTools.UnitTesting;

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

    private static ExcelWorkbook Workbook(string value)
    {
        var workbook = new ExcelWorkbook();
        var sheet = new ExcelSheet();
        sheet.Rows.Add(0, new ExcelRow(0, new[] { new ExcelCell(value, 0, 0) }));
        workbook.Sheets.Add("Sheet1", sheet);
        return workbook;
    }

    private static string TempPath() =>
        Path.Combine(Path.GetTempPath(), $"excelmerge-result-{Guid.NewGuid():N}.xlsx");
}

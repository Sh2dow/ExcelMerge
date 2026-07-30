using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.XSSF.UserModel;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class ExcelWorkbookTests
{
    [TestMethod]
    public void CreateLoadsWorkbookSheetsAndValues()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmerge-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XSSFWorkbook())
            {
                var data = workbook.CreateSheet("Data");
                data.CreateRow(0).CreateCell(0).SetCellValue("value");
                workbook.CreateSheet("Other");
                using var output = File.Create(path);
                workbook.Write(output);
            }

            var loaded = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());

            CollectionAssert.AreEqual(new[] { "Data", "Other" }, loaded.Sheets.Keys.ToArray());
            Assert.AreEqual("value", loaded.Sheets["Data"].Rows[0].Cells[0].Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void CreateAcceptsUppercaseCsvExtension()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmerge-{Guid.NewGuid():N}.CSV");
        try
        {
            File.WriteAllText(path, "Id,Name\n1,Alice");

            var workbook = ExcelWorkbook.Create(path, new ExcelSheetReadConfig());

            Assert.AreEqual(1, workbook.Sheets.Count);
            Assert.AreEqual("Alice", workbook.Sheets.Values.Single().Rows[1].Cells[1].Value);
            Assert.AreEqual(1, workbook.Sheets.Values.Single().Rows[1].Cells[1].OriginalColumnIndex);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [TestMethod]
    public void CreateLoadsOnlyPhysicalRowsFromSparseWorksheet()
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmerge-{Guid.NewGuid():N}.xlsx");
        try
        {
            using (var workbook = new XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("Sparse");
                sheet.CreateRow(0).CreateCell(0).SetCellValue("first");
                sheet.CreateRow(100000).CreateCell(0).SetCellValue("last");
                using var output = File.Create(path);
                workbook.Write(output);
            }

            var loaded = ExcelWorkbook.Create(path, new ExcelSheetReadConfig()).Sheets["Sparse"];

            Assert.AreEqual(2, loaded.Rows.Count);
            Assert.AreEqual(0, loaded.Rows[0].Cells[0].OriginalRowIndex);
            Assert.AreEqual(100000, loaded.Rows[1].Cells[0].OriginalRowIndex);
        }
        finally
        {
            File.Delete(path);
        }
    }
}

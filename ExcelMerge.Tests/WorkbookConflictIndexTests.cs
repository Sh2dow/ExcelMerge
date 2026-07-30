using ExcelMerge.Avalonia;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class WorkbookConflictIndexTests
{
    [TestMethod]
    public void FindWorkbookConflictsIndexesEverySheet()
    {
        var conflicts = MainWindow.FindWorkbookConflicts(
            Workbook(("Inventory", "10"), ("Team", "Engineer"), ("Notes", "Base")),
            Workbook(("Inventory", "11"), ("Team", "Lead"), ("Notes", "Local")),
            Workbook(("Inventory", "12"), ("Team", "Architect"), ("Notes", "Base")));

        Assert.AreEqual(1, conflicts["Inventory"].Count);
        Assert.AreEqual(new MergeCellKey("Inventory", 0, 0), conflicts["Inventory"][0]);
        Assert.AreEqual(1, conflicts["Team"].Count);
        Assert.AreEqual(new MergeCellKey("Team", 0, 0), conflicts["Team"][0]);
        Assert.AreEqual(0, conflicts["Notes"].Count);
    }

    [TestMethod]
    public void FindWorkbookConflictsDoesNotReuseSoleBaseSheetForNewSheet()
    {
        var conflicts = MainWindow.FindWorkbookConflicts(
            Workbook(("Existing", "Base")),
            Workbook(("Existing", "Base"), ("SharedNew", "Base")),
            Workbook(("Existing", "Base"), ("SharedNew", "Remote")));

        Assert.AreEqual(1, conflicts["SharedNew"].Count);
        Assert.AreEqual(new MergeCellKey("SharedNew", 0, 0), conflicts["SharedNew"][0]);
    }

    private static ExcelWorkbook Workbook(params (string Name, string Value)[] sheets)
    {
        var workbook = new ExcelWorkbook();
        foreach (var (name, value) in sheets)
        {
            var sheet = new ExcelSheet();
            sheet.Rows.Add(0, new ExcelRow(0, new[] { new ExcelCell(value, 0, 0) }));
            workbook.Sheets.Add(name, sheet);
        }
        return workbook;
    }
}

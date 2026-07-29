using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class ExcelSheetDiffTests
{
    [TestMethod]
    public void DiffReportsModifiedCells()
    {
        var left = CreateSheet(
            new[] { "Id", "Name" },
            new[] { "1", "Alice" });
        var right = CreateSheet(
            new[] { "Id", "Name" },
            new[] { "1", "Bob" });

        var diff = ExcelSheet.Diff(left, right, HeaderConfig());

        Assert.AreEqual(1, diff.CreateSummary().ModifiedCellCount);
        Assert.IsTrue(diff.Rows.Values.SelectMany(row => row.Cells.Values)
            .Any(cell => cell.Status == ExcelCellStatus.Modified && cell.SrcCell.Value == "Alice" && cell.DstCell.Value == "Bob"));
    }

    [TestMethod]
    public void DiffDoesNotMutateInputsWhenColumnsAreInserted()
    {
        var left = CreateSheet(
            new[] { "Id", "Name" },
            new[] { "1", "Alice" });
        var right = CreateSheet(
            new[] { "Id", "Email", "Name" },
            new[] { "1", "alice@example.com", "Alice" });
        var leftBefore = Snapshot(left);
        var rightBefore = Snapshot(right);

        var first = ExcelSheet.Diff(left, right, HeaderConfig()).CreateSummary();
        var second = ExcelSheet.Diff(left, right, HeaderConfig()).CreateSummary();

        CollectionAssert.AreEqual(leftBefore, Snapshot(left));
        CollectionAssert.AreEqual(rightBefore, Snapshot(right));
        Assert.AreEqual(first.ModifiedCellCount, second.ModifiedCellCount);
        Assert.AreEqual(first.AddedRowCount, second.AddedRowCount);
        Assert.AreEqual(first.RemovedRowCount, second.RemovedRowCount);
    }

    [TestMethod]
    public void DiffHandlesEmptySheets()
    {
        var diff = ExcelSheet.Diff(new ExcelSheet(), new ExcelSheet(), new ExcelSheetDiffConfig());

        Assert.AreEqual(0, diff.Rows.Count);
        Assert.AreEqual(0, diff.CreateSummary().ModifiedCellCount);
    }

    private static ExcelSheet CreateSheet(params string[][] values)
    {
        var sheet = new ExcelSheet();
        for (var rowIndex = 0; rowIndex < values.Length; rowIndex++)
        {
            var cells = values[rowIndex]
                .Select((value, columnIndex) => new ExcelCell(value, columnIndex, rowIndex));
            sheet.Rows.Add(rowIndex, new ExcelRow(rowIndex, cells));
        }

        return sheet;
    }

    private static ExcelSheetDiffConfig HeaderConfig() => new()
    {
        SrcHeaderIndex = 0,
        DstHeaderIndex = 0,
    };

    private static string[] Snapshot(ExcelSheet sheet) => sheet.Rows.Values
        .Select(row => string.Join("\u001f", row.Cells.Select(cell => cell.Value)))
        .ToArray();
}

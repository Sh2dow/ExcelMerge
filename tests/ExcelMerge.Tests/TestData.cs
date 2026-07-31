using ExcelMerge.Domain;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

internal static class TestData
{
    public static CellValue Text(string value, string? displayText = null) =>
        CellValue.FromText(value, displayText);

    public static CellValue Number(double value, string? displayText = null) =>
        CellValue.FromNumber(value, displayText);

    public static RowRecord Row(
        int rowIndex,
        params (int ColumnIndex, CellValue Value)[] cells) =>
        new(
            rowIndex,
            cells.Select(cell => new CellRecord(
                new CellAddress(rowIndex, cell.ColumnIndex),
                cell.Value)).ToArray());

    public static WorksheetSnapshot Sheet(params RowRecord[] rows) =>
        Sheet(new SheetMetadata("rId1", "Sheet1", 0), rows);

    public static WorksheetSnapshot Sheet(SheetMetadata metadata, params RowRecord[] rows) =>
        new(metadata, rows);

    public static WorkbookMetadata Workbook(string displayName = "book.xlsx") =>
        new(displayName, WorkbookFormat.Xlsx, [new SheetMetadata("rId1", "Sheet1", 0)]);

    public static void AssertRowsEqual(in RowRecord expected, in RowRecord actual)
    {
        Assert.AreEqual(expected.RowIndex, actual.RowIndex);
        Assert.AreEqual(expected.Height, actual.Height);
        Assert.AreEqual(expected.IsHidden, actual.IsHidden);
        Assert.AreEqual(expected.StyleIndex, actual.StyleIndex);
        Assert.AreEqual(expected.CellCount, actual.CellCount);

        for (var index = 0; index < expected.CellCount; index++)
        {
            Assert.AreEqual(expected.Cells.Span[index], actual.Cells.Span[index]);
        }
    }
}

internal sealed class TestDirectory : IDisposable
{
    public TestDirectory()
    {
        Path = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            "ExcelMerge.Tests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    public string Path { get; }

    public string GetPath(string fileName) => System.IO.Path.Combine(Path, fileName);

    public void Dispose()
    {
        if (Directory.Exists(Path))
        {
            Directory.Delete(Path, recursive: true);
        }
    }
}

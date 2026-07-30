using ExcelMerge.Avalonia;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class DiffRowTests
{
    [TestMethod]
    public void DiffRowCachesStructuralLookupsWhileResolutionRemainsDynamic()
    {
        var source = new ExcelCell("local", 2, 4);
        var destination = new ExcelCell("remote", 2, 4);
        var rowDiff = new ExcelRowDiff(0);
        rowDiff.CreateCell(source, destination, 2, ExcelCellStatus.Modified);
        var row = new DiffRow(rowDiff, new Dictionary<(int Row, int Column), string>
        {
            [(4, 2)] = "base",
        });
        var cell = row.Cells.Single();
        var cellChanges = new List<string?>();
        var rowChanges = new List<string?>();
        cell.PropertyChanged += (_, args) => cellChanges.Add(args.PropertyName);
        row.PropertyChanged += (_, args) => rowChanges.Add(args.PropertyName);

        Assert.AreSame(cell, row.GetCell(2));
        Assert.IsNull(row.GetCell(1));
        Assert.AreSame(row.ConflictCells, row.ConflictCells);
        Assert.IsTrue(row.HasConflict);
        Assert.IsFalse(row.IsResolved);
        Assert.AreSame(cell.InlineDiff, cell.InlineDiff);

        row.Resolve(MergeResolution.Both);

        Assert.IsTrue(row.IsResolved);
        Assert.AreEqual("1 x2", row.DisplayIndex);
        CollectionAssert.Contains(cellChanges, nameof(DiffCell.LocalBackground));
        CollectionAssert.Contains(cellChanges, nameof(DiffCell.RemoteDisplayValue));
        CollectionAssert.Contains(rowChanges, nameof(DiffRow.DisplayIndex));
    }
}

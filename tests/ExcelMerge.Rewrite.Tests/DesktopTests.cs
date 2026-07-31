using ExcelMerge.Application;
using ExcelMerge.Desktop;
using ExcelMerge.Engine;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Rewrite.Tests;

[TestClass]
public sealed class DesktopTests
{
    [TestMethod]
    public async Task Merge_grid_accepts_zero_conflict_rows_with_negative_start_sentinel()
    {
        var baseSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("unchanged"))),
            TestData.Row(1, (0, TestData.Text("base"))));
        var localSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("unchanged"))),
            TestData.Row(1, (0, TestData.Text("local"))));
        var remoteSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("unchanged"))),
            TestData.Row(1, (0, TestData.Text("remote"))));
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet,
            new ThreeWayMergeOptions { SheetGroupId = "sheet" });
        var result = new MergeSheetResult(
            "sheet",
            baseSheet,
            localSheet,
            remoteSheet,
            merge);

        Assert.AreEqual(-1, merge.ViewRows.Span[0].ConflictStartIndex);
        Assert.AreEqual(0, merge.ViewRows.Span[0].ConflictCount);

        var document = GridDocument.FromMerge(result, new HashSet<long>());
        Assert.AreEqual(2, document.RowCount);
        Assert.AreEqual(GridRowVisualState.Unchanged, document.Rows[0].State);
        Assert.AreEqual(GridRowVisualState.Conflict, document.Rows[1].State);

        var resolved = GridDocument.FromMerge(
            result,
            new HashSet<long> { merge.Conflicts.Span[0].Id },
            hideUnchanged: true);
        Assert.AreEqual(1, resolved.RowCount);
        Assert.AreEqual(GridRowVisualState.Resolved, resolved.Rows[0].State);
    }
}

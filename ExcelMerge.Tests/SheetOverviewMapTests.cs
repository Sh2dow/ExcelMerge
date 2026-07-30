using ExcelMerge;
using ExcelMerge.Avalonia;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class SheetOverviewMapTests
{
    [TestMethod]
    public void ChangedCellsAreYellowMarkers()
    {
        Assert.AreEqual(
            SheetOverviewMarkerKind.Changed,
            SheetOverviewMap.Classify(ExcelCellStatus.Modified, false, MergeResolution.Unresolved));
        Assert.AreEqual(
            SheetOverviewMarkerKind.None,
            SheetOverviewMap.Classify(ExcelCellStatus.None, false, MergeResolution.Unresolved));
    }

    [TestMethod]
    public void ConflictStateOverridesTheChangeMarker()
    {
        Assert.AreEqual(
            SheetOverviewMarkerKind.UnresolvedConflict,
            SheetOverviewMap.Classify(ExcelCellStatus.Modified, true, MergeResolution.Unresolved));
        Assert.AreEqual(
            SheetOverviewMarkerKind.ResolvedConflict,
            SheetOverviewMap.Classify(ExcelCellStatus.Modified, true, MergeResolution.Local));
    }

    [TestMethod]
    public void FullSheetViewportKeepsItsHeightWhileMoving()
    {
        var rows = Enumerable.Range(0, 100).ToList();

        var first = MainWindow.MapVisibleViewportToSheet(new NormalizedViewport(.1, .3), rows, 100);
        var second = MainWindow.MapVisibleViewportToSheet(new NormalizedViewport(.6, .8), rows, 100);

        Assert.AreEqual(.2, first.End - first.Start, 0.000001);
        Assert.AreEqual(first.End - first.Start, second.End - second.Start, 0.000001);
    }

    [TestMethod]
    public void FilteredViewportMapsBackToSheetRows()
    {
        var viewport = MainWindow.MapVisibleViewportToSheet(
            new NormalizedViewport(0, 1),
            new[] { 9, 49, 89 },
            100);

        Assert.AreEqual(.09, viewport.Start, 0.000001);
        Assert.AreEqual(.9, viewport.End, 0.000001);
    }
}

using ExcelMerge.Avalonia;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class TableMetricsTests
{
    [TestMethod]
    public void DefaultFontUsesCompactCellMetrics()
    {
        Assert.AreEqual(120d, MainWindow.CellWidth(11));
        Assert.AreEqual(28d, MainWindow.CellHeight(11));
    }

    [TestMethod]
    public void CellMetricsScaleWithFontSize()
    {
        Assert.AreEqual(96d, MainWindow.CellWidth(8));
        Assert.AreEqual(24d, MainWindow.CellHeight(8));
        Assert.AreEqual(192d, MainWindow.CellWidth(20));
        Assert.AreEqual(46d, MainWindow.CellHeight(20));
    }
}

using ExcelMerge.Avalonia;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class SplitScrollSynchronizerTests
{
    [TestMethod]
    public void MapAlignedOffsetKeepsTheSameContentOffset()
    {
        var result = SplitScrollSynchronizer.MapAlignedOffset(128, 0, 500, 0, 700);

        Assert.AreEqual(128, result);
    }

    [TestMethod]
    public void MapAlignedOffsetAlignsBothEndsWhenRangesDiffer()
    {
        Assert.AreEqual(0, SplitScrollSynchronizer.MapAlignedOffset(0, 0, 500, 0, 700));
        Assert.AreEqual(700, SplitScrollSynchronizer.MapAlignedOffset(500, 0, 500, 0, 700));
    }

    [TestMethod]
    public void MapAlignedOffsetClampsToTheTargetRange()
    {
        var result = SplitScrollSynchronizer.MapAlignedOffset(400, 0, 500, 0, 250);

        Assert.AreEqual(250, result);
    }
}

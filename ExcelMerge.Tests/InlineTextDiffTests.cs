using ExcelMerge.Avalonia;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class InlineTextDiffTests
{
    [TestMethod]
    public void CreateHighlightsCharacterReplacement()
    {
        var result = InlineTextDiff.Create("abc", "axc");

        Assert.AreEqual("abc", Reconstruct(result.Local));
        Assert.AreEqual("axc", Reconstruct(result.Remote));
        Assert.AreEqual("b", string.Concat(result.Local.Where(segment => segment.IsChanged).Select(segment => segment.Text)));
        Assert.AreEqual("x", string.Concat(result.Remote.Where(segment => segment.IsChanged).Select(segment => segment.Text)));
    }

    [TestMethod]
    public void CreateHighlightsOnlyInsertedSide()
    {
        var result = InlineTextDiff.Create("abc", "abXc");

        Assert.IsFalse(result.Local.Any(segment => segment.IsChanged));
        Assert.AreEqual("X", string.Concat(result.Remote.Where(segment => segment.IsChanged).Select(segment => segment.Text)));
    }

    [TestMethod]
    public void CreatePreservesMultilineValues()
    {
        const string local = "first\nlocal value\nlast";
        const string remote = "first\nremote value\nlast";

        var result = InlineTextDiff.Create(local, remote);

        Assert.AreEqual(local, Reconstruct(result.Local));
        Assert.AreEqual(remote, Reconstruct(result.Remote));
        Assert.IsTrue(result.Local.Any(segment => segment.IsChanged));
        Assert.IsTrue(result.Remote.Any(segment => segment.IsChanged));
    }

    [TestMethod]
    public void CreateDoesNotHighlightEqualValues()
    {
        var result = InlineTextDiff.Create("unchanged", "unchanged");

        Assert.AreEqual("unchanged", Reconstruct(result.Local));
        Assert.IsFalse(result.Local.Any(segment => segment.IsChanged));
        Assert.IsFalse(result.Remote.Any(segment => segment.IsChanged));
    }

    [TestMethod]
    public void CreateFallsBackSafelyForVeryLongValues()
    {
        var local = new string('L', 10001);
        var remote = new string('R', 10001);

        var result = InlineTextDiff.Create(local, remote);

        Assert.AreEqual(local, Reconstruct(result.Local));
        Assert.AreEqual(remote, Reconstruct(result.Remote));
        Assert.IsTrue(result.Local.All(segment => segment.IsChanged));
        Assert.IsTrue(result.Remote.All(segment => segment.IsChanged));
    }

    private static string Reconstruct(IEnumerable<InlineDiffSegment> segments) =>
        string.Concat(segments.Select(segment => segment.Text));
}

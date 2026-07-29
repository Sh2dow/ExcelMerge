using ExcelMerge.Avalonia;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class CommandLineOptionsTests
{
    [TestMethod]
    public void DiffAcceptsLegacySourceAndDestinationOptions()
    {
        var options = CommandLineOptions.Parse(new[]
        {
            "diff", "-s", "local.xlsx", "-d", "remote.xlsx",
        });

        Assert.AreEqual(ApplicationMode.Diff, options.Mode);
        Assert.AreEqual("local.xlsx", options.LocalPath);
        Assert.AreEqual("remote.xlsx", options.RemotePath);
        Assert.IsNull(options.BasePath);
    }

    [TestMethod]
    public void MergeAcceptsBaseLocalAndRemoteOptions()
    {
        var options = CommandLineOptions.Parse(new[]
        {
            "merge", "-b", "base.xlsx", "-l", "local.xlsx", "-r", "remote.xlsx",
        });

        Assert.AreEqual(ApplicationMode.Merge, options.Mode);
        Assert.AreEqual("base.xlsx", options.BasePath);
        Assert.AreEqual("local.xlsx", options.LocalPath);
        Assert.AreEqual("remote.xlsx", options.RemotePath);
    }

    [TestMethod]
    public void MergeAcceptsLongEqualsOptions()
    {
        var options = CommandLineOptions.Parse(new[]
        {
            "merge", "--base-path=base.xlsx", "--local-path=local.xlsx", "--remote-path=remote.xlsx",
        });

        Assert.AreEqual("base.xlsx", options.BasePath);
        Assert.AreEqual("local.xlsx", options.LocalPath);
        Assert.AreEqual("remote.xlsx", options.RemotePath);
    }

    [TestMethod]
    public void MergeAcceptsThreePositionalPaths()
    {
        var options = CommandLineOptions.Parse(new[]
        {
            "merge", "base.xlsx", "local.xlsx", "remote.xlsx",
        });

        Assert.AreEqual(ApplicationMode.Merge, options.Mode);
        Assert.AreEqual("base.xlsx", options.BasePath);
        Assert.AreEqual("local.xlsx", options.LocalPath);
        Assert.AreEqual("remote.xlsx", options.RemotePath);
    }

    [TestMethod]
    public void MissingOptionValueIsRejected()
    {
        Assert.ThrowsException<ArgumentException>(() =>
            CommandLineOptions.Parse(new[] { "merge", "--base-path" }));
    }

    [TestMethod]
    public void DiffIgnoresLegacyWpfIntegrationOptions()
    {
        var options = CommandLineOptions.Parse(new[]
        {
            "diff", "-s", "local.xlsx", "-d", "remote.xlsx",
            "-c", "WinMerge", "-i", "-w", "-v", "-e", "empty", "-k",
        });

        Assert.AreEqual("local.xlsx", options.LocalPath);
        Assert.AreEqual("remote.xlsx", options.RemotePath);
    }
}

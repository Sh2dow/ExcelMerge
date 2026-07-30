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
    public void MergeDriverAcceptsGitDriverArguments()
    {
        var options = CommandLineOptions.Parse(new[]
        {
            "merge-driver", "base.tmp", "ours.tmp", "theirs.tmp", "7", "workbooks/report.xlsx",
        });

        Assert.AreEqual(ApplicationMode.Merge, options.Mode);
        Assert.AreEqual("base.tmp", options.BasePath);
        Assert.AreEqual("ours.tmp", options.LocalPath);
        Assert.AreEqual("theirs.tmp", options.RemotePath);
        Assert.AreEqual("ours.tmp", options.OutputPath);
        Assert.AreEqual("workbooks/report.xlsx", options.RepositoryPath);
        Assert.AreEqual(7, options.ConflictMarkerSize);
        Assert.IsTrue(options.IsMergeDriver);
    }

    [TestMethod]
    public void MergeDriverRemovesNestedShellQuotesFromGitPaths()
    {
        var options = CommandLineOptions.Parse(new[]
        {
            "merge-driver", ".merge_base", ".merge_ours", ".merge_theirs", "7", "'workbooks/report.xlsx'",
        });

        Assert.AreEqual("workbooks/report.xlsx", options.RepositoryPath);
    }

    [TestMethod]
    public void MergeAcceptsNamedGitDriverOptions()
    {
        var options = CommandLineOptions.Parse(new[]
        {
            "merge",
            "--base", "base.tmp",
            "--ours", "ours.tmp",
            "--theirs", "theirs.tmp",
            "--output", "ours.tmp",
            "--path", "workbooks/report.xlsx",
            "--marker-size", "7",
        });

        Assert.AreEqual("ours.tmp", options.LocalPath);
        Assert.AreEqual("theirs.tmp", options.RemotePath);
        Assert.AreEqual("ours.tmp", options.OutputPath);
        Assert.AreEqual("workbooks/report.xlsx", options.RepositoryPath);
        Assert.IsTrue(options.IsMergeDriver);
    }

    [TestMethod]
    public void MergeDriverRejectsInvalidMarkerSize()
    {
        Assert.ThrowsException<ArgumentException>(() => CommandLineOptions.Parse(new[]
        {
            "merge-driver", "base.tmp", "ours.tmp", "theirs.tmp", "invalid", "report.xlsx",
        }));
    }

    [TestMethod]
    public void MergeDriverReturnsFailureUntilUserConfirmsResult()
    {
        var options = CommandLineOptions.Parse(new[]
        {
            "merge-driver", "base.tmp", "ours.tmp", "theirs.tmp", "7", "report.xlsx",
        });

        Assert.AreEqual(1, Program.ResolveExitCode(options, 0, mergeDriverCompleted: false));
        Assert.AreEqual(0, Program.ResolveExitCode(options, 0, mergeDriverCompleted: true));
        Assert.AreEqual(1, Program.ResolveExitCode(options, 3, mergeDriverCompleted: false));
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

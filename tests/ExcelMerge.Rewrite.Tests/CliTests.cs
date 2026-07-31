using ExcelMerge.Cli;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Rewrite.Tests;

[TestClass]
public sealed class CliTests
{
    [TestMethod]
    public void Parser_accepts_positional_alias_and_git_driver_contracts()
    {
        var positional = CliCommandLine.Parse(["diff", "local.xlsx", "remote.xlsx"]);
        var aliases = CliCommandLine.Parse([
            "merge",
            "--base", "base.csv",
            "--ours", "local.csv",
            "--theirs", "remote.csv",
            "--output", "result.csv",
            "--marker-size", "7",
            "--repository-path", "data.csv",
        ]);
        var driver = CliCommandLine.Parse([
            "merge-driver",
            "\"base file\"",
            "'ours file'",
            "theirs",
            "9",
            "folder/data.xlsx",
        ]);

        Assert.AreEqual(CliCommandKind.Diff, positional.Kind);
        Assert.AreEqual("local.xlsx", positional.LocalPath);
        Assert.AreEqual(CliCommandKind.Merge, aliases.Kind);
        Assert.AreEqual("result.csv", aliases.OutputPath);
        Assert.AreEqual(7, aliases.MarkerSize);
        Assert.AreEqual(CliCommandKind.MergeDriver, driver.Kind);
        Assert.AreEqual("base file", driver.BasePath);
        Assert.AreEqual("ours file", driver.LocalPath);
        Assert.AreEqual(driver.LocalPath, driver.OutputPath);
        Assert.AreEqual(9, driver.MarkerSize);
        var leadingDash = CliCommandLine.Parse([
            "merge-driver", "-base", "-ours", "-theirs", "7", "-report.xlsx",
        ]);
        Assert.AreEqual("-report.xlsx", leadingDash.RepositoryPath);
    }

    [TestMethod]
    public void Parser_rejects_missing_output_unknown_options_and_invalid_marker()
    {
        Assert.ThrowsException<CliUsageException>(() =>
            CliCommandLine.Parse(["merge", "base.csv", "local.csv", "remote.csv"]));
        Assert.ThrowsException<CliUsageException>(() =>
            CliCommandLine.Parse(["diff", "--unknown", "value", "a.csv", "b.csv"]));
        Assert.ThrowsException<CliUsageException>(() =>
            CliCommandLine.Parse(["merge-driver", "base", "ours", "theirs", "0", "data.csv"]));
    }

    [TestMethod]
    public async Task Runner_returns_conventional_diff_exit_codes_without_diagnostics()
    {
        using var temporaryDirectory = new TestDirectory();
        var localPath = temporaryDirectory.GetPath("local.csv");
        var samePath = temporaryDirectory.GetPath("same.csv");
        var changedPath = temporaryDirectory.GetPath("changed.csv");
        await File.WriteAllTextAsync(localPath, "value");
        await File.WriteAllTextAsync(samePath, "value");
        await File.WriteAllTextAsync(changedPath, "changed");
        using var output = new StringWriter();
        using var error = new StringWriter();

        var same = await CliRunner.RunAsync(
            ["diff", localPath, samePath],
            output,
            error);
        var changed = await CliRunner.RunAsync(
            ["diff", localPath, changedPath],
            output,
            error);

        Assert.AreEqual(CliRunner.SuccessExitCode, same);
        Assert.AreEqual(CliRunner.DifferenceOrFailureExitCode, changed);
        Assert.AreEqual(string.Empty, output.ToString());
        Assert.AreEqual(string.Empty, error.ToString());
    }

    [TestMethod]
    public async Task Runner_writes_automatic_merge_and_rejects_unresolved_merge()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = temporaryDirectory.GetPath("base.csv");
        var localPath = temporaryDirectory.GetPath("local.csv");
        var remotePath = temporaryDirectory.GetPath("remote.csv");
        var resultPath = temporaryDirectory.GetPath("result.csv");
        await File.WriteAllTextAsync(basePath, "base");
        await File.WriteAllTextAsync(localPath, "base");
        await File.WriteAllTextAsync(remotePath, "remote");
        using var output = new StringWriter();
        using var error = new StringWriter();

        var success = await CliRunner.RunAsync(
            ["merge", basePath, localPath, remotePath, "--output", resultPath],
            output,
            error);

        Assert.AreEqual(CliRunner.SuccessExitCode, success);
        Assert.AreEqual("remote\r\n", await File.ReadAllTextAsync(resultPath));
        await File.WriteAllTextAsync(localPath, "LOCAL");
        await File.WriteAllTextAsync(remotePath, "REMOTE");
        File.Delete(resultPath);
        var failure = await CliRunner.RunAsync(
            ["merge", basePath, localPath, remotePath, "--output", resultPath],
            output,
            error);

        Assert.AreEqual(CliRunner.DifferenceOrFailureExitCode, failure);
        Assert.IsFalse(File.Exists(resultPath));
        StringAssert.Contains(error.ToString(), "unresolved conflict");
    }

    [TestMethod]
    public async Task Git_merge_driver_stages_extensionless_inputs_and_replaces_ours()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = temporaryDirectory.GetPath("base-temp");
        var oursPath = temporaryDirectory.GetPath("ours-temp");
        var theirsPath = temporaryDirectory.GetPath("theirs-temp");
        await File.WriteAllTextAsync(basePath, "base");
        await File.WriteAllTextAsync(oursPath, "base");
        await File.WriteAllTextAsync(theirsPath, "remote");
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = await CliRunner.RunAsync(
            ["merge-driver", basePath, oursPath, theirsPath, "7", "sheets/data.csv"],
            output,
            error);

        Assert.AreEqual(CliRunner.SuccessExitCode, exitCode);
        Assert.AreEqual("remote\r\n", await File.ReadAllTextAsync(oursPath));
        Assert.AreEqual(string.Empty, output.ToString());
        Assert.AreEqual(string.Empty, error.ToString());
    }

    [TestMethod]
    public async Task Runner_reports_usage_errors_on_stderr_with_exit_code_two()
    {
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = await CliRunner.RunAsync(
            ["merge", "only-one-path"],
            output,
            error);

        Assert.AreEqual(CliRunner.UsageExitCode, exitCode);
        Assert.AreEqual(string.Empty, output.ToString());
        StringAssert.Contains(error.ToString(), "requires BASE, LOCAL, REMOTE");
        StringAssert.Contains(error.ToString(), "Usage:");
    }
}

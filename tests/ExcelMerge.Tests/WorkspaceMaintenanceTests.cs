using ExcelMerge.Storage;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class WorkspaceMaintenanceTests
{
    [TestMethod]
    public async Task Cleanup_removes_only_stale_owned_inactive_workspaces()
    {
        using var temporaryDirectory = new TestDirectory();
        var options = CreateOptions(temporaryDirectory.Path);

        string stalePath;
        await using (var stale = new Workspace(options))
        {
            stalePath = stale.DirectoryPath;
        }

        string youngPath;
        await using (var young = new Workspace(options))
        {
            youngPath = young.DirectoryPath;
        }

        await using var active = new Workspace(options);
        var foreignPath = Path.Combine(temporaryDirectory.Path, "maintenance-foreign");
        Directory.CreateDirectory(foreignPath);
        await File.WriteAllTextAsync(
            Path.Combine(foreignPath, options.OwnershipMarkerFileName),
            "another-application");

        var oldTimestamp = DateTime.UtcNow - TimeSpan.FromDays(2);
        Directory.SetLastWriteTimeUtc(stalePath, oldTimestamp);
        Directory.SetLastWriteTimeUtc(active.DirectoryPath, oldTimestamp);
        Directory.SetLastWriteTimeUtc(foreignPath, oldTimestamp);

        var result = await WorkspaceMaintenance.CleanupStaleAsync(
            options,
            minimumAge: TimeSpan.FromDays(1));

        Assert.IsFalse(Directory.Exists(stalePath));
        Assert.IsTrue(Directory.Exists(youngPath));
        Assert.IsTrue(Directory.Exists(active.DirectoryPath));
        Assert.IsTrue(Directory.Exists(foreignPath));
        Assert.AreEqual(4, result.ScannedCount);
        Assert.AreEqual(1, result.DeletedCount);
        Assert.AreEqual(1, result.ActiveCount);
        Assert.AreEqual(1, result.TooYoungCount);
        Assert.AreEqual(1, result.ForeignCount);
        Assert.AreEqual(0, result.FailedCount);
    }

    [TestMethod]
    public async Task DeleteOnDispose_false_leaves_an_owned_workspace_for_later_maintenance()
    {
        using var temporaryDirectory = new TestDirectory();
        var options = CreateOptions(temporaryDirectory.Path);
        string workspacePath;
        string markerPath;
        string leasePath;

        await using (var workspace = new Workspace(options))
        {
            workspacePath = workspace.DirectoryPath;
            markerPath = workspace.OwnershipMarkerPath;
            leasePath = workspace.ActiveLeasePath;
        }

        Assert.IsTrue(Directory.Exists(workspacePath));
        Assert.IsTrue(File.Exists(markerPath));
        Assert.IsTrue(File.Exists(leasePath));

        var result = await WorkspaceMaintenance.CleanupStaleAsync(
            options,
            minimumAge: TimeSpan.Zero);

        Assert.AreEqual(1, result.DeletedCount);
        Assert.IsFalse(Directory.Exists(workspacePath));
    }

    [TestMethod]
    public void Temporary_disk_estimation_and_evaluation_are_overflow_safe()
    {
        var estimate = TemporaryDiskPreflight.Estimate(
            sourceBytes: 100,
            temporaryCopyCount: 3,
            reserveBytes: 50);

        Assert.AreEqual(300L, estimate.TemporaryBytes);
        Assert.AreEqual(50L, estimate.ReserveBytes);
        Assert.AreEqual(350L, estimate.RequiredBytes);
        Assert.IsFalse(estimate.Overflowed);
        Assert.AreEqual(
            TemporaryDiskPreflightStatus.Sufficient,
            TemporaryDiskPreflight.Evaluate(estimate, availableBytes: 350).Status);
        Assert.AreEqual(
            TemporaryDiskPreflightStatus.Insufficient,
            TemporaryDiskPreflight.Evaluate(estimate, availableBytes: 349).Status);
        Assert.AreEqual(
            TemporaryDiskPreflightStatus.UnknownVolume,
            TemporaryDiskPreflight.Evaluate(estimate, availableBytes: null).Status);

        var overflow = TemporaryDiskPreflight.Estimate(
            sourceBytes: long.MaxValue,
            temporaryCopyCount: 2,
            reserveBytes: 1);
        Assert.IsTrue(overflow.Overflowed);
        Assert.AreEqual(long.MaxValue, overflow.RequiredBytes);
        Assert.AreEqual(
            TemporaryDiskPreflightStatus.EstimateOverflow,
            TemporaryDiskPreflight.Evaluate(overflow, long.MaxValue).Status);

        var reserveOverflow = TemporaryDiskPreflight.Estimate(
            sourceBytes: long.MaxValue - 1,
            temporaryCopyCount: 1,
            reserveBytes: 2);
        Assert.AreEqual(long.MaxValue - 1, reserveOverflow.TemporaryBytes);
        Assert.IsTrue(reserveOverflow.Overflowed);
    }

    [TestMethod]
    public async Task Cleanup_honors_pre_cancellation_without_deleting_workspaces()
    {
        using var temporaryDirectory = new TestDirectory();
        var options = CreateOptions(temporaryDirectory.Path);
        string workspacePath;

        await using (var workspace = new Workspace(options))
        {
            workspacePath = workspace.DirectoryPath;
        }

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsExceptionAsync<OperationCanceledException>(() =>
            WorkspaceMaintenance.CleanupStaleAsync(
                options,
                minimumAge: TimeSpan.Zero,
                cancellation.Token));
        Assert.IsTrue(Directory.Exists(workspacePath));
    }

    private static WorkspaceOptions CreateOptions(string baseDirectory) =>
        new()
        {
            BaseDirectory = baseDirectory,
            DirectoryPrefix = "maintenance-",
            OwnershipMarkerFileName = ".owner",
            OwnershipMarkerValue = "ExcelMerge.Tests.Workspace/v1",
            ActiveLeaseFileName = ".lease",
            DeleteOnDispose = false,
            CleanupRetryDelay = TimeSpan.FromMilliseconds(1),
        };
}

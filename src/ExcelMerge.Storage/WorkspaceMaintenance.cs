using System.Text;

namespace ExcelMerge.Storage;

/// <summary>
/// Summarizes one stale-workspace maintenance pass.
/// </summary>
public sealed record WorkspaceMaintenanceResult(
    int ScannedCount,
    int DeletedCount,
    int ActiveCount,
    int TooYoungCount,
    int ForeignCount,
    int FailedCount,
    int RetryCount)
{
    public int RetainedCount => ActiveCount + TooYoungCount + ForeignCount + FailedCount;
}

/// <summary>
/// Removes abandoned workspaces that can be positively identified as belonging to this application.
/// </summary>
public static class WorkspaceMaintenance
{
    public static TimeSpan DefaultMinimumAge { get; } = TimeSpan.FromDays(1);

    /// <summary>
    /// Deletes owned, inactive workspaces older than <paramref name="minimumAge"/>.
    /// Directories without the exact configured marker, and directories with an active lease,
    /// are retained.
    /// </summary>
    public static async Task<WorkspaceMaintenanceResult> CleanupStaleAsync(
        WorkspaceOptions? options = null,
        TimeSpan? minimumAge = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var settings = (options ?? new WorkspaceOptions()).Validate();
        var requiredAge = minimumAge ?? DefaultMinimumAge;
        if (requiredAge < TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(minimumAge), "The minimum workspace age cannot be negative.");
        }

        if (!Directory.Exists(settings.BaseDirectory))
        {
            return new WorkspaceMaintenanceResult(0, 0, 0, 0, 0, 0, 0);
        }

        var expectedMarker = Encoding.UTF8.GetBytes(settings.OwnershipMarkerValue);
        var comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        var now = DateTime.UtcNow;
        var scanned = 0;
        var deleted = 0;
        var active = 0;
        var tooYoung = 0;
        var foreign = 0;
        var failed = 0;
        var retries = 0;

        foreach (var directoryPath in Directory.EnumerateDirectories(settings.BaseDirectory))
        {
            cancellationToken.ThrowIfCancellationRequested();

            var directoryName = Path.GetFileName(directoryPath);
            if (!directoryName.StartsWith(settings.DirectoryPrefix, comparison))
            {
                continue;
            }

            scanned++;
            var ownership = CheckOwnership(directoryPath, settings, expectedMarker);
            if (ownership == OwnershipStatus.Foreign)
            {
                foreign++;
                continue;
            }

            if (ownership == OwnershipStatus.Failed)
            {
                failed++;
                continue;
            }

            var lease = ProbeLease(directoryPath, settings.ActiveLeaseFileName);
            if (lease == LeaseStatus.Active)
            {
                active++;
                continue;
            }

            if (lease == LeaseStatus.Failed)
            {
                failed++;
                continue;
            }

            DateTime lastWriteTimeUtc;
            try
            {
                lastWriteTimeUtc = Directory.GetLastWriteTimeUtc(directoryPath);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                failed++;
                continue;
            }

            if (lastWriteTimeUtc > now || now - lastWriteTimeUtc < requiredAge)
            {
                tooYoung++;
                continue;
            }

            var deleteResult = await DeleteWithRetriesAsync(
                directoryPath,
                settings,
                expectedMarker,
                cancellationToken).ConfigureAwait(false);
            retries += deleteResult.RetryCount;

            switch (deleteResult.Status)
            {
                case DeleteStatus.Deleted:
                    deleted++;
                    break;
                case DeleteStatus.Active:
                    active++;
                    break;
                case DeleteStatus.Foreign:
                    foreign++;
                    break;
                default:
                    failed++;
                    break;
            }
        }

        return new WorkspaceMaintenanceResult(
            scanned,
            deleted,
            active,
            tooYoung,
            foreign,
            failed,
            retries);
    }

    private static async Task<DeleteResult> DeleteWithRetriesAsync(
        string directoryPath,
        WorkspaceSettings settings,
        byte[] expectedMarker,
        CancellationToken cancellationToken)
    {
        for (var attempt = 0; ; attempt++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                var attributes = File.GetAttributes(directoryPath);
                if ((attributes & FileAttributes.Directory) == 0 ||
                    (attributes & FileAttributes.ReparsePoint) != 0)
                {
                    return new DeleteResult(DeleteStatus.Foreign, attempt);
                }

                if (attempt == 0)
                {
                    var ownership = CheckOwnership(directoryPath, settings, expectedMarker);
                    if (ownership != OwnershipStatus.Owned)
                    {
                        return new DeleteResult(
                            ownership == OwnershipStatus.Foreign ? DeleteStatus.Foreign : DeleteStatus.Failed,
                            attempt);
                    }

                    var lease = ProbeLease(directoryPath, settings.ActiveLeaseFileName);
                    if (lease != LeaseStatus.Inactive)
                    {
                        return new DeleteResult(
                            lease == LeaseStatus.Active ? DeleteStatus.Active : DeleteStatus.Failed,
                            attempt);
                    }
                }

                Workspace.MakeTreeWritable(directoryPath, cancellationToken);
                Directory.Delete(directoryPath, recursive: true);
                return new DeleteResult(DeleteStatus.Deleted, attempt);
            }
            catch (Exception exception) when (
                exception is FileNotFoundException or DirectoryNotFoundException)
            {
                return new DeleteResult(DeleteStatus.Deleted, attempt);
            }
            catch (Exception exception) when (
                exception is IOException or UnauthorizedAccessException &&
                attempt < settings.CleanupRetryCount)
            {
                if (settings.CleanupRetryDelay > TimeSpan.Zero)
                {
                    await Task.Delay(settings.CleanupRetryDelay, cancellationToken).ConfigureAwait(false);
                }
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                return new DeleteResult(DeleteStatus.Failed, attempt);
            }
        }
    }

    private static OwnershipStatus CheckOwnership(
        string directoryPath,
        WorkspaceSettings settings,
        byte[] expectedMarker)
    {
        try
        {
            var directoryAttributes = File.GetAttributes(directoryPath);
            if ((directoryAttributes & FileAttributes.ReparsePoint) != 0)
            {
                return OwnershipStatus.Foreign;
            }

            var markerPath = Path.Combine(directoryPath, settings.OwnershipMarkerFileName);
            var markerAttributes = File.GetAttributes(markerPath);
            if ((markerAttributes & (FileAttributes.Directory | FileAttributes.ReparsePoint)) != 0)
            {
                return OwnershipStatus.Foreign;
            }

            using var marker = new FileStream(
                markerPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read);
            if (marker.Length != expectedMarker.Length)
            {
                return OwnershipStatus.Foreign;
            }

            var actualMarker = new byte[expectedMarker.Length];
            marker.ReadExactly(actualMarker);
            return actualMarker.AsSpan().SequenceEqual(expectedMarker)
                ? OwnershipStatus.Owned
                : OwnershipStatus.Foreign;
        }
        catch (Exception exception) when (
            exception is FileNotFoundException or DirectoryNotFoundException)
        {
            return OwnershipStatus.Foreign;
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return OwnershipStatus.Failed;
        }
    }

    private static LeaseStatus ProbeLease(string directoryPath, string leaseFileName)
    {
        var leasePath = Path.Combine(directoryPath, leaseFileName);
        try
        {
            var attributes = File.GetAttributes(leasePath);
            if ((attributes & (FileAttributes.Directory | FileAttributes.ReparsePoint)) != 0)
            {
                return LeaseStatus.Failed;
            }

            using var lease = new FileStream(
                leasePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.None);
            return LeaseStatus.Inactive;
        }
        catch (Exception exception) when (
            exception is FileNotFoundException or DirectoryNotFoundException)
        {
            return LeaseStatus.Inactive;
        }
        catch (UnauthorizedAccessException)
        {
            return LeaseStatus.Failed;
        }
        catch (IOException)
        {
            // Sharing violations are reported as IOException on all supported platforms.
            return LeaseStatus.Active;
        }
    }

    private enum OwnershipStatus
    {
        Owned,
        Foreign,
        Failed,
    }

    private enum LeaseStatus
    {
        Inactive,
        Active,
        Failed,
    }

    private enum DeleteStatus
    {
        Deleted,
        Active,
        Foreign,
        Failed,
    }

    private readonly record struct DeleteResult(DeleteStatus Status, int RetryCount);
}

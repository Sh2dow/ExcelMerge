namespace ExcelMerge.Storage;

/// <summary>
/// Configures temporary workspaces and the row stores created in them.
/// </summary>
public sealed record WorkspaceOptions
{
    public const long DefaultChunkSizeBytes = 256L * 1024 * 1024;
    public const int DefaultMaximumRowSizeBytes = 128 * 1024 * 1024;
    public const long DefaultRowCacheByteLimit = 64L * 1024 * 1024;
    public const string DefaultOwnershipMarkerFileName = ".excelmerge-workspace";
    public const string DefaultOwnershipMarkerValue = "ExcelMerge.Storage.Workspace/v1";
    public const string DefaultActiveLeaseFileName = ".excelmerge-workspace.active";

    /// <summary>
    /// Gets the directory in which the unique workspace directory is created.
    /// The operating system temporary directory is used when this is null.
    /// </summary>
    public string? BaseDirectory { get; init; }

    /// <summary>
    /// Gets the prefix used for the unique workspace directory name.
    /// </summary>
    public string DirectoryPrefix { get; init; } = "excelmerge-";

    /// <summary>
    /// Gets the file name used to mark a directory as owned by this application.
    /// </summary>
    public string OwnershipMarkerFileName { get; init; } = DefaultOwnershipMarkerFileName;

    /// <summary>
    /// Gets the exact marker value required before startup maintenance may delete a directory.
    /// Applications that customize this value must use the same value for creation and maintenance.
    /// </summary>
    public string OwnershipMarkerValue { get; init; } = DefaultOwnershipMarkerValue;

    /// <summary>
    /// Gets the file name held with an exclusive lease while a workspace is active.
    /// </summary>
    public string ActiveLeaseFileName { get; init; } = DefaultActiveLeaseFileName;

    /// <summary>
    /// Gets whether the workspace directory is recursively deleted on disposal.
    /// </summary>
    public bool DeleteOnDispose { get; init; } = true;

    /// <summary>
    /// Gets the target maximum size of each row-data chunk. A single oversized
    /// row is kept intact and may exceed this value.
    /// </summary>
    public long ChunkSizeBytes { get; init; } = DefaultChunkSizeBytes;

    /// <summary>
    /// Gets the maximum serialized size accepted for one row.
    /// </summary>
    public int MaximumRowSizeBytes { get; init; } = DefaultMaximumRowSizeBytes;

    /// <summary>
    /// Gets the maximum number of decoded rows retained by each store cache.
    /// Set to zero to disable the cache.
    /// </summary>
    public int RowCacheCapacity { get; init; } = 256;

    /// <summary>
    /// Gets the approximate in-memory byte limit of each decoded-row cache.
    /// Set to zero to disable the cache.
    /// </summary>
    public long RowCacheByteLimit { get; init; } = DefaultRowCacheByteLimit;

    /// <summary>
    /// Gets the number of cleanup retries after the initial delete attempt.
    /// </summary>
    public int CleanupRetryCount { get; init; } = 3;

    /// <summary>
    /// Gets the delay between workspace cleanup attempts.
    /// </summary>
    public TimeSpan CleanupRetryDelay { get; init; } = TimeSpan.FromMilliseconds(50);

    internal WorkspaceSettings Validate()
    {
        if (BaseDirectory is not null && string.IsNullOrWhiteSpace(BaseDirectory))
        {
            throw new ArgumentException("The workspace base directory cannot be empty.", nameof(BaseDirectory));
        }

        if (string.IsNullOrWhiteSpace(DirectoryPrefix))
        {
            throw new ArgumentException("The workspace directory prefix cannot be empty.", nameof(DirectoryPrefix));
        }

        if (DirectoryPrefix.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
            DirectoryPrefix.Contains(Path.DirectorySeparatorChar) ||
            DirectoryPrefix.Contains(Path.AltDirectorySeparatorChar))
        {
            throw new ArgumentException("The workspace directory prefix must be a valid file-name prefix.", nameof(DirectoryPrefix));
        }

        ValidateMetadataFileName(OwnershipMarkerFileName, nameof(OwnershipMarkerFileName));
        ValidateMetadataFileName(ActiveLeaseFileName, nameof(ActiveLeaseFileName));

        if (string.IsNullOrWhiteSpace(OwnershipMarkerValue))
        {
            throw new ArgumentException("The workspace ownership marker value cannot be empty.", nameof(OwnershipMarkerValue));
        }

        var fileNameComparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        if (string.Equals(OwnershipMarkerFileName, ActiveLeaseFileName, fileNameComparison))
        {
            throw new ArgumentException("The ownership marker and active lease file names must be different.");
        }

        if (ChunkSizeBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(ChunkSizeBytes), "Chunk size must be positive.");
        }

        if (MaximumRowSizeBytes <= 0 || MaximumRowSizeBytes > int.MaxValue - 8)
        {
            throw new ArgumentOutOfRangeException(
                nameof(MaximumRowSizeBytes),
                $"Maximum row size must be between 1 and {int.MaxValue - 8:N0} bytes.");
        }

        if (RowCacheCapacity < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(RowCacheCapacity), "Cache capacity cannot be negative.");
        }

        if (RowCacheByteLimit < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(RowCacheByteLimit), "Cache byte limit cannot be negative.");
        }

        if (CleanupRetryCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(CleanupRetryCount), "Cleanup retry count cannot be negative.");
        }

        if (CleanupRetryDelay < TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(CleanupRetryDelay), "Cleanup retry delay cannot be negative.");
        }

        var baseDirectory = BaseDirectory is null
            ? Path.GetTempPath()
            : Path.GetFullPath(BaseDirectory);

        return new WorkspaceSettings(
            baseDirectory,
            DirectoryPrefix,
            OwnershipMarkerFileName,
            OwnershipMarkerValue,
            ActiveLeaseFileName,
            DeleteOnDispose,
            ChunkSizeBytes,
            MaximumRowSizeBytes,
            RowCacheCapacity,
            RowCacheByteLimit,
            CleanupRetryCount,
            CleanupRetryDelay);
    }

    private static void ValidateMetadataFileName(string fileName, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(fileName) ||
            fileName is "." or ".." ||
            fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
            fileName.Contains(Path.DirectorySeparatorChar) ||
            fileName.Contains(Path.AltDirectorySeparatorChar))
        {
            throw new ArgumentException("Workspace metadata names must be valid single file names.", parameterName);
        }
    }
}

internal sealed record WorkspaceSettings(
    string BaseDirectory,
    string DirectoryPrefix,
    string OwnershipMarkerFileName,
    string OwnershipMarkerValue,
    string ActiveLeaseFileName,
    bool DeleteOnDispose,
    long ChunkSizeBytes,
    int MaximumRowSizeBytes,
    int RowCacheCapacity,
    long RowCacheByteLimit,
    int CleanupRetryCount,
    TimeSpan CleanupRetryDelay);

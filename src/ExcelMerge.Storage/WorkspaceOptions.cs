namespace ExcelMerge.Storage;

/// <summary>
/// Configures temporary workspaces and the row stores created in them.
/// </summary>
public sealed record WorkspaceOptions
{
    public const long DefaultChunkSizeBytes = 256L * 1024 * 1024;
    public const int DefaultMaximumRowSizeBytes = 128 * 1024 * 1024;
    public const long DefaultRowCacheByteLimit = 64L * 1024 * 1024;

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
            DeleteOnDispose,
            ChunkSizeBytes,
            MaximumRowSizeBytes,
            RowCacheCapacity,
            RowCacheByteLimit,
            CleanupRetryCount,
            CleanupRetryDelay);
    }
}

internal sealed record WorkspaceSettings(
    string BaseDirectory,
    string DirectoryPrefix,
    bool DeleteOnDispose,
    long ChunkSizeBytes,
    int MaximumRowSizeBytes,
    int RowCacheCapacity,
    long RowCacheByteLimit,
    int CleanupRetryCount,
    TimeSpan CleanupRetryDelay);

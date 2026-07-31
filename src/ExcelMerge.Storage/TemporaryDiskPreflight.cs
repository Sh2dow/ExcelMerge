namespace ExcelMerge.Storage;

public enum TemporaryDiskPreflightStatus
{
    Sufficient,
    Insufficient,
    UnknownVolume,
    EstimateOverflow,
}

/// <summary>
/// A checked estimate of temporary data plus the free-space reserve that must remain.
/// </summary>
public readonly record struct TemporaryDiskSpaceEstimate
{
    internal TemporaryDiskSpaceEstimate(
        long temporaryBytes,
        long reserveBytes,
        long requiredBytes,
        bool overflowed)
    {
        TemporaryBytes = temporaryBytes;
        ReserveBytes = reserveBytes;
        RequiredBytes = requiredBytes;
        Overflowed = overflowed;
    }

    public long TemporaryBytes { get; }

    public long ReserveBytes { get; }

    public long RequiredBytes { get; }

    public bool Overflowed { get; }
}

/// <summary>
/// Reports whether a volume can accommodate an estimate while preserving its reserve.
/// </summary>
public readonly record struct TemporaryDiskPreflightResult(
    TemporaryDiskPreflightStatus Status,
    TemporaryDiskSpaceEstimate Estimate,
    long? AvailableBytes,
    string? VolumeRoot)
{
    public bool IsSufficient => Status == TemporaryDiskPreflightStatus.Sufficient;
}

/// <summary>
/// Provides overflow-safe temporary-space estimation and volume preflight checks.
/// </summary>
public static class TemporaryDiskPreflight
{
    /// <summary>
    /// Estimates storage for a number of equally sized temporary copies and a free-space reserve.
    /// Overflow is represented by <see cref="TemporaryDiskSpaceEstimate.Overflowed"/>.
    /// </summary>
    public static TemporaryDiskSpaceEstimate Estimate(
        long sourceBytes,
        int temporaryCopyCount = 1,
        long reserveBytes = 0)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(sourceBytes);
        ArgumentOutOfRangeException.ThrowIfNegative(temporaryCopyCount);
        ArgumentOutOfRangeException.ThrowIfNegative(reserveBytes);

        long temporaryBytes;
        try
        {
            temporaryBytes = checked(sourceBytes * temporaryCopyCount);
        }
        catch (OverflowException)
        {
            return Overflow(reserveBytes);
        }

        try
        {
            var requiredBytes = checked(temporaryBytes + reserveBytes);
            return new TemporaryDiskSpaceEstimate(
                temporaryBytes,
                reserveBytes,
                requiredBytes,
                overflowed: false);
        }
        catch (OverflowException)
        {
            return Overflow(reserveBytes, temporaryBytes);
        }
    }

    /// <summary>
    /// Sums independently estimated temporary components and adds a free-space reserve.
    /// </summary>
    public static TemporaryDiskSpaceEstimate Estimate(
        IEnumerable<long> temporaryComponentSizesBytes,
        long reserveBytes = 0)
    {
        ArgumentNullException.ThrowIfNull(temporaryComponentSizesBytes);
        ArgumentOutOfRangeException.ThrowIfNegative(reserveBytes);

        var temporaryBytes = 0L;
        try
        {
            foreach (var componentSize in temporaryComponentSizesBytes)
            {
                ArgumentOutOfRangeException.ThrowIfNegative(componentSize);
                temporaryBytes = checked(temporaryBytes + componentSize);
            }

        }
        catch (OverflowException)
        {
            return Overflow(reserveBytes);
        }

        try
        {
            var requiredBytes = checked(temporaryBytes + reserveBytes);
            return new TemporaryDiskSpaceEstimate(
                temporaryBytes,
                reserveBytes,
                requiredBytes,
                overflowed: false);
        }
        catch (OverflowException)
        {
            return Overflow(reserveBytes, temporaryBytes);
        }
    }

    /// <summary>
    /// Evaluates an estimate against a known free-space value. A null value represents
    /// a volume whose free space could not be determined.
    /// </summary>
    public static TemporaryDiskPreflightResult Evaluate(
        TemporaryDiskSpaceEstimate estimate,
        long? availableBytes)
    {
        if (availableBytes < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(availableBytes), "Available space cannot be negative.");
        }

        if (estimate.Overflowed)
        {
            return new TemporaryDiskPreflightResult(
                TemporaryDiskPreflightStatus.EstimateOverflow,
                estimate,
                availableBytes,
                VolumeRoot: null);
        }

        if (availableBytes is null)
        {
            return new TemporaryDiskPreflightResult(
                TemporaryDiskPreflightStatus.UnknownVolume,
                estimate,
                AvailableBytes: null,
                VolumeRoot: null);
        }

        var status = availableBytes.Value >= estimate.RequiredBytes
            ? TemporaryDiskPreflightStatus.Sufficient
            : TemporaryDiskPreflightStatus.Insufficient;
        return new TemporaryDiskPreflightResult(status, estimate, availableBytes, VolumeRoot: null);
    }

    /// <summary>
    /// Resolves the volume containing <paramref name="path"/> and checks its available space.
    /// Unresolvable, unavailable, and inaccessible volumes return <see cref="TemporaryDiskPreflightStatus.UnknownVolume"/>.
    /// </summary>
    public static TemporaryDiskPreflightResult Check(
        string path,
        TemporaryDiskSpaceEstimate estimate)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        if (estimate.Overflowed)
        {
            return Evaluate(estimate, availableBytes: null);
        }

        string? root = null;
        try
        {
            root = Path.GetPathRoot(Path.GetFullPath(path));
            if (string.IsNullOrEmpty(root))
            {
                return UnknownVolume(estimate, root);
            }

            var drive = new DriveInfo(root);
            if (!drive.IsReady)
            {
                return UnknownVolume(estimate, root);
            }

            var evaluated = Evaluate(estimate, drive.AvailableFreeSpace);
            return evaluated with { VolumeRoot = root };
        }
        catch (Exception exception) when (
            exception is ArgumentException or IOException or UnauthorizedAccessException or NotSupportedException)
        {
            return UnknownVolume(estimate, root);
        }
    }

    private static TemporaryDiskPreflightResult UnknownVolume(
        TemporaryDiskSpaceEstimate estimate,
        string? root) =>
        new(
            TemporaryDiskPreflightStatus.UnknownVolume,
            estimate,
            AvailableBytes: null,
            VolumeRoot: root);

    private static TemporaryDiskSpaceEstimate Overflow(
        long reserveBytes,
        long temporaryBytes = long.MaxValue) =>
        new(
            temporaryBytes,
            reserveBytes,
            long.MaxValue,
            overflowed: true);
}

using System.Security.Cryptography;
using ExcelMerge.Domain;

namespace ExcelMerge.OpenXml;

public readonly record struct OpenXmlSourceFingerprint
{
    public OpenXmlSourceFingerprint(
        long contentLength,
        long lastWriteTimeUtcTicks,
        string sha256)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(contentLength);
        if (lastWriteTimeUtcTicks < DateTime.MinValue.Ticks ||
            lastWriteTimeUtcTicks > DateTime.MaxValue.Ticks)
        {
            throw new ArgumentOutOfRangeException(nameof(lastWriteTimeUtcTicks));
        }

        ArgumentException.ThrowIfNullOrWhiteSpace(sha256);
        if (sha256.Length != 64 || !sha256.All(static value => char.IsAsciiHexDigit(value)))
        {
            throw new ArgumentException("The SHA-256 value must contain 64 hexadecimal characters.", nameof(sha256));
        }

        ContentLength = contentLength;
        LastWriteTimeUtcTicks = lastWriteTimeUtcTicks;
        Sha256 = sha256.ToUpperInvariant();
    }

    public long ContentLength { get; }

    public long LastWriteTimeUtcTicks { get; }

    public string Sha256 { get; }
}

public sealed record OpenXmlWorkbookSource
{
    public OpenXmlWorkbookSource(string filePath, OpenXmlSourceFingerprint fingerprint)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        FilePath = Path.GetFullPath(filePath);
        Fingerprint = fingerprint;
    }

    public string FilePath { get; }

    public OpenXmlSourceFingerprint Fingerprint { get; }

    public static async ValueTask<OpenXmlWorkbookSource> CaptureAsync(
        string filePath,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        var fullPath = Path.GetFullPath(filePath);
        await using var stream = OpenRead(fullPath);
        var fingerprint = await ComputeFingerprintAsync(
            fullPath,
            stream,
            cancellationToken).ConfigureAwait(false);
        return new OpenXmlWorkbookSource(fullPath, fingerprint);
    }

    internal static FileStream OpenRead(string fullPath) => new(
        fullPath,
        FileMode.Open,
        FileAccess.Read,
        FileShare.Read | FileShare.Delete,
        bufferSize: 128 * 1024,
        FileOptions.Asynchronous | FileOptions.SequentialScan);

    internal static async ValueTask<OpenXmlSourceFingerprint> ComputeFingerprintAsync(
        string fullPath,
        FileStream stream,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var before = new FileInfo(fullPath);
        before.Refresh();
        if (!before.Exists)
        {
            throw new FileNotFoundException("The workbook source does not exist.", fullPath);
        }

        stream.Position = 0;
        var hash = await SHA256.HashDataAsync(stream, cancellationToken).ConfigureAwait(false);
        var after = new FileInfo(fullPath);
        after.Refresh();
        if (!after.Exists ||
            before.Length != after.Length ||
            before.LastWriteTimeUtc != after.LastWriteTimeUtc ||
            stream.Length != after.Length)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.SourceChanged,
                "A workbook source changed while its fingerprint was being calculated.",
                fullPath);
        }

        return new OpenXmlSourceFingerprint(
            after.Length,
            after.LastWriteTimeUtc.Ticks,
            Convert.ToHexString(hash));
    }
}

public sealed record OpenXmlWriteRequest
{
    public OpenXmlWriteRequest(
        MergePlan plan,
        OpenXmlWorkbookSource baseSource,
        OpenXmlWorkbookSource localSource,
        OpenXmlWorkbookSource remoteSource,
        string destinationPath)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(baseSource);
        ArgumentNullException.ThrowIfNull(localSource);
        ArgumentNullException.ThrowIfNull(remoteSource);
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);

        Plan = plan;
        BaseSource = baseSource;
        LocalSource = localSource;
        RemoteSource = remoteSource;
        DestinationPath = Path.GetFullPath(destinationPath);
    }

    public MergePlan Plan { get; }

    public OpenXmlWorkbookSource BaseSource { get; }

    public OpenXmlWorkbookSource LocalSource { get; }

    public OpenXmlWorkbookSource RemoteSource { get; }

    public string DestinationPath { get; }
}

public sealed record OpenXmlWriterOptions
{
    public bool Overwrite { get; init; } = true;

    public bool ValidatePackage { get; init; } = true;

    public int MaximumValidationErrors { get; init; } = 100;

    public long MinimumFreeSpaceReserveBytes { get; init; } = 16L * 1024 * 1024;
}

public enum OpenXmlWriterStage
{
    Preflight,
    CopyingLocal,
    ApplyingPlan,
    Validating,
    Committing,
    Completed,
}

public readonly record struct OpenXmlWriterProgress(
    OpenXmlWriterStage Stage,
    int CompletedWorksheetCount = 0,
    int TotalWorksheetCount = 0,
    string? SheetId = null);

public sealed record OpenXmlWriteResult(
    string DestinationPath,
    long ContentLength,
    OpenXmlSourceFingerprint Fingerprint);

public enum OpenXmlWriterError
{
    InvalidPlan,
    UnresolvedConflicts,
    SourceChanged,
    UnsupportedOperation,
    DestinationExists,
    InsufficientTemporarySpace,
    InvalidPackage,
    ValidationFailed,
    IoFailure,
}

public sealed record OpenXmlValidationError(
    string? Id,
    string Description,
    string? PartUri,
    string? Path);

public sealed class OpenXmlWriterException : IOException
{
    public OpenXmlWriterException(
        OpenXmlWriterError error,
        string message,
        string? path = null,
        IReadOnlyList<OpenXmlValidationError>? validationErrors = null,
        Exception? innerException = null)
        : base(message, innerException)
    {
        Error = error;
        Path = path;
        ValidationErrors = validationErrors ?? Array.Empty<OpenXmlValidationError>();
    }

    public OpenXmlWriterError Error { get; }

    public string? Path { get; }

    public IReadOnlyList<OpenXmlValidationError> ValidationErrors { get; }
}

public interface IOpenXmlWorkbookWriter
{
    ValueTask<OpenXmlWriteResult> WriteAsync(
        OpenXmlWriteRequest request,
        OpenXmlWriterOptions? options = null,
        IProgress<OpenXmlWriterProgress>? progress = null,
        CancellationToken cancellationToken = default);
}

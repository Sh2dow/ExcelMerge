namespace ExcelMerge.Storage;

/// <summary>
/// Thrown when a workspace row or index record is incomplete or invalid.
/// </summary>
public sealed class StorageCorruptionException : IOException
{
    public StorageCorruptionException(string message)
        : base(message)
    {
    }

    public StorageCorruptionException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

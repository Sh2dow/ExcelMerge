namespace ExcelMerge.Storage;

/// <summary>
/// Identifies the physical record for a row. ByteOffset points to the record
/// header, and ByteLength includes that header.
/// </summary>
public readonly record struct RowOffset
{
    public RowOffset(
        long rowOrdinal,
        int chunkNumber,
        long byteOffset,
        int byteLength,
        int cellCount,
        int sourceRowIndex)
    {
        if (rowOrdinal < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rowOrdinal));
        }

        if (chunkNumber < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(chunkNumber));
        }

        if (byteOffset < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(byteOffset));
        }

        if (byteLength <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(byteLength));
        }

        if (cellCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(cellCount));
        }

        if (sourceRowIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(sourceRowIndex));
        }

        RowOrdinal = rowOrdinal;
        ChunkNumber = chunkNumber;
        ByteOffset = byteOffset;
        ByteLength = byteLength;
        CellCount = cellCount;
        SourceRowIndex = sourceRowIndex;
    }

    public long RowOrdinal { get; }

    public int ChunkNumber { get; }

    public long ByteOffset { get; }

    public int ByteLength { get; }

    public int CellCount { get; }

    public int SourceRowIndex { get; }

    public long EndByteOffset => checked(ByteOffset + ByteLength);
}

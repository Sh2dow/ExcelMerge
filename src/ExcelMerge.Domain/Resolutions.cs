namespace ExcelMerge.Domain;

public enum ResolutionKind
{
    Unresolved,
    Local,
    Remote,
    Both,
    Custom,
}

/// <summary>
/// A cell-level decision. BOTH is deliberately invalid here because it duplicates a complete row.
/// </summary>
public readonly record struct CellResolution
{
    public CellResolution(
        long conflictId,
        ResolutionKind kind,
        CellValue? customValue = null)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(conflictId);
        if (kind == ResolutionKind.Both)
            throw new ArgumentException("BOTH is only valid for a row resolution.", nameof(kind));
        ValidateCustomValue(kind, customValue.HasValue);

        ConflictId = conflictId;
        Kind = kind;
        CustomValue = customValue;
    }

    public long ConflictId { get; }

    public ResolutionKind Kind { get; }

    public CellValue? CustomValue { get; }

    public bool IsResolved => Kind != ResolutionKind.Unresolved;

    public static CellResolution Unresolved(long conflictId) =>
        new(conflictId, ResolutionKind.Unresolved);

    public static CellResolution UseLocal(long conflictId) =>
        new(conflictId, ResolutionKind.Local);

    public static CellResolution UseRemote(long conflictId) =>
        new(conflictId, ResolutionKind.Remote);

    public static CellResolution Custom(long conflictId, CellValue value) =>
        new(conflictId, ResolutionKind.Custom, value);

    private static void ValidateCustomValue(ResolutionKind kind, bool hasCustomValue)
    {
        if (kind == ResolutionKind.Custom && !hasCustomValue)
            throw new ArgumentException("A custom resolution requires a custom value.", nameof(kind));
        if (kind != ResolutionKind.Custom && hasCustomValue)
            throw new ArgumentException("A custom value is only valid for a custom resolution.", nameof(kind));
    }
}

public readonly record struct RowResolution
{
    public RowResolution(
        long conflictId,
        ResolutionKind kind,
        RowRecord? customRow = null)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(conflictId);
        if (kind == ResolutionKind.Custom && !customRow.HasValue)
            throw new ArgumentException("A custom resolution requires a custom row.", nameof(kind));
        if (kind != ResolutionKind.Custom && customRow.HasValue)
            throw new ArgumentException("A custom row is only valid for a custom resolution.", nameof(kind));

        ConflictId = conflictId;
        Kind = kind;
        CustomRow = customRow;
    }

    public long ConflictId { get; }

    public ResolutionKind Kind { get; }

    public RowRecord? CustomRow { get; }

    public bool IsResolved => Kind != ResolutionKind.Unresolved;

    public static RowResolution Unresolved(long conflictId) =>
        new(conflictId, ResolutionKind.Unresolved);

    public static RowResolution UseLocal(long conflictId) =>
        new(conflictId, ResolutionKind.Local);

    public static RowResolution UseRemote(long conflictId) =>
        new(conflictId, ResolutionKind.Remote);

    public static RowResolution UseBoth(long conflictId) =>
        new(conflictId, ResolutionKind.Both);

    public static RowResolution Custom(long conflictId, RowRecord row) =>
        new(conflictId, ResolutionKind.Custom, row);
}

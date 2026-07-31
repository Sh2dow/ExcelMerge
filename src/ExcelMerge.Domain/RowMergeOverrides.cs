namespace ExcelMerge.Domain;

/// <summary>A complete-row result choice that covers every conflict on the same aligned row.</summary>
public readonly record struct RowMergeOverride
{
    public RowMergeOverride(
        string sheetId,
        int? baseRowIndex,
        int? localRowIndex,
        int? remoteRowIndex,
        ResolutionKind kind,
        RowRecord? customRow = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetId);
        if (!baseRowIndex.HasValue && !localRowIndex.HasValue && !remoteRowIndex.HasValue)
        {
            throw new ArgumentException("At least one row coordinate is required.", nameof(baseRowIndex));
        }

        if (baseRowIndex < 0 || localRowIndex < 0 || remoteRowIndex < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(baseRowIndex));
        }

        if (!Enum.IsDefined(kind) || kind == ResolutionKind.Unresolved)
        {
            throw new ArgumentException("A row override must contain a resolved choice.", nameof(kind));
        }

        if (kind == ResolutionKind.Custom != customRow.HasValue)
        {
            throw new ArgumentException(
                "A custom row is required only for a custom row override.",
                nameof(customRow));
        }

        SheetId = sheetId;
        BaseRowIndex = baseRowIndex;
        LocalRowIndex = localRowIndex;
        RemoteRowIndex = remoteRowIndex;
        Kind = kind;
        CustomRow = customRow;
    }

    public string SheetId { get; }

    public int? BaseRowIndex { get; }

    public int? LocalRowIndex { get; }

    public int? RemoteRowIndex { get; }

    public ResolutionKind Kind { get; }

    public RowRecord? CustomRow { get; }

    public bool Covers(ConflictLocation location) =>
        string.Equals(SheetId, location.SheetId, StringComparison.Ordinal) &&
        BaseRowIndex == location.BaseRowIndex &&
        LocalRowIndex == location.LocalRowIndex &&
        RemoteRowIndex == location.RemoteRowIndex;
}

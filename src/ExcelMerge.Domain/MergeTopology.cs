namespace ExcelMerge.Domain;

/// <summary>A complete row mapping in result order, independent of viewport filtering.</summary>
public readonly record struct MergeRowMapping
{
    public MergeRowMapping(
        int? baseRowIndex,
        int? localRowIndex,
        int? remoteRowIndex)
    {
        if (!baseRowIndex.HasValue && !localRowIndex.HasValue && !remoteRowIndex.HasValue)
        {
            throw new ArgumentException("At least one row coordinate is required.");
        }

        if (baseRowIndex < 0 || localRowIndex < 0 || remoteRowIndex < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(baseRowIndex),
                "Row coordinates cannot be negative.");
        }

        BaseRowIndex = baseRowIndex;
        LocalRowIndex = localRowIndex;
        RemoteRowIndex = remoteRowIndex;
    }

    public int? BaseRowIndex { get; }

    public int? LocalRowIndex { get; }

    public int? RemoteRowIndex { get; }
}

/// <summary>A paired worksheet and all engine output required to compile its result.</summary>
public sealed record SheetMergePlan
{
    public SheetMergePlan(
        string id,
        SheetMetadata? baseSheet,
        SheetMetadata? localSheet,
        SheetMetadata? remoteSheet,
        SheetChange localChange,
        SheetChange remoteChange,
        ReadOnlyMemory<MergeRowMapping> rows,
        ReadOnlyMemory<AutomaticMergeDecision> automaticDecisions)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(id);
        if (baseSheet is null && localSheet is null && remoteSheet is null)
        {
            throw new ArgumentException("At least one worksheet is required.", nameof(baseSheet));
        }

        Id = id;
        BaseSheet = baseSheet;
        LocalSheet = localSheet;
        RemoteSheet = remoteSheet;
        LocalChange = localChange;
        RemoteChange = remoteChange;
        Rows = rows;
        AutomaticDecisions = automaticDecisions;
    }

    public string Id { get; }

    public SheetMetadata? BaseSheet { get; }

    public SheetMetadata? LocalSheet { get; }

    public SheetMetadata? RemoteSheet { get; }

    public SheetChange LocalChange { get; }

    public SheetChange RemoteChange { get; }

    public ReadOnlyMemory<MergeRowMapping> Rows { get; }

    public ReadOnlyMemory<AutomaticMergeDecision> AutomaticDecisions { get; }
}

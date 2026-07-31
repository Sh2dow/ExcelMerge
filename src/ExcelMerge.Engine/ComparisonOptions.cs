using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

/// <summary>Controls exact typed cell comparison.</summary>
public sealed record CellComparisonOptions
{
    /// <summary>Includes a formula's typed cached value as well as its formula text.</summary>
    public bool CompareFormulaCachedValues { get; init; } = true;

    /// <summary>Includes adapter-provided display text. Raw typed values are always compared.</summary>
    public bool CompareDisplayText { get; init; }
}

/// <summary>Bounds and semantic choices for one selected worksheet operation.</summary>
public sealed record WorksheetComparisonOptions
{
    public IReadOnlyList<int> KeyColumns { get; init; } = Array.Empty<int>();

    public CellComparisonOptions CellValues { get; init; } = new();

    public bool CompareCellStyles { get; init; }

    public bool CompareRowMetadata { get; init; }

    public bool CompareSheetMetadata { get; init; } = true;

    /// <summary>Makes an unstyled explicit blank equivalent to an absent sparse cell.</summary>
    public bool TreatExplicitBlankCellsAsMissing { get; init; } = true;

    /// <summary>Prevents an all-blank configured key from becoming a row anchor.</summary>
    public bool RequireNonBlankKey { get; init; } = true;

    public bool IncludeUnchangedViewRows { get; init; } = true;

    /// <summary>Physical row span requested from a snapshot at one time.</summary>
    public int ReadBatchRowSpan { get; init; } = 4_096;

    public int MaxMaterializedRows { get; init; } = 1_048_576;

    public long MaxMaterializedCells { get; init; } = 5_000_000;

    /// <summary>Maximum physical row span scanned when loading a sparse selected sheet.</summary>
    public long MaxScannedRowSpan { get; init; } = 10_000_000;

    /// <summary>Maximum length on either side of a gap handled by dynamic programming.</summary>
    public int MaxAlignmentGapRows { get; init; } = 4_096;

    /// <summary>Maximum cells in a gap alignment trace matrix.</summary>
    public long MaxDynamicProgrammingCells { get; init; } = 2_000_000;

    public int AlignmentLookaheadRows { get; init; } = 64;

    /// <summary>Maximum inner-loop work between explicit cancellation checks.</summary>
    public int CancellationCheckInterval { get; init; } = 1_024;
}

public sealed record ThreeWayMergeOptions
{
    public WorksheetComparisonOptions Comparison { get; init; } = new();

    /// <summary>First deterministic conflict identifier allocated by this operation.</summary>
    public long FirstConflictId { get; init; }
}

/// <summary>Raised before a selected worksheet exceeds an explicit engine bound.</summary>
public sealed class WorksheetLimitExceededException : InvalidOperationException
{
    public WorksheetLimitExceededException(string limitName, long limit, long observed)
        : base($"Worksheet limit '{limitName}' is {limit}, but at least {observed} was required.")
    {
        LimitName = limitName;
        Limit = limit;
        Observed = observed;
    }

    public string LimitName { get; }

    public long Limit { get; }

    public long Observed { get; }
}

/// <summary>Raised when an <see cref="IWorksheetSnapshot"/> violates its ordering or range contract.</summary>
public sealed class InvalidWorksheetSnapshotException : IOException
{
    public InvalidWorksheetSnapshotException(string message)
        : base(message)
    {
    }
}

internal sealed class NormalizedComparisonOptions
{
    public NormalizedComparisonOptions(WorksheetComparisonOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(options.CellValues);
        ArgumentNullException.ThrowIfNull(options.KeyColumns);

        ValidatePositive(options.ReadBatchRowSpan, nameof(options.ReadBatchRowSpan));
        ValidatePositive(options.MaxMaterializedRows, nameof(options.MaxMaterializedRows));
        ValidatePositive(options.MaxMaterializedCells, nameof(options.MaxMaterializedCells));
        ValidatePositive(options.MaxScannedRowSpan, nameof(options.MaxScannedRowSpan));
        ValidatePositive(options.MaxAlignmentGapRows, nameof(options.MaxAlignmentGapRows));
        ValidatePositive(options.MaxDynamicProgrammingCells, nameof(options.MaxDynamicProgrammingCells));
        ValidatePositive(options.AlignmentLookaheadRows, nameof(options.AlignmentLookaheadRows));
        ValidatePositive(options.CancellationCheckInterval, nameof(options.CancellationCheckInterval));

        var keys = options.KeyColumns.ToArray();
        Array.Sort(keys);
        for (var index = 0; index < keys.Length; index++)
        {
            ArgumentOutOfRangeException.ThrowIfNegative(keys[index], nameof(options.KeyColumns));
            if (index > 0 && keys[index - 1] == keys[index])
            {
                throw new ArgumentException(
                    $"Key column {keys[index]} is specified more than once.",
                    nameof(options));
            }
        }

        Source = options;
        KeyColumns = keys;
        CellComparer = new CellValueComparer(options.CellValues);
    }

    public WorksheetComparisonOptions Source { get; }

    public int[] KeyColumns { get; }

    public CellValueComparer CellComparer { get; }

    private static void ValidatePositive(long value, string parameterName)
    {
        if (value <= 0)
            throw new ArgumentOutOfRangeException(parameterName, value, "The value must be positive.");
    }
}

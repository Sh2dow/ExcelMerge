using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

/// <summary>A contiguous run of columns with the same cell change kind.</summary>
public readonly record struct ColumnChangeRun(
    int StartColumnIndex,
    int ColumnCount,
    ChangeKind Kind)
{
    public int EndColumnIndexExclusive => checked(StartColumnIndex + ColumnCount);
}

/// <summary>A compact row mapping for a synchronized two-way view.</summary>
public readonly record struct DiffViewRow(
    int? BaseRowIndex,
    int? SourceRowIndex,
    ChangeKind Kind,
    ReadOnlyMemory<ColumnChangeRun> ChangedColumnRuns);

public readonly record struct DiffStatistics(
    int ViewRowCount,
    int ChangedRowCount,
    int AddedRowCount,
    int RemovedRowCount,
    int ModifiedRowCount,
    long ChangedCellCount);

/// <summary>A compact view plus its writer-compatible domain change.</summary>
public sealed record TwoWayDiffResult(
    SheetChange SheetChange,
    ReadOnlyMemory<DiffViewRow> ViewRows,
    DiffStatistics Statistics);

/// <summary>A compact physical-row mapping produced without constructing cell changes.</summary>
public readonly record struct RowAlignmentEntry(
    int? BaseRowIndex,
    int? SourceRowIndex,
    ChangeKind Kind);

public sealed record RowAlignmentResult(
    SheetMetadata BaseSheet,
    SheetMetadata SourceSheet,
    ReadOnlyMemory<RowAlignmentEntry> Rows);

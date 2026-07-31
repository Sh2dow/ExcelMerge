using ExcelMerge.Domain;
using ExcelMerge.Engine;

namespace ExcelMerge.Rewrite.Benchmarks;

internal static class BenchmarkData
{
    public const int SmokeCellCount = 200_000;
    public const int StandardCellCount = 1_000_000;
    public const int CapacityCellCount = 5_000_000;
    public const string ScaleEnvironmentVariable = "EXCELMERGE_BENCHMARK_SCALES";

    public static IReadOnlyList<int> EnabledCellCounts { get; } = ReadEnabledCellCounts();

    private const int CellsPerRow = 20;
    private const int ChangedColumnIndex = 7;
    private const int ChangedRowInterval = 100;

    public static RowRecord[] CreateRows(int cellCount, double changedValueOffset = 0)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(cellCount);

        var rowCount = checked((cellCount + CellsPerRow - 1) / CellsPerRow);
        var rows = new RowRecord[rowCount];
        var remainingCellCount = cellCount;

        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var cellsInRow = Math.Min(CellsPerRow, remainingCellCount);
            var cells = new CellRecord[cellsInRow];
            for (var columnIndex = 0; columnIndex < cells.Length; columnIndex++)
            {
                var value = ((long)rowIndex * CellsPerRow) + columnIndex;
                var offset = changedValueOffset != 0 &&
                    rowIndex % ChangedRowInterval == 0 &&
                    columnIndex == ChangedColumnIndex
                    ? changedValueOffset
                    : 0;
                cells[columnIndex] = new CellRecord(
                    new CellAddress(rowIndex, columnIndex),
                    CellValue.FromNumber(value + offset));
            }

            rows[rowIndex] = new RowRecord(rowIndex, cells);
            remainingCellCount -= cellsInRow;
        }

        return rows;
    }

    public static EngineBenchmarkData CreateEngineData(int cellCount)
    {
        var baseRows = CreateRows(cellCount);
        var localRows = CreateRows(cellCount, changedValueOffset: 0.25);
        var remoteRows = CreateRows(cellCount, changedValueOffset: 0.5);
        var rowCount = baseRows.Length;
        var metadata = new SheetMetadata(
            "smoke-sheet",
            "Smoke",
            0,
            FirstRowIndex: 0,
            LastRowIndex: rowCount - 1,
            FirstColumnIndex: 0,
            LastColumnIndex: Math.Min(CellsPerRow, cellCount) - 1,
            NonEmptyCellCount: cellCount);
        var comparison = new WorksheetComparisonOptions
        {
            KeyColumns = [0],
            IncludeUnchangedViewRows = false,
            MaxMaterializedRows = rowCount,
            MaxMaterializedCells = cellCount,
            MaxScannedRowSpan = rowCount,
        };

        return new EngineBenchmarkData(
            new WorksheetSnapshot(metadata, baseRows),
            new WorksheetSnapshot(metadata, localRows),
            new WorksheetSnapshot(metadata, remoteRows),
            comparison,
            new ThreeWayMergeOptions
            {
                Comparison = comparison,
                SheetGroupId = metadata.Id,
            });
    }

    private static IReadOnlyList<int> ReadEnabledCellCounts()
    {
        var configuredScales = Environment.GetEnvironmentVariable(ScaleEnvironmentVariable);
        if (string.IsNullOrWhiteSpace(configuredScales))
        {
            return [SmokeCellCount];
        }

        var cellCounts = new SortedSet<int>();
        foreach (var configuredScale in configuredScales.Split(
            [',', ';'],
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            switch (configuredScale.ToLowerInvariant())
            {
                case "smoke":
                case "200k":
                case "200000":
                    cellCounts.Add(SmokeCellCount);
                    break;
                case "standard":
                case "1m":
                case "1000000":
                    cellCounts.Add(StandardCellCount);
                    break;
                case "capacity":
                case "5m":
                case "5000000":
                    cellCounts.Add(CapacityCellCount);
                    break;
                case "all":
                    cellCounts.Add(SmokeCellCount);
                    cellCounts.Add(StandardCellCount);
                    cellCounts.Add(CapacityCellCount);
                    break;
                default:
                    throw new InvalidOperationException(
                        $"Unknown benchmark scale '{configuredScale}' in {ScaleEnvironmentVariable}. " +
                        "Use smoke, standard, capacity, a comma-separated combination, or all.");
            }
        }

        if (cellCounts.Count == 0)
        {
            throw new InvalidOperationException(
                $"{ScaleEnvironmentVariable} did not contain a benchmark scale.");
        }

        return cellCounts.ToArray();
    }
}

internal sealed record EngineBenchmarkData(
    WorksheetSnapshot BaseSheet,
    WorksheetSnapshot LocalSheet,
    WorksheetSnapshot RemoteSheet,
    WorksheetComparisonOptions Comparison,
    ThreeWayMergeOptions MergeOptions);

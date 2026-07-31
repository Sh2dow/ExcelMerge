using BenchmarkDotNet.Attributes;
using ExcelMerge.Engine;

namespace ExcelMerge.Benchmarks;

[BenchmarkCategory("Engine")]
public class EngineBenchmarks
{
    private readonly RowAlignmentEngine _rowAlignmentEngine = new();
    private readonly TwoWayDiffEngine _twoWayDiffEngine = new();
    private readonly ThreeWayMergeEngine _threeWayMergeEngine = new();
    private EngineBenchmarkData _data = null!;

    [Params(BenchmarkData.SmokeCellCount)]
    public int CellCount { get; set; }

    [GlobalSetup]
    public void Setup() => _data = BenchmarkData.CreateEngineData(CellCount);

    [Benchmark(Description = "WorksheetSnapshot / row alignment")]
    [BenchmarkCategory("RowAlignment")]
    public async Task<int> AlignRows()
    {
        var result = await _rowAlignmentEngine.AlignAsync(
            _data.BaseSheet,
            _data.LocalSheet,
            _data.Comparison).ConfigureAwait(false);
        return result.Rows.Length;
    }

    [Benchmark(Description = "Two-way diff")]
    [BenchmarkCategory("TwoWayDiff")]
    public async Task<long> TwoWayDiff()
    {
        var result = await _twoWayDiffEngine.CompareAsync(
            _data.BaseSheet,
            _data.LocalSheet,
            _data.Comparison).ConfigureAwait(false);
        return result.Statistics.ChangedCellCount;
    }

    [Benchmark(Description = "Three-way conflict scan")]
    [BenchmarkCategory("ThreeWayConflict")]
    public async Task<int> ThreeWayConflictScan()
    {
        var result = await _threeWayMergeEngine.MergeAsync(
            _data.BaseSheet,
            _data.LocalSheet,
            _data.RemoteSheet,
            _data.MergeOptions).ConfigureAwait(false);
        return result.Conflicts.Length;
    }
}

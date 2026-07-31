using BenchmarkDotNet.Attributes;
using ExcelMerge.Domain;
using ExcelMerge.Storage;

namespace ExcelMerge.Rewrite.Benchmarks;

[BenchmarkCategory("Storage", "Append")]
public class ChunkedCellStoreAppendBenchmarks
{
    private RowRecord[] _rows = null!;
    private CellStoreFixture? _fixture;

    [Params(BenchmarkData.SmokeCellCount)]
    public int CellCount { get; set; }

    [GlobalSetup]
    public void Setup() => _rows = BenchmarkData.CreateRows(CellCount);

    [IterationSetup]
    public void SetupIteration()
    {
        CleanupIteration();
        _fixture = CellStoreFixture.Create("excelmerge-benchmark-append-");
    }

    [Benchmark(Description = "ChunkedCellStore append")]
    public async Task<long> AppendRows()
    {
        var store = _fixture?.Store ?? throw new InvalidOperationException("The store is not initialized.");
        foreach (var row in _rows)
        {
            await store.AppendRowAsync(row).ConfigureAwait(false);
        }

        return store.DataLengthBytes + store.IndexLengthBytes;
    }

    [IterationCleanup]
    public void CleanupIteration()
    {
        var fixture = _fixture;
        _fixture = null;
        fixture?.Dispose();
    }

    [GlobalCleanup]
    public void Cleanup() => CleanupIteration();
}

[BenchmarkCategory("Storage", "Read")]
public class ChunkedCellStoreReadBenchmarks
{
    private CellStoreFixture? _fixture;

    [Params(BenchmarkData.SmokeCellCount)]
    public int CellCount { get; set; }

    [GlobalSetup]
    public async Task Setup()
    {
        _fixture = CellStoreFixture.Create("excelmerge-benchmark-read-");

        try
        {
            foreach (var row in BenchmarkData.CreateRows(CellCount))
            {
                await _fixture.Store.AppendRowAsync(row).ConfigureAwait(false);
            }

            await _fixture.Store.FlushAsync().ConfigureAwait(false);
        }
        catch
        {
            Cleanup();
            throw;
        }
    }

    [Benchmark(Description = "ChunkedCellStore read")]
    public async Task<long> ReadRows()
    {
        var store = _fixture?.Store ?? throw new InvalidOperationException("The store is not initialized.");
        long checksum = 0;
        for (long rowOrdinal = 0; rowOrdinal < store.RowCount; rowOrdinal++)
        {
            var row = await store.ReadRowAsync(rowOrdinal).ConfigureAwait(false);
            checksum = unchecked((checksum * 31) + row.RowIndex + row.CellCount);
        }

        return checksum;
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        var fixture = _fixture;
        _fixture = null;
        fixture?.Dispose();
    }
}

internal sealed class CellStoreFixture : IDisposable
{
    private Workspace? _workspace;

    private CellStoreFixture(Workspace workspace, ChunkedCellStore store)
    {
        _workspace = workspace;
        Store = store;
    }

    public ChunkedCellStore Store { get; }

    public static CellStoreFixture Create(string directoryPrefix)
    {
        Workspace? workspace = null;
        try
        {
            workspace = new Workspace(new WorkspaceOptions
            {
                DirectoryPrefix = directoryPrefix,
                RowCacheCapacity = 0,
                RowCacheByteLimit = 0,
            });
            return new CellStoreFixture(workspace, workspace.CreateCellStore());
        }
        catch
        {
            workspace?.Dispose();
            throw;
        }
    }

    public void Dispose()
    {
        var workspace = Interlocked.Exchange(ref _workspace, null);
        workspace?.Dispose();
    }
}

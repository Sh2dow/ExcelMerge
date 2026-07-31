using BenchmarkDotNet.Attributes;
using ExcelMerge.Domain;
using ExcelMerge.OpenXml;
using ExcelMerge.Storage;

namespace ExcelMerge.Benchmarks;

[BenchmarkCategory("OpenXml", "Reader", "Metadata")]
public class OpenXmlMetadataBenchmarks
{
    private const int MetadataWorksheetCount = 16;

    private readonly OpenXmlWorkbookReader _reader = new();
    private XlsxBenchmarkCorpus? _corpus;

    public IEnumerable<int> CellCounts => BenchmarkData.EnabledCellCounts;

    [ParamsSource(nameof(CellCounts))]
    public int CellCount { get; set; }

    [GlobalSetup]
    public void Setup() => _corpus = XlsxBenchmarkCorpus.Create(
        CellCount,
        XlsxCorpusLayout.Dense,
        XlsxSharedStringDistribution.Repeated,
        MetadataWorksheetCount);

    [Benchmark(Description = "XLSX metadata-only load")]
    public async Task<long> LoadMetadata()
    {
        var corpus = _corpus ?? throw new InvalidOperationException("The corpus is not initialized.");
        var metadata = await _reader.ReadMetadataAsync(corpus.FilePath).ConfigureAwait(false);
        return metadata.ContentLength.GetValueOrDefault() + metadata.Sheets.Count;
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        var corpus = _corpus;
        _corpus = null;
        corpus?.Dispose();
    }
}

[BenchmarkCategory("OpenXml", "Reader", "Indexing")]
public class OpenXmlIndexingBenchmarks
{
    private static readonly OpenXmlReaderOptions ReaderOptions = new()
    {
        IncludeDisplayText = true,
    };

    private readonly OpenXmlWorkbookReader _reader = new();
    private XlsxBenchmarkCorpus? _corpus;
    private Workspace? _workspace;

    public IEnumerable<int> CellCounts => BenchmarkData.EnabledCellCounts;

    [ParamsSource(nameof(CellCounts))]
    public int CellCount { get; set; }

    [Params(XlsxCorpusLayout.Dense, XlsxCorpusLayout.Sparse)]
    public XlsxCorpusLayout Layout { get; set; }

    [Params(XlsxSharedStringDistribution.Repeated, XlsxSharedStringDistribution.Unique)]
    public XlsxSharedStringDistribution SharedStrings { get; set; }

    [GlobalSetup]
    public void Setup() => _corpus = XlsxBenchmarkCorpus.Create(CellCount, Layout, SharedStrings);

    [IterationSetup]
    public void SetupIteration()
    {
        if (_workspace is not null)
        {
            throw new InvalidOperationException("The prior indexing workspace was not cleaned up.");
        }

        var corpus = _corpus ?? throw new InvalidOperationException("The corpus is not initialized.");
        _workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = corpus.DirectoryPath,
            DirectoryPrefix = "excelmerge-benchmark-index-",
        });
    }

    [Benchmark(Description = "Streaming XLSX index")]
    public async Task<long> IndexWorkbook()
    {
        var corpus = _corpus ?? throw new InvalidOperationException("The corpus is not initialized.");
        var workspace = _workspace ?? throw new InvalidOperationException("The workspace is not initialized.");
        var result = await _reader.IndexAsync(
            corpus.FilePath,
            workspace,
            ReaderOptions).ConfigureAwait(false);
        return result.Worksheets[0].Metadata.NonEmptyCellCount.GetValueOrDefault();
    }

    [IterationCleanup]
    public void CleanupIteration()
    {
        var workspace = _workspace;
        _workspace = null;
        workspace?.Dispose();
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        try
        {
            CleanupIteration();
        }
        finally
        {
            var corpus = _corpus;
            _corpus = null;
            corpus?.Dispose();
        }
    }
}

[BenchmarkCategory("OpenXml", "Reader", "Viewport")]
public class OpenXmlViewportBenchmarks
{
    private const int ViewportRowCount = 128;
    private const int ViewportPositionCount = 16;

    private readonly OpenXmlWorkbookReader _reader = new();
    private XlsxBenchmarkCorpus? _corpus;
    private Workspace? _workspace;
    private OpenXmlReaderWorksheet? _worksheet;
    private int _nextViewport;

    public IEnumerable<int> CellCounts => BenchmarkData.EnabledCellCounts;

    [ParamsSource(nameof(CellCounts))]
    public int CellCount { get; set; }

    [Params(XlsxCorpusLayout.Dense, XlsxCorpusLayout.Sparse)]
    public XlsxCorpusLayout Layout { get; set; }

    [GlobalSetup]
    public async Task Setup()
    {
        try
        {
            _corpus = XlsxBenchmarkCorpus.Create(
                CellCount,
                Layout,
                XlsxSharedStringDistribution.Repeated);
            _workspace = new Workspace(new WorkspaceOptions
            {
                BaseDirectory = _corpus.DirectoryPath,
                DirectoryPrefix = "excelmerge-benchmark-viewport-",
                RowCacheCapacity = 256,
                RowCacheByteLimit = 64L * 1024 * 1024,
            });
            var result = await _reader.IndexAsync(_corpus.FilePath, _workspace).ConfigureAwait(false);
            _worksheet = result.Worksheets[0];
            _nextViewport = -1;
        }
        catch
        {
            try
            {
                Cleanup();
            }
            catch
            {
                // Preserve the setup failure after attempting full cleanup.
            }

            throw;
        }
    }

    [Benchmark(Description = "Indexed viewport row fetch")]
    public ValueTask<WorksheetRowBatch> FetchViewport()
    {
        var worksheet = _worksheet ?? throw new InvalidOperationException("The worksheet is not initialized.");
        _nextViewport = (_nextViewport + 1) % ViewportPositionCount;
        var lastRowIndex = worksheet.Metadata.LastRowIndex.GetValueOrDefault();
        var maximumStartRow = Math.Max(0, lastRowIndex - ViewportRowCount + 1);
        var startRowIndex = checked((int)(
            ((long)maximumStartRow * _nextViewport) / (ViewportPositionCount - 1)));
        return worksheet.ReadRowsAsync(startRowIndex, ViewportRowCount);
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        _worksheet = null;
        var workspace = _workspace;
        _workspace = null;
        try
        {
            workspace?.Dispose();
        }
        finally
        {
            var corpus = _corpus;
            _corpus = null;
            corpus?.Dispose();
        }
    }
}

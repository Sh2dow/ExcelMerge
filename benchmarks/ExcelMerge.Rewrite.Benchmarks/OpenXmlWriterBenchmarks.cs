using BenchmarkDotNet.Attributes;
using ExcelMerge.Domain;
using ExcelMerge.OpenXml;

namespace ExcelMerge.Rewrite.Benchmarks;

[BenchmarkCategory("OpenXml", "Writer", "Transactional")]
public class OpenXmlTransactionalWriterBenchmarks
{
    private const string SheetGroupId = "benchmark-sheet";

    private static readonly OpenXmlWriterOptions WriterOptions = new()
    {
        MinimumFreeSpaceReserveBytes = 0,
        ValidatePackage = true,
    };

    private readonly OpenXmlWorkbookReader _reader = new();
    private readonly OpenXmlWorkbookWriter _writer = new();
    private XlsxBenchmarkCorpus? _corpus;
    private OpenXmlWriteRequest? _request;
    private string? _outputDirectory;

    public IEnumerable<int> CellCounts => BenchmarkData.EnabledCellCounts;

    [ParamsSource(nameof(CellCounts))]
    public int CellCount { get; set; }

    [GlobalSetup]
    public async Task Setup()
    {
        try
        {
            _corpus = XlsxBenchmarkCorpus.Create(
                CellCount,
                XlsxCorpusLayout.Dense,
                XlsxSharedStringDistribution.Repeated);
            var metadata = await _reader.ReadMetadataAsync(_corpus.FilePath).ConfigureAwait(false);
            var source = await OpenXmlWorkbookSource.CaptureAsync(_corpus.FilePath).ConfigureAwait(false);
            _outputDirectory = Path.Combine(_corpus.DirectoryPath, "writer-output");
            var destinationPath = Path.Combine(_outputDirectory, "result.xlsx");
            _request = new OpenXmlWriteRequest(
                CreateSingleCellPatchPlan(metadata),
                source,
                source,
                source,
                destinationPath);
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

    [IterationSetup]
    public void SetupIteration()
    {
        var outputDirectory = _outputDirectory ??
            throw new InvalidOperationException("The output directory is not initialized.");
        BenchmarkFileSystem.DeleteDirectory(outputDirectory);
        Directory.CreateDirectory(outputDirectory);

        var request = _request ?? throw new InvalidOperationException("The write request is not initialized.");
        File.WriteAllBytes(request.DestinationPath, new byte[] { 0 });
    }

    [Benchmark(Description = "Transactional XLSX package write")]
    public ValueTask<OpenXmlWriteResult> WritePackage()
    {
        var request = _request ?? throw new InvalidOperationException("The write request is not initialized.");
        return _writer.WriteAsync(request, WriterOptions);
    }

    [IterationCleanup]
    public void CleanupIteration()
    {
        if (_outputDirectory is not null)
        {
            BenchmarkFileSystem.DeleteDirectory(_outputDirectory);
        }
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
            _request = null;
            _outputDirectory = null;
            var corpus = _corpus;
            _corpus = null;
            corpus?.Dispose();
        }
    }

    private static MergePlan CreateSingleCellPatchPlan(WorkbookMetadata metadata)
    {
        var sheet = metadata.Sheets[0];
        var localChange = new SheetChange(
            SourceSide.Local,
            ChangeKind.Unchanged,
            sheet,
            sheet,
            ReadOnlyMemory<RowChange>.Empty);
        var remoteChange = new SheetChange(
            SourceSide.Remote,
            ChangeKind.Modified,
            sheet,
            sheet,
            ReadOnlyMemory<RowChange>.Empty);
        var worksheet = new SheetMergePlan(
            SheetGroupId,
            sheet,
            sheet,
            sheet,
            localChange,
            remoteChange,
            new MergeRowMapping[] { new(0, 0, 0) },
            new AutomaticMergeDecision[] { new(
                MergeDecisionScope.Cell,
                AutomaticMergeKind.UseRemote,
                new ConflictLocation(SheetGroupId, 0, 0, 0, 0),
                CellValue.FromText("transactional-writer-benchmark")) });
        return new MergePlan(
            metadata,
            metadata,
            metadata,
            new SheetMergePlan[] { worksheet },
            ReadOnlyMemory<ConflictRecord>.Empty,
            ReadOnlyMemory<CellResolution>.Empty,
            ReadOnlyMemory<RowResolution>.Empty);
    }
}

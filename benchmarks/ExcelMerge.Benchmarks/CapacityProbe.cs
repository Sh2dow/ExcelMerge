using System.Diagnostics;
using ExcelMerge.Domain;
using ExcelMerge.OpenXml;
using ExcelMerge.Storage;

namespace ExcelMerge.Benchmarks;

internal static class CapacityProbe
{
    public static async Task RunAsync()
    {
        var validatePackage = !string.Equals(
            Environment.GetEnvironmentVariable("EXCELMERGE_CAPACITY_VALIDATE"),
            "false",
            StringComparison.OrdinalIgnoreCase);
        using var corpus = XlsxBenchmarkCorpus.Create(
            BenchmarkData.CapacityCellCount,
            XlsxCorpusLayout.Sparse,
            XlsxSharedStringDistribution.Unique);
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = corpus.DirectoryPath,
            DirectoryPrefix = "excelmerge-capacity-probe-",
        });

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        var process = Process.GetCurrentProcess();
        long peakManagedBytes = GC.GetTotalMemory(forceFullCollection: false);
        long peakWorkingSetBytes = process.WorkingSet64;
        using var monitorCancellation = new CancellationTokenSource();
        var monitor = Task.Run(async () =>
        {
            try
            {
                while (true)
                {
                    UpdateMaximum(ref peakManagedBytes, GC.GetTotalMemory(forceFullCollection: false));
                    process.Refresh();
                    UpdateMaximum(ref peakWorkingSetBytes, process.WorkingSet64);
                    await Task.Delay(50, monitorCancellation.Token).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException) when (monitorCancellation.IsCancellationRequested)
            {
            }
        });

        var stopwatch = Stopwatch.StartNew();
        OpenXmlReaderResult result;
        try
        {
            result = await new OpenXmlWorkbookReader().IndexAsync(corpus.FilePath, workspace)
                .ConfigureAwait(false);
        }
        finally
        {
            stopwatch.Stop();
            monitorCancellation.Cancel();
            await monitor.ConfigureAwait(false);
        }

        process.Refresh();
        UpdateMaximum(ref peakManagedBytes, GC.GetTotalMemory(forceFullCollection: false));
        UpdateMaximum(ref peakWorkingSetBytes, process.WorkingSet64);
        var worksheet = result.Worksheets.Single();
        var indexedBytes = worksheet.CellStore.DataLengthBytes + worksheet.CellStore.IndexLengthBytes;
        Console.WriteLine($"Cells={BenchmarkData.CapacityCellCount}");
        Console.WriteLine("Layout=Sparse SharedStrings=Unique");
        Console.WriteLine($"ElapsedSeconds={stopwatch.Elapsed.TotalSeconds:F3}");
        Console.WriteLine($"PeakManagedBytes={peakManagedBytes}");
        Console.WriteLine($"PeakWorkingSetBytes={peakWorkingSetBytes}");
        Console.WriteLine($"IndexedStoreBytes={indexedBytes}");

        await workspace.DisposeAsync().ConfigureAwait(false);
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        peakManagedBytes = GC.GetTotalMemory(forceFullCollection: false);
        process.Refresh();
        peakWorkingSetBytes = process.WorkingSet64;
        using var writerMonitorCancellation = new CancellationTokenSource();
        var writerMonitor = MonitorMemoryAsync(
            process,
            writerMonitorCancellation.Token,
            value => UpdateMaximum(ref peakManagedBytes, value),
            value => UpdateMaximum(ref peakWorkingSetBytes, value));
        var metadata = await new OpenXmlWorkbookReader().ReadMetadataAsync(corpus.FilePath)
            .ConfigureAwait(false);
        var source = await OpenXmlWorkbookSource.CaptureAsync(corpus.FilePath).ConfigureAwait(false);
        var destinationPath = Path.Combine(corpus.DirectoryPath, "capacity-result.xlsx");
        var request = new OpenXmlWriteRequest(
            CreateSingleCellPatchPlan(metadata),
            source,
            source,
            source,
            destinationPath);
        stopwatch.Restart();
        try
        {
            await new OpenXmlWorkbookWriter().WriteAsync(
                request,
                new OpenXmlWriterOptions
                {
                    MinimumFreeSpaceReserveBytes = 0,
                    ValidatePackage = validatePackage,
                }).ConfigureAwait(false);
        }
        finally
        {
            stopwatch.Stop();
            writerMonitorCancellation.Cancel();
            await writerMonitor.ConfigureAwait(false);
        }

        Console.WriteLine($"WriterElapsedSeconds={stopwatch.Elapsed.TotalSeconds:F3}");
        Console.WriteLine($"WriterValidatePackage={validatePackage}");
        Console.WriteLine($"WriterPeakManagedBytes={peakManagedBytes}");
        Console.WriteLine($"WriterPeakWorkingSetBytes={peakWorkingSetBytes}");
    }

    private static async Task MonitorMemoryAsync(
        Process process,
        CancellationToken cancellationToken,
        Action<long> reportManaged,
        Action<long> reportWorkingSet)
    {
        try
        {
            while (true)
            {
                reportManaged(GC.GetTotalMemory(forceFullCollection: false));
                process.Refresh();
                reportWorkingSet(process.WorkingSet64);
                await Task.Delay(50, cancellationToken).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
        }
    }

    private static MergePlan CreateSingleCellPatchPlan(WorkbookMetadata metadata)
    {
        const string sheetGroupId = "capacity-sheet";
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
            sheetGroupId,
            sheet,
            sheet,
            sheet,
            localChange,
            remoteChange,
            new MergeRowMapping[] { new(0, 0, 0) },
            new AutomaticMergeDecision[] { new(
                MergeDecisionScope.Cell,
                AutomaticMergeKind.UseRemote,
                new ConflictLocation(sheetGroupId, 0, 0, 0, 0),
                CellValue.FromText("capacity-probe")) });
        return new MergePlan(
            metadata,
            metadata,
            metadata,
            new SheetMergePlan[] { worksheet },
            ReadOnlyMemory<ConflictRecord>.Empty,
            ReadOnlyMemory<CellResolution>.Empty,
            ReadOnlyMemory<RowResolution>.Empty);
    }

    private static void UpdateMaximum(ref long target, long value)
    {
        while (true)
        {
            var current = Volatile.Read(ref target);
            if (value <= current || Interlocked.CompareExchange(ref target, value, current) == current)
            {
                return;
            }
        }
    }
}

using ExcelMerge.Delimited;
using ExcelMerge.Domain;
using ExcelMerge.Engine;
using ExcelMerge.OpenXml;
using ExcelMerge.Storage;

namespace ExcelMerge.Application;

public sealed class MergeSession : IApplicationSession
{
    private readonly Workspace _workspace;
    private readonly IndexedApplicationWorkbook _baseSource;
    private readonly IndexedApplicationWorkbook _localSource;
    private readonly IndexedApplicationWorkbook _remoteSource;
    private readonly Action<IApplicationSession> _onDisposed;
    private readonly Dictionary<long, ConflictRecord> _conflicts;
    private readonly Dictionary<long, CellResolution> _cellResolutions;
    private readonly Dictionary<long, RowResolution> _rowResolutions;
    private readonly Dictionary<RowOverrideKey, RowMergeOverride> _rowOverrides = [];
    private readonly SemaphoreSlim _saveGate = new(1, 1);
    private readonly CancellationTokenSource _lifetimeCancellation = new();
    private readonly Lazy<Task> _disposeTask;
    private readonly object _sync = new();
    private int _disposed;

    internal MergeSession(
        Workspace workspace,
        IndexedApplicationWorkbook baseSource,
        IndexedApplicationWorkbook localSource,
        IndexedApplicationWorkbook remoteSource,
        IReadOnlyList<MergeSheetResult> sheets,
        Action<IApplicationSession> onDisposed)
    {
        _workspace = workspace;
        _baseSource = baseSource;
        _localSource = localSource;
        _remoteSource = remoteSource;
        _onDisposed = onDisposed;
        Sheets = sheets;
        Base = baseSource.Metadata;
        Local = localSource.Metadata;
        Remote = remoteSource.Metadata;
        _conflicts = sheets
            .SelectMany(static sheet => sheet.Merge.Conflicts.ToArray())
            .ToDictionary(static conflict => conflict.Id);
        _cellResolutions = sheets
            .SelectMany(static sheet => sheet.Merge.CellResolutions.ToArray())
            .ToDictionary(static resolution => resolution.ConflictId);
        _rowResolutions = sheets
            .SelectMany(static sheet => sheet.Merge.RowResolutions.ToArray())
            .ToDictionary(static resolution => resolution.ConflictId);
        Conflicts = _conflicts.Values.OrderBy(static conflict => conflict.Id).ToArray();
        _disposeTask = new Lazy<Task>(DisposeCoreAsync, LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public WorkbookMetadata Base { get; }

    public WorkbookMetadata Local { get; }

    public WorkbookMetadata Remote { get; }

    public IReadOnlyList<MergeSheetResult> Sheets { get; }

    public IReadOnlyList<ConflictRecord> Conflicts { get; }

    public bool IsDisposed => Volatile.Read(ref _disposed) != 0;

    public bool HasUnresolvedConflicts
    {
        get
        {
            lock (_sync)
            {
                ThrowIfDisposed();
                return BuildPlanCore().HasUnresolvedResolutions;
            }
        }
    }

    public void ResolveCell(
        long conflictId,
        ResolutionKind kind,
        CellValue? customValue = null)
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            var conflict = GetConflict(conflictId);
            if (conflict.Kind is not (ConflictKind.CellValue or ConflictKind.CellDeleteEdit))
            {
                throw InvalidResolution(conflictId, "The conflict requires a row-level resolution.");
            }

            try
            {
                _cellResolutions[conflictId] = new CellResolution(conflictId, kind, customValue);
            }
            catch (ArgumentException exception)
            {
                throw InvalidResolution(conflictId, exception.Message, exception);
            }
        }
    }

    public void ResolveRow(
        long conflictId,
        ResolutionKind kind,
        RowRecord? customRow = null)
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            var conflict = GetConflict(conflictId);
            if (conflict.Kind is ConflictKind.CellValue or ConflictKind.CellDeleteEdit)
            {
                throw InvalidResolution(conflictId, "The conflict requires a cell-level resolution.");
            }

            try
            {
                _rowResolutions[conflictId] = new RowResolution(conflictId, kind, customRow);
            }
            catch (ArgumentException exception)
            {
                throw InvalidResolution(conflictId, exception.Message, exception);
            }
        }
    }

    public void SetRowOverride(RowMergeOverride rowOverride)
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            var sheet = Sheets.FirstOrDefault(candidate => string.Equals(
                candidate.Id,
                rowOverride.SheetId,
                StringComparison.Ordinal));
            if (sheet is null || !sheet.Merge.RowMappings.Span.Contains(new MergeRowMapping(
                    rowOverride.BaseRowIndex,
                    rowOverride.LocalRowIndex,
                    rowOverride.RemoteRowIndex)))
            {
                throw new ExcelMergeApplicationException(
                    ApplicationError.InvalidResolution,
                    "The row override does not identify a row in this merge session.");
            }

            _rowOverrides[RowOverrideKey.From(rowOverride)] = rowOverride;
        }
    }

    public bool RemoveRowOverride(
        string sheetId,
        int? baseRowIndex,
        int? localRowIndex,
        int? remoteRowIndex)
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            return _rowOverrides.Remove(new RowOverrideKey(
                sheetId,
                baseRowIndex,
                localRowIndex,
                remoteRowIndex));
        }
    }

    public MergePlan BuildPlan()
    {
        lock (_sync)
        {
            ThrowIfDisposed();
            return BuildPlanCore();
        }
    }

    public async ValueTask<ApplicationSaveResult> SaveAsync(
        SaveRequest request,
        IProgress<ApplicationProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentException.ThrowIfNullOrWhiteSpace(request.DestinationPath);
        ArgumentOutOfRangeException.ThrowIfNegative(request.MinimumFreeSpaceReserveBytes);
        ThrowIfDisposed();
        using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken,
            _lifetimeCancellation.Token);
        var token = linkedCancellation.Token;
        await _saveGate.WaitAsync(token).ConfigureAwait(false);
        try
        {
            ThrowIfDisposed();
            MergePlan plan;
            lock (_sync)
            {
                plan = BuildPlanCore();
            }

            if (plan.HasUnresolvedResolutions)
            {
                throw new ExcelMergeApplicationException(
                    ApplicationError.UnresolvedConflicts,
                    "Every merge conflict must be resolved before saving.");
            }

            ValidateOutputExtension(request.DestinationPath, _localSource.Format);
            progress?.Report(new ApplicationProgress(ApplicationStage.Saving));
            ApplicationSaveResult result;
            if (_localSource.Format == WorkbookFormat.Xlsx)
            {
                var writeRequest = new OpenXmlWriteRequest(
                    plan,
                    _baseSource.OpenXmlSource!,
                    _localSource.OpenXmlSource!,
                    _remoteSource.OpenXmlSource!,
                    request.DestinationPath);
                var written = await new OpenXmlWorkbookWriter().WriteAsync(
                    writeRequest,
                    new OpenXmlWriterOptions
                    {
                        Overwrite = request.Overwrite,
                        ValidatePackage = request.ValidatePackage,
                        MinimumFreeSpaceReserveBytes = request.MinimumFreeSpaceReserveBytes,
                    },
                    new InlineProgress<OpenXmlWriterProgress>(value => progress?.Report(
                        new ApplicationProgress(
                            ApplicationStage.Saving,
                            Completed: value.CompletedWorksheetCount,
                            Total: value.TotalWorksheetCount,
                            SheetId: value.SheetId))),
                    token).ConfigureAwait(false);
                result = new ApplicationSaveResult(
                    written.DestinationPath,
                    WorkbookFormat.Xlsx,
                    written.ContentLength);
            }
            else
            {
                await _baseSource.Fingerprint.VerifyAsync(_baseSource.FilePath, token)
                    .ConfigureAwait(false);
                await _localSource.Fingerprint.VerifyAsync(_localSource.FilePath, token)
                    .ConfigureAwait(false);
                await _remoteSource.Fingerprint.VerifyAsync(_remoteSource.FilePath, token)
                    .ConfigureAwait(false);
                if (Sheets.Count != 1)
                {
                    throw new ExcelMergeApplicationException(
                        ApplicationError.UnsupportedOutput,
                        "Delimited merge output supports exactly one worksheet.");
                }

                var written = await new DelimitedContentWriter().WriteAsync(
                    request.DestinationPath,
                    DelimitedMergeAssembler.AssembleAsync(plan, Sheets[0], token),
                    new DelimitedWriterOptions
                    {
                        Format = _localSource.Format == WorkbookFormat.Csv
                            ? DelimitedFileFormat.Csv
                            : DelimitedFileFormat.Tsv,
                        Overwrite = request.Overwrite,
                    },
                    new InlineProgress<DelimitedWriterProgress>(value => progress?.Report(
                        new ApplicationProgress(
                            ApplicationStage.Saving,
                            Completed: value.RecordsWritten,
                            Rows: value.RecordsWritten,
                            Cells: value.CellsWritten))),
                    token).ConfigureAwait(false);
                result = new ApplicationSaveResult(
                    written.DestinationPath,
                    _localSource.Format,
                    written.ContentLength);
            }

            progress?.Report(new ApplicationProgress(ApplicationStage.Completed, Completed: 1, Total: 1));
            return result;
        }
        finally
        {
            _saveGate.Release();
        }
    }

    public ValueTask DisposeAsync() => new(_disposeTask.Value);

    private async Task DisposeCoreAsync()
    {
        Interlocked.Exchange(ref _disposed, 1);
        _lifetimeCancellation.Cancel();
        await _saveGate.WaitAsync().ConfigureAwait(false);
        try
        {
            await _workspace.DisposeAsync().ConfigureAwait(false);
        }
        finally
        {
            _saveGate.Release();
            _saveGate.Dispose();
            _lifetimeCancellation.Dispose();
            _onDisposed(this);
        }
    }

    private MergePlan BuildPlanCore()
    {
        var results = Sheets.Select(sheet => sheet.Merge with
        {
            CellResolutions = sheet.Merge.CellResolutions.ToArray()
                .Select(resolution => _cellResolutions[resolution.ConflictId])
                .ToArray(),
            RowResolutions = sheet.Merge.RowResolutions.ToArray()
                .Select(resolution => _rowResolutions[resolution.ConflictId])
                .ToArray(),
        });
        return MergePlanFactory.Create(Base, Local, Remote, results) with
        {
            RowOverrides = _rowOverrides.Values.ToArray(),
        };
    }

    private ConflictRecord GetConflict(long conflictId) =>
        _conflicts.TryGetValue(conflictId, out var conflict)
            ? conflict
            : throw InvalidResolution(conflictId, "The conflict does not belong to this session.");

    private void ThrowIfDisposed()
    {
        if (IsDisposed)
        {
            throw new ObjectDisposedException(nameof(MergeSession));
        }
    }

    private static ExcelMergeApplicationException InvalidResolution(
        long conflictId,
        string message,
        Exception? innerException = null) =>
        new(
            ApplicationError.InvalidResolution,
            $"Conflict {conflictId}: {message}",
            innerException: innerException);

    private static void ValidateOutputExtension(string destinationPath, WorkbookFormat format)
    {
        var expected = format switch
        {
            WorkbookFormat.Xlsx => ".xlsx",
            WorkbookFormat.Csv => ".csv",
            WorkbookFormat.Tsv => ".tsv",
            _ => throw new ExcelMergeApplicationException(
                ApplicationError.UnsupportedOutput,
                $"Workbook format '{format}' cannot be saved."),
        };
        if (!Path.GetExtension(destinationPath).Equals(expected, StringComparison.OrdinalIgnoreCase))
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.UnsupportedOutput,
                $"The result path must use the '{expected}' extension.",
                path: Path.GetFullPath(destinationPath));
        }
    }

    private readonly record struct RowOverrideKey(
        string SheetId,
        int? BaseRowIndex,
        int? LocalRowIndex,
        int? RemoteRowIndex)
    {
        public static RowOverrideKey From(in RowMergeOverride value) =>
            new(value.SheetId, value.BaseRowIndex, value.LocalRowIndex, value.RemoteRowIndex);
    }
}

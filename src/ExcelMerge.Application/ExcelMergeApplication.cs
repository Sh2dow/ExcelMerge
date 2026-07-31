using ExcelMerge.Domain;
using ExcelMerge.Engine;
using ExcelMerge.Storage;

namespace ExcelMerge.Application;

public sealed class ExcelMergeApplication : IAsyncDisposable
{
    private readonly ExcelMergeApplicationOptions _options;
    private readonly ApplicationWorkbookLoader _loader;
    private readonly HashSet<IApplicationSession> _sessions = [];
    private readonly object _sync = new();
    private readonly Lazy<Task> _disposeTask;
    private bool _disposed;

    public ExcelMergeApplication(ExcelMergeApplicationOptions? options = null)
    {
        _options = options ?? new ExcelMergeApplicationOptions();
        ArgumentNullException.ThrowIfNull(_options.Workspace);
        ArgumentNullException.ThrowIfNull(_options.Comparison);
        if (_options.ProgressIntervalRows <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(_options.ProgressIntervalRows));
        }

        if (_options.StaleWorkspaceMinimumAge < TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(_options.StaleWorkspaceMinimumAge));
        }

        _loader = new ApplicationWorkbookLoader(_options.ProgressIntervalRows);
        _disposeTask = new Lazy<Task>(DisposeCoreAsync, LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public async ValueTask<CompareSession> OpenCompareAsync(
        CompareRequest request,
        IProgress<ApplicationProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ThrowIfDisposed();
        progress?.Report(new ApplicationProgress(ApplicationStage.Validating));
        var localFormat = _loader.DetectFormat(request.LocalPath, WorkbookSide.Local);
        var remoteFormat = _loader.DetectFormat(request.RemotePath, WorkbookSide.Remote);
        EnsureSameFormat(localFormat, remoteFormat);

        var workspace = await Workspace.CreateAsync(_options.Workspace, cancellationToken)
            .ConfigureAwait(false);
        try
        {
            var local = await _loader.LoadAsync(
                request.LocalPath,
                WorkbookSide.Local,
                localFormat,
                workspace,
                progress,
                cancellationToken).ConfigureAwait(false);
            var remote = await _loader.LoadAsync(
                request.RemotePath,
                WorkbookSide.Remote,
                remoteFormat,
                workspace,
                progress,
                cancellationToken).ConfigureAwait(false);
            var pairs = ApplicationSheetPairing.PairCompare(
                local,
                remote,
                request.LocalSheet,
                request.RemoteSheet);
            var results = new List<CompareSheetResult>(pairs.Count);
            for (var index = 0; index < pairs.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(new ApplicationProgress(
                    ApplicationStage.Comparing,
                    Completed: index,
                    Total: pairs.Count));
                var pair = pairs[index];
                var difference = await new TwoWayDiffEngine().CompareAsync(
                    pair.Local,
                    pair.Remote,
                    SourceSide.Remote,
                    request.Options ?? _options.Comparison,
                    cancellationToken).ConfigureAwait(false);
                results.Add(new CompareSheetResult(
                    $"compare-{index:D4}",
                    pair.Local,
                    pair.Remote,
                    difference));
            }

            var session = new CompareSession(
                workspace,
                local.Metadata,
                remote.Metadata,
                results,
                Unregister);
            Register(session);
            progress?.Report(new ApplicationProgress(
                ApplicationStage.Completed,
                Completed: results.Count,
                Total: results.Count));
            return session;
        }
        catch
        {
            await workspace.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    public async ValueTask<MergeSession> OpenMergeAsync(
        MergeRequest request,
        IProgress<ApplicationProgress>? progress = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        ThrowIfDisposed();
        progress?.Report(new ApplicationProgress(ApplicationStage.Validating));
        var baseFormat = _loader.DetectFormat(request.BasePath, WorkbookSide.Base);
        var localFormat = _loader.DetectFormat(request.LocalPath, WorkbookSide.Local);
        var remoteFormat = _loader.DetectFormat(request.RemotePath, WorkbookSide.Remote);
        EnsureSameFormat(baseFormat, localFormat, remoteFormat);

        var workspace = await Workspace.CreateAsync(_options.Workspace, cancellationToken)
            .ConfigureAwait(false);
        try
        {
            var @base = await _loader.LoadAsync(
                request.BasePath,
                WorkbookSide.Base,
                baseFormat,
                workspace,
                progress,
                cancellationToken).ConfigureAwait(false);
            var local = await _loader.LoadAsync(
                request.LocalPath,
                WorkbookSide.Local,
                localFormat,
                workspace,
                progress,
                cancellationToken).ConfigureAwait(false);
            var remote = await _loader.LoadAsync(
                request.RemotePath,
                WorkbookSide.Remote,
                remoteFormat,
                workspace,
                progress,
                cancellationToken).ConfigureAwait(false);
            var groups = ApplicationSheetPairing.PairMerge(@base, local, remote);
            var results = new List<MergeSheetResult>(groups.Count);
            long nextConflictId = 0;
            for (var index = 0; index < groups.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(new ApplicationProgress(
                    ApplicationStage.Merging,
                    Completed: index,
                    Total: groups.Count));
                var group = groups[index];
                var id = $"sheet-{index:D4}";
                var merge = await new ThreeWayMergeEngine().MergeAsync(
                    group.Base,
                    group.Local,
                    group.Remote,
                    new ThreeWayMergeOptions
                    {
                        Comparison = request.Options ?? _options.Comparison,
                        SheetGroupId = id,
                        FirstConflictId = nextConflictId,
                    },
                    cancellationToken).ConfigureAwait(false);
                nextConflictId = merge.NextConflictId;
                results.Add(new MergeSheetResult(
                    id,
                    group.Base,
                    group.Local,
                    group.Remote,
                    merge));
            }

            var session = new MergeSession(
                workspace,
                @base,
                local,
                remote,
                results,
                Unregister);
            Register(session);
            progress?.Report(new ApplicationProgress(
                ApplicationStage.Completed,
                Completed: results.Count,
                Total: results.Count));
            return session;
        }
        catch
        {
            await workspace.DisposeAsync().ConfigureAwait(false);
            throw;
        }
    }

    public Task<WorkspaceMaintenanceResult> CleanupStaleWorkspacesAsync(
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        return WorkspaceMaintenance.CleanupStaleAsync(
            _options.Workspace,
            _options.StaleWorkspaceMinimumAge,
            cancellationToken);
    }

    public ValueTask DisposeAsync() => new(_disposeTask.Value);

    private async Task DisposeCoreAsync()
    {
        IApplicationSession[] sessions;
        lock (_sync)
        {
            _disposed = true;
            sessions = _sessions.ToArray();
        }

        foreach (var session in sessions)
        {
            await session.DisposeAsync().ConfigureAwait(false);
        }
    }

    private void Register(IApplicationSession session)
    {
        lock (_sync)
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ExcelMergeApplication));
            }

            _sessions.Add(session);
        }
    }

    private void Unregister(IApplicationSession session)
    {
        lock (_sync)
        {
            _sessions.Remove(session);
        }
    }

    private void ThrowIfDisposed()
    {
        lock (_sync)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
        }
    }

    private static void EnsureSameFormat(params WorkbookFormat[] formats)
    {
        if (formats.Length == 0 || formats.Any(format => format != formats[0]))
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.MixedFormats,
                "All files in one operation must use the same workbook format.");
        }
    }
}

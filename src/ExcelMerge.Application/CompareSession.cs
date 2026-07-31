using ExcelMerge.Domain;
using ExcelMerge.Storage;

namespace ExcelMerge.Application;

internal interface IApplicationSession : IAsyncDisposable;

public sealed class CompareSession : IApplicationSession
{
    private readonly Workspace _workspace;
    private readonly Action<IApplicationSession> _onDisposed;
    private readonly Lazy<Task> _disposeTask;
    private int _disposed;

    internal CompareSession(
        Workspace workspace,
        WorkbookMetadata local,
        WorkbookMetadata remote,
        IReadOnlyList<CompareSheetResult> sheets,
        Action<IApplicationSession> onDisposed)
    {
        _workspace = workspace;
        _onDisposed = onDisposed;
        Local = local;
        Remote = remote;
        Sheets = sheets;
        _disposeTask = new Lazy<Task>(DisposeCoreAsync, LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public WorkbookMetadata Local { get; }

    public WorkbookMetadata Remote { get; }

    public IReadOnlyList<CompareSheetResult> Sheets { get; }

    public bool IsDisposed => Volatile.Read(ref _disposed) != 0;

    public ValueTask DisposeAsync() => new(_disposeTask.Value);

    private async Task DisposeCoreAsync()
    {
        Interlocked.Exchange(ref _disposed, 1);
        try
        {
            await _workspace.DisposeAsync().ConfigureAwait(false);
        }
        finally
        {
            _onDisposed(this);
        }
    }
}

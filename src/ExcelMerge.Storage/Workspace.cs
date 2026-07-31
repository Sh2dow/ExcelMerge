using System.Runtime.ExceptionServices;

namespace ExcelMerge.Storage;

/// <summary>
/// Owns an isolated temporary directory and all row stores created in it.
/// </summary>
public sealed class Workspace : IDisposable, IAsyncDisposable
{
    private readonly WorkspaceSettings _settings;
    private readonly HashSet<ChunkedCellStore> _stores = new();
    private readonly object _sync = new();
    private readonly Lazy<Task> _disposeTask;
    private int _nextStoreNumber;
    private bool _disposeRequested;

    public Workspace(WorkspaceOptions? options = null)
    {
        var requestedOptions = options ?? new WorkspaceOptions();
        _settings = requestedOptions.Validate();

        Directory.CreateDirectory(_settings.BaseDirectory);
        DirectoryPath = CreateUniqueDirectory(_settings.BaseDirectory, _settings.DirectoryPrefix);
        Options = requestedOptions with { BaseDirectory = _settings.BaseDirectory };
        _disposeTask = new Lazy<Task>(DisposeCoreAsync, LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public WorkspaceOptions Options { get; }

    public string DirectoryPath { get; }

    public bool IsDisposed
    {
        get
        {
            lock (_sync)
            {
                return _disposeRequested;
            }
        }
    }

    public static ValueTask<Workspace> CreateAsync(
        WorkspaceOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return ValueTask.FromResult(new Workspace(options));
    }

    /// <summary>
    /// Resolves a relative path and rejects paths that escape the workspace.
    /// </summary>
    public string ResolvePath(string relativePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(relativePath);
        ThrowIfDisposed();

        if (Path.IsPathRooted(relativePath))
        {
            throw new ArgumentException("The workspace path must be relative.", nameof(relativePath));
        }

        var fullPath = Path.GetFullPath(Path.Combine(DirectoryPath, relativePath));
        var workspacePrefix = DirectoryPath.EndsWith(Path.DirectorySeparatorChar)
            ? DirectoryPath
            : DirectoryPath + Path.DirectorySeparatorChar;
        var comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

        if (!fullPath.StartsWith(workspacePrefix, comparison))
        {
            throw new ArgumentException("The path escapes the workspace directory.", nameof(relativePath));
        }

        return fullPath;
    }

    public string CreateSubdirectory(string relativePath)
    {
        lock (_sync)
        {
            ThrowIfDisposedLocked();
            var fullPath = ResolvePath(relativePath);
            Directory.CreateDirectory(fullPath);
            return fullPath;
        }
    }

    public ChunkedCellStore CreateCellStore(string? name = null)
    {
        lock (_sync)
        {
            ThrowIfDisposedLocked();

            name ??= $"store-{_nextStoreNumber++:D4}";
            ValidateStoreName(name);

            var storeDirectory = Path.Combine(DirectoryPath, name);
            if (Directory.Exists(storeDirectory) || File.Exists(storeDirectory))
            {
                throw new IOException($"A workspace entry named '{name}' already exists.");
            }

            Directory.CreateDirectory(storeDirectory);

            try
            {
                var store = new ChunkedCellStore(storeDirectory, _settings, UnregisterStore);
                _stores.Add(store);
                return store;
            }
            catch
            {
                TryDeleteDirectory(storeDirectory);
                throw;
            }
        }
    }

    public ValueTask<ChunkedCellStore> CreateCellStoreAsync(
        string? name = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return ValueTask.FromResult(CreateCellStore(name));
    }

    public void Dispose() => _disposeTask.Value.GetAwaiter().GetResult();

    public ValueTask DisposeAsync() => new(_disposeTask.Value);

    private async Task DisposeCoreAsync()
    {
        ChunkedCellStore[] stores;
        lock (_sync)
        {
            _disposeRequested = true;
            stores = _stores.ToArray();
        }

        List<Exception>? errors = null;
        foreach (var store in stores)
        {
            try
            {
                await store.DisposeAsync().ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                (errors ??= []).Add(exception);
            }
        }

        if (_settings.DeleteOnDispose)
        {
            try
            {
                await DeleteWorkspaceDirectoryAsync().ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                (errors ??= []).Add(exception);
            }
        }

        if (errors is { Count: 1 })
        {
            ExceptionDispatchInfo.Capture(errors[0]).Throw();
        }

        if (errors is { Count: > 1 })
        {
            throw new AggregateException("One or more workspace resources could not be disposed.", errors);
        }
    }

    private async Task DeleteWorkspaceDirectoryAsync()
    {
        for (var attempt = 0; ; attempt++)
        {
            try
            {
                if (Directory.Exists(DirectoryPath))
                {
                    MakeTreeWritable(DirectoryPath);
                    Directory.Delete(DirectoryPath, recursive: true);
                }

                return;
            }
            catch (Exception exception) when (
                exception is IOException or UnauthorizedAccessException &&
                attempt < _settings.CleanupRetryCount)
            {
                if (_settings.CleanupRetryDelay > TimeSpan.Zero)
                {
                    await Task.Delay(_settings.CleanupRetryDelay).ConfigureAwait(false);
                }
            }
        }
    }

    private void UnregisterStore(ChunkedCellStore store)
    {
        lock (_sync)
        {
            _stores.Remove(store);
        }
    }

    private void ThrowIfDisposed()
    {
        lock (_sync)
        {
            ThrowIfDisposedLocked();
        }
    }

    private void ThrowIfDisposedLocked()
    {
        ObjectDisposedException.ThrowIf(_disposeRequested, this);
    }

    private static string CreateUniqueDirectory(string baseDirectory, string prefix)
    {
        for (var attempt = 0; attempt < 10; attempt++)
        {
            var path = Path.Combine(baseDirectory, prefix + Guid.NewGuid().ToString("N"));
            if (Directory.Exists(path) || File.Exists(path))
            {
                continue;
            }

            Directory.CreateDirectory(path);
            return Path.GetFullPath(path);
        }

        throw new IOException("Could not allocate a unique workspace directory.");
    }

    private static void ValidateStoreName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) ||
            name is "." or ".." ||
            name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
            name.Contains(Path.DirectorySeparatorChar) ||
            name.Contains(Path.AltDirectorySeparatorChar))
        {
            throw new ArgumentException("The store name must be a valid single directory name.", nameof(name));
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch
        {
            // Preserve the store-creation exception. Workspace disposal retries cleanup.
        }
    }

    private static void MakeTreeWritable(string rootPath)
    {
        var directories = new Stack<string>();
        directories.Push(rootPath);

        while (directories.TryPop(out var directory))
        {
            foreach (var entry in Directory.EnumerateFileSystemEntries(directory))
            {
                var attributes = File.GetAttributes(entry);
                if ((attributes & FileAttributes.ReparsePoint) != 0)
                {
                    continue;
                }

                if ((attributes & FileAttributes.ReadOnly) != 0)
                {
                    File.SetAttributes(entry, attributes & ~FileAttributes.ReadOnly);
                }

                if ((attributes & FileAttributes.Directory) != 0)
                {
                    directories.Push(entry);
                }
            }
        }

        var rootAttributes = File.GetAttributes(rootPath);
        if ((rootAttributes & FileAttributes.ReadOnly) != 0)
        {
            File.SetAttributes(rootPath, rootAttributes & ~FileAttributes.ReadOnly);
        }
    }
}

using System.Text.Json;
using System.Text.Json.Nodes;
using System.Collections.Concurrent;

namespace ExcelMerge.Application;

public sealed record ApplicationSettings
{
    public int MaximumRecentSessions { get; init; } = 20;

    public string Theme { get; init; } = "System";

    public IReadOnlyList<int> KeyColumns { get; init; } = Array.Empty<int>();

    public bool CompareFormulaCachedValues { get; init; } = true;

    public bool CompareDisplayText { get; init; }

    public bool CompareCellStyles { get; init; }

    public bool CompareRowMetadata { get; init; }

    public bool TreatExplicitBlankCellsAsMissing { get; init; } = true;
}

public enum RecentSessionMode
{
    Compare,
    Merge,
}

public sealed record RecentSessionDescriptor(
    Guid Id,
    RecentSessionMode Mode,
    string? BasePath,
    string LocalPath,
    string RemotePath,
    DateTimeOffset LastOpenedUtc,
    string? LocalSheetId = null,
    string? RemoteSheetId = null,
    IReadOnlyList<int>? KeyColumns = null);

public interface IApplicationSettingsStore
{
    ValueTask<ApplicationSettings> LoadAsync(CancellationToken cancellationToken = default);

    ValueTask SaveAsync(
        ApplicationSettings settings,
        CancellationToken cancellationToken = default);
}

public interface IRecentSessionStore
{
    ValueTask<IReadOnlyList<RecentSessionDescriptor>> ListAsync(
        CancellationToken cancellationToken = default);

    ValueTask RecordAsync(
        RecentSessionDescriptor session,
        CancellationToken cancellationToken = default);

    ValueTask<bool> RemoveAsync(
        Guid id,
        CancellationToken cancellationToken = default);
}

public sealed class JsonApplicationStateStore :
    IApplicationSettingsStore,
    IRecentSessionStore,
    IAsyncDisposable
{
    private const int FormatVersion = 1;
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> FileGates =
        new(OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
    private readonly SemaphoreSlim _gate;
    private int _disposed;

    public JsonApplicationStateStore(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        FilePath = Path.GetFullPath(filePath);
        _gate = FileGates.GetOrAdd(FilePath, static _ => new SemaphoreSlim(1, 1));
    }

    public string FilePath { get; }

    public async ValueTask<ApplicationSettings> LoadAsync(
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return (await ReadStateAsync(cancellationToken).ConfigureAwait(false)).Settings;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async ValueTask SaveAsync(
        ApplicationSettings settings,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(settings);
        ValidateSettings(settings);
        ThrowIfDisposed();
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var state = await ReadStateAsync(cancellationToken).ConfigureAwait(false);
            var recent = state.Recent
                .Take(settings.MaximumRecentSessions)
                .ToArray();
            await WriteStateAsync(
                new StoredState(settings, recent),
                cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    public async ValueTask<IReadOnlyList<RecentSessionDescriptor>> ListAsync(
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return (await ReadStateAsync(cancellationToken).ConfigureAwait(false)).Recent;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async ValueTask RecordAsync(
        RecentSessionDescriptor session,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(session);
        ValidateSession(session);
        ThrowIfDisposed();
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var state = await ReadStateAsync(cancellationToken).ConfigureAwait(false);
            var recent = state.Recent
                .Where(item => item.Id != session.Id)
                .Prepend(session with
                {
                    BasePath = NormalizeOptionalPath(session.BasePath),
                    LocalPath = Path.GetFullPath(session.LocalPath),
                    RemotePath = Path.GetFullPath(session.RemotePath),
                    KeyColumns = session.KeyColumns?.ToArray() ?? Array.Empty<int>(),
                })
                .OrderByDescending(static item => item.LastOpenedUtc)
                .Take(state.Settings.MaximumRecentSessions)
                .ToArray();
            await WriteStateAsync(
                state with { Recent = recent },
                cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            _gate.Release();
        }
    }

    public async ValueTask<bool> RemoveAsync(
        Guid id,
        CancellationToken cancellationToken = default)
    {
        ThrowIfDisposed();
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var state = await ReadStateAsync(cancellationToken).ConfigureAwait(false);
            var recent = state.Recent.Where(item => item.Id != id).ToArray();
            if (recent.Length == state.Recent.Count)
            {
                return false;
            }

            await WriteStateAsync(
                state with { Recent = recent },
                cancellationToken).ConfigureAwait(false);
            return true;
        }
        finally
        {
            _gate.Release();
        }
    }

    public ValueTask DisposeAsync()
    {
        if (Interlocked.Exchange(ref _disposed, 1) == 0)
        {
            // File gates are process-wide because multiple store instances may share one state file.
        }

        return ValueTask.CompletedTask;
    }

    private async ValueTask<StoredState> ReadStateAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(FilePath))
        {
            return StoredState.Default;
        }

        await using var stream = new FileStream(
            FilePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 16 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        var root = await JsonNode.ParseAsync(
            stream,
            documentOptions: default,
            cancellationToken: cancellationToken).ConfigureAwait(false) as JsonObject ??
            throw InvalidState("The application state root must be a JSON object.");
        if (ReadInt(root, "version", required: true) != FormatVersion)
        {
            throw InvalidState("The application state version is not supported.");
        }

        var settingsNode = root["settings"] as JsonObject ??
            throw InvalidState("The application state has no settings object.");
        var settings = new ApplicationSettings
        {
            MaximumRecentSessions = ReadInt(settingsNode, "maximumRecentSessions", required: true),
            Theme = ReadTheme(settingsNode),
            KeyColumns = ReadIntArray(settingsNode, "keyColumns"),
            CompareFormulaCachedValues = ReadBool(settingsNode, "compareFormulaCachedValues"),
            CompareDisplayText = ReadBool(settingsNode, "compareDisplayText"),
            CompareCellStyles = ReadBool(settingsNode, "compareCellStyles"),
            CompareRowMetadata = ReadBool(settingsNode, "compareRowMetadata"),
            TreatExplicitBlankCellsAsMissing = ReadBool(
                settingsNode,
                "treatExplicitBlankCellsAsMissing"),
        };
        ValidateSettings(settings);

        var recent = new List<RecentSessionDescriptor>();
        if (root["recent"] is JsonArray recentArray)
        {
            foreach (var node in recentArray)
            {
                if (node is not JsonObject item ||
                    !Guid.TryParse(ReadString(item, "id", required: true), out var id) ||
                    !Enum.TryParse<RecentSessionMode>(
                        ReadString(item, "mode", required: true),
                        ignoreCase: true,
                        out var mode) ||
                    !DateTimeOffset.TryParse(
                        ReadString(item, "lastOpenedUtc", required: true),
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.RoundtripKind,
                        out var lastOpenedUtc))
                {
                    throw InvalidState("A recent-session descriptor is invalid.");
                }

                var descriptor = new RecentSessionDescriptor(
                    id,
                    mode,
                    ReadString(item, "basePath", required: false),
                    ReadString(item, "localPath", required: true)!,
                    ReadString(item, "remotePath", required: true)!,
                    lastOpenedUtc,
                    ReadString(item, "localSheetId", required: false),
                    ReadString(item, "remoteSheetId", required: false),
                    ReadIntArray(item, "keyColumns"));
                ValidateSession(descriptor);
                recent.Add(descriptor);
            }
        }

        return new StoredState(
            settings,
            recent.OrderByDescending(static item => item.LastOpenedUtc)
                .Take(settings.MaximumRecentSessions)
                .ToArray());
    }

    private async ValueTask WriteStateAsync(
        StoredState state,
        CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(FilePath);
        if (string.IsNullOrEmpty(directory))
        {
            throw new DirectoryNotFoundException("The application state directory could not be determined.");
        }

        Directory.CreateDirectory(directory);
        var temporaryPath = Path.Combine(
            directory,
            $".{Path.GetFileName(FilePath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            var root = new JsonObject
            {
                ["version"] = FormatVersion,
                ["settings"] = WriteSettings(state.Settings),
                ["recent"] = new JsonArray(state.Recent.Select(WriteRecent).ToArray()),
            };
            await using (var stream = new FileStream(
                temporaryPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                bufferSize: 16 * 1024,
                FileOptions.Asynchronous | FileOptions.SequentialScan))
            {
                using var jsonWriter = new Utf8JsonWriter(stream);
                root.WriteTo(jsonWriter);
                await jsonWriter.FlushAsync(cancellationToken).ConfigureAwait(false);
                await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                stream.Flush(flushToDisk: true);
            }

            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporaryPath, FilePath, overwrite: true);
        }
        finally
        {
            try
            {
                if (File.Exists(temporaryPath))
                {
                    File.Delete(temporaryPath);
                }
            }
            catch
            {
                // Preserve the persistence result; startup maintenance can remove abandoned temp files.
            }
        }
    }

    private static JsonObject WriteSettings(ApplicationSettings settings) =>
        new()
        {
            ["maximumRecentSessions"] = settings.MaximumRecentSessions,
            ["theme"] = settings.Theme,
            ["keyColumns"] = new JsonArray(settings.KeyColumns
                .Select(static value => (JsonNode?)JsonValue.Create(value))
                .ToArray()),
            ["compareFormulaCachedValues"] = settings.CompareFormulaCachedValues,
            ["compareDisplayText"] = settings.CompareDisplayText,
            ["compareCellStyles"] = settings.CompareCellStyles,
            ["compareRowMetadata"] = settings.CompareRowMetadata,
            ["treatExplicitBlankCellsAsMissing"] = settings.TreatExplicitBlankCellsAsMissing,
        };

    private static JsonNode WriteRecent(RecentSessionDescriptor session) =>
        new JsonObject
        {
            ["id"] = session.Id.ToString("D"),
            ["mode"] = session.Mode.ToString(),
            ["basePath"] = session.BasePath,
            ["localPath"] = session.LocalPath,
            ["remotePath"] = session.RemotePath,
            ["lastOpenedUtc"] = session.LastOpenedUtc.ToString("O"),
            ["localSheetId"] = session.LocalSheetId,
            ["remoteSheetId"] = session.RemoteSheetId,
            ["keyColumns"] = new JsonArray(
                (session.KeyColumns ?? Array.Empty<int>())
                    .Select(static value => (JsonNode?)JsonValue.Create(value))
                    .ToArray()),
        };

    private static string ReadTheme(JsonObject source) =>
        source["theme"]?.GetValue<string>() is { } theme && theme is "Light" or "Dark"
            ? theme
            : "System";

    private static int ReadInt(JsonObject source, string name, bool required) =>
        source[name]?.GetValue<int>() ?? (required
            ? throw InvalidState($"State value '{name}' is missing.")
            : default);

    private static bool ReadBool(JsonObject source, string name) =>
        source[name]?.GetValue<bool>() ?? throw InvalidState($"State value '{name}' is missing.");

    private static string? ReadString(JsonObject source, string name, bool required) =>
        source[name]?.GetValue<string>() ?? (required
            ? throw InvalidState($"State value '{name}' is missing.")
            : null);

    private static int[] ReadIntArray(JsonObject source, string name) =>
        source[name] is JsonArray values
            ? values.Select(value => value?.GetValue<int>() ??
                throw InvalidState($"State array '{name}' contains a null value.")).ToArray()
            : throw InvalidState($"State array '{name}' is missing.");

    private static void ValidateSettings(ApplicationSettings settings)
    {
        if (settings.MaximumRecentSessions is < 0 or > 1000 ||
            settings.KeyColumns is null ||
            settings.Theme is not ("System" or "Light" or "Dark"))
        {
            throw new ArgumentOutOfRangeException(nameof(settings));
        }

        var previous = -1;
        foreach (var column in settings.KeyColumns.Order())
        {
            if (column < 0 || column == previous)
            {
                throw new ArgumentException("Settings key columns must be unique and non-negative.");
            }

            previous = column;
        }
    }

    private static void ValidateSession(RecentSessionDescriptor session)
    {
        if (session.Id == Guid.Empty ||
            !Enum.IsDefined(session.Mode) ||
            string.IsNullOrWhiteSpace(session.LocalPath) ||
            string.IsNullOrWhiteSpace(session.RemotePath) ||
            session.Mode == RecentSessionMode.Merge && string.IsNullOrWhiteSpace(session.BasePath) ||
            session.KeyColumns?.Any(static column => column < 0) == true)
        {
            throw new ArgumentException("The recent-session descriptor is invalid.", nameof(session));
        }
    }

    private static string? NormalizeOptionalPath(string? path) =>
        string.IsNullOrWhiteSpace(path) ? null : Path.GetFullPath(path);

    private void ThrowIfDisposed() =>
        ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed) != 0, this);

    private static InvalidDataException InvalidState(string message) => new(message);

    private sealed record StoredState(
        ApplicationSettings Settings,
        IReadOnlyList<RecentSessionDescriptor> Recent)
    {
        public static StoredState Default { get; } = new(new ApplicationSettings(), []);
    }
}

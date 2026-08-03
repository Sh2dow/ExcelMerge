using System.Collections.ObjectModel;
using ExcelMerge.Application;
using ExcelMerge.Domain;
using ExcelMerge.Engine;

namespace ExcelMerge.Desktop;

public enum DesktopMode
{
    Compare,
    Merge,
}

public sealed class MainWindowViewModel : ObservableObject, IAsyncDisposable
{
    private readonly ExcelMergeApplication _application;
    private readonly IApplicationSettingsStore _settingsStore;
    private readonly IRecentSessionStore _recentStore;
    private CompareSession? _compareSession;
    private MergeSession? _mergeSession;
    private CancellationTokenSource? _operationCancellation;
    private DesktopMode _mode;
    private string _basePath = string.Empty;
    private string _localPath = string.Empty;
    private string _remotePath = string.Empty;
    private bool _isBusy;
    private string _status = LocalizationService.Get("NoSession");
    private double _progressValue;
    private bool _progressIndeterminate;
    private SheetItemViewModel? _selectedSheet;
    private ConflictItemViewModel? _selectedConflict;
    private GridDocument? _gridDocument;
    private string _searchText = string.Empty;
    private bool _searchExact;
    private bool _searchCaseSensitive;
    private bool _searchRegex;
    private string _searchSummary = string.Empty;
    private string _customValue = string.Empty;
    private int _currentGridRow = -1;
    private IReadOnlyList<GridSearchResult> _searchResults = Array.Empty<GridSearchResult>();
    private int _searchResultIndex = -1;
    private ApplicationSettings _settings = new();
    private bool _hideUnchanged;
    private bool _inputsExpanded = true;
    private string? _suggestedOutputPath;
    private bool _suggestedOutputSaved;
    private bool _selectingConflictFromGrid;
    private int _cellSelectionVersion;

    public MainWindowViewModel(
        ExcelMergeApplication application,
        IApplicationSettingsStore settingsStore,
        IRecentSessionStore recentStore)
    {
        _application = application;
        _settingsStore = settingsStore;
        _recentStore = recentStore;
        RunCommand = new AsyncCommand(RunAsync, CanRun);
        CancelCommand = new RelayCommand(Cancel, () => IsBusy);
        SearchCommand = new AsyncCommand(SearchAsync, () => GridDocument is not null && !IsBusy);
        PreviousChangeCommand = new RelayCommand(() => NavigateChange(-1), () => GridDocument is not null);
        NextChangeCommand = new RelayCommand(() => NavigateChange(1), () => GridDocument is not null);
        PreviousSearchCommand = new RelayCommand(() => NavigateSearch(-1), () => _searchResults.Count != 0);
        NextSearchCommand = new RelayCommand(() => NavigateSearch(1), () => _searchResults.Count != 0);
        UseLocalCommand = new RelayCommand(
            () => ResolveSelected(ResolutionKind.Local),
            () => CanResolveSelected);
        UseRemoteCommand = new RelayCommand(
            () => ResolveSelected(ResolutionKind.Remote),
            () => CanResolveSelected);
        UseBothCommand = new RelayCommand(
            () => ResolveSelected(ResolutionKind.Both),
            () => CanUseBoth);
        UseCustomCommand = new RelayCommand(
            ResolveCustom,
            () => CanUseCustom);
        SwapSidesCommand = new RelayCommand(SwapSides, () => !IsBusy);
        ToggleInputsCommand = new RelayCommand(() => InputsExpanded = !InputsExpanded);
        _ = InitializeAsync();
    }

    public ObservableCollection<SheetItemViewModel> Sheets { get; } = [];

    public ObservableCollection<ConflictItemViewModel> Conflicts { get; } = [];

    public ObservableCollection<RecentItemViewModel> RecentSessions { get; } = [];

    public ObservableCollection<string> Logs { get; } = [];

    public InspectorViewModel Inspector { get; } = new();

    public AsyncCommand RunCommand { get; }

    public RelayCommand CancelCommand { get; }

    public AsyncCommand SearchCommand { get; }

    public RelayCommand PreviousChangeCommand { get; }

    public RelayCommand NextChangeCommand { get; }

    public RelayCommand PreviousSearchCommand { get; }

    public RelayCommand NextSearchCommand { get; }

    public RelayCommand UseLocalCommand { get; }

    public RelayCommand UseRemoteCommand { get; }

    public RelayCommand UseBothCommand { get; }

    public RelayCommand UseCustomCommand { get; }

    public RelayCommand SwapSidesCommand { get; }

    public RelayCommand ToggleInputsCommand { get; }

    public event Action<int, int>? NavigationRequested;

    public event Action<ConflictItemViewModel>? ConflictSelectionRequested;

    public DesktopMode Mode
    {
        get => _mode;
        set
        {
            if (SetProperty(ref _mode, value))
            {
                OnPropertyChanged(nameof(IsMergeMode));
                OnPropertyChanged(nameof(IsCompareMode));
                RunCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsMergeMode => Mode == DesktopMode.Merge;

    public bool IsCompareMode => Mode == DesktopMode.Compare;

    public ApplicationSettings CurrentSettings => _settings;

    public bool HideUnchanged
    {
        get => _hideUnchanged;
        set
        {
            if (SetProperty(ref _hideUnchanged, value))
            {
                RefreshSelectedSheet();
            }
        }
    }

    public bool InputsExpanded
    {
        get => _inputsExpanded;
        set
        {
            if (SetProperty(ref _inputsExpanded, value))
            {
                OnPropertyChanged(nameof(InputToggleText));
            }
        }
    }

    public string InputToggleText => InputsExpanded
        ? LocalizationService.Get("Collapse")
        : LocalizationService.Get("Expand");

    public string? SuggestedOutputPath
    {
        get => _suggestedOutputPath;
        private set
        {
            if (SetProperty(ref _suggestedOutputPath, value))
            {
                SuggestedOutputSaved = false;
            }
        }
    }

    public bool SuggestedOutputSaved
    {
        get => _suggestedOutputSaved;
        private set => SetProperty(ref _suggestedOutputSaved, value);
    }

    public string BasePath
    {
        get => _basePath;
        set
        {
            if (SetProperty(ref _basePath, value))
            {
                RunCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public string LocalPath
    {
        get => _localPath;
        set
        {
            if (SetProperty(ref _localPath, value))
            {
                RunCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public string RemotePath
    {
        get => _remotePath;
        set
        {
            if (SetProperty(ref _remotePath, value))
            {
                RunCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public bool IsBusy
    {
        get => _isBusy;
        private set
        {
            if (SetProperty(ref _isBusy, value))
            {
                RunCommand.RaiseCanExecuteChanged();
                CancelCommand.RaiseCanExecuteChanged();
                SearchCommand.RaiseCanExecuteChanged();
                SwapSidesCommand.RaiseCanExecuteChanged();
                OnPropertyChanged(nameof(CanSave));
            }
        }
    }

    public string Status
    {
        get => _status;
        private set => SetProperty(ref _status, value);
    }

    public double ProgressValue
    {
        get => _progressValue;
        private set => SetProperty(ref _progressValue, value);
    }

    public bool ProgressIndeterminate
    {
        get => _progressIndeterminate;
        private set => SetProperty(ref _progressIndeterminate, value);
    }

    public SheetItemViewModel? SelectedSheet
    {
        get => _selectedSheet;
        set
        {
            if (!SetProperty(ref _selectedSheet, value))
            {
                return;
            }

            RefreshSelectedSheet();
        }
    }

    public ConflictItemViewModel? SelectedConflict
    {
        get => _selectedConflict;
        set
        {
            if (!_selectingConflictFromGrid)
            {
                _cellSelectionVersion++;
            }

            if (!SetProperty(ref _selectedConflict, value))
            {
                return;
            }

            OnPropertyChanged(nameof(CanResolveSelected));
            OnPropertyChanged(nameof(CanUseBoth));
            OnPropertyChanged(nameof(CanUseCustom));
            RaiseResolutionCanExecute();
            if (value is not null && !_selectingConflictFromGrid)
            {
                NavigateToConflict(value.Conflict);
            }
        }
    }

    public GridDocument? GridDocument
    {
        get => _gridDocument;
        private set
        {
            if (SetProperty(ref _gridDocument, value))
            {
                SearchCommand.RaiseCanExecuteChanged();
                PreviousChangeCommand.RaiseCanExecuteChanged();
                NextChangeCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public string SearchText
    {
        get => _searchText;
        set => SetProperty(ref _searchText, value);
    }

    public bool SearchExact
    {
        get => _searchExact;
        set => SetProperty(ref _searchExact, value);
    }

    public bool SearchCaseSensitive
    {
        get => _searchCaseSensitive;
        set => SetProperty(ref _searchCaseSensitive, value);
    }

    public bool SearchRegex
    {
        get => _searchRegex;
        set => SetProperty(ref _searchRegex, value);
    }

    public string SearchSummary
    {
        get => _searchSummary;
        private set => SetProperty(ref _searchSummary, value);
    }

    public string CustomValue
    {
        get => _customValue;
        set
        {
            if (SetProperty(ref _customValue, value))
            {
                UseCustomCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public bool CanResolveSelected => _mergeSession is not null && SelectedConflict is not null && !IsBusy;

    public bool CanUseBoth => CanResolveSelected && SelectedConflict!.SupportsBoth;

    public bool CanUseCustom =>
        CanResolveSelected && !SelectedConflict!.IsRowConflict && CustomValue.Length != 0;

    public bool CanSave =>
        _mergeSession is not null && !_mergeSession.HasUnresolvedConflicts && !IsBusy;

    public string ConflictSummary
    {
        get
        {
            if (_mergeSession is null)
            {
                return string.Empty;
            }

            var plan = _mergeSession.BuildPlan();
            var resolutions = BuildEffectiveResolutions(plan);
            var unresolved = resolutions.Values.Count(static kind => kind == ResolutionKind.Unresolved);
            var total = plan.Conflicts.Length;
            return $"{unresolved} {LocalizationService.Get("Unresolved")} / " +
                $"{total - unresolved} {LocalizationService.Get("Resolved")}";
        }
    }

    public void SetMode(DesktopMode mode) => Mode = mode;

    public void CancelPendingCellSelection() => _cellSelectionVersion++;

    public void ApplyStartupArguments(IReadOnlyList<string> arguments)
    {
        if (arguments.Count == 3 &&
            string.Equals(arguments[0], "diff", StringComparison.OrdinalIgnoreCase))
        {
            Mode = DesktopMode.Compare;
            SuggestedOutputPath = null;
            LocalPath = arguments[1];
            RemotePath = arguments[2];
            RunCommand.Execute(null);
        }
        else if (arguments.Count is 4 or 5 &&
            string.Equals(arguments[0], "merge", StringComparison.OrdinalIgnoreCase))
        {
            Mode = DesktopMode.Merge;
            BasePath = arguments[1];
            LocalPath = arguments[2];
            RemotePath = arguments[3];
            SuggestedOutputPath = arguments.Count == 5
                ? Path.GetFullPath(arguments[4])
                : null;
            RunCommand.Execute(null);
        }
    }

    public async Task SaveSettingsAsync(ApplicationSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        await _settingsStore.SaveAsync(settings);
        _settings = settings;
        AddLog(LocalizationService.Get("SettingsSaved"));
    }

    public async Task<bool> SaveAsync(string destinationPath)
    {
        if (_mergeSession is null || !CanSave)
        {
            return false;
        }

        var saved = await RunOperationAsync(async token =>
        {
            Status = LocalizationService.Get("Saving");
            AddLog(Status);
            await _mergeSession.SaveAsync(
                new SaveRequest(destinationPath),
                new Progress<ApplicationProgress>(HandleProgress),
                token);
            Status = LocalizationService.Get("Completed");
            AddLog($"{Status}: {destinationPath}");
        });
        if (saved &&
            SuggestedOutputPath is not null &&
            string.Equals(
                Path.GetFullPath(destinationPath),
                SuggestedOutputPath,
                OperatingSystem.IsWindows()
                    ? StringComparison.OrdinalIgnoreCase
                    : StringComparison.Ordinal))
        {
            SuggestedOutputSaved = true;
        }

        return saved;
    }

    public async Task SelectCellAsync(GridCellSelection selection)
    {
        var document = GridDocument;
        if (document is null || selection.ViewRowIndex >= document.RowCount)
        {
            return;
        }

        var selectionVersion = ++_cellSelectionVersion;
        _currentGridRow = selection.ViewRowIndex;
        var row = await document.LoadRowAsync(selection.ViewRowIndex);
        if (selectionVersion != _cellSelectionVersion || !ReferenceEquals(document, GridDocument))
        {
            return;
        }

        Inspector.Address = GridDocument.FormatAddress(
            row.Descriptor.LocalRowIndex ?? row.Descriptor.RemoteRowIndex ??
                row.Descriptor.BaseRowIndex ?? 0,
            selection.ColumnIndex);
        Inspector.BaseValue = DescribeCell(GridDocument.FindCell(row.BaseRow, selection.ColumnIndex));
        Inspector.LocalValue = DescribeCell(GridDocument.FindCell(row.LocalRow, selection.ColumnIndex));
        Inspector.RemoteValue = DescribeCell(GridDocument.FindCell(row.RemoteRow, selection.ColumnIndex));
        Inspector.ResultValue = selection.Pane == GridPane.Local
            ? Inspector.LocalValue
            : Inspector.RemoteValue;
        var selected = selection.Pane == GridPane.Local
            ? GridDocument.FindCell(row.LocalRow, selection.ColumnIndex)
            : GridDocument.FindCell(row.RemoteRow, selection.ColumnIndex);
        Inspector.Formula = selected?.Value.Formula ?? string.Empty;
        Inspector.ValueType = selected?.Value.Kind.ToString() ?? CellKind.Blank.ToString();

        var conflict = Conflicts.FirstOrDefault(item =>
            item.Matches(row.Descriptor, selection.ColumnIndex));
        if (conflict is not null)
        {
            if (!ReferenceEquals(SelectedConflict, conflict))
            {
                _selectingConflictFromGrid = true;
                try
                {
                    SelectedConflict = conflict;
                }
                finally
                {
                    _selectingConflictFromGrid = false;
                }
            }

            ConflictSelectionRequested?.Invoke(conflict);
        }
    }

    public string GetSelectedCopyText(char delimiter)
    {
        var values = new[] { Inspector.LocalValue, Inspector.RemoteValue };
        return string.Join(delimiter, values.Select(value => Escape(value, delimiter)));
    }

    public void ApplyRecent(RecentItemViewModel recent)
    {
        Mode = recent.Descriptor.Mode == RecentSessionMode.Merge
            ? DesktopMode.Merge
            : DesktopMode.Compare;
        BasePath = recent.Descriptor.BasePath ?? string.Empty;
        LocalPath = recent.Descriptor.LocalPath;
        RemotePath = recent.Descriptor.RemotePath;
    }

    public async ValueTask DisposeAsync()
    {
        _operationCancellation?.Cancel();
        await DisposeSessionsAsync();
        await _application.DisposeAsync();
        if (ReferenceEquals(_settingsStore, _recentStore))
        {
            if (_settingsStore is IAsyncDisposable shared)
            {
                await shared.DisposeAsync();
            }
        }
        else
        {
            if (_settingsStore is IAsyncDisposable settings)
                await settings.DisposeAsync();
            if (_recentStore is IAsyncDisposable recent)
                await recent.DisposeAsync();
        }
    }

    private async Task InitializeAsync()
    {
        try
        {
            await _application.CleanupStaleWorkspacesAsync();
            _settings = await _settingsStore.LoadAsync();
            await RefreshRecentAsync();
        }
        catch (Exception exception)
        {
            AddLog(exception.Message);
        }
    }

    private bool CanRun() =>
        !IsBusy &&
        !string.IsNullOrWhiteSpace(LocalPath) &&
        !string.IsNullOrWhiteSpace(RemotePath) &&
        (Mode == DesktopMode.Compare || !string.IsNullOrWhiteSpace(BasePath));

    private async Task RunAsync()
    {
        await RunOperationAsync(async token =>
        {
            await DisposeSessionsAsync();
            Sheets.Clear();
            Conflicts.Clear();
            GridDocument = null;
            Inspector.Clear();
            Status = LocalizationService.Get("Loading");
            AddLog(Status);
            var progress = new Progress<ApplicationProgress>(HandleProgress);
            if (Mode == DesktopMode.Compare)
            {
                _compareSession = await _application.OpenCompareAsync(
                    new CompareRequest(LocalPath, RemotePath, Options: CreateComparisonOptions()),
                    progress,
                    token);
                foreach (var sheet in _compareSession.Sheets)
                {
                    Sheets.Add(SheetItemViewModel.FromCompare(sheet));
                }
            }
            else
            {
                _mergeSession = await _application.OpenMergeAsync(
                    new MergeRequest(BasePath, LocalPath, RemotePath, CreateComparisonOptions()),
                    progress,
                    token);
                foreach (var sheet in _mergeSession.Sheets)
                {
                    Sheets.Add(SheetItemViewModel.FromMerge(sheet));
                }
            }

            SelectedSheet = Sheets.FirstOrDefault();
            Status = LocalizationService.Get("Completed");
            AddLog(Status);
            await RecordRecentAsync();
            OnPropertyChanged(nameof(CanSave));
            OnPropertyChanged(nameof(ConflictSummary));
        });
    }

    private async Task<bool> RunOperationAsync(Func<CancellationToken, Task> operation)
    {
        if (IsBusy)
        {
            return false;
        }

        _operationCancellation?.Dispose();
        _operationCancellation = new CancellationTokenSource();
        IsBusy = true;
        ProgressIndeterminate = true;
        ProgressValue = 0;
        try
        {
            await operation(_operationCancellation.Token);
            return true;
        }
        catch (OperationCanceledException)
        {
            Status = LocalizationService.Get("Cancel");
            AddLog(Status);
            return false;
        }
        catch (Exception exception)
        {
            Status = $"{LocalizationService.Get("Failed")}: {DescribeException(exception)}";
            AddLog(Status);
            return false;
        }
        finally
        {
            IsBusy = false;
            ProgressIndeterminate = false;
            RunCommand.RaiseCanExecuteChanged();
            RaiseResolutionCanExecute();
        }
    }

    private void Cancel() => _operationCancellation?.Cancel();

    private void HandleProgress(ApplicationProgress progress)
    {
        ProgressIndeterminate = !progress.Total.HasValue || progress.Total.Value == 0;
        if (progress.Total is > 0)
        {
            ProgressValue = Math.Clamp(
                (double)progress.Completed / progress.Total.Value * 100,
                0,
                100);
        }

        Status = progress.Stage switch
        {
            ApplicationStage.Saving => LocalizationService.Get("Saving"),
            ApplicationStage.Completed => LocalizationService.Get("Completed"),
            _ => LocalizationService.Get("Loading"),
        };
    }

    private void RefreshSelectedSheet()
    {
        Conflicts.Clear();
        SelectedConflict = null;
        if (SelectedSheet is null)
        {
            GridDocument = null;
            return;
        }

        if (SelectedSheet.Compare is not null)
        {
            GridDocument = GridDocument.FromCompare(SelectedSheet.Compare, HideUnchanged);
        }
        else if (SelectedSheet.Merge is not null)
        {
            var resolutions = _mergeSession is null
                ? new Dictionary<long, ResolutionKind>()
                : BuildEffectiveResolutions(_mergeSession.BuildPlan());
            foreach (var item in ConflictsForSheet(SelectedSheet.Merge, resolutions))
            {
                Conflicts.Add(item);
            }

            GridDocument = GridDocument.FromMerge(SelectedSheet.Merge, resolutions, HideUnchanged);
            SelectedConflict = Conflicts.FirstOrDefault(static conflict => !conflict.IsResolved) ??
                Conflicts.FirstOrDefault();
        }

        OnPropertyChanged(nameof(ConflictSummary));
        OnPropertyChanged(nameof(CanSave));
    }

    private static IEnumerable<ConflictItemViewModel> ConflictsForSheet(
        MergeSheetResult sheet,
        IReadOnlyDictionary<long, ResolutionKind> resolutions)
    {
        foreach (var conflict in sheet.Merge.Conflicts.ToArray())
        {
            var kind = resolutions.TryGetValue(conflict.Id, out var resolution)
                ? resolution
                : ResolutionKind.Unresolved;
            yield return new ConflictItemViewModel(conflict, kind);
        }
    }

    private void ResolveSelected(ResolutionKind kind)
    {
        if (_mergeSession is null || SelectedConflict is null)
        {
            return;
        }

        try
        {
            var conflict = SelectedConflict.Conflict;
            if (kind == ResolutionKind.Both)
            {
                if (!SelectedConflict.SupportsBoth)
                {
                    return;
                }

                _mergeSession.SetRowOverride(new RowMergeOverride(
                    conflict.Location.SheetId,
                    conflict.Location.BaseRowIndex,
                    conflict.Location.LocalRowIndex,
                    conflict.Location.RemoteRowIndex,
                    ResolutionKind.Both));
            }
            else if (SelectedConflict.IsRowConflict)
            {
                _mergeSession.ResolveRow(conflict.Id, kind);
                RemoveRowOverride(_mergeSession, conflict.Location);
            }
            else
            {
                _mergeSession.ResolveCell(conflict.Id, kind);
                RemoveRowOverride(_mergeSession, conflict.Location);
            }

            RefreshAfterResolution();
        }
        catch (Exception exception)
        {
            Status = exception.Message;
            AddLog(exception.Message);
        }
    }

    private void ResolveCustom()
    {
        if (_mergeSession is null || SelectedConflict is null || SelectedConflict.IsRowConflict)
        {
            return;
        }

        _mergeSession.ResolveCell(
            SelectedConflict.Conflict.Id,
            ResolutionKind.Custom,
            CellValue.FromText(CustomValue, CustomValue));
        RemoveRowOverride(_mergeSession, SelectedConflict.Conflict.Location);
        RefreshAfterResolution();
    }

    private void RefreshAfterResolution()
    {
        if (_mergeSession is not null)
        {
            var resolutions = BuildEffectiveResolutions(_mergeSession.BuildPlan());
            foreach (var conflict in Conflicts)
            {
                conflict.Resolution = resolutions.GetValueOrDefault(
                    conflict.Conflict.Id,
                    ResolutionKind.Unresolved);
            }

            if (SelectedSheet?.Merge is { } merge)
            {
                GridDocument = GridDocument.FromMerge(merge, resolutions, HideUnchanged);
            }
        }

        OnPropertyChanged(nameof(ConflictSummary));
        OnPropertyChanged(nameof(CanSave));
        RaiseResolutionCanExecute();
    }

    private static void RemoveRowOverride(MergeSession session, ConflictLocation location) =>
        session.RemoveRowOverride(
            location.SheetId,
            location.BaseRowIndex,
            location.LocalRowIndex,
            location.RemoteRowIndex);

    private static Dictionary<long, ResolutionKind> BuildEffectiveResolutions(MergePlan plan)
    {
        var resolutions = plan.CellResolutions.ToArray()
            .ToDictionary(static item => item.ConflictId, static item => item.Kind);
        foreach (var resolution in plan.RowResolutions.Span)
        {
            resolutions[resolution.ConflictId] = resolution.Kind;
        }

        var rowOverrides = plan.RowOverrides.ToArray().ToDictionary(
            static item => (
                item.SheetId,
                item.BaseRowIndex,
                item.LocalRowIndex,
                item.RemoteRowIndex),
            static item => item.Kind);
        foreach (var conflict in plan.Conflicts.Span)
        {
            if (rowOverrides.TryGetValue((
                    conflict.Location.SheetId,
                    conflict.Location.BaseRowIndex,
                    conflict.Location.LocalRowIndex,
                    conflict.Location.RemoteRowIndex),
                out var kind))
            {
                resolutions[conflict.Id] = kind;
            }
        }

        return resolutions;
    }

    private async Task SearchAsync()
    {
        if (GridDocument is null)
        {
            return;
        }

        await RunOperationAsync(async token =>
        {
            _searchResults = await GridDocument.SearchAsync(
                SearchText,
                SearchExact,
                SearchCaseSensitive,
                SearchRegex,
                token);
            _searchResultIndex = _searchResults.Count == 0 ? -1 : 0;
            SearchSummary = _searchResults.Count.ToString(System.Globalization.CultureInfo.CurrentCulture);
            PreviousSearchCommand.RaiseCanExecuteChanged();
            NextSearchCommand.RaiseCanExecuteChanged();
            if (_searchResultIndex >= 0)
            {
                var result = _searchResults[_searchResultIndex];
                NavigationRequested?.Invoke(result.ViewRowIndex, result.ColumnIndex);
            }
        });
    }

    private void NavigateChange(int direction)
    {
        if (GridDocument is null)
        {
            return;
        }

        var row = GridDocument.FindNextChangedRow(_currentGridRow, direction);
        if (row >= 0)
        {
            _currentGridRow = row;
            NavigationRequested?.Invoke(row, 0);
        }
    }

    private void NavigateSearch(int direction)
    {
        if (_searchResults.Count == 0)
        {
            return;
        }

        _searchResultIndex = (_searchResultIndex + direction) % _searchResults.Count;
        if (_searchResultIndex < 0)
        {
            _searchResultIndex += _searchResults.Count;
        }

        var result = _searchResults[_searchResultIndex];
        NavigationRequested?.Invoke(result.ViewRowIndex, result.ColumnIndex);
    }

    private void NavigateToConflict(ConflictRecord conflict)
    {
        if (SelectedSheet?.Merge is null)
        {
            return;
        }

        var index = GridDocument?.FindViewRow(
            conflict.Location.BaseRowIndex,
            conflict.Location.LocalRowIndex,
            conflict.Location.RemoteRowIndex) ?? -1;
        if (index >= 0)
        {
            _currentGridRow = index;
            NavigationRequested?.Invoke(index, conflict.Location.ColumnIndex ?? 0);
        }
    }

    private async Task DisposeSessionsAsync()
    {
        if (_compareSession is not null)
        {
            await _compareSession.DisposeAsync();
            _compareSession = null;
        }

        if (_mergeSession is not null)
        {
            await _mergeSession.DisposeAsync();
            _mergeSession = null;
        }
    }

    private async Task RecordRecentAsync()
    {
        var descriptor = new RecentSessionDescriptor(
            Guid.NewGuid(),
            Mode == DesktopMode.Merge ? RecentSessionMode.Merge : RecentSessionMode.Compare,
            Mode == DesktopMode.Merge ? BasePath : null,
            LocalPath,
            RemotePath,
            DateTimeOffset.UtcNow,
            KeyColumns: _settings.KeyColumns);
        await _recentStore.RecordAsync(descriptor);
        await RefreshRecentAsync();
    }

    private async Task RefreshRecentAsync()
    {
        var recent = await _recentStore.ListAsync();
        RecentSessions.Clear();
        foreach (var descriptor in recent)
        {
            RecentSessions.Add(new RecentItemViewModel(descriptor));
        }
    }

    private void RaiseResolutionCanExecute()
    {
        UseLocalCommand.RaiseCanExecuteChanged();
        UseRemoteCommand.RaiseCanExecuteChanged();
        UseBothCommand.RaiseCanExecuteChanged();
        UseCustomCommand.RaiseCanExecuteChanged();
    }

    private void AddLog(string message)
    {
        Logs.Insert(0, $"{DateTime.Now:HH:mm:ss}  {message}");
        while (Logs.Count > 200)
        {
            Logs.RemoveAt(Logs.Count - 1);
        }
    }

    private void SwapSides()
    {
        (LocalPath, RemotePath) = (RemotePath, LocalPath);
    }

    private WorksheetComparisonOptions CreateComparisonOptions() => new()
    {
        KeyColumns = _settings.KeyColumns,
        CellValues = new CellComparisonOptions
        {
            CompareFormulaCachedValues = _settings.CompareFormulaCachedValues,
            CompareDisplayText = _settings.CompareDisplayText,
        },
        CompareCellStyles = _settings.CompareCellStyles,
        CompareRowMetadata = _settings.CompareRowMetadata,
        TreatExplicitBlankCellsAsMissing = _settings.TreatExplicitBlankCellsAsMissing,
    };

    private static string DescribeCell(CellRecord? cell) =>
        cell.HasValue ? GridDocument.FormatValue(cell.Value.Value) : string.Empty;

    private static string DescribeException(Exception exception) =>
        exception is ExcelMergeApplicationException
        {
            Error: ApplicationError.UnsupportedLegacyWorkbook,
        }
            ? LocalizationService.Get("LegacyWorkbookUnsupported")
            : exception.Message;

    private static string Escape(string value, char delimiter) =>
        value.IndexOfAny([delimiter, '"', '\r', '\n']) < 0
            ? value
            : '"' + value.Replace("\"", "\"\"", StringComparison.Ordinal) + '"';
}

public sealed class SheetItemViewModel
{
    private SheetItemViewModel(
        string id,
        string name,
        string summary,
        CompareSheetResult? compare,
        MergeSheetResult? merge)
    {
        Id = id;
        Name = name;
        Summary = summary;
        Compare = compare;
        Merge = merge;
    }

    public string Id { get; }

    public string Name { get; }

    public string Summary { get; }

    public CompareSheetResult? Compare { get; }

    public MergeSheetResult? Merge { get; }

    public static SheetItemViewModel FromCompare(CompareSheetResult result) =>
        new(
            result.Id,
            result.LocalWorksheet?.Metadata.Name ?? result.RemoteWorksheet?.Metadata.Name ?? result.Id,
            $"{result.Difference.Statistics.ChangedRowCount} / {result.Difference.Statistics.ChangedCellCount}",
            result,
            null);

    public static SheetItemViewModel FromMerge(MergeSheetResult result) =>
        new(
            result.Id,
            result.LocalWorksheet?.Metadata.Name ??
                result.RemoteWorksheet?.Metadata.Name ??
                result.BaseWorksheet?.Metadata.Name ?? result.Id,
            result.Merge.Conflicts.Length == 0
                ? LocalizationService.Get("Changes")
                : $"{result.Merge.Conflicts.Length} {LocalizationService.Get("Conflicts")}",
            null,
            result);
}

public sealed class ConflictItemViewModel : ObservableObject
{
    private ResolutionKind _resolution;

    public ConflictItemViewModel(ConflictRecord conflict, ResolutionKind resolution)
    {
        Conflict = conflict;
        _resolution = resolution;
    }

    public ConflictRecord Conflict { get; }

    public string Location
    {
        get
        {
            var row = Conflict.Location.LocalRowIndex ??
                Conflict.Location.RemoteRowIndex ??
                Conflict.Location.BaseRowIndex ?? 0;
            return Conflict.Location.ColumnIndex is { } column
                ? GridDocument.FormatAddress(row, column)
                : $"R{row + 1}";
        }
    }

    public string Kind => Conflict.Kind.ToString();

    public bool IsRowConflict => Conflict.Kind is not (ConflictKind.CellValue or ConflictKind.CellDeleteEdit);

    public bool SupportsBoth =>
        Conflict.Location.LocalRowIndex.HasValue && Conflict.Location.RemoteRowIndex.HasValue;

    public bool Matches(GridRowDescriptor row, int columnIndex) =>
        Conflict.Location.ColumnIndex == columnIndex &&
        row.MatchesLocation(
            Conflict.Location.BaseRowIndex,
            Conflict.Location.LocalRowIndex,
            Conflict.Location.RemoteRowIndex);

    public string BaseValue => FormatValue(Conflict.BaseValue);

    public string LocalValue => FormatValue(Conflict.LocalValue);

    public string RemoteValue => FormatValue(Conflict.RemoteValue);

    public ResolutionKind Resolution
    {
        get => _resolution;
        set
        {
            if (SetProperty(ref _resolution, value))
            {
                OnPropertyChanged(nameof(IsResolved));
                OnPropertyChanged(nameof(ResolutionText));
            }
        }
    }

    public bool IsResolved => Resolution != ResolutionKind.Unresolved;

    public string ResolutionText => IsResolved ? Resolution.ToString() : LocalizationService.Get("Unresolved");

    private static string FormatValue(CellValue? value) =>
        value.HasValue ? GridDocument.FormatValue(value.Value) : string.Empty;

}

public sealed class InspectorViewModel : ObservableObject
{
    private string _address = string.Empty;
    private string _baseValue = string.Empty;
    private string _localValue = string.Empty;
    private string _remoteValue = string.Empty;
    private string _resultValue = string.Empty;
    private string _formula = string.Empty;
    private string _valueType = string.Empty;

    public string Address { get => _address; set => SetProperty(ref _address, value); }
    public string BaseValue { get => _baseValue; set => SetProperty(ref _baseValue, value); }
    public string LocalValue { get => _localValue; set => SetProperty(ref _localValue, value); }
    public string RemoteValue { get => _remoteValue; set => SetProperty(ref _remoteValue, value); }
    public string ResultValue { get => _resultValue; set => SetProperty(ref _resultValue, value); }
    public string Formula { get => _formula; set => SetProperty(ref _formula, value); }
    public string ValueType { get => _valueType; set => SetProperty(ref _valueType, value); }

    public void Clear()
    {
        Address = BaseValue = LocalValue = RemoteValue = ResultValue = Formula = ValueType = string.Empty;
    }
}

public sealed record RecentItemViewModel(RecentSessionDescriptor Descriptor)
{
    public string DisplayName =>
        $"{Path.GetFileName(Descriptor.LocalPath)}  /  {Path.GetFileName(Descriptor.RemotePath)}";

    public string Timestamp => Descriptor.LastOpenedUtc.LocalDateTime.ToString("g");
}

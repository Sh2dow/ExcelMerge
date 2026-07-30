using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Selection;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Avalonia.VisualTree;
using ExcelMerge;

namespace ExcelMerge.Avalonia;

public partial class MainWindow : Window
{
    private const double DefaultCellFontSize = 11;
    private const double DefaultCellWidth = 120;
    private const double DefaultCellHeight = 28;
    private static readonly IBrush ReadySheetBackground = new SolidColorBrush(Color.Parse("#ECF8F0"));
    private static readonly IBrush ConflictSheetBackground = new SolidColorBrush(Color.Parse("#FFF0F2"));
    private static readonly IBrush ReadySheetForeground = new SolidColorBrush(Color.Parse("#176B38"));
    private static readonly IBrush ConflictSheetForeground = new SolidColorBrush(Color.Parse("#9B2638"));
    private static readonly IBrush SelectedSheetBorder = new SolidColorBrush(Color.Parse("#2563EB"));
    private static readonly IBrush ReadySheetBorder = new SolidColorBrush(Color.Parse("#B9DFC7"));
    private static readonly IBrush ConflictSheetBorder = new SolidColorBrush(Color.Parse("#E8C4CA"));
    private static readonly IBrush SheetNameForeground = new SolidColorBrush(Color.Parse("#253047"));
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xlsx", ".xls", ".csv", ".tsv",
    };

    private readonly MainViewModel _viewModel;
    private readonly CommandLineOptions _options;
    private ExcelWorkbook? _baseWorkbook;
    private ExcelWorkbook? _localWorkbook;
    private ExcelWorkbook? _remoteWorkbook;
    private ExcelSheetDiff? _diff;
    private List<DiffRow> _allRows = new();
    private List<DiffRow> _changedRows = new();
    private List<DiffRow> _conflictRows = new();
    private List<DiffRow> _visibleRows = new();
    private List<int> _visibleRowPositions = new();
    private readonly Dictionary<DiffRow, int> _rowPositions = new();
    private readonly Dictionary<int, DataGridColumn> _localColumnsByIndex = new();
    private readonly Dictionary<int, DataGridColumn> _remoteColumnsByIndex = new();
    private readonly Dictionary<MergeCellKey, MergeCellResolution> _mergeResolutions = new();
    private readonly Dictionary<MergeRowKey, MergeResolution> _rowResolutions = new();
    private IReadOnlyDictionary<string, IReadOnlyList<MergeCellKey>> _workbookConflicts =
        new Dictionary<string, IReadOnlyList<MergeCellKey>>(StringComparer.Ordinal);
    private readonly List<MergeCellKey> _workbookConflictTargets = new();
    private readonly Dictionary<MergeCellKey, int> _workbookConflictPositions = new();
    private readonly Dictionary<MergeCellKey, bool> _workbookConflictResolutionStates = new();
    private readonly Dictionary<string, int> _resolvedConflictCountsBySheet = new(StringComparer.Ordinal);
    private int _workbookResolvedConflictCount;
    private readonly List<ConflictTarget> _conflictTargets = new();
    private readonly Dictionary<DiffCell, int> _conflictPositions = new();
    private readonly Dictionary<MergeCellKey, ConflictTarget> _sheetConflictTargets = new();
    private readonly Dictionary<string, SheetConflictCard> _sheetConflictCards = new(StringComparer.Ordinal);
    private string? _cachedDiffLocalSheetName;
    private string? _cachedDiffRemoteSheetName;
    private ExcelSheetDiff? _cachedSheetDiff;
    private bool _workbookConflictsLoaded;
    private int _changeIndex = -1;
    private int _conflictIndex = -1;
    private DiffRow? _currentConflictRow;
    private DiffCell? _currentConflictCell;
    private DiffCell? _dialogConflictCell;
    private bool _isRowConflictDialog;
    private DiffCell? _selectedRowCustomCell;
    private readonly Dictionary<DiffCell, string> _rowCustomDrafts = new();
    private readonly Dictionary<DiffCell, SelectableTextBlock> _rowResultValueBoxes = new();
    private bool _updatingRowCustomEditor;
    private string? _loadedBasePath;
    private string? _loadedLocalPath;
    private string? _loadedRemotePath;
    private string? _currentMergeSheetName;
    private readonly SplitScrollSynchronizer _scrollSynchronizer;
    private ColumnWidthSynchronizer? _columnWidthSynchronizer;
    private bool _synchronizingSelection;
    private bool _updatingSheetSelections;
    private double _cellFontSize = DefaultCellFontSize;
    private bool _overviewUpdatePending;
    private bool _isBusy;
    private bool _isClosed;
    private CancellationTokenSource? _operationCancellation;

    private enum WorkbookInput
    {
        Base,
        Local,
        Remote,
    }

    private sealed record SheetChoice(string Name, bool Exists, string Side)
    {
        public override string ToString() => Exists ? Name : $"{Name} (missing in {Side})";
    }

    private readonly record struct ConflictTarget(DiffRow Row, DiffCell Cell);

    private sealed record WorkbookConflictScan(
        IReadOnlyDictionary<string, IReadOnlyList<MergeCellKey>> Conflicts,
        string? CachedSheetName,
        ExcelSheetDiff? CachedDiff);

    private sealed record SheetConflictCard(Button Button, TextBlock Status);

    public MainWindow() : this(Array.Empty<string>())
    {
    }

    public MainWindow(string[] args)
    {
        InitializeComponent();

        _options = CommandLineOptions.Parse(args);

        _viewModel = new MainViewModel(CompareFiles, _options.Mode)
        {
            BasePath = _options.BasePath ?? string.Empty,
            LocalPath = _options.LocalPath ?? string.Empty,
            RemotePath = _options.RemotePath ?? string.Empty,
        };
        DataContext = _viewModel;
        SheetOverviewTitle.Text = _viewModel.IsMergeMode ? "SHEET STATUS" : "SHEET PAIR";
        MergeSheetCardPanel.IsVisible = _viewModel.IsMergeMode;
        DiffSheetCardPanel.IsVisible = !_viewModel.IsMergeMode;
        _scrollSynchronizer = new SplitScrollSynchronizer(LocalGrid, RemoteGrid);
        _scrollSynchronizer.ViewportChanged += ScrollViewportChanged;
        Title = _options.Mode == ApplicationMode.Merge ? "ExcelMerge - Merge" : "ExcelMerge - Diff";
        if (_options.IsMergeDriver)
        {
            SaveResultButton.Content = "Complete Git merge";
            RepositoryPathText.Text = _options.RepositoryPath ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(_options.RepositoryPath))
                Title += $" - {_options.RepositoryPath}";
        }

        Closed += (_, _) =>
        {
            _isClosed = true;
            _operationCancellation?.Cancel();
            _scrollSynchronizer.ViewportChanged -= ScrollViewportChanged;
            _scrollSynchronizer.Dispose();
            _columnWidthSynchronizer?.Dispose();
        };

        if (CanAutoCompare())
        {
            Opened += async (_, _) => await CompareFiles();
        }
    }

    private async void BrowseBase(object? sender, RoutedEventArgs e) => await PickFile(WorkbookInput.Base);
    private async void BrowseLocal(object? sender, RoutedEventArgs e) => await PickFile(WorkbookInput.Local);
    private async void BrowseRemote(object? sender, RoutedEventArgs e) => await PickFile(WorkbookInput.Remote);

    private void ToggleFileDrawer(object? sender, RoutedEventArgs e)
    {
        var expanded = !TableConfigPanel.IsVisible;
        TableConfigPanel.IsVisible = expanded;
        FileDrawerToggle.Content = expanded ? "Hide files ▲" : "Show files ▼";
        ToolTip.SetTip(FileDrawerToggle, expanded
            ? "Hide BASE, LOCAL, and REMOTE file inputs"
            : "Show BASE, LOCAL, and REMOTE file inputs");
    }

    private void CellFontSizeChanged(object? sender, RangeBaseValueChangedEventArgs e)
    {
        if (LocalGrid == null || RemoteGrid == null || CellFontSizeText == null)
            return;
        var size = Math.Round(e.NewValue);
        _cellFontSize = size;
        Resources["TableCellFontSize"] = size;
        LocalGrid.FontSize = size;
        RemoteGrid.FontSize = size;
        CellFontSizeText.Text = $"{size:0} px";
        ApplyDefaultCellMetrics();
    }

    private async Task PickFile(WorkbookInput input)
    {
        var picked = false;
        var succeeded = await RunOperation(async cancellationToken =>
        {
            var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                AllowMultiple = false,
                FileTypeFilter = new[]
                {
                    new FilePickerFileType("Spreadsheet files")
                    {
                        Patterns = new[] { "*.xlsx", "*.xls", "*.csv", "*.tsv" },
                    },
                },
            });
            if (files.Count == 0)
                return;

            var path = files[0].TryGetLocalPath() ?? files[0].Path.LocalPath;
            await LoadAndAssignFiles(new[] { (input, path) }, cancellationToken);
            picked = true;
        });

        if (succeeded && picked)
            await CompareIfReady();
    }

    private void FileDragOver(object? sender, DragEventArgs e)
    {
        e.DragEffects = GetDroppedFilePaths(e).Count > 0
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    private async void DropBase(object? sender, DragEventArgs e)
    {
        e.Handled = true;
        await DropInto(WorkbookInput.Base, e);
    }

    private async void DropLocal(object? sender, DragEventArgs e)
    {
        e.Handled = true;
        await DropInto(WorkbookInput.Local, e);
    }

    private async void DropRemote(object? sender, DragEventArgs e)
    {
        e.Handled = true;
        await DropInto(WorkbookInput.Remote, e);
    }

    private async Task DropInto(WorkbookInput input, DragEventArgs e)
    {
        var paths = GetDroppedFilePaths(e);
        if (paths.Count == 0)
        {
            Error.Text = "Drop an .xlsx, .xls, .csv, or .tsv file.";
            return;
        }

        if (paths.Count >= 2)
        {
            await OpenDroppedFiles(new[]
            {
                (WorkbookInput.Local, paths[0]),
                (WorkbookInput.Remote, paths[1]),
            });
        }
        else
        {
            await OpenDroppedFiles(new[] { (input, paths[0]) });
        }
    }

    private async void WindowDrop(object? sender, DragEventArgs e)
    {
        e.Handled = true;
        var paths = GetDroppedFilePaths(e);
        if (paths.Count == 0)
        {
            Error.Text = "Drop an .xlsx, .xls, .csv, or .tsv file.";
            return;
        }

        if (paths.Count >= 2)
        {
            await OpenDroppedFiles(new[]
            {
                (WorkbookInput.Local, paths[0]),
                (WorkbookInput.Remote, paths[1]),
            });
            return;
        }

        if (string.IsNullOrWhiteSpace(_viewModel.LocalPath))
            await OpenDroppedFiles(new[] { (WorkbookInput.Local, paths[0]) });
        else if (string.IsNullOrWhiteSpace(_viewModel.RemotePath))
            await OpenDroppedFiles(new[] { (WorkbookInput.Remote, paths[0]) });
        else
            Error.Text = "Drop one file onto LOCAL or REMOTE, or drop two files together.";
    }

    private async Task OpenDroppedFiles(IReadOnlyList<(WorkbookInput Input, string Path)> files)
    {
        var succeeded = await RunOperation(cancellationToken => LoadAndAssignFiles(files, cancellationToken));

        if (succeeded)
            await CompareIfReady();
    }

    private async Task LoadAndAssignFiles(
        IReadOnlyList<(WorkbookInput Input, string Path)> files,
        CancellationToken cancellationToken)
    {
        var loaded = new List<(WorkbookInput Input, string Path, ExcelWorkbook Workbook)>();
        foreach (var file in files)
        {
            EnsureSupportedFile(file.Path);
            var workbook = await LoadWorkbook(file.Input, file.Path, cancellationToken);
            loaded.Add((file.Input, file.Path, workbook));
        }

        if (loaded.Any(file => !string.Equals(GetLoadedPath(file.Input), file.Path, StringComparison.Ordinal)))
            ClearDiff();
        foreach (var file in loaded)
            CommitWorkbook(file.Input, file.Path, file.Workbook);
    }

    private string? GetLoadedPath(WorkbookInput input)
    {
        return input switch
        {
            WorkbookInput.Base => _loadedBasePath,
            WorkbookInput.Local => _loadedLocalPath,
            WorkbookInput.Remote => _loadedRemotePath,
            _ => throw new ArgumentOutOfRangeException(nameof(input)),
        };
    }

    private static IReadOnlyList<string> GetDroppedFilePaths(DragEventArgs e)
    {
        var paths = e.DataTransfer.TryGetFiles()?
            .OfType<IStorageFile>()
            .Select(file => file.TryGetLocalPath() ?? file.Path.LocalPath)
            .Where(path => SupportedExtensions.Contains(Path.GetExtension(path)))
            .Distinct(StringComparer.Ordinal)
            .ToList();
        if (paths == null)
            return Array.Empty<string>();
        return paths;
    }

    private static void EnsureSupportedFile(string path)
    {
        if (!SupportedExtensions.Contains(Path.GetExtension(path)))
            throw new NotSupportedException($"Unsupported file type: {Path.GetExtension(path)}");
    }

    private void AssignPath(WorkbookInput input, string path)
    {
        switch (input)
        {
            case WorkbookInput.Base:
                _viewModel.BasePath = path;
                BasePath.Text = path;
                break;
            case WorkbookInput.Local:
                _viewModel.LocalPath = path;
                LocalPath.Text = path;
                break;
            case WorkbookInput.Remote:
                _viewModel.RemotePath = path;
                RemotePath.Text = path;
                break;
        }
    }

    private async Task<bool> CompareIfReady()
    {
        return CanAutoCompare() && await CompareFiles();
    }

    private bool CanAutoCompare()
    {
        return File.Exists(_viewModel.LocalPath)
            && File.Exists(_viewModel.RemotePath)
            && (!_viewModel.IsMergeMode || File.Exists(_viewModel.BasePath));
    }

    private async Task<bool> CompareFiles()
    {
        return await RunOperation(CompareCurrentFiles);
    }

    private async Task CompareCurrentFiles(CancellationToken cancellationToken)
    {
        _viewModel.BasePath = BasePath.Text?.Trim() ?? string.Empty;
        _viewModel.LocalPath = LocalPath.Text?.Trim() ?? string.Empty;
        _viewModel.RemotePath = RemotePath.Text?.Trim() ?? string.Empty;
        var inputsChanged = !string.Equals(_loadedLocalPath, _viewModel.LocalPath, StringComparison.Ordinal)
            || !string.Equals(_loadedRemotePath, _viewModel.RemotePath, StringComparison.Ordinal)
            || _viewModel.IsMergeMode && !string.Equals(_loadedBasePath, _viewModel.BasePath, StringComparison.Ordinal);
        if (inputsChanged)
            ClearDiff();
        Progress.Text = "Loading...";

        if (!File.Exists(_viewModel.LocalPath) || !File.Exists(_viewModel.RemotePath))
            throw new FileNotFoundException("Both LOCAL and REMOTE files must exist.");

        if (_viewModel.IsMergeMode)
        {
            if (!File.Exists(_viewModel.BasePath))
                throw new FileNotFoundException("The BASE file must exist in merge mode.");
            await EnsureWorkbookLoaded(WorkbookInput.Base, cancellationToken);
        }

        await EnsureWorkbookLoaded(WorkbookInput.Local, cancellationToken);
        await EnsureWorkbookLoaded(WorkbookInput.Remote, cancellationToken);
        var localWorkbook = _localWorkbook ?? throw new InvalidOperationException("The LOCAL workbook could not be loaded.");
        var remoteWorkbook = _remoteWorkbook ?? throw new InvalidOperationException("The REMOTE workbook could not be loaded.");
        var localChoice = LocalSheet.SelectedItem as SheetChoice
            ?? throw new InvalidOperationException("Select a LOCAL sheet.");
        var remoteChoice = RemoteSheet.SelectedItem as SheetChoice
            ?? throw new InvalidOperationException("Select a REMOTE sheet.");
        var localName = localChoice.Name;
        var remoteName = remoteChoice.Name;
        if (_viewModel.IsMergeMode
            && (_baseWorkbook?.Sheets.Count > 1 || localWorkbook.Sheets.Count > 1 || remoteWorkbook.Sheets.Count > 1)
            && !string.Equals(localName, remoteName, StringComparison.Ordinal))
            throw new WorkbookMergeException("Select the same sheet name in LOCAL and REMOTE before merging.");
        if (_viewModel.IsMergeMode && !_workbookConflictsLoaded)
        {
            var baseWorkbook = _baseWorkbook
                ?? throw new InvalidOperationException("The BASE workbook could not be loaded.");
            Progress.Text = "Scanning sheet conflicts...";
            var scan = await Task.Run(
                () => ScanWorkbookConflicts(
                    baseWorkbook,
                    localWorkbook,
                    remoteWorkbook,
                    string.Equals(localName, remoteName, StringComparison.Ordinal) ? localName : null),
                cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            _workbookConflicts = scan.Conflicts;
            _cachedDiffLocalSheetName = scan.CachedSheetName;
            _cachedDiffRemoteSheetName = scan.CachedSheetName;
            _cachedSheetDiff = scan.CachedDiff;
            RebuildWorkbookConflictTargets();
            _workbookConflictsLoaded = true;
        }
        var allowSoleSheetFallback = _baseWorkbook?.Sheets.Count == 1
            && localWorkbook.Sheets.Count == 1
            && remoteWorkbook.Sheets.Count == 1;
        var baseValues = _viewModel.IsMergeMode
            ? CreateBaseValueMap(_baseWorkbook, localName, remoteName, allowSoleSheetFallback)
            : null;
        var localSheet = localChoice.Exists ? localWorkbook.Sheets[localName] : new ExcelSheet();
        var remoteSheet = remoteChoice.Exists ? remoteWorkbook.Sheets[remoteName] : new ExcelSheet();

        ExcelSheetDiff diff;
        if (_cachedSheetDiff != null
            && string.Equals(_cachedDiffLocalSheetName, localName, StringComparison.Ordinal)
            && string.Equals(_cachedDiffRemoteSheetName, remoteName, StringComparison.Ordinal))
        {
            diff = _cachedSheetDiff;
        }
        else
        {
            diff = await Task.Run(
                () => ExcelSheet.Diff(localSheet, remoteSheet, new ExcelSheetDiffConfig()),
                cancellationToken);
            _cachedDiffLocalSheetName = localName;
            _cachedDiffRemoteSheetName = remoteName;
            _cachedSheetDiff = diff;
        }
        cancellationToken.ThrowIfCancellationRequested();
        _currentMergeSheetName = localName;
        _diff = diff;
        BuildRows(baseValues);
        UpdateSheetOverview();
        var prefix = _viewModel.IsMergeMode ? "Merge preview · " : string.Empty;
        var missingSheet = !localChoice.Exists
            ? "LOCAL sheet missing · "
            : !remoteChoice.Exists
                ? "REMOTE sheet missing · "
                : string.Empty;
        Summary.Text = prefix + missingSheet + FormatSummary(_diff.CreateSummary());
        Progress.Text = $"{_allRows.Count} rows";
        if (_options.IsMergeDriver)
        {
            var hasConflicts = _workbookConflicts.Values.Any(conflicts => conflicts.Count > 0);
            SaveResultButton.Content = hasConflicts ? "Complete Git merge" : "Auto Merge";
            if (!hasConflicts)
                Progress.Text = "Automatic merge ready for review";
        }
    }

    private void ClearDiff()
    {
        _diff = null;
        _allRows.Clear();
        _changedRows.Clear();
        _conflictRows.Clear();
        _visibleRows.Clear();
        _visibleRowPositions.Clear();
        _rowPositions.Clear();
        _localColumnsByIndex.Clear();
        _remoteColumnsByIndex.Clear();
        _mergeResolutions.Clear();
        _rowResolutions.Clear();
        _workbookConflicts = new Dictionary<string, IReadOnlyList<MergeCellKey>>(StringComparer.Ordinal);
        _workbookConflictTargets.Clear();
        _workbookConflictPositions.Clear();
        _workbookConflictResolutionStates.Clear();
        _resolvedConflictCountsBySheet.Clear();
        _workbookResolvedConflictCount = 0;
        _conflictTargets.Clear();
        _conflictPositions.Clear();
        _sheetConflictTargets.Clear();
        _sheetConflictCards.Clear();
        _cachedDiffLocalSheetName = null;
        _cachedDiffRemoteSheetName = null;
        _cachedSheetDiff = null;
        _workbookConflictsLoaded = false;
        _currentMergeSheetName = null;
        _columnWidthSynchronizer?.Dispose();
        _columnWidthSynchronizer = null;
        LocalGrid.Columns.Clear();
        RemoteGrid.Columns.Clear();
        LocalGrid.ItemsSource = null;
        RemoteGrid.ItemsSource = null;
        SheetOverviewMap.SetRows(Array.Empty<DiffRow>());
        SheetOverviewMap.SetViewport(new Rect(0, 0, 1, 1));
        Summary.Text = "No comparison loaded.";
        _changeIndex = -1;
        _conflictIndex = -1;
        _currentConflictRow = null;
        _currentConflictCell = null;
        _dialogConflictCell = null;
        ConflictDialogOverlay.IsVisible = false;
        SheetConflictItems.Children.Clear();
        LocalSheetItems.Children.Clear();
        RemoteSheetItems.Children.Clear();
        WorkbookConflictSummary.Text = string.Empty;
        SaveResultButton.IsEnabled = false;
        UpdateConflictPanel();
    }

    private static IReadOnlyDictionary<(int Row, int Column), string> CreateBaseValueMap(
        ExcelWorkbook? workbook,
        string localSheetName,
        string remoteSheetName,
        bool allowSoleSheetFallback = true)
    {
        var values = new Dictionary<(int Row, int Column), string>();
        if (workbook == null)
            return values;

        if (!workbook.Sheets.TryGetValue(localSheetName, out var sheet)
            && !workbook.Sheets.TryGetValue(remoteSheetName, out sheet))
        {
            if (!allowSoleSheetFallback || workbook.Sheets.Count != 1)
                return values;
            sheet = workbook.Sheets.Values.Single();
        }

        foreach (var row in sheet.Rows.Values)
        {
            foreach (var cell in row.Cells)
                values[(cell.OriginalRowIndex, cell.OriginalColumnIndex)] = cell.Value;
        }

        return values;
    }

    internal static IReadOnlyDictionary<string, IReadOnlyList<MergeCellKey>> FindWorkbookConflicts(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook)
    {
        return ScanWorkbookConflicts(baseWorkbook, localWorkbook, remoteWorkbook, null).Conflicts;
    }

    private static WorkbookConflictScan ScanWorkbookConflicts(
        ExcelWorkbook baseWorkbook,
        ExcelWorkbook localWorkbook,
        ExcelWorkbook remoteWorkbook,
        string? cachedSheetName)
    {
        var names = localWorkbook.Sheets.Keys
            .Concat(remoteWorkbook.Sheets.Keys)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        var result = new Dictionary<string, IReadOnlyList<MergeCellKey>>(StringComparer.Ordinal);
        ExcelSheetDiff? cachedDiff = null;
        foreach (var name in names)
        {
            var localSheet = localWorkbook.Sheets.TryGetValue(name, out var local)
                ? local
                : new ExcelSheet();
            var remoteSheet = remoteWorkbook.Sheets.TryGetValue(name, out var remote)
                ? remote
                : new ExcelSheet();
            var baseValues = CreateBaseValueMap(baseWorkbook, name, name, allowSoleSheetFallback: false);
            var diff = ExcelSheet.Diff(localSheet, remoteSheet, new ExcelSheetDiffConfig());
            if (string.Equals(name, cachedSheetName, StringComparison.Ordinal))
                cachedDiff = diff;

            var conflicts = new List<MergeCellKey>();
            foreach (var cell in diff.Rows.Values.SelectMany(row => row.Cells.Values))
            {
                var source = cell.Status == ExcelCellStatus.Added ? cell.DstCell : cell.SrcCell;
                baseValues.TryGetValue((source.OriginalRowIndex, source.OriginalColumnIndex), out var baseValue);
                baseValue ??= string.Empty;
                if (!string.Equals(cell.SrcCell.Value, cell.DstCell.Value, StringComparison.Ordinal)
                    && !string.Equals(cell.SrcCell.Value, baseValue, StringComparison.Ordinal)
                    && !string.Equals(cell.DstCell.Value, baseValue, StringComparison.Ordinal))
                {
                    conflicts.Add(new MergeCellKey(
                        name,
                        source.OriginalRowIndex,
                        source.OriginalColumnIndex));
                }
            }
            result[name] = conflicts;
        }
        return new WorkbookConflictScan(result, cachedSheetName, cachedDiff);
    }

    private void RebuildWorkbookConflictTargets()
    {
        _workbookConflictTargets.Clear();
        _workbookConflictPositions.Clear();
        _workbookConflictResolutionStates.Clear();
        _resolvedConflictCountsBySheet.Clear();
        _workbookResolvedConflictCount = 0;
        foreach (var key in _workbookConflicts.Values.SelectMany(conflicts => conflicts))
        {
            if (!_workbookConflictPositions.ContainsKey(key))
                _workbookConflictPositions[key] = _workbookConflictTargets.Count;
            _workbookConflictTargets.Add(key);
            var resolved = IsConflictResolved(key);
            _workbookConflictResolutionStates[key] = resolved;
            if (resolved)
            {
                _workbookResolvedConflictCount++;
                _resolvedConflictCountsBySheet[key.SheetName] =
                    _resolvedConflictCountsBySheet.GetValueOrDefault(key.SheetName) + 1;
            }
        }
    }

    private void RefreshConflictResolutionProgress(DiffRow row)
    {
        if (_currentMergeSheetName == null)
            return;
        foreach (var cell in row.ConflictCells)
        {
            var key = new MergeCellKey(
                _currentMergeSheetName,
                cell.OriginalRowIndex,
                cell.OriginalColumnIndex);
            if (!_workbookConflictResolutionStates.TryGetValue(key, out var wasResolved))
                continue;
            var isResolved = IsConflictResolved(key);
            if (wasResolved == isResolved)
                continue;
            _workbookConflictResolutionStates[key] = isResolved;
            var change = isResolved ? 1 : -1;
            _workbookResolvedConflictCount += change;
            _resolvedConflictCountsBySheet[key.SheetName] =
                _resolvedConflictCountsBySheet.GetValueOrDefault(key.SheetName) + change;
        }
    }

    private async Task EnsureWorkbookLoaded(WorkbookInput input, CancellationToken cancellationToken)
    {
        var path = input switch
        {
            WorkbookInput.Base => _viewModel.BasePath,
            WorkbookInput.Local => _viewModel.LocalPath,
            WorkbookInput.Remote => _viewModel.RemotePath,
            _ => throw new ArgumentOutOfRangeException(nameof(input)),
        };
        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}", path);

        var workbook = await LoadWorkbook(input, path, cancellationToken);
        CommitWorkbook(input, path, workbook);
    }

    private async Task<ExcelWorkbook> LoadWorkbook(
        WorkbookInput input,
        string path,
        CancellationToken cancellationToken)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"File not found: {path}", path);

        var cachedWorkbook = input switch
        {
            WorkbookInput.Base => _baseWorkbook,
            WorkbookInput.Local => _localWorkbook,
            WorkbookInput.Remote => _remoteWorkbook,
            _ => throw new ArgumentOutOfRangeException(nameof(input)),
        };
        var cachedPath = input switch
        {
            WorkbookInput.Base => _loadedBasePath,
            WorkbookInput.Local => _loadedLocalPath,
            WorkbookInput.Remote => _loadedRemotePath,
            _ => throw new ArgumentOutOfRangeException(nameof(input)),
        };
        if (cachedWorkbook != null && string.Equals(cachedPath, path, StringComparison.Ordinal))
            return cachedWorkbook;

        var workbook = await Task.Run(() => ExcelWorkbook.Create(path, new ExcelSheetReadConfig()), cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return workbook;
    }

    private void CommitWorkbook(WorkbookInput input, string path, ExcelWorkbook workbook)
    {
        AssignPath(input, path);

        if (input == WorkbookInput.Base)
        {
            _baseWorkbook = workbook;
            _loadedBasePath = path;
            return;
        }

        if (input == WorkbookInput.Local)
        {
            _localWorkbook = workbook;
            _loadedLocalPath = path;
        }
        else
        {
            _remoteWorkbook = workbook;
            _loadedRemotePath = path;
        }
        RebuildSheetChoices();
    }

    private void RebuildSheetChoices()
    {
        var previousLocal = (LocalSheet.SelectedItem as SheetChoice)?.Name;
        var previousRemote = (RemoteSheet.SelectedItem as SheetChoice)?.Name;
        if (_localWorkbook?.Sheets.Count == 1 && _remoteWorkbook?.Sheets.Count == 1)
        {
            var singleLocalChoices = _localWorkbook.Sheets.Keys
                .Select(name => new SheetChoice(name, true, "LOCAL"))
                .ToList();
            var singleRemoteChoices = _remoteWorkbook.Sheets.Keys
                .Select(name => new SheetChoice(name, true, "REMOTE"))
                .ToList();
            _updatingSheetSelections = true;
            LocalSheet.ItemsSource = singleLocalChoices;
            RemoteSheet.ItemsSource = singleRemoteChoices;
            LocalSheet.SelectedItem = singleLocalChoices[0];
            RemoteSheet.SelectedItem = singleRemoteChoices[0];
            _updatingSheetSelections = false;
            UpdateSheetOverview();
            return;
        }

        var names = (_localWorkbook?.Sheets.Keys ?? Enumerable.Empty<string>())
            .Concat(_remoteWorkbook?.Sheets.Keys ?? Enumerable.Empty<string>())
            .Distinct(StringComparer.Ordinal)
            .ToList();
        var localChoices = names.Select(name => new SheetChoice(
            name,
            _localWorkbook?.Sheets.ContainsKey(name) == true,
            "LOCAL")).ToList();
        var remoteChoices = names.Select(name => new SheetChoice(
            name,
            _remoteWorkbook?.Sheets.ContainsKey(name) == true,
            "REMOTE")).ToList();
        var selectedName = previousLocal ?? previousRemote
            ?? localChoices.FirstOrDefault(choice => choice.Exists)?.Name
            ?? remoteChoices.FirstOrDefault(choice => choice.Exists)?.Name;
        var selectedLocalName = _viewModel.IsMergeMode
            ? selectedName
            : localChoices.FirstOrDefault(choice => choice.Name == previousLocal && choice.Exists)?.Name
                ?? localChoices.FirstOrDefault(choice => choice.Exists)?.Name;
        var selectedRemoteName = _viewModel.IsMergeMode
            ? selectedName
            : remoteChoices.FirstOrDefault(choice => choice.Name == previousRemote && choice.Exists)?.Name
                ?? remoteChoices.FirstOrDefault(choice => choice.Exists)?.Name;

        _updatingSheetSelections = true;
        LocalSheet.ItemsSource = localChoices;
        RemoteSheet.ItemsSource = remoteChoices;
        LocalSheet.SelectedItem = localChoices.FirstOrDefault(choice => choice.Name == selectedLocalName);
        RemoteSheet.SelectedItem = remoteChoices.FirstOrDefault(choice => choice.Name == selectedRemoteName);
        _updatingSheetSelections = false;
        UpdateSheetOverview();
    }

    private async void SheetSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_updatingSheetSelections || sender is not ComboBox source || source.SelectedItem is not SheetChoice selected)
            return;

        if (_viewModel.IsMergeMode)
        {
            var target = ReferenceEquals(source, LocalSheet) ? RemoteSheet : LocalSheet;
            var matching = target.ItemsSource?.Cast<SheetChoice>().FirstOrDefault(choice => choice.Name == selected.Name);
            if (matching != null)
            {
                _updatingSheetSelections = true;
                target.SelectedItem = matching;
                _updatingSheetSelections = false;
            }
        }
        UpdateSheetOverview();
        await CompareIfReady();
    }

    private async Task<bool> SelectSheet(string name)
    {
        var local = LocalSheet.ItemsSource?.Cast<SheetChoice>().FirstOrDefault(choice => choice.Name == name);
        var remote = RemoteSheet.ItemsSource?.Cast<SheetChoice>().FirstOrDefault(choice => choice.Name == name);
        if (local == null || remote == null)
            return false;

        _updatingSheetSelections = true;
        LocalSheet.SelectedItem = local;
        RemoteSheet.SelectedItem = remote;
        _updatingSheetSelections = false;
        return await CompareIfReady()
            && string.Equals(_currentMergeSheetName, name, StringComparison.Ordinal);
    }

    private async Task SelectDiffSheet(ComboBox selector, string name)
    {
        var choice = selector.ItemsSource?.Cast<SheetChoice>()
            .FirstOrDefault(candidate => candidate.Name == name && candidate.Exists);
        if (choice == null)
            return;

        _updatingSheetSelections = true;
        selector.SelectedItem = choice;
        _updatingSheetSelections = false;
        UpdateSheetOverview();
        await CompareIfReady();
    }

    private async Task<bool> RunOperation(Func<CancellationToken, Task> operation)
    {
        if (_isBusy)
            return false;

        _isBusy = true;
        SetInputEnabled(false);
        Error.Text = string.Empty;
        using var cancellation = new CancellationTokenSource();
        _operationCancellation = cancellation;
        try
        {
            await operation(cancellation.Token);
            return true;
        }
        catch (OperationCanceledException) when (_isClosed)
        {
            return false;
        }
        catch (Exception ex)
        {
            Error.Text = ex.Message;
            Progress.Text = string.Empty;
            if (_options.IsMergeDriver)
                Console.Error.WriteLine(ex.Message);
            return false;
        }
        finally
        {
            if (ReferenceEquals(_operationCancellation, cancellation))
                _operationCancellation = null;
            _isBusy = false;
            if (!_isClosed)
                SetInputEnabled(true);
        }
    }

    private void SetInputEnabled(bool enabled)
    {
        FileInputPanel.IsEnabled = enabled;
        SheetConflictOverview.IsEnabled = enabled;
        MergeConflictPanel.IsEnabled = enabled;
    }

    private void BuildRows(IReadOnlyDictionary<(int Row, int Column), string>? baseValues)
    {
        _currentConflictRow = null;
        _currentConflictCell = null;
        _conflictIndex = -1;
        _allRows = _diff!.Rows.Values.Select(row => new DiffRow(row, baseValues)).ToList();
        _changedRows = _allRows.Where(row => row.IsChanged).ToList();
        _conflictRows = _allRows.Where(row => row.HasConflict).ToList();
        _rowPositions.Clear();
        for (var index = 0; index < _allRows.Count; index++)
            _rowPositions[_allRows[index]] = index;
        var defaultRowHeight = CellHeight(_cellFontSize);
        foreach (var row in _allRows)
            row.SetDefaultRowHeight(defaultRowHeight);
        if (_currentMergeSheetName != null)
        {
            foreach (var row in _allRows.Where(row => row.HasConflict))
            {
                if (_rowResolutions.TryGetValue(RowKey(row), out var rowResolution))
                {
                    row.Resolve(rowResolution);
                    continue;
                }
                foreach (var cell in row.ConflictCells)
                {
                    if (_mergeResolutions.TryGetValue(CellKey(row, cell), out var resolution))
                        cell.Resolve(resolution.Resolution, resolution.CustomValue);
                }
            }
        }
        _conflictTargets.Clear();
        _conflictPositions.Clear();
        _sheetConflictTargets.Clear();
        foreach (var row in _conflictRows)
        {
            foreach (var cell in row.ConflictCells)
            {
                var target = new ConflictTarget(row, cell);
                _conflictPositions[cell] = _conflictTargets.Count;
                _conflictTargets.Add(target);
                if (_currentMergeSheetName != null)
                {
                    _sheetConflictTargets[new MergeCellKey(
                        _currentMergeSheetName,
                        cell.OriginalRowIndex,
                        cell.OriginalColumnIndex)] = target;
                }
            }
        }
        var columns = _allRows.SelectMany(row => row.Cells).Select(cell => cell.ColumnIndex).Distinct().OrderBy(index => index).ToList();
        _columnWidthSynchronizer?.Dispose();
        _columnWidthSynchronizer = null;
        LocalGrid.Columns.Clear();
        RemoteGrid.Columns.Clear();
        _localColumnsByIndex.Clear();
        _remoteColumnsByIndex.Clear();
        LocalGrid.Columns.Add(RowNumberColumn());
        RemoteGrid.Columns.Add(RowNumberColumn());
        foreach (var column in columns)
        {
            var localColumn = CellColumn(column, false);
            var remoteColumn = CellColumn(column, true);
            _localColumnsByIndex[column] = localColumn;
            _remoteColumnsByIndex[column] = remoteColumn;
            LocalGrid.Columns.Add(localColumn);
            RemoteGrid.Columns.Add(remoteColumn);
        }
        _columnWidthSynchronizer = new ColumnWidthSynchronizer(LocalGrid, RemoteGrid);
        ApplyFilter();
        UpdateConflictPanel();
    }

    private DataGridTemplateColumn RowNumberColumn() => new()
    {
        Header = "#",
        Width = new DataGridLength(55),
        IsReadOnly = true,
        CellTemplate = new FuncDataTemplate<DiffRow>((row, _) => new RowResizeCell(row, ResizeRow, OpenRowConflictDialog)),
    };

    private DataGridTemplateColumn CellColumn(int index, bool remote)
    {
        return new DataGridTemplateColumn
        {
            Header = ColumnName(index),
            Width = new DataGridLength(CellWidth(_cellFontSize)),
            CellTemplate = new FuncDataTemplate<DiffRow>((row, _) =>
            {
                var cell = row.GetCell(index);
                var value = new TextBlock();
                if (cell != null)
                {
                    value.Bind(
                        TextBlock.TextProperty,
                        new Binding(remote ? nameof(DiffCell.RemoteDisplayValue) : nameof(DiffCell.LocalDisplayValue))
                        {
                            Source = cell,
                        });
                }
                value.Bind(
                    TextBlock.FontSizeProperty,
                    CellFontSizeSlider.GetObservable(RangeBase.ValueProperty));
                var border = new Border
                {
                    Padding = new Thickness(7, 3),
                    Child = value,
                };
                if (cell != null)
                {
                    border.Bind(
                        Border.BackgroundProperty,
                        new Binding(remote ? nameof(DiffCell.RemoteBackground) : nameof(DiffCell.LocalBackground))
                        {
                            Source = cell,
                        });
                }
                if (cell?.IsConflict == true)
                {
                    border.Cursor = new Cursor(StandardCursorType.Hand);
                    border.Bind(
                        ToolTip.TipProperty,
                        new Binding(nameof(DiffCell.ResolutionToolTip)) { Source = cell });
                    border.PointerPressed += (_, e) =>
                    {
                        e.Handled = true;
                        OpenConflictDialog(row, cell);
                    };
                }
                return border;
            }),
        };
    }

    private static string ColumnName(int index)
    {
        var name = string.Empty;
        do
        {
            name = (char)('A' + index % 26) + name;
            index = index / 26 - 1;
        }
        while (index >= 0);
        return name;
    }

    private void FilterChanged(object? sender, RoutedEventArgs e) => ApplyFilter();

    private void ApplyFilter()
    {
        _visibleRows = HideUnchanged.IsChecked == true ? _changedRows.ToList() : _allRows.ToList();
        _visibleRowPositions = _visibleRows.Select(row => _rowPositions[row]).ToList();
        LocalGrid.ItemsSource = _visibleRows;
        RemoteGrid.ItemsSource = _visibleRows;
        SheetOverviewMap.SetRows(_allRows);
        _changeIndex = -1;
        ScheduleOverviewViewportUpdate();
    }

    private void SheetOverviewPositionRequested(object? sender, SheetOverviewPositionRequestedEventArgs e)
    {
        if (_allRows.Count == 0 || _visibleRows.Count == 0)
            return;

        var requestedRow = (int)Math.Round(e.Position.Y * (_allRows.Count - 1));
        var visibleIndex = FindNearestVisibleRow(requestedRow);
        var verticalPosition = _visibleRows.Count <= 1
            ? 0
            : (double)visibleIndex / (_visibleRows.Count - 1);
        _scrollSynchronizer.ScrollTo(e.Position.X, verticalPosition);
        ScheduleOverviewViewportUpdate();
    }

    private int FindNearestVisibleRow(int requestedRow)
    {
        var low = 0;
        var high = _visibleRows.Count;
        while (low < high)
        {
            var middle = low + (high - low) / 2;
            if (_visibleRowPositions[middle] < requestedRow)
                low = middle + 1;
            else
                high = middle;
        }

        if (low == 0)
            return 0;
        if (low == _visibleRows.Count)
            return _visibleRows.Count - 1;
        var previousDistance = requestedRow - _visibleRowPositions[low - 1];
        var nextDistance = _visibleRowPositions[low] - requestedRow;
        return previousDistance <= nextDistance ? low - 1 : low;
    }

    private void ScrollViewportChanged(object? sender, EventArgs e) => ScheduleOverviewViewportUpdate();

    private void GridSizeChanged(object? sender, SizeChangedEventArgs e) => ScheduleOverviewViewportUpdate();

    private void ScheduleOverviewViewportUpdate()
    {
        if (_overviewUpdatePending)
            return;

        _overviewUpdatePending = true;
        Dispatcher.UIThread.Post(() =>
        {
            Dispatcher.UIThread.Post(() =>
            {
                _overviewUpdatePending = false;
                UpdateOverviewViewport();
            }, DispatcherPriority.Loaded);
        }, DispatcherPriority.Loaded);
    }

    private void UpdateOverviewViewport()
    {
        if (_allRows.Count == 0 || _visibleRows.Count == 0)
        {
            SheetOverviewMap.SetViewport(new Rect(0, 0, 1, 1));
            return;
        }

        var vertical = MapVisibleViewportToSheet(
            _scrollSynchronizer.VerticalViewport,
            _visibleRowPositions,
            _allRows.Count);
        var horizontal = _scrollSynchronizer.HorizontalViewport;
        SheetOverviewMap.SetViewport(new Rect(
            horizontal.Start,
            vertical.Start,
            Math.Max(0, horizontal.End - horizontal.Start),
            Math.Max(0, vertical.End - vertical.Start)));
    }

    internal static NormalizedViewport MapVisibleViewportToSheet(
        NormalizedViewport viewport,
        IReadOnlyList<int> visibleRowPositions,
        int rowCount)
    {
        if (rowCount <= 0 || visibleRowPositions.Count == 0)
            return new NormalizedViewport(0, 1);

        return new NormalizedViewport(
            MapVisibleBoundary(viewport.Start, visibleRowPositions, rowCount),
            MapVisibleBoundary(viewport.End, visibleRowPositions, rowCount));
    }

    private static double MapVisibleBoundary(
        double position,
        IReadOnlyList<int> visibleRowPositions,
        int rowCount)
    {
        var scaled = Math.Clamp(position, 0, 1) * visibleRowPositions.Count;
        var boundary = Math.Min((int)Math.Floor(scaled), visibleRowPositions.Count - 1);
        var fraction = scaled - boundary;
        var start = (double)visibleRowPositions[boundary] / rowCount;
        var end = boundary + 1 < visibleRowPositions.Count
            ? (double)visibleRowPositions[boundary + 1] / rowCount
            : (double)(visibleRowPositions[^1] + 1) / rowCount;
        return Math.Clamp(start + (end - start) * fraction, 0, 1);
    }

    private void PreviousChange(object? sender, RoutedEventArgs e) => Navigate(-1);
    private void NextChange(object? sender, RoutedEventArgs e) => Navigate(1);

    private void Navigate(int direction)
    {
        if (_changedRows.Count == 0)
            return;

        _changeIndex = (_changeIndex + direction + _changedRows.Count) % _changedRows.Count;
        var target = _changedRows[_changeIndex];
        _synchronizingSelection = true;
        LocalGrid.SelectedItem = target;
        RemoteGrid.SelectedItem = target;
        _synchronizingSelection = false;
        LocalGrid.ScrollIntoView(target, null);
    }

    private void GridSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_synchronizingSelection || sender is not DataGrid source)
            return;

        _synchronizingSelection = true;
        var target = ReferenceEquals(source, LocalGrid) ? RemoteGrid : LocalGrid;
        target.SelectedItem = source.SelectedItem;
        _synchronizingSelection = false;
        if (source.SelectedItem is DiffRow row && row.HasConflict)
            UpdateConflictPanel(row);
    }

    private async void PreviousConflict(object? sender, RoutedEventArgs e) => await NavigateConflict(-1);
    private async void NextConflict(object? sender, RoutedEventArgs e) => await NavigateConflict(1);
    private async Task NavigateConflict(int direction)
    {
        if (_workbookConflictsLoaded)
        {
            if (_workbookConflictTargets.Count == 0)
                return;
            var currentKey = _currentMergeSheetName != null && _currentConflictCell != null
                ? new MergeCellKey(
                    _currentMergeSheetName,
                    _currentConflictCell.OriginalRowIndex,
                    _currentConflictCell.OriginalColumnIndex)
                : (MergeCellKey?)null;
            var currentIndex = currentKey.HasValue
                && _workbookConflictPositions.TryGetValue(currentKey.Value, out var position)
                    ? position
                    : -1;
            var targetIndex = currentIndex < 0
                ? direction >= 0 ? 0 : _workbookConflictTargets.Count - 1
                : (currentIndex + direction + _workbookConflictTargets.Count) % _workbookConflictTargets.Count;
            var workbookTarget = _workbookConflictTargets[targetIndex];
            if (!string.Equals(_currentMergeSheetName, workbookTarget.SheetName, StringComparison.Ordinal))
            {
                if (!await SelectSheet(workbookTarget.SheetName))
                    return;
            }

            if (_sheetConflictTargets.TryGetValue(workbookTarget, out var workbookConflict))
                SelectConflict(workbookConflict.Row, workbookConflict.Cell);
            return;
        }

        if (_conflictTargets.Count == 0)
            return;

        _conflictIndex = _conflictIndex < 0
            ? direction >= 0 ? 0 : _conflictTargets.Count - 1
            : (_conflictIndex + direction + _conflictTargets.Count) % _conflictTargets.Count;
        var target = _conflictTargets[_conflictIndex];
        SelectConflict(target.Row, target.Cell);
    }

    private void SelectConflict(DiffRow row, DiffCell? cell = null)
    {
        cell ??= row.ConflictCells.FirstOrDefault(candidate => candidate.Resolution == MergeResolution.Unresolved)
            ?? row.ConflictCells.FirstOrDefault();
        _currentConflictRow = row;
        _currentConflictCell = cell;
        _synchronizingSelection = true;
        SelectGridCell(LocalGrid, row, cell);
        SelectGridCell(RemoteGrid, row, cell);
        _synchronizingSelection = false;
        _conflictIndex = cell != null && _conflictPositions.TryGetValue(cell, out var position)
            ? position
            : -1;
        UpdateConflictPanel(row);
    }

    private void SelectGridCell(DataGrid grid, DiffRow row, DiffCell? cell)
    {
        grid.SelectedItem = row;
        DataGridColumn? column = null;
        if (cell != null)
        {
            var columns = ReferenceEquals(grid, LocalGrid) ? _localColumnsByIndex : _remoteColumnsByIndex;
            columns.TryGetValue(cell.ColumnIndex, out column);
        }
        if (column != null)
            grid.CurrentColumn = column;
        grid.ScrollIntoView(row, column);
    }

    private void OpenConflictDialog(DiffRow row, DiffCell cell)
    {
        if (!_viewModel.IsMergeMode || !cell.IsConflict)
            return;

        SelectConflict(row, cell);
        _isRowConflictDialog = false;
        _dialogConflictCell = cell;
        ConflictDialogTitle.Text = "Resolve cell conflict";
        ConflictDialogAddress.Text = $"{_currentMergeSheetName}!{ColumnName(cell.OriginalColumnIndex)}{cell.OriginalRowIndex + 1}";
        CellConflictPanel.IsVisible = true;
        CellBasePanel.IsVisible = true;
        RowConflictPanel.IsVisible = false;
        CustomChoice.IsEnabled = true;
        RowCustomEditor.IsVisible = false;
        RowColumnSelectionBorder.IsVisible = false;
        var inlineDiff = cell.InlineDiff;
        InlineTextDiff.Apply(ConflictLocalValue, inlineDiff.Local);
        InlineTextDiff.Apply(ConflictRemoteValue, inlineDiff.Remote);
        ConflictBaseValue.Text = cell.BaseValue;
        UseLocalChoice.IsChecked = cell.Resolution == MergeResolution.Local || cell.Resolution == MergeResolution.Unresolved;
        UseRemoteChoice.IsChecked = cell.Resolution == MergeResolution.Remote;
        BothChoice.IsChecked = cell.Resolution == MergeResolution.Both;
        CustomChoice.IsChecked = cell.Resolution == MergeResolution.Custom;
        CustomConflictValue.Text = cell.CustomValue ?? string.Empty;
        CustomConflictValue.IsVisible = CustomChoice.IsChecked == true;
        CustomConflictValue.IsEnabled = CustomConflictValue.IsVisible;
        BothResolutionHint.IsVisible = BothChoice.IsChecked == true;
        ConflictDialogOverlay.IsVisible = true;
    }

    private void OpenRowConflictDialog(DiffRow row)
    {
        if (!_viewModel.IsMergeMode || !row.HasConflict)
            return;

        SelectConflict(row);
        _isRowConflictDialog = true;
        _dialogConflictCell = null;
        _selectedRowCustomCell = null;
        _rowCustomDrafts.Clear();
        ConflictDialogTitle.Text = "Resolve row conflicts";
        ConflictDialogAddress.Text = $"{_currentMergeSheetName}!Row {row.ConflictRowIndex + 1} · {row.ConflictCells.Count} conflict cells";
        CellConflictPanel.IsVisible = false;
        CellBasePanel.IsVisible = false;
        RowConflictPanel.IsVisible = true;
        CustomConflictValue.IsVisible = false;
        CustomChoice.IsEnabled = true;
        RowCustomEditor.IsVisible = false;
        RowColumnSelectionBorder.IsVisible = false;

        foreach (var cell in row.ConflictCells)
            _rowCustomDrafts[cell] = GetResolvedValue(cell);
        PopulateRowSourceCells(row);
        var resolution = _rowResolutions.TryGetValue(RowKey(row), out var selected)
            ? selected
            : row.ConflictCells.All(cell => cell.Resolution == MergeResolution.Custom)
                ? MergeResolution.Custom
                : MergeResolution.Local;
        UseLocalChoice.IsChecked = resolution == MergeResolution.Local;
        UseRemoteChoice.IsChecked = resolution == MergeResolution.Remote;
        BothChoice.IsChecked = resolution == MergeResolution.Both;
        CustomChoice.IsChecked = resolution == MergeResolution.Custom;
        BothResolutionHint.IsVisible = resolution == MergeResolution.Both;
        UpdateRowResultPreview();
        if (resolution == MergeResolution.Custom)
            SelectRowCustomColumn(row.ConflictCells.First());
        ConflictDialogOverlay.IsVisible = true;
    }

    private void PopulateRowSourceCells(DiffRow row)
    {
        RowBaseCells.Children.Clear();
        RowLocalCells.Children.Clear();
        RowRemoteCells.Children.Clear();
        RowResultCells.Children.Clear();
        RowSecondResultCells.Children.Clear();
        _rowResultValueBoxes.Clear();
        foreach (var cell in row.Cells)
        {
            var inlineDiff = string.Equals(cell.LocalValue, cell.RemoteValue, StringComparison.Ordinal)
                ? null
                : cell.InlineDiff;
            RowBaseCells.Children.Add(RowCellBox(cell, cell.BaseValue, "#F8FAFC"));
            RowLocalCells.Children.Add(RowCellBox(cell, cell.LocalValue, "#EAF7EE", inlineDiff?.Local));
            RowRemoteCells.Children.Add(RowCellBox(cell, cell.RemoteValue, "#FCEDEF", inlineDiff?.Remote));
        }
    }

    private Border RowCellBox(
        DiffCell cell,
        string value,
        string background,
        IReadOnlyList<InlineDiffSegment>? inlineDiff = null)
    {
        var conflictColor = new SolidColorBrush(Color.Parse("#D6455D"));
        var valueBox = new SelectableTextBlock
        {
            Height = 30,
            Padding = new Thickness(5, 3),
            TextWrapping = TextWrapping.NoWrap,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Background = Brushes.White,
            Foreground = new SolidColorBrush(Color.Parse("#1F2937")),
        };
        if (inlineDiff == null)
            valueBox.Text = value;
        else
            InlineTextDiff.Apply(valueBox, inlineDiff);
        var border = new Border
        {
            Width = 140,
            Height = 64,
            Padding = new Thickness(7, 5),
            CornerRadius = new CornerRadius(4),
            Background = new SolidColorBrush(Color.Parse(background)),
            BorderBrush = cell.IsConflict
                ? conflictColor
                : new SolidColorBrush(Color.Parse("#C7CED8")),
            BorderThickness = new Thickness(cell.IsConflict ? 2 : 1),
            Child = new StackPanel
            {
                Spacing = 3,
                Children =
                {
                    new TextBlock
                    {
                        Text = cell.IsConflict
                            ? $"{ColumnName(cell.OriginalColumnIndex)} · CONFLICT"
                            : ColumnName(cell.OriginalColumnIndex),
                        FontSize = 11,
                        FontWeight = FontWeight.SemiBold,
                        Foreground = cell.IsConflict
                            ? conflictColor
                            : new SolidColorBrush(Color.Parse("#5B6472")),
                    },
                    valueBox,
                },
            },
        };
        ToolTip.SetTip(border, value);
        if (cell.IsConflict)
        {
            border.Cursor = new Cursor(StandardCursorType.Hand);
            border.AddHandler(
                PointerPressedEvent,
                (_, e) =>
                {
                    if (!_isRowConflictDialog || CustomChoice.IsChecked != true)
                        return;
                    if (ReferenceEquals(_selectedRowCustomCell, cell))
                        return;
                    SelectRowCustomColumn(cell);
                    e.Handled = true;
                },
                RoutingStrategies.Bubble,
                handledEventsToo: true);
        }
        border.Tag = valueBox;
        return border;
    }

    private static string GetResolvedValue(DiffCell cell) => cell.Resolution switch
    {
        MergeResolution.Remote => cell.RemoteValue,
        MergeResolution.Custom => cell.CustomValue ?? string.Empty,
        _ => cell.LocalValue,
    };

    private void UpdateRowResultPreview()
    {
        if (!_isRowConflictDialog || _currentConflictRow == null)
            return;

        RowResultCells.Children.Clear();
        RowSecondResultCells.Children.Clear();
        _rowResultValueBoxes.Clear();
        var keepBoth = BothChoice.IsChecked == true;
        var custom = CustomChoice.IsChecked == true;
        RowResultLabel.Text = keepBoth ? "RESULT 1 · LOCAL" : "RESULT";
        RowSecondResultLabel.IsVisible = keepBoth;
        RowSecondResultCells.IsVisible = keepBoth;
        foreach (var cell in _currentConflictRow.Cells)
        {
            var value = UseRemoteChoice.IsChecked == true
                ? cell.RemoteValue
                : custom
                    ? cell.IsConflict && _rowCustomDrafts.TryGetValue(cell, out var draft)
                        ? draft
                        : AutomaticMergedValue(cell)
                    : cell.LocalValue;
            var resultCell = RowCellBox(cell, value, "#E8F2FC");
            _rowResultValueBoxes[cell] = (SelectableTextBlock)resultCell.Tag!;
            RowResultCells.Children.Add(resultCell);
            if (keepBoth)
                RowSecondResultCells.Children.Add(RowCellBox(cell, cell.RemoteValue, "#E8F2FC"));
        }
    }

    private static string AutomaticMergedValue(DiffCell cell)
    {
        if (cell.LocalValue == cell.RemoteValue)
            return cell.LocalValue;
        if (cell.LocalValue == cell.BaseValue)
            return cell.RemoteValue;
        return cell.LocalValue;
    }

    private void SelectRowCustomColumn(DiffCell cell)
    {
        if (!_isRowConflictDialog || CustomChoice.IsChecked != true || !cell.IsConflict || _currentConflictRow == null)
            return;

        _selectedRowCustomCell = cell;
        var index = 0;
        while (index < _currentConflictRow.Cells.Count
            && !ReferenceEquals(_currentConflictRow.Cells[index], cell))
            index++;
        RowColumnSelectionBorder.Margin = new Thickness(index * 146 - 5, -5, -5, -5);
        RowColumnSelectionBorder.IsVisible = true;
        RowCustomEditor.IsVisible = true;
        RowCustomColumnTitle.Text = $"Edit conflict · Column {ColumnName(cell.OriginalColumnIndex)}";
        RowCustomBaseValue.Text = cell.BaseValue;
        var inlineDiff = cell.InlineDiff;
        InlineTextDiff.Apply(RowCustomLocalValue, inlineDiff.Local);
        InlineTextDiff.Apply(RowCustomRemoteValue, inlineDiff.Remote);
        _updatingRowCustomEditor = true;
        RowCustomResultValue.Text = _rowCustomDrafts[cell];
        _updatingRowCustomEditor = false;
        RowCustomResultValue.Focus();
        RowCustomResultValue.SelectAll();
    }

    private void RowCustomResultChanged(object? sender, TextChangedEventArgs e)
    {
        if (_updatingRowCustomEditor || _selectedRowCustomCell == null || !_isRowConflictDialog)
            return;

        var value = RowCustomResultValue.Text ?? string.Empty;
        _rowCustomDrafts[_selectedRowCustomCell] = value;
        if (_rowResultValueBoxes.TryGetValue(_selectedRowCustomCell, out var preview))
            preview.Text = value;
    }

    private void ResolutionChoiceChanged(object? sender, RoutedEventArgs e)
    {
        if (_isRowConflictDialog)
        {
            UpdateRowResultPreview();
            BothResolutionHint.IsVisible = BothChoice.IsChecked == true;
            if (CustomChoice.IsChecked == true)
                SelectRowCustomColumn(_selectedRowCustomCell ?? _currentConflictRow!.ConflictCells.First());
            else
            {
                RowCustomEditor.IsVisible = false;
                RowColumnSelectionBorder.IsVisible = false;
            }
            return;
        }

        CustomConflictValue.IsVisible = CustomChoice.IsChecked == true;
        CustomConflictValue.IsEnabled = CustomConflictValue.IsVisible;
        BothResolutionHint.IsVisible = BothChoice.IsChecked == true;
        if (CustomConflictValue.IsEnabled)
        {
            CustomConflictValue.Focus();
            CustomConflictValue.SelectAll();
        }
    }

    private void CancelConflictDialog(object? sender, RoutedEventArgs e)
    {
        ConflictDialogOverlay.IsVisible = false;
        _dialogConflictCell = null;
        _selectedRowCustomCell = null;
        _rowCustomDrafts.Clear();
        _rowResultValueBoxes.Clear();
    }

    private void ConfirmConflictDialog(object? sender, RoutedEventArgs e)
    {
        var cell = _dialogConflictCell;
        var row = _currentConflictRow;
        if (row == null || _currentMergeSheetName == null
            || !_isRowConflictDialog && (cell == null || !row.ConflictCells.Contains(cell)))
            return;

        var resolution = UseRemoteChoice.IsChecked == true
            ? MergeResolution.Remote
            : BothChoice.IsChecked == true
                ? MergeResolution.Both
                : CustomChoice.IsChecked == true
                    ? MergeResolution.Custom
                    : MergeResolution.Local;
        if (_isRowConflictDialog)
        {
            if (resolution == MergeResolution.Custom)
            {
                _rowResolutions.Remove(RowKey(row));
                foreach (var conflictCell in row.ConflictCells)
                    SetCellResolution(row, conflictCell, new MergeCellResolution(
                        MergeResolution.Custom,
                        _rowCustomDrafts[conflictCell]));
            }
            else
            {
                _rowResolutions[RowKey(row)] = resolution;
                foreach (var conflictCell in row.ConflictCells)
                {
                    _mergeResolutions.Remove(CellKey(row, conflictCell));
                    conflictCell.Resolve(resolution);
                }
            }
        }
        else if (resolution == MergeResolution.Both)
        {
            _rowResolutions[RowKey(row)] = MergeResolution.Both;
            foreach (var conflictCell in row.ConflictCells)
            {
                _mergeResolutions.Remove(CellKey(row, conflictCell));
                conflictCell.Resolve(MergeResolution.Both);
            }
        }
        else
        {
            var customValue = resolution == MergeResolution.Custom ? CustomConflictValue.Text ?? string.Empty : null;

            if (_rowResolutions.Remove(RowKey(row)))
            {
                foreach (var conflictCell in row.ConflictCells)
                    conflictCell.Resolve(MergeResolution.Unresolved);
            }

            if (row.ConflictCells.Any(conflictCell => conflictCell.Resolution == MergeResolution.Both))
            {
                foreach (var conflictCell in row.ConflictCells)
                {
                    conflictCell.Resolve(MergeResolution.Unresolved);
                    _mergeResolutions.Remove(CellKey(row, conflictCell));
                }
            }
            SetCellResolution(row, cell!, new MergeCellResolution(resolution, customValue));
        }

        ConflictDialogOverlay.IsVisible = false;
        _dialogConflictCell = null;
        _selectedRowCustomCell = null;
        _rowCustomDrafts.Clear();
        _rowResultValueBoxes.Clear();
        RefreshConflictResolutionProgress(row);
        if (_rowPositions.TryGetValue(row, out var rowPosition))
            SheetOverviewMap.UpdateRow(rowPosition, row);
        SelectConflict(row, _isRowConflictDialog ? null : cell);
    }

    private void SetCellResolution(DiffRow row, DiffCell cell, MergeCellResolution resolution)
    {
        cell.Resolve(resolution.Resolution, resolution.CustomValue);
        _mergeResolutions[CellKey(row, cell)] = resolution;
    }

    private MergeCellKey CellKey(DiffRow row, DiffCell cell) =>
        new(_currentMergeSheetName ?? string.Empty, cell.OriginalRowIndex, cell.OriginalColumnIndex);

    private MergeRowKey RowKey(DiffRow row) =>
        new(_currentMergeSheetName ?? string.Empty, row.ConflictRowIndex);

    private async void SaveResult(object? sender, RoutedEventArgs e)
    {
        if (_baseWorkbook == null || _localWorkbook == null || _remoteWorkbook == null)
        {
            Error.Text = "Compare BASE, LOCAL, and REMOTE before saving RESULT.";
            return;
        }

        var saved = false;
        var succeeded = await RunOperation(async cancellationToken =>
        {
            var path = _options.OutputPath;
            if (path == null)
            {
                var localExtension = Path.GetExtension(_loadedLocalPath ?? _viewModel.LocalPath);
                if (!localExtension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                    && !localExtension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                    localExtension = ".xlsx";
                var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
                {
                    Title = "Save merged RESULT workbook",
                    SuggestedFileName = $"merged-result{localExtension}",
                    DefaultExtension = localExtension.TrimStart('.'),
                    FileTypeChoices = new[]
                    {
                        new FilePickerFileType("Excel workbook") { Patterns = new[] { $"*{localExtension}" } },
                    },
                });
                if (file == null)
                    return;

                path = file.TryGetLocalPath() ?? file.Path.LocalPath;
                if (string.IsNullOrEmpty(Path.GetExtension(path)))
                    path = Path.ChangeExtension(path, ".xlsx");
            }

            await SaveMergedWorkbook(path, cancellationToken);
            Progress.Text = $"Saved RESULT: {path}";
            saved = true;
        });

        if (succeeded && saved && _options.IsMergeDriver)
            CompleteMergeDriver();
    }

    private async Task SaveMergedWorkbook(string outputPath, CancellationToken cancellationToken)
    {
        var baseWorkbook = _baseWorkbook
            ?? throw new InvalidOperationException("The BASE workbook could not be loaded.");
        var localWorkbook = _localWorkbook
            ?? throw new InvalidOperationException("The LOCAL workbook could not be loaded.");
        var remoteWorkbook = _remoteWorkbook
            ?? throw new InvalidOperationException("The REMOTE workbook could not be loaded.");
        var localSourcePath = _loadedLocalPath
            ?? throw new InvalidOperationException("The LOCAL source path could not be loaded.");
        var remoteSourcePath = _loadedRemotePath
            ?? throw new InvalidOperationException("The REMOTE source path could not be loaded.");
        var rowResolutions = new Dictionary<MergeRowKey, MergeResolution>(_rowResolutions);
        var cellResolutions = new Dictionary<MergeCellKey, MergeCellResolution>(_mergeResolutions);
        string? temporaryPath = null;
        var serviceOutputPath = outputPath;
        if (_options.IsMergeDriver)
        {
            var extension = GetMergeDriverOutputExtension(outputPath);
            temporaryPath = Path.Combine(Path.GetTempPath(), $"excelmerge-{Guid.NewGuid():N}{extension}");
            serviceOutputPath = temporaryPath;
        }

        try
        {
            var service = new WorkbookMergeService();
            await Task.Run(() => service.Save(
                baseWorkbook,
                localWorkbook,
                remoteWorkbook,
                rowResolutions,
                cellResolutions,
                localSourcePath,
                remoteSourcePath,
                serviceOutputPath), cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            if (temporaryPath != null)
                File.Copy(temporaryPath, outputPath, overwrite: true);
        }
        finally
        {
            if (temporaryPath != null)
            {
                try
                {
                    File.Delete(temporaryPath);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    // Temp cleanup must not turn a completed merge into a Git conflict.
                }
            }
        }
    }

    private string GetMergeDriverOutputExtension(string outputPath)
    {
        foreach (var candidate in new[] { outputPath, _options.RepositoryPath, _options.LocalPath })
        {
            var extension = Path.GetExtension(candidate ?? string.Empty) ?? string.Empty;
            if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                || extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                return extension;
        }

        throw new WorkbookMergeException(
            "Git merge driver output must identify an .xlsx or .xls file through --path or --output.");
    }

    private void CompleteMergeDriver()
    {
        Program.MarkMergeDriverCompleted();
        Close();
    }

    private void UpdateConflictPanel(DiffRow? preferredRow = null)
    {
        if (!_viewModel.IsMergeMode)
            return;

        if (_diff == null)
        {
            SetConflictPanelReady(false);
            ConflictSummary.Text = "Compare files to detect conflicts.";
            RemainingConflictCount.Text = "—";
            PreviousConflictButton.IsEnabled = false;
            NextConflictButton.IsEnabled = false;
            SaveResultButton.IsEnabled = false;
            UpdateWorkbookConflictOverview();
            return;
        }

        var resolvedCount = _workbookConflictsLoaded && _currentMergeSheetName != null
            ? _resolvedConflictCountsBySheet.GetValueOrDefault(_currentMergeSheetName)
            : _conflictTargets.Count(target => target.Cell.Resolution != MergeResolution.Unresolved);
        var currentSheetRemaining = _conflictTargets.Count - resolvedCount;
        var currentSheetReady = resolvedCount == _conflictTargets.Count;
        var (workbookTotal, workbookResolved) = GetWorkbookConflictProgress();
        var workbookRemaining = Math.Max(0, workbookTotal - workbookResolved);
        var workbookReady = _workbookConflictsLoaded && workbookResolved == workbookTotal;
        RemainingConflictCount.Text = !_workbookConflictsLoaded
            ? "SCANNING"
            : workbookTotal == 0
                ? "NO CONFLICTS"
                : workbookRemaining == 0 ? "ALL RESOLVED" : $"{workbookRemaining} LEFT";
        SetConflictPanelReady(workbookReady);
        SaveResultButton.IsEnabled = workbookReady;
        UpdateWorkbookConflictOverview((workbookTotal, workbookResolved));
        var hasConflicts = _conflictTargets.Count > 0;
        var hasNavigableConflicts = _workbookConflictsLoaded ? workbookTotal > 0 : hasConflicts;
        PreviousConflictButton.IsEnabled = hasNavigableConflicts;
        NextConflictButton.IsEnabled = hasNavigableConflicts;

        if (!hasConflicts)
        {
            _currentConflictRow = null;
            _currentConflictCell = null;
            _conflictIndex = -1;
            if (!workbookReady)
                ConflictSummary.Text = $"No conflicts in {_currentMergeSheetName} · {workbookRemaining} unresolved in other sheets";
            else if (workbookTotal > 0)
                ConflictSummary.Text = $"All {workbookTotal} workbook conflicts resolved · Ready to merge";
            else
                ConflictSummary.Text = _options.IsMergeDriver
                    ? "No conflicts. Review the changes, then select Auto Merge."
                    : "No conflicts. Changes can be merged automatically.";
            return;
        }

        ConflictSummary.Text = workbookReady
            ? $"All {workbookTotal} workbook conflicts resolved · Ready to merge"
            : currentSheetReady
                ? $"All {_conflictTargets.Count} conflicts in {_currentMergeSheetName} resolved · {workbookRemaining} unresolved in other sheets"
                : $"{currentSheetRemaining} unresolved in {_currentMergeSheetName} · {workbookRemaining} remaining in workbook";

        var row = preferredRow?.HasConflict == true
            ? preferredRow
            : _currentConflictRow?.HasConflict == true && _conflictRows.Contains(_currentConflictRow)
                ? _currentConflictRow
                : _conflictRows[0];
        var cell = ReferenceEquals(row, _currentConflictRow)
                && _currentConflictCell != null
                && row.ConflictCells.Contains(_currentConflictCell)
            ? _currentConflictCell
            : row.ConflictCells.FirstOrDefault(candidate => candidate.Resolution == MergeResolution.Unresolved)
                ?? row.ConflictCells[0];
        _currentConflictRow = row;
        _currentConflictCell = cell;
        _conflictIndex = _conflictPositions.TryGetValue(cell, out var position) ? position : -1;
    }

    private (int Total, int Resolved) GetWorkbookConflictProgress()
    {
        return (_workbookConflictTargets.Count, _workbookResolvedConflictCount);
    }

    private bool IsConflictResolved(MergeCellKey key)
    {
        return _rowResolutions.TryGetValue(new MergeRowKey(key.SheetName, key.RowIndex), out var rowResolution)
                && rowResolution != MergeResolution.Unresolved
            || _mergeResolutions.TryGetValue(key, out var cellResolution)
                && cellResolution.Resolution != MergeResolution.Unresolved;
    }

    private void UpdateSheetOverview()
    {
        if (_viewModel.IsMergeMode)
            UpdateWorkbookConflictOverview();
        else
            UpdateDiffSheetOverview();
    }

    private void UpdateDiffSheetOverview()
    {
        LocalSheetItems.Children.Clear();
        RemoteSheetItems.Children.Clear();
        AddDiffSheetCards(LocalSheet, LocalSheetItems, true);
        AddDiffSheetCards(RemoteSheet, RemoteSheetItems, false);

        var localName = (LocalSheet.SelectedItem as SheetChoice)?.Name;
        var remoteName = (RemoteSheet.SelectedItem as SheetChoice)?.Name;
        WorkbookConflictSummary.Text = localName == null || remoteName == null
            ? "Load both files to select sheets"
            : $"LOCAL {localName} · REMOTE {remoteName}";
    }

    private void AddDiffSheetCards(ComboBox selector, StackPanel panel, bool local)
    {
        var selectedName = (selector.SelectedItem as SheetChoice)?.Name;
        var choices = selector.ItemsSource?.Cast<SheetChoice>().Where(choice => choice.Exists)
            ?? Enumerable.Empty<SheetChoice>();
        foreach (var choice in choices)
        {
            var selected = string.Equals(choice.Name, selectedName, StringComparison.Ordinal);
            var accent = local ? "#167444" : "#A93B4C";
            var background = selected
                ? local ? "#ECF8F0" : "#FFF0F2"
                : "#FFFFFF";
            var button = new Button
            {
                MinWidth = 120,
                Padding = new Thickness(12, 6),
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Background = new SolidColorBrush(Color.Parse(background)),
                BorderBrush = new SolidColorBrush(Color.Parse(selected ? accent : "#CBD5E1")),
                BorderThickness = new Thickness(selected ? 2 : 1),
                Content = new TextBlock
                {
                    Text = choice.Name,
                    FontWeight = FontWeight.SemiBold,
                    Foreground = new SolidColorBrush(Color.Parse(selected ? accent : "#536174")),
                },
            };
            var name = choice.Name;
            button.Click += async (_, _) => await SelectDiffSheet(selector, name);
            panel.Children.Add(button);
        }
    }

    private void UpdateWorkbookConflictOverview((int Total, int Resolved)? progress = null)
    {
        if (!_workbookConflictsLoaded)
        {
            SheetConflictItems.Children.Clear();
            _sheetConflictCards.Clear();
            WorkbookConflictSummary.Text = "Compare files to scan all sheets";
            return;
        }
        if (_sheetConflictCards.Count != _workbookConflicts.Count
            || _workbookConflicts.Keys.Any(name => !_sheetConflictCards.ContainsKey(name)))
            RebuildWorkbookConflictCards();

        var (workbookTotal, workbookResolved) = progress ?? GetWorkbookConflictProgress();
        var readySheets = 0;
        foreach (var (name, keys) in _workbookConflicts)
        {
            var resolved = _resolvedConflictCountsBySheet.GetValueOrDefault(name);
            var unresolved = keys.Count - resolved;
            var ready = unresolved == 0;
            if (ready)
                readySheets++;

            var selected = string.Equals(name, _currentMergeSheetName, StringComparison.Ordinal);
            var status = keys.Count == 0
                ? "NO CONFLICTS"
                : ready ? $"{keys.Count} RESOLVED" : $"{unresolved} UNRESOLVED";
            var card = _sheetConflictCards[name];
            card.Button.Background = ready ? ReadySheetBackground : ConflictSheetBackground;
            card.Button.BorderBrush = selected
                ? SelectedSheetBorder
                : ready ? ReadySheetBorder : ConflictSheetBorder;
            card.Button.BorderThickness = new Thickness(selected ? 2 : 1);
            card.Status.Text = status;
            card.Status.Foreground = ready ? ReadySheetForeground : ConflictSheetForeground;
            ToolTip.SetTip(card.Button, ready
                ? $"{name} is ready"
                : $"Open {name} to resolve {unresolved} conflicts");
        }

        WorkbookConflictSummary.Text = workbookTotal == 0
            ? $"{_workbookConflicts.Count} sheets · No conflicts"
            : $"{readySheets}/{_workbookConflicts.Count} sheets ready · {workbookResolved}/{workbookTotal} conflicts resolved";
    }

    private void RebuildWorkbookConflictCards()
    {
        SheetConflictItems.Children.Clear();
        _sheetConflictCards.Clear();
        foreach (var name in _workbookConflicts.Keys)
        {
            var status = new TextBlock
            {
                FontSize = 11,
                FontWeight = FontWeight.Bold,
            };
            var button = new Button
            {
                MinWidth = 150,
                Padding = new Thickness(12, 7),
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Content = new StackPanel
                {
                    Spacing = 1,
                    Children =
                    {
                        new TextBlock
                        {
                            Text = name,
                            FontWeight = FontWeight.SemiBold,
                            Foreground = SheetNameForeground,
                        },
                        status,
                    },
                },
            };
            var sheetName = name;
            button.Click += async (_, _) => await SelectSheet(sheetName);
            _sheetConflictCards.Add(name, new SheetConflictCard(button, status));
            SheetConflictItems.Children.Add(button);
        }
    }

    private void SetConflictPanelReady(bool ready)
    {
        var (background, border, foreground) = ready
            ? ("#ECF8F0", "#B9DFC7", "#176B38")
            : ("#FFF3F4", "#F0C6CC", "#963547");
        MergeConflictPanel.Background = new SolidColorBrush(Color.Parse(background));
        MergeConflictPanel.BorderBrush = new SolidColorBrush(Color.Parse(border));
        ConflictSummary.Foreground = new SolidColorBrush(Color.Parse(foreground));
        RemainingConflictBadge.Background = new SolidColorBrush(Color.Parse(ready ? "#D9F2E3" : "#FBDDE2"));
        RemainingConflictBadge.BorderBrush = new SolidColorBrush(Color.Parse(ready ? "#A9D5B9" : "#E8B9C1"));
        RemainingConflictCount.Foreground = new SolidColorBrush(Color.Parse(foreground));
    }

    private void GridLoadingRow(object? sender, DataGridRowEventArgs e)
    {
        e.Row.Height = e.Row.DataContext is DiffRow row ? row.RowHeight : double.NaN;
    }

    private void GridUnloadingRow(object? sender, DataGridRowEventArgs e)
    {
        e.Row.ClearValue(Layoutable.HeightProperty);
    }

    private void ResizeRow(DiffRow row, double height)
    {
        row.SetCustomRowHeight(height);
        ApplyRowHeight(LocalGrid, row, height);
        ApplyRowHeight(RemoteGrid, row, height);
        ScheduleOverviewViewportUpdate();
    }

    private void ApplyDefaultCellMetrics()
    {
        var width = new DataGridLength(CellWidth(_cellFontSize));
        foreach (var column in LocalGrid.Columns.Skip(1))
            column.Width = width;
        foreach (var column in RemoteGrid.Columns.Skip(1))
            column.Width = width;

        var height = CellHeight(_cellFontSize);
        foreach (var row in _allRows)
        {
            if (row.HasCustomRowHeight)
                continue;
            row.SetDefaultRowHeight(height);
        }
        ApplyRealizedRowHeights(LocalGrid);
        ApplyRealizedRowHeights(RemoteGrid);
        ScheduleOverviewViewportUpdate();
    }

    private static void ApplyRealizedRowHeights(DataGrid grid)
    {
        foreach (var realizedRow in grid.GetVisualDescendants().OfType<DataGridRow>())
        {
            if (realizedRow.DataContext is DiffRow row)
                realizedRow.Height = row.RowHeight;
        }
        grid.InvalidateMeasure();
    }

    internal static double CellWidth(double fontSize) =>
        Math.Clamp(DefaultCellWidth + (fontSize - DefaultCellFontSize) * 8, 96, 192);

    internal static double CellHeight(double fontSize) =>
        Math.Clamp(DefaultCellHeight + (fontSize - DefaultCellFontSize) * 2, 24, 46);

    private static void ApplyRowHeight(DataGrid grid, DiffRow row, double height)
    {
        foreach (var realizedRow in grid.GetVisualDescendants().OfType<DataGridRow>())
        {
            if (ReferenceEquals(realizedRow.DataContext, row))
                realizedRow.Height = height;
        }
        grid.InvalidateMeasure();
    }

    private static string FormatSummary(ExcelSheetDiffSummary summary) =>
        $"{summary.ModifiedCellCount} changed cells  ·  {summary.AddedRowCount} added  ·  {summary.RemovedRowCount} removed  ·  {summary.ModifiedRowCount} changed rows";
}

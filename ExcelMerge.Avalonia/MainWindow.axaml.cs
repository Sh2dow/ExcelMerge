using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Selection;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using ExcelMerge;

namespace ExcelMerge.Avalonia;

public partial class MainWindow : Window
{
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xlsx", ".xls", ".csv", ".tsv",
    };

    private readonly MainViewModel _viewModel;
    private ExcelWorkbook? _baseWorkbook;
    private ExcelWorkbook? _localWorkbook;
    private ExcelWorkbook? _remoteWorkbook;
    private ExcelSheetDiff? _diff;
    private List<DiffRow> _allRows = new();
    private readonly Dictionary<MergeRowKey, MergeResolution> _mergeResolutions = new();
    private int _changeIndex = -1;
    private int _conflictIndex = -1;
    private DiffRow? _currentConflictRow;
    private string? _loadedBasePath;
    private string? _loadedLocalPath;
    private string? _loadedRemotePath;
    private string? _currentMergeSheetName;
    private readonly SplitScrollSynchronizer _scrollSynchronizer;
    private ColumnWidthSynchronizer? _columnWidthSynchronizer;
    private bool _synchronizingSelection;
    private bool _isBusy;
    private bool _isClosed;
    private CancellationTokenSource? _operationCancellation;

    private enum WorkbookInput
    {
        Base,
        Local,
        Remote,
    }

    public MainWindow() : this(Array.Empty<string>())
    {
    }

    public MainWindow(string[] args)
    {
        InitializeComponent();

        var options = CommandLineOptions.Parse(args);

        _viewModel = new MainViewModel(CompareFiles, options.Mode)
        {
            BasePath = options.BasePath ?? string.Empty,
            LocalPath = options.LocalPath ?? string.Empty,
            RemotePath = options.RemotePath ?? string.Empty,
        };
        DataContext = _viewModel;
        _scrollSynchronizer = new SplitScrollSynchronizer(LocalGrid, RemoteGrid);
        Title = options.Mode == ApplicationMode.Merge ? "ExcelMerge - Merge" : "ExcelMerge - Diff";

        Closed += (_, _) =>
        {
            _isClosed = true;
            _operationCancellation?.Cancel();
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

    private async Task CompareIfReady()
    {
        if (CanAutoCompare())
            await CompareFiles();
    }

    private bool CanAutoCompare()
    {
        return File.Exists(_viewModel.LocalPath)
            && File.Exists(_viewModel.RemotePath)
            && (!_viewModel.IsMergeMode || File.Exists(_viewModel.BasePath));
    }

    private async Task CompareFiles()
    {
        await RunOperation(CompareCurrentFiles);
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
        var localName = LocalSheet.SelectedItem as string ?? localWorkbook.Sheets.Keys.First();
        var remoteName = RemoteSheet.SelectedItem as string ?? remoteWorkbook.Sheets.Keys.First();
        if (_viewModel.IsMergeMode
            && (_baseWorkbook?.Sheets.Count > 1 || localWorkbook.Sheets.Count > 1 || remoteWorkbook.Sheets.Count > 1)
            && !string.Equals(localName, remoteName, StringComparison.Ordinal))
            throw new WorkbookMergeException("Select the same sheet name in LOCAL and REMOTE before merging.");
        _currentMergeSheetName = localName;
        var baseValues = _viewModel.IsMergeMode
            ? CreateBaseValueMap(_baseWorkbook, localName, remoteName)
            : null;

        _diff = await Task.Run(
            () => ExcelSheet.Diff(localWorkbook.Sheets[localName], remoteWorkbook.Sheets[remoteName], new ExcelSheetDiffConfig()),
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        BuildRows(baseValues);
        var prefix = _viewModel.IsMergeMode ? "Merge preview · " : string.Empty;
        Summary.Text = prefix + FormatSummary(_diff.CreateSummary());
        Progress.Text = $"{_allRows.Count} rows";
        SaveResultButton.IsEnabled = _viewModel.IsMergeMode;
    }

    private void ClearDiff()
    {
        _diff = null;
        _allRows.Clear();
        _mergeResolutions.Clear();
        _currentMergeSheetName = null;
        _columnWidthSynchronizer?.Dispose();
        _columnWidthSynchronizer = null;
        LocalGrid.Columns.Clear();
        RemoteGrid.Columns.Clear();
        LocalGrid.ItemsSource = null;
        RemoteGrid.ItemsSource = null;
        Summary.Text = "No comparison loaded.";
        _changeIndex = -1;
        _conflictIndex = -1;
        _currentConflictRow = null;
        SaveResultButton.IsEnabled = false;
        UpdateConflictPanel();
    }

    private static IReadOnlyDictionary<(int Row, int Column), string> CreateBaseValueMap(
        ExcelWorkbook? workbook,
        string localSheetName,
        string remoteSheetName)
    {
        var values = new Dictionary<(int Row, int Column), string>();
        if (workbook == null)
            return values;

        if (!workbook.Sheets.TryGetValue(localSheetName, out var sheet)
            && !workbook.Sheets.TryGetValue(remoteSheetName, out sheet))
        {
            if (workbook.Sheets.Count != 1)
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

        var isLocal = input == WorkbookInput.Local;
        var selector = isLocal ? LocalSheet : RemoteSheet;
        var previousSelection = selector.SelectedItem as string;
        var sheetNames = workbook.Sheets.Keys.ToList();
        selector.ItemsSource = sheetNames;
        selector.SelectedItem = previousSelection != null && sheetNames.Contains(previousSelection)
            ? previousSelection
            : sheetNames.FirstOrDefault();

        if (isLocal)
        {
            _localWorkbook = workbook;
            _loadedLocalPath = path;
        }
        else
        {
            _remoteWorkbook = workbook;
            _loadedRemotePath = path;
        }
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
        DiffOptionsPanel.IsEnabled = enabled;
        MergeConflictPanel.IsEnabled = enabled;
    }

    private void BuildRows(IReadOnlyDictionary<(int Row, int Column), string>? baseValues)
    {
        _currentConflictRow = null;
        _conflictIndex = -1;
        _allRows = _diff!.Rows.Values.Select(row => new DiffRow(row, baseValues)).ToList();
        if (_currentMergeSheetName != null)
        {
            foreach (var row in _allRows.Where(row => row.HasConflict))
            {
                if (_mergeResolutions.TryGetValue(
                    new MergeRowKey(_currentMergeSheetName, row.ConflictRowIndex),
                    out var resolution))
                    row.Resolve(resolution);
            }
        }
        var columns = _allRows.SelectMany(row => row.Cells).Select(cell => cell.ColumnIndex).Distinct().OrderBy(index => index).ToList();
        _columnWidthSynchronizer?.Dispose();
        _columnWidthSynchronizer = null;
        LocalGrid.Columns.Clear();
        RemoteGrid.Columns.Clear();
        LocalGrid.Columns.Add(RowNumberColumn());
        RemoteGrid.Columns.Add(RowNumberColumn());
        foreach (var column in columns)
        {
            LocalGrid.Columns.Add(CellColumn(column, false));
            RemoteGrid.Columns.Add(CellColumn(column, true));
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
        CellTemplate = new FuncDataTemplate<DiffRow>((row, _) => new RowResizeCell(row, ResizeRow)),
    };

    private static DataGridTemplateColumn CellColumn(int index, bool remote)
    {
        return new DataGridTemplateColumn
        {
            Header = ColumnName(index),
            Width = new DataGridLength(140),
            CellTemplate = new FuncDataTemplate<DiffRow>((row, _) =>
            {
                var cell = row.Cells.FirstOrDefault(candidate => candidate.ColumnIndex == index);
                return new Border
                {
                    Padding = new Thickness(7, 3),
                    Background = cell?.Brush(remote),
                    Child = new TextBlock { Text = cell?.Value(remote) ?? string.Empty },
                };
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
        var rows = HideUnchanged.IsChecked == true ? _allRows.Where(row => row.IsChanged) : _allRows;
        var visibleRows = rows.ToList();
        LocalGrid.ItemsSource = visibleRows;
        RemoteGrid.ItemsSource = visibleRows;
        _changeIndex = -1;
    }

    private void PreviousChange(object? sender, RoutedEventArgs e) => Navigate(-1);
    private void NextChange(object? sender, RoutedEventArgs e) => Navigate(1);

    private void Navigate(int direction)
    {
        var changed = _allRows.Where(row => row.IsChanged).ToList();
        if (changed.Count == 0)
            return;

        _changeIndex = (_changeIndex + direction + changed.Count) % changed.Count;
        var target = changed[_changeIndex];
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

    private void PreviousConflict(object? sender, RoutedEventArgs e) => NavigateConflict(-1);
    private void NextConflict(object? sender, RoutedEventArgs e) => NavigateConflict(1);
    private void UseLocal(object? sender, RoutedEventArgs e) => ResolveCurrentConflict(MergeResolution.Local);
    private void UseRemote(object? sender, RoutedEventArgs e) => ResolveCurrentConflict(MergeResolution.Remote);
    private void KeepBoth(object? sender, RoutedEventArgs e) => ResolveCurrentConflict(MergeResolution.Both);

    private void NavigateConflict(int direction)
    {
        var conflicts = _allRows.Where(row => row.HasConflict).ToList();
        if (conflicts.Count == 0)
            return;

        _conflictIndex = (_conflictIndex + direction + conflicts.Count) % conflicts.Count;
        SelectConflict(conflicts[_conflictIndex]);
    }

    private void SelectConflict(DiffRow row)
    {
        _currentConflictRow = row;
        _synchronizingSelection = true;
        LocalGrid.SelectedItem = row;
        RemoteGrid.SelectedItem = row;
        _synchronizingSelection = false;
        LocalGrid.ScrollIntoView(row, null);
        UpdateConflictPanel(row);
    }

    private void ResolveCurrentConflict(MergeResolution resolution)
    {
        var row = LocalGrid.SelectedItem as DiffRow;
        if (row?.HasConflict != true || !_allRows.Contains(row))
            row = _currentConflictRow;
        if (row?.HasConflict != true || !_allRows.Contains(row))
            return;

        row.Resolve(resolution);
        if (_currentMergeSheetName != null)
            _mergeResolutions[new MergeRowKey(_currentMergeSheetName, row.ConflictRowIndex)] = resolution;
        ApplyFilter();
        SelectConflict(row);
    }

    private async void SaveResult(object? sender, RoutedEventArgs e)
    {
        if (_baseWorkbook == null || _localWorkbook == null || _remoteWorkbook == null)
        {
            Error.Text = "Compare BASE, LOCAL, and REMOTE before saving RESULT.";
            return;
        }

        await RunOperation(async cancellationToken =>
        {
            var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = "Save merged RESULT workbook",
                SuggestedFileName = "merged-result.xlsx",
                DefaultExtension = "xlsx",
                FileTypeChoices = new[]
                {
                    new FilePickerFileType("Excel workbook") { Patterns = new[] { "*.xlsx", "*.xls" } },
                },
            });
            if (file == null)
                return;

            var path = file.TryGetLocalPath() ?? file.Path.LocalPath;
            if (string.IsNullOrEmpty(Path.GetExtension(path)))
                path = Path.ChangeExtension(path, ".xlsx");
            var service = new WorkbookMergeService();
            await Task.Run(() => service.Save(
                _baseWorkbook,
                _localWorkbook,
                _remoteWorkbook,
                new Dictionary<MergeRowKey, MergeResolution>(_mergeResolutions),
                path), cancellationToken);
            Progress.Text = $"Saved RESULT: {path}";
        });
    }

    private void UpdateConflictPanel(DiffRow? preferredRow = null)
    {
        if (!_viewModel.IsMergeMode)
            return;

        var conflicts = _allRows.Where(row => row.HasConflict).ToList();
        var resolvedCount = conflicts.Count(row => row.Resolution != MergeResolution.Unresolved);
        ConflictSummary.Text = $"{conflicts.Count} conflicts · {resolvedCount} resolved";
        var hasConflicts = conflicts.Count > 0;
        PreviousConflictButton.IsEnabled = hasConflicts;
        NextConflictButton.IsEnabled = hasConflicts;

        if (!hasConflicts)
        {
            _currentConflictRow = null;
            _conflictIndex = -1;
            ConflictDetails.Text = "No conflicts. Changes can be merged automatically.";
            SetResolutionButtonsEnabled(false);
            return;
        }

        var row = preferredRow?.HasConflict == true
            ? preferredRow
            : _currentConflictRow?.HasConflict == true && conflicts.Contains(_currentConflictRow)
                ? _currentConflictRow
                : conflicts[0];
        _currentConflictRow = row;
        _conflictIndex = conflicts.IndexOf(row);
        var cell = row.ConflictCells.First();
        var address = $"{ColumnName(cell.ColumnIndex)}{row.Index}";
        ConflictDetails.Text = $"{address} · BASE: {FormatConflictValue(cell.BaseValue)} · LOCAL: {FormatConflictValue(cell.LocalValue)} · REMOTE: {FormatConflictValue(cell.RemoteValue)} · Resolution: {row.Resolution}";
        SetResolutionButtonsEnabled(true);
    }

    private void SetResolutionButtonsEnabled(bool enabled)
    {
        UseLocalButton.IsEnabled = enabled;
        UseRemoteButton.IsEnabled = enabled;
        KeepBothButton.IsEnabled = enabled;
    }

    private static string FormatConflictValue(string value)
    {
        if (string.IsNullOrEmpty(value))
            return "<empty>";
        return $"\"{value.Replace("\r", " ").Replace("\n", " ")}\"";
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
        row.RowHeight = height;
        ApplyRowHeight(LocalGrid, row, height);
        ApplyRowHeight(RemoteGrid, row, height);
    }

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

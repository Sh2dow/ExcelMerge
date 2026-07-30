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
    private readonly CommandLineOptions _options;
    private ExcelWorkbook? _baseWorkbook;
    private ExcelWorkbook? _localWorkbook;
    private ExcelWorkbook? _remoteWorkbook;
    private ExcelSheetDiff? _diff;
    private List<DiffRow> _allRows = new();
    private readonly Dictionary<MergeCellKey, MergeCellResolution> _mergeResolutions = new();
    private readonly Dictionary<MergeRowKey, MergeResolution> _rowResolutions = new();
    private int _changeIndex = -1;
    private int _conflictIndex = -1;
    private DiffRow? _currentConflictRow;
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
        _scrollSynchronizer = new SplitScrollSynchronizer(LocalGrid, RemoteGrid);
        Title = _options.Mode == ApplicationMode.Merge ? "ExcelMerge - Merge" : "ExcelMerge - Diff";
        if (_options.IsMergeDriver)
        {
            SaveResultButton.Content = "Complete Git merge";
            if (!string.IsNullOrWhiteSpace(_options.RepositoryPath))
                Title += $" - {_options.RepositoryPath}";
        }

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
        _currentMergeSheetName = localName;
        var baseValues = _viewModel.IsMergeMode
            ? CreateBaseValueMap(_baseWorkbook, localName, remoteName)
            : null;
        var localSheet = localChoice.Exists ? localWorkbook.Sheets[localName] : new ExcelSheet();
        var remoteSheet = remoteChoice.Exists ? remoteWorkbook.Sheets[remoteName] : new ExcelSheet();

        _diff = await Task.Run(
            () => ExcelSheet.Diff(localSheet, remoteSheet, new ExcelSheetDiffConfig()),
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        BuildRows(baseValues);
        var prefix = _viewModel.IsMergeMode ? "Merge preview · " : string.Empty;
        var missingSheet = !localChoice.Exists
            ? "LOCAL sheet missing · "
            : !remoteChoice.Exists
                ? "REMOTE sheet missing · "
                : string.Empty;
        Summary.Text = prefix + missingSheet + FormatSummary(_diff.CreateSummary());
        Progress.Text = $"{_allRows.Count} rows";
        SaveResultButton.IsEnabled = _viewModel.IsMergeMode;

        if (_options.IsMergeDriver && !_allRows.Any(row => row.HasConflict))
        {
            Progress.Text = "Writing automatic Git merge result...";
            await SaveMergedWorkbook(_options.OutputPath!, cancellationToken);
            CompleteMergeDriver();
        }
    }

    private void ClearDiff()
    {
        _diff = null;
        _allRows.Clear();
        _mergeResolutions.Clear();
        _rowResolutions.Clear();
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
        _dialogConflictCell = null;
        ConflictDialogOverlay.IsVisible = false;
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

        _updatingSheetSelections = true;
        LocalSheet.ItemsSource = localChoices;
        RemoteSheet.ItemsSource = remoteChoices;
        LocalSheet.SelectedItem = localChoices.FirstOrDefault(choice => choice.Name == selectedName);
        RemoteSheet.SelectedItem = remoteChoices.FirstOrDefault(choice => choice.Name == selectedName);
        _updatingSheetSelections = false;
    }

    private async void SheetSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (_updatingSheetSelections || sender is not ComboBox source || source.SelectedItem is not SheetChoice selected)
            return;

        var target = ReferenceEquals(source, LocalSheet) ? RemoteSheet : LocalSheet;
        var matching = target.ItemsSource?.Cast<SheetChoice>().FirstOrDefault(choice => choice.Name == selected.Name);
        if (matching != null)
        {
            _updatingSheetSelections = true;
            target.SelectedItem = matching;
            _updatingSheetSelections = false;
        }
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
        CellTemplate = new FuncDataTemplate<DiffRow>((row, _) => new RowResizeCell(row, ResizeRow, OpenRowConflictDialog)),
    };

    private DataGridTemplateColumn CellColumn(int index, bool remote)
    {
        return new DataGridTemplateColumn
        {
            Header = ColumnName(index),
            Width = new DataGridLength(140),
            CellTemplate = new FuncDataTemplate<DiffRow>((row, _) =>
            {
                var cell = row.Cells.FirstOrDefault(candidate => candidate.ColumnIndex == index);
                var border = new Border
                {
                    Padding = new Thickness(7, 3),
                    Background = cell?.Brush(remote),
                    Child = new TextBlock { Text = cell?.Value(remote) ?? string.Empty },
                };
                if (cell?.IsConflict == true)
                {
                    border.Cursor = new Cursor(StandardCursorType.Hand);
                    ToolTip.SetTip(border, cell.Resolution == MergeResolution.Unresolved
                        ? "Click to resolve this conflict"
                        : $"Resolved: {cell.Resolution}. Click to change.");
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

    private void OpenConflictDialog(DiffRow row, DiffCell cell)
    {
        if (!_viewModel.IsMergeMode || !cell.IsConflict)
            return;

        SelectConflict(row);
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
        var inlineDiff = InlineTextDiff.Create(cell.LocalValue, cell.RemoteValue);
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
        foreach (var cell in row.Cells.OrderBy(cell => cell.ColumnIndex))
        {
            var inlineDiff = InlineTextDiff.Create(cell.LocalValue, cell.RemoteValue);
            RowBaseCells.Children.Add(RowCellBox(cell, cell.BaseValue, "#F8FAFC"));
            RowLocalCells.Children.Add(RowCellBox(cell, cell.LocalValue, "#EAF7EE", inlineDiff.Local));
            RowRemoteCells.Children.Add(RowCellBox(cell, cell.RemoteValue, "#FCEDEF", inlineDiff.Remote));
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
        foreach (var cell in _currentConflictRow.Cells.OrderBy(cell => cell.ColumnIndex))
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
        var index = _currentConflictRow.Cells.OrderBy(candidate => candidate.ColumnIndex).ToList().IndexOf(cell);
        RowColumnSelectionBorder.Margin = new Thickness(index * 146 - 5, -5, -5, -5);
        RowColumnSelectionBorder.IsVisible = true;
        RowCustomEditor.IsVisible = true;
        RowCustomColumnTitle.Text = $"Edit conflict · Column {ColumnName(cell.OriginalColumnIndex)}";
        RowCustomBaseValue.Text = cell.BaseValue;
        var inlineDiff = InlineTextDiff.Create(cell.LocalValue, cell.RemoteValue);
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
            foreach (var conflictCell in row.ConflictCells)
                SetCellResolution(row, conflictCell, new MergeCellResolution(MergeResolution.Both));
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
        ApplyFilter();
        SelectConflict(row);
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

        var conflicts = _allRows.Where(row => row.HasConflict).ToList();
        var conflictCells = conflicts.SelectMany(row => row.ConflictCells).ToList();
        var resolvedCount = conflictCells.Count(cell => cell.Resolution != MergeResolution.Unresolved);
        ConflictSummary.Text = $"{conflictCells.Count} conflict cells · {resolvedCount} resolved · Click a red cell to resolve and preview RESULT";
        var hasConflicts = conflicts.Count > 0;
        PreviousConflictButton.IsEnabled = hasConflicts;
        NextConflictButton.IsEnabled = hasConflicts;

        if (!hasConflicts)
        {
            _currentConflictRow = null;
            _conflictIndex = -1;
            ConflictSummary.Text = "No conflicts. Changes can be merged automatically.";
            return;
        }

        var row = preferredRow?.HasConflict == true
            ? preferredRow
            : _currentConflictRow?.HasConflict == true && conflicts.Contains(_currentConflictRow)
                ? _currentConflictRow
                : conflicts[0];
        _currentConflictRow = row;
        _conflictIndex = conflicts.IndexOf(row);
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

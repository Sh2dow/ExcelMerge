using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Platform.Storage;
using Avalonia.Threading;

namespace ExcelMerge.Desktop;

public sealed partial class MainWindow : Window
{
    private MainWindowViewModel? _subscribedViewModel;
    private bool _updatingHorizontalScrollBars;
    private bool _updatingVerticalScrollBar;

    public MainWindow()
    {
        InitializeComponent();
        DiffGridControl.CellSelected += GridCellSelected;
        DiffGridControl.HorizontalScrollChanged += GridHorizontalScrollChanged;
        DiffGridControl.VerticalScrollChanged += GridVerticalScrollChanged;
        ChangeMapControl.PositionRequested += ChangeMapPositionRequested;
        SynchronizeHorizontalScrollBars();
        SynchronizeVerticalNavigation();
        DataContextChanged += (_, _) => SubscribeViewModel();
    }

    private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

    private VirtualDiffGrid DiffGridControl => this.FindControl<VirtualDiffGrid>("DiffGrid")!;

    private ScrollBar LocalHorizontalScrollBarControl =>
        this.FindControl<ScrollBar>("LocalHorizontalScrollBar")!;

    private ScrollBar RemoteHorizontalScrollBarControl =>
        this.FindControl<ScrollBar>("RemoteHorizontalScrollBar")!;

    private ScrollBar VerticalScrollBarControl =>
        this.FindControl<ScrollBar>("VerticalScrollBar")!;

    private DiffChangeMap ChangeMapControl => this.FindControl<DiffChangeMap>("ChangeMap")!;

    private ListBox ConflictNavigatorControl => this.FindControl<ListBox>("ConflictNavigator")!;

    private void SubscribeViewModel()
    {
        if (_subscribedViewModel is not null)
        {
            _subscribedViewModel.NavigationRequested -= NavigateGrid;
            _subscribedViewModel.ConflictSelectionRequested -= RevealConflict;
        }

        _subscribedViewModel = DataContext as MainWindowViewModel;
        if (_subscribedViewModel is not null)
        {
            _subscribedViewModel.NavigationRequested += NavigateGrid;
            _subscribedViewModel.ConflictSelectionRequested += RevealConflict;
        }
    }

    private void NavigateGrid(int row, int column) => DiffGridControl.ScrollTo(row, column);

    private void RevealConflict(ConflictItemViewModel conflict)
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (!ReferenceEquals(_subscribedViewModel?.SelectedConflict, conflict))
            {
                return;
            }

            ConflictNavigatorControl.ScrollIntoView(conflict);
        });
    }

    private void ConflictNavigatorPointerPressed(object? sender, PointerPressedEventArgs e) =>
        _subscribedViewModel?.CancelPendingCellSelection();

    private async void GridCellSelected(object? sender, GridCellSelection selection)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            await viewModel.SelectCellAsync(selection);
        }
    }

    private void GridHorizontalScrollChanged(object? sender, EventArgs e) =>
        SynchronizeHorizontalScrollBars();

    private void GridVerticalScrollChanged(object? sender, EventArgs e) =>
        SynchronizeVerticalNavigation();

    private void HorizontalScrollValueChanged(
        object? sender,
        RangeBaseValueChangedEventArgs e)
    {
        if (!_updatingHorizontalScrollBars)
        {
            DiffGridControl.ScrollHorizontalTo(e.NewValue);
        }
    }

    private void VerticalScrollValueChanged(
        object? sender,
        RangeBaseValueChangedEventArgs e)
    {
        if (!_updatingVerticalScrollBar)
        {
            DiffGridControl.ScrollVerticalTo(e.NewValue);
        }
    }

    private void ChangeMapPositionRequested(double horizontalRatio, double verticalRatio) =>
        DiffGridControl.CenterOnMapPosition(horizontalRatio, verticalRatio);

    private void SynchronizeHorizontalScrollBars()
    {
        _updatingHorizontalScrollBars = true;
        try
        {
            Synchronize(LocalHorizontalScrollBarControl);
            Synchronize(RemoteHorizontalScrollBarControl);
        }
        finally
        {
            _updatingHorizontalScrollBars = false;
        }

        ChangeMapControl.FirstVisibleColumn = DiffGridControl.FirstVisibleColumn;
        ChangeMapControl.VisibleColumnCount = DiffGridControl.HorizontalViewportSize;

        void Synchronize(ScrollBar scrollBar)
        {
            scrollBar.Maximum = DiffGridControl.HorizontalScrollMaximum;
            scrollBar.ViewportSize = DiffGridControl.HorizontalViewportSize;
            scrollBar.LargeChange = Math.Max(1, DiffGridControl.HorizontalViewportSize);
            scrollBar.Value = DiffGridControl.FirstVisibleColumn;
            scrollBar.IsEnabled = DiffGridControl.HorizontalScrollMaximum > 0;
        }
    }

    private void SynchronizeVerticalNavigation()
    {
        _updatingVerticalScrollBar = true;
        try
        {
            VerticalScrollBarControl.Maximum = DiffGridControl.VerticalScrollMaximum;
            VerticalScrollBarControl.ViewportSize = DiffGridControl.VerticalViewportSize;
            VerticalScrollBarControl.LargeChange = Math.Max(1, DiffGridControl.VerticalViewportSize);
            VerticalScrollBarControl.Value = DiffGridControl.FirstVisibleRow;
            VerticalScrollBarControl.IsEnabled = DiffGridControl.VerticalScrollMaximum > 0;
            ChangeMapControl.FirstVisibleRow = DiffGridControl.FirstVisibleRow;
            ChangeMapControl.VisibleRowCount = DiffGridControl.VerticalViewportSize;
        }
        finally
        {
            _updatingVerticalScrollBar = false;
        }
    }

    private void SelectCompareMode(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            viewModel.SetMode(DesktopMode.Compare);
            this.FindControl<ToggleButton>("CompareModeButton")!.IsChecked = true;
            this.FindControl<ToggleButton>("MergeModeButton")!.IsChecked = false;
        }
    }

    private void SelectMergeMode(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            viewModel.SetMode(DesktopMode.Merge);
            this.FindControl<ToggleButton>("CompareModeButton")!.IsChecked = false;
            this.FindControl<ToggleButton>("MergeModeButton")!.IsChecked = true;
        }
    }

    private async void BrowseBase(object? sender, RoutedEventArgs e) =>
        await BrowseAsync(path => ((MainWindowViewModel)DataContext!).BasePath = path);

    private async void BrowseLocal(object? sender, RoutedEventArgs e) =>
        await BrowseAsync(path => ((MainWindowViewModel)DataContext!).LocalPath = path);

    private async void BrowseRemote(object? sender, RoutedEventArgs e) =>
        await BrowseAsync(path => ((MainWindowViewModel)DataContext!).RemotePath = path);

    private async Task BrowseAsync(Action<string> apply)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            AllowMultiple = false,
            FileTypeFilter = [WorkbookFileType],
        });
        var path = files.FirstOrDefault()?.TryGetLocalPath();
        if (!string.IsNullOrEmpty(path))
        {
            apply(path);
        }
    }

    private async void SaveResult(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel viewModel || !viewModel.CanSave)
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(viewModel.SuggestedOutputPath))
        {
            if (await viewModel.SaveAsync(viewModel.SuggestedOutputPath))
            {
                Close();
            }
            return;
        }

        var extension = Path.GetExtension(viewModel.LocalPath).ToLowerInvariant();
        var files = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            SuggestedFileName = "result" + extension,
            DefaultExtension = extension.TrimStart('.'),
            FileTypeChoices = [WorkbookFileType],
        });
        var path = files?.TryGetLocalPath();
        if (!string.IsNullOrEmpty(path))
        {
            await viewModel.SaveAsync(path);
        }
    }

    private void OpenSettings(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            new SettingsWindow(viewModel).ShowDialog(this);
        }
    }

    private void OpenDiagnostics(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            new DiagnosticsWindow { DataContext = viewModel }.ShowDialog(this);
        }
    }

    private void RecentSelected(object? sender, SelectionChangedEventArgs e)
    {
        if (DataContext is MainWindowViewModel viewModel &&
            sender is ListBox { SelectedItem: RecentItemViewModel recent })
        {
            viewModel.ApplyRecent(recent);
        }
    }

    private async void CopyTsv(object? sender, RoutedEventArgs e) => await CopyAsync('\t');

    private async void CopyCsv(object? sender, RoutedEventArgs e) => await CopyAsync(',');

    private async Task CopyAsync(char delimiter)
    {
        if (DataContext is MainWindowViewModel viewModel && Clipboard is not null)
        {
            await Clipboard.SetTextAsync(viewModel.GetSelectedCopyText(delimiter));
        }
    }

    private void WindowKeyDown(object? sender, KeyEventArgs e)
    {
        if (DataContext is not MainWindowViewModel viewModel)
        {
            return;
        }

        if (e.KeyModifiers.HasFlag(KeyModifiers.Control) && e.Key == Key.F)
        {
            this.FindControl<TextBox>("SearchBox")!.Focus();
            e.Handled = true;
        }
        else if (e.KeyModifiers.HasFlag(KeyModifiers.Control) && e.Key == Key.S)
        {
            SaveResult(sender, new RoutedEventArgs());
            e.Handled = true;
        }
        else if (e.Key == Key.F3)
        {
            if (e.KeyModifiers.HasFlag(KeyModifiers.Shift))
                viewModel.PreviousSearchCommand.Execute(null);
            else
                viewModel.NextSearchCommand.Execute(null);
            e.Handled = true;
        }
        else if (e.Key == Key.F7)
        {
            if (e.KeyModifiers.HasFlag(KeyModifiers.Shift))
                viewModel.PreviousChangeCommand.Execute(null);
            else
                viewModel.NextChangeCommand.Execute(null);
            e.Handled = true;
        }
    }

    private void WindowDragOver(object? sender, DragEventArgs e)
    {
        e.DragEffects = e.DataTransfer.TryGetFiles()?.Any() == true
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    private void WindowDrop(object? sender, DragEventArgs e)
    {
        if (DataContext is not MainWindowViewModel viewModel)
        {
            return;
        }

        var paths = (e.DataTransfer.TryGetFiles() ?? [])
            .Select(static file => file.TryGetLocalPath())
            .Where(static path => !string.IsNullOrEmpty(path))
            .Cast<string>()
            .Take(3)
            .ToArray();
        if (paths.Length == 2)
        {
            viewModel.Mode = DesktopMode.Compare;
            viewModel.LocalPath = paths[0];
            viewModel.RemotePath = paths[1];
        }
        else if (paths.Length >= 3)
        {
            viewModel.Mode = DesktopMode.Merge;
            viewModel.BasePath = paths[0];
            viewModel.LocalPath = paths[1];
            viewModel.RemotePath = paths[2];
        }
    }

    private static FilePickerFileType WorkbookFileType { get; } = new("ExcelMerge workbooks")
    {
        Patterns = ["*.xlsx", "*.csv", "*.tsv"],
    };

}

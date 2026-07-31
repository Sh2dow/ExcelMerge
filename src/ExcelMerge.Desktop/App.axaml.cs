using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using ExcelMerge.Application;

namespace ExcelMerge.Desktop;

public sealed partial class App : Avalonia.Application
{
    private MainWindowViewModel? _viewModel;

    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
        LocalizationService.Apply("en-US");
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var statePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelMerge",
                "state.json");
            var stateStore = new JsonApplicationStateStore(statePath);
            _viewModel = new MainWindowViewModel(
                new ExcelMergeApplication(),
                stateStore,
                stateStore);
            var mainWindow = new MainWindow
            {
                DataContext = _viewModel,
            };
            mainWindow.Opened += (_, _) => _viewModel.ApplyStartupArguments(desktop.Args ?? []);
            desktop.MainWindow = mainWindow;
            desktop.Exit += (_, _) => _viewModel.DisposeAsync().AsTask().GetAwaiter().GetResult();
        }

        base.OnFrameworkInitializationCompleted();
    }
}

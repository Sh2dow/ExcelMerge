using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Styling;
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

    public static void ApplyTheme(string? theme) =>
        Current!.RequestedThemeVariant = theme switch
        {
            "Light" => ThemeVariant.Light,
            "Dark" => ThemeVariant.Dark,
            _ => ThemeVariant.Default,
        };

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
            desktop.Exit += (_, eventArgs) =>
            {
                _viewModel.DisposeAsync().AsTask().GetAwaiter().GetResult();
                if (_viewModel is
                    {
                        SuggestedOutputPath: not null,
                        SuggestedOutputSaved: false,
                    })
                {
                    eventArgs.ApplicationExitCode = 1;
                }
            };
        }

        base.OnFrameworkInitializationCompleted();
    }
}

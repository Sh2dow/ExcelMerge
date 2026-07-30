using Avalonia;

namespace ExcelMerge.Avalonia;

internal static class Program
{
    private static bool _mergeDriverCompleted;

    [STAThread]
    public static int Main(string[] args)
    {
        CommandLineOptions options;
        try
        {
            options = CommandLineOptions.Parse(args);
        }
        catch (ArgumentException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 2;
        }

        _mergeDriverCompleted = false;
        int applicationExitCode;
        try
        {
            applicationExitCode = BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
        catch (Exception ex) when (options.IsMergeDriver)
        {
            Console.Error.WriteLine(ex.Message);
            return 1;
        }
        return ResolveExitCode(options, applicationExitCode, _mergeDriverCompleted);
    }

    internal static void MarkMergeDriverCompleted() => _mergeDriverCompleted = true;

    internal static int ResolveExitCode(
        CommandLineOptions options,
        int applicationExitCode,
        bool mergeDriverCompleted)
    {
        if (!options.IsMergeDriver)
            return applicationExitCode;
        return mergeDriverCompleted ? 0 : 1;
    }

    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<App>().UsePlatformDetect().LogToTrace();
}

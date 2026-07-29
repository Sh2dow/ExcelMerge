using Avalonia;

namespace ExcelMerge.Avalonia;

internal static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        try
        {
            CommandLineOptions.Parse(args);
        }
        catch (ArgumentException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 2;
        }

        return BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<App>().UsePlatformDetect().LogToTrace();
}

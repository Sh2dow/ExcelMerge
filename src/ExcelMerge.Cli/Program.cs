namespace ExcelMerge.Cli;

public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        using var cancellation = new CancellationTokenSource();
        ConsoleCancelEventHandler handler = (_, eventArgs) =>
        {
            eventArgs.Cancel = true;
            cancellation.Cancel();
        };
        Console.CancelKeyPress += handler;
        try
        {
            return await CliRunner.RunAsync(
                args,
                Console.Out,
                Console.Error,
                cancellation.Token).ConfigureAwait(false);
        }
        finally
        {
            Console.CancelKeyPress -= handler;
        }
    }
}

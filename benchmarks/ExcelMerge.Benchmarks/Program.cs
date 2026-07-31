using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Running;

namespace ExcelMerge.Benchmarks;

internal static class Program
{
    public static async Task Main(string[] args)
    {
        if (args is ["--capacity-probe"])
        {
            await CapacityProbe.RunAsync().ConfigureAwait(false);
            return;
        }

        var config = ManualConfig.Create(DefaultConfig.Instance)
            .AddDiagnoser(MemoryDiagnoser.Default);
        if (!HasExplicitJob(args))
        {
            config.AddJob(Job.ShortRun
                .WithId("quick")
                .WithInvocationCount(1)
                .WithUnrollFactor(1));
        }

        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args, config);
    }

    private static bool HasExplicitJob(string[] args)
    {
        for (var index = 0; index < args.Length; index++)
        {
            if (args[index] is "-j" or "--job" || args[index].StartsWith("--job=", StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }
}

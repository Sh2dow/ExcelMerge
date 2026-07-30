namespace ExcelMerge.Avalonia;

public enum ApplicationMode
{
    Diff,
    Merge,
}

public sealed record CommandLineOptions(
    ApplicationMode Mode,
    string? BasePath,
    string? LocalPath,
    string? RemotePath,
    string? OutputPath = null,
    string? RepositoryPath = null,
    int? ConflictMarkerSize = null)
{
    public bool IsMergeDriver => Mode == ApplicationMode.Merge && OutputPath != null;

    public static CommandLineOptions Parse(IReadOnlyList<string> args)
    {
        if (args.Count > 0 && args[0].Equals("merge-driver", StringComparison.OrdinalIgnoreCase))
            return ParseMergeDriver(args);

        var mode = ApplicationMode.Diff;
        var index = 0;
        if (args.Count > 0 && Enum.TryParse<ApplicationMode>(args[0], true, out var parsedMode))
        {
            mode = parsedMode;
            index++;
        }

        string? basePath = null;
        string? localPath = null;
        string? remotePath = null;
        string? outputPath = null;
        string? repositoryPath = null;
        int? conflictMarkerSize = null;
        var positional = new List<string>();

        while (index < args.Count)
        {
            var argument = args[index++];
            if (!argument.StartsWith('-'))
            {
                positional.Add(argument);
                continue;
            }

            var separatorIndex = argument.IndexOf('=');
            var name = separatorIndex >= 0 ? argument[..separatorIndex] : argument;
            var normalizedName = name.ToLowerInvariant();
            if (IsLegacyBooleanOption(normalizedName))
                continue;

            var value = separatorIndex >= 0 ? argument[(separatorIndex + 1)..] : null;
            if (value == null)
            {
                if (index >= args.Count)
                    throw new ArgumentException($"Missing value for command-line option '{name}'.");
                value = args[index++];
            }

            switch (normalizedName)
            {
                case "-b":
                case "--base":
                case "--base-path":
                    basePath = value;
                    break;
                case "-l":
                case "--local":
                case "--local-path":
                case "--ours":
                case "-s":
                case "--src-path":
                    localPath = value;
                    break;
                case "-r":
                case "--remote":
                case "--remote-path":
                case "--theirs":
                case "-d":
                case "--dst-path":
                    remotePath = value;
                    break;
                case "-o":
                case "--output":
                case "--output-path":
                    outputPath = value;
                    break;
                case "--path":
                case "--repository-path":
                    repositoryPath = value;
                    break;
                case "--marker-size":
                    conflictMarkerSize = ParseMarkerSize(value);
                    break;
                case "-c":
                case "--external-cmd":
                case "-e":
                case "--empty-file-name":
                    // These legacy WPF options no longer affect the Avalonia application.
                    break;
                default:
                    throw new ArgumentException($"Unknown command-line option '{name}'.");
            }
        }

        if (positional.Count > 0)
        {
            var expectedCount = mode == ApplicationMode.Merge ? 3 : 2;
            if (positional.Count != expectedCount)
                throw new ArgumentException($"{mode} expects {expectedCount} positional file paths.");

            if (mode == ApplicationMode.Merge)
            {
                basePath ??= positional[0];
                localPath ??= positional[1];
                remotePath ??= positional[2];
            }
            else
            {
                localPath ??= positional[0];
                remotePath ??= positional[1];
            }
        }

        if (outputPath != null)
        {
            if (mode != ApplicationMode.Merge)
                throw new ArgumentException("--output is only valid in merge mode.");
            if (basePath == null || localPath == null || remotePath == null)
                throw new ArgumentException("Git merge driver mode requires BASE, OURS, and THEIRS file paths.");
        }

        return new CommandLineOptions(
            mode,
            basePath,
            localPath,
            remotePath,
            outputPath,
            repositoryPath,
            conflictMarkerSize);
    }

    private static CommandLineOptions ParseMergeDriver(IReadOnlyList<string> args)
    {
        if (args.Count != 6)
        {
            throw new ArgumentException(
                "merge-driver expects: <base> <ours-output> <theirs> <marker-size> <repository-path>.");
        }

        return new CommandLineOptions(
            ApplicationMode.Merge,
            UnquoteMergeDriverArgument(args[1]),
            UnquoteMergeDriverArgument(args[2]),
            UnquoteMergeDriverArgument(args[3]),
            UnquoteMergeDriverArgument(args[2]),
            UnquoteMergeDriverArgument(args[5]),
            ParseMarkerSize(UnquoteMergeDriverArgument(args[4])));
    }

    private static string UnquoteMergeDriverArgument(string value)
    {
        if (value.Length >= 2
            && (value[0] == '\'' && value[^1] == '\''
                || value[0] == '"' && value[^1] == '"'))
            return value[1..^1];
        return value;
    }

    private static int ParseMarkerSize(string value)
    {
        if (!int.TryParse(value, out var markerSize) || markerSize <= 0)
            throw new ArgumentException($"Invalid conflict marker size '{value}'.");
        return markerSize;
    }

    private static bool IsLegacyBooleanOption(string name)
    {
        return name is "-i" or "--immediately-execute-external-cmd"
            or "-w" or "--wait-external-cmd"
            or "-v" or "--validate-extension"
            or "-k" or "--keep-file-history";
    }
}

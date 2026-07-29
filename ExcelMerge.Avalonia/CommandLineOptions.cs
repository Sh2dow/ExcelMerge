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
    string? RemotePath)
{
    public static CommandLineOptions Parse(IReadOnlyList<string> args)
    {
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
                case "-s":
                case "--src-path":
                    localPath = value;
                    break;
                case "-r":
                case "--remote":
                case "--remote-path":
                case "-d":
                case "--dst-path":
                    remotePath = value;
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

        return new CommandLineOptions(mode, basePath, localPath, remotePath);
    }

    private static bool IsLegacyBooleanOption(string name)
    {
        return name is "-i" or "--immediately-execute-external-cmd"
            or "-w" or "--wait-external-cmd"
            or "-v" or "--validate-extension"
            or "-k" or "--keep-file-history";
    }
}

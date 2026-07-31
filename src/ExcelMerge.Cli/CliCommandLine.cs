namespace ExcelMerge.Cli;

public enum CliCommandKind
{
    Help,
    Version,
    Diff,
    Merge,
    MergeDriver,
}

public sealed record CliCommand(
    CliCommandKind Kind,
    string? BasePath = null,
    string? LocalPath = null,
    string? RemotePath = null,
    string? OutputPath = null,
    int? MarkerSize = null,
    string? RepositoryPath = null);

public sealed class CliUsageException(string message) : ArgumentException(message);

public static class CliCommandLine
{
    public const string Usage =
        "Usage:\n" +
        "  excelmerge diff LOCAL REMOTE\n" +
        "  excelmerge diff --local LOCAL --remote REMOTE\n" +
        "  excelmerge merge BASE LOCAL REMOTE --output RESULT\n" +
        "  excelmerge merge-driver BASE OURS THEIRS MARKER_SIZE REPOSITORY_PATH\n" +
        "\n" +
        "Supported formats: .xlsx, .csv, .tsv";

    public static CliCommand Parse(IReadOnlyList<string> arguments)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        if (arguments.Count == 0)
        {
            throw new CliUsageException("A command is required.");
        }

        var commandName = arguments[0].ToLowerInvariant();
        if (commandName is "-h" or "--help" or "help")
        {
            return new CliCommand(CliCommandKind.Help);
        }

        if (commandName is "-v" or "--version" or "version")
        {
            return new CliCommand(CliCommandKind.Version);
        }

        if (arguments.Count == 2 && arguments[1] is "-h" or "--help")
        {
            return new CliCommand(CliCommandKind.Help);
        }

        return commandName switch
        {
            "diff" => ParseDiff(arguments.Skip(1).ToArray()),
            "merge" => ParseMerge(arguments.Skip(1).ToArray()),
            "merge-driver" => ParseMergeDriver(arguments.Skip(1).ToArray()),
            _ => throw new CliUsageException($"Unknown command '{arguments[0]}'."),
        };
    }

    private static CliCommand ParseDiff(string[] arguments)
    {
        var parsed = ParseOptions(arguments);
        var local = TakeOption(parsed, "local", "local-path", "l", "source", "s");
        var remote = TakeOption(parsed, "remote", "remote-path", "r", "destination", "d");
        FillPositionals(parsed.Positionals, ref local, ref remote);
        EnsureNoUnknownOptions(parsed);
        if (string.IsNullOrWhiteSpace(local) || string.IsNullOrWhiteSpace(remote))
        {
            throw new CliUsageException("The diff command requires LOCAL and REMOTE paths.");
        }

        return new CliCommand(
            CliCommandKind.Diff,
            LocalPath: Unquote(local),
            RemotePath: Unquote(remote));
    }

    private static CliCommand ParseMerge(string[] arguments)
    {
        var parsed = ParseOptions(arguments);
        var basePath = TakeOption(parsed, "base", "base-path", "b");
        var local = TakeOption(parsed, "local", "local-path", "ours", "l");
        var remote = TakeOption(parsed, "remote", "remote-path", "theirs", "r");
        var output = TakeOption(parsed, "output", "output-path", "o");
        var marker = TakeOption(parsed, "marker-size");
        var repositoryPath = TakeOption(parsed, "repository-path", "path");
        FillPositionals(parsed.Positionals, ref basePath, ref local, ref remote);
        EnsureNoUnknownOptions(parsed);
        if (string.IsNullOrWhiteSpace(basePath) ||
            string.IsNullOrWhiteSpace(local) ||
            string.IsNullOrWhiteSpace(remote) ||
            string.IsNullOrWhiteSpace(output))
        {
            throw new CliUsageException(
                "The merge command requires BASE, LOCAL, REMOTE, and --output paths.");
        }

        return new CliCommand(
            CliCommandKind.Merge,
            Unquote(basePath),
            Unquote(local),
            Unquote(remote),
            Unquote(output),
            ParseMarkerSize(marker),
            string.IsNullOrWhiteSpace(repositoryPath) ? null : Unquote(repositoryPath));
    }

    private static CliCommand ParseMergeDriver(string[] arguments)
    {
        if (arguments.Length != 5)
        {
            throw new CliUsageException(
                "The merge-driver command requires BASE OURS THEIRS MARKER_SIZE REPOSITORY_PATH.");
        }

        var markerSize = ParseMarkerSize(arguments[3]) ?? throw new CliUsageException(
            "The merge-driver marker size is required.");
        return new CliCommand(
            CliCommandKind.MergeDriver,
            Unquote(arguments[0]),
            Unquote(arguments[1]),
            Unquote(arguments[2]),
            Unquote(arguments[1]),
            markerSize,
            Unquote(arguments[4]));
    }

    private static ParsedArguments ParseOptions(string[] arguments)
    {
        var options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var positionals = new List<string>();
        var endOfOptions = false;
        for (var index = 0; index < arguments.Length; index++)
        {
            var argument = arguments[index];
            if (!endOfOptions && argument == "--")
            {
                endOfOptions = true;
                continue;
            }

            if (endOfOptions || !argument.StartsWith('-') || argument == "-")
            {
                positionals.Add(argument);
                continue;
            }

            var separator = argument.IndexOf('=');
            var name = NormalizeOption(separator >= 0 ? argument[..separator] : argument);
            var value = separator >= 0
                ? argument[(separator + 1)..]
                : ++index < arguments.Length
                    ? arguments[index]
                    : throw new CliUsageException($"Option '{argument}' requires a value.");
            if (value.Length == 0 || !options.TryAdd(name, value))
            {
                throw new CliUsageException($"Option '{argument}' is empty or duplicated.");
            }
        }

        return new ParsedArguments(options, positionals);
    }

    private static string? TakeOption(ParsedArguments arguments, params string[] names)
    {
        string? result = null;
        foreach (var name in names)
        {
            if (!arguments.Options.Remove(name, out var value))
            {
                continue;
            }

            if (result is not null)
            {
                throw new CliUsageException($"More than one alias was supplied for '--{names[0]}'.");
            }

            result = value;
        }

        return result;
    }

    private static void FillPositionals(
        List<string> positionals,
        ref string? first,
        ref string? second)
    {
        var values = new Queue<string>(positionals);
        first ??= values.Count == 0 ? null : values.Dequeue();
        second ??= values.Count == 0 ? null : values.Dequeue();
        if (values.Count != 0)
        {
            throw new CliUsageException("Too many positional arguments were supplied.");
        }
    }

    private static void FillPositionals(
        List<string> positionals,
        ref string? first,
        ref string? second,
        ref string? third)
    {
        var values = new Queue<string>(positionals);
        first ??= values.Count == 0 ? null : values.Dequeue();
        second ??= values.Count == 0 ? null : values.Dequeue();
        third ??= values.Count == 0 ? null : values.Dequeue();
        if (values.Count != 0)
        {
            throw new CliUsageException("Too many positional arguments were supplied.");
        }
    }

    private static void EnsureNoUnknownOptions(ParsedArguments arguments)
    {
        if (arguments.Options.Count != 0)
        {
            throw new CliUsageException(
                $"Unknown option '--{arguments.Options.Keys.Order().First()}'.");
        }
    }

    private static int? ParseMarkerSize(string? value)
    {
        if (value is null)
        {
            return null;
        }

        if (!int.TryParse(
                Unquote(value),
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out var markerSize) || markerSize <= 0)
        {
            throw new CliUsageException("The marker size must be a positive integer.");
        }

        return markerSize;
    }

    private static string NormalizeOption(string value) => value.TrimStart('-');

    private static string Unquote(string value) =>
        value.Length >= 2 &&
        value[0] == value[^1] &&
        value[0] is '\'' or '"'
            ? value[1..^1]
            : value;

    private sealed record ParsedArguments(
        Dictionary<string, string> Options,
        List<string> Positionals);
}

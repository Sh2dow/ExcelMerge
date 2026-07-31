using ExcelMerge.Application;
using ExcelMerge.Domain;

namespace ExcelMerge.Cli;

public static class CliRunner
{
    public const int SuccessExitCode = 0;
    public const int DifferenceOrFailureExitCode = 1;
    public const int UsageExitCode = 2;

    public static async Task<int> RunAsync(
        IReadOnlyList<string> arguments,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(standardOutput);
        ArgumentNullException.ThrowIfNull(standardError);
        try
        {
            var command = CliCommandLine.Parse(arguments);
            if (command.Kind == CliCommandKind.Help)
            {
                await standardOutput.WriteLineAsync(CliCommandLine.Usage).ConfigureAwait(false);
                return SuccessExitCode;
            }

            if (command.Kind == CliCommandKind.Version)
            {
                var version = typeof(CliRunner).Assembly.GetName().Version?.ToString() ?? "0.0.0";
                await standardOutput.WriteLineAsync(version).ConfigureAwait(false);
                return SuccessExitCode;
            }

            await using var application = new ExcelMergeApplication();
            return command.Kind switch
            {
                CliCommandKind.Diff => await RunDiffAsync(
                    application,
                    command,
                    cancellationToken).ConfigureAwait(false),
                CliCommandKind.Merge => await RunMergeAsync(
                    application,
                    command,
                    cancellationToken).ConfigureAwait(false),
                CliCommandKind.MergeDriver => await RunMergeDriverAsync(
                    application,
                    command,
                    cancellationToken).ConfigureAwait(false),
                _ => throw new CliUsageException("The command is not executable."),
            };
        }
        catch (CliUsageException exception)
        {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(CliCommandLine.Usage).ConfigureAwait(false);
            return UsageExitCode;
        }
        catch (OperationCanceledException)
        {
            await standardError.WriteLineAsync("Operation cancelled.").ConfigureAwait(false);
            return DifferenceOrFailureExitCode;
        }
        catch (Exception exception)
        {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return DifferenceOrFailureExitCode;
        }
    }

    private static async Task<int> RunDiffAsync(
        ExcelMergeApplication application,
        CliCommand command,
        CancellationToken cancellationToken)
    {
        await using var session = await application.OpenCompareAsync(
            new CompareRequest(command.LocalPath!, command.RemotePath!),
            cancellationToken: cancellationToken).ConfigureAwait(false);
        return session.Sheets.Any(static sheet =>
            sheet.Difference.SheetChange.Kind != ChangeKind.Unchanged)
                ? DifferenceOrFailureExitCode
                : SuccessExitCode;
    }

    private static async Task<int> RunMergeAsync(
        ExcelMergeApplication application,
        CliCommand command,
        CancellationToken cancellationToken)
    {
        await using var session = await application.OpenMergeAsync(
            new MergeRequest(command.BasePath!, command.LocalPath!, command.RemotePath!),
            cancellationToken: cancellationToken).ConfigureAwait(false);
        if (session.HasUnresolvedConflicts)
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.UnresolvedConflicts,
                $"The merge has {session.Conflicts.Count} unresolved conflict(s).");
        }

        await session.SaveAsync(
            new SaveRequest(command.OutputPath!),
            cancellationToken: cancellationToken).ConfigureAwait(false);
        return SuccessExitCode;
    }

    private static async Task<int> RunMergeDriverAsync(
        ExcelMergeApplication application,
        CliCommand command,
        CancellationToken cancellationToken)
    {
        using var staging = GitMergeDriverStaging.Create(command);
        await using var session = await application.OpenMergeAsync(
            new MergeRequest(staging.BasePath, staging.LocalPath, staging.RemotePath),
            cancellationToken: cancellationToken).ConfigureAwait(false);
        if (session.HasUnresolvedConflicts)
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.UnresolvedConflicts,
                $"The merge has {session.Conflicts.Count} unresolved conflict(s).");
        }

        await session.SaveAsync(
            new SaveRequest(staging.ResultPath),
            cancellationToken: cancellationToken).ConfigureAwait(false);
        staging.Commit(command.OutputPath!, cancellationToken);
        return SuccessExitCode;
    }
}

internal sealed class GitMergeDriverStaging : IDisposable
{
    private GitMergeDriverStaging(string directoryPath, string extension)
    {
        DirectoryPath = directoryPath;
        BasePath = Path.Combine(directoryPath, "base" + extension);
        LocalPath = Path.Combine(directoryPath, "local" + extension);
        RemotePath = Path.Combine(directoryPath, "remote" + extension);
        ResultPath = Path.Combine(directoryPath, "result" + extension);
    }

    public string DirectoryPath { get; }

    public string BasePath { get; }

    public string LocalPath { get; }

    public string RemotePath { get; }

    public string ResultPath { get; }

    public static GitMergeDriverStaging Create(CliCommand command)
    {
        var extension = Path.GetExtension(command.RepositoryPath!).ToLowerInvariant();
        if (extension is not (".xlsx" or ".csv" or ".tsv"))
        {
            throw new ExcelMergeApplicationException(
                ApplicationError.UnsupportedFormat,
                "The repository path must identify an .xlsx, .csv, or .tsv file.");
        }

        var directory = Path.Combine(
            Path.GetTempPath(),
            "ExcelMerge.Cli",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var staging = new GitMergeDriverStaging(directory, extension);
        try
        {
            File.Copy(command.BasePath!, staging.BasePath);
            File.Copy(command.LocalPath!, staging.LocalPath);
            File.Copy(command.RemotePath!, staging.RemotePath);
            return staging;
        }
        catch
        {
            staging.Dispose();
            throw;
        }
    }

    public void Commit(string destinationPath, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var fullPath = Path.GetFullPath(destinationPath);
        var directory = Path.GetDirectoryName(fullPath) ?? throw new DirectoryNotFoundException(
            "The merge-driver output directory could not be determined.");
        var transactionPath = Path.Combine(
            directory,
            $".{Path.GetFileName(fullPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            File.Copy(ResultPath, transactionPath);
            using (var stream = new FileStream(
                transactionPath,
                FileMode.Open,
                FileAccess.Write,
                FileShare.None,
                bufferSize: 1,
                FileOptions.WriteThrough))
            {
                stream.Flush(flushToDisk: true);
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (File.Exists(fullPath))
            {
                try
                {
                    File.Replace(transactionPath, fullPath, destinationBackupFileName: null);
                }
                catch (PlatformNotSupportedException)
                {
                    File.Move(transactionPath, fullPath, overwrite: true);
                }
            }
            else
            {
                File.Move(transactionPath, fullPath);
            }
        }
        finally
        {
            try
            {
                if (File.Exists(transactionPath))
                {
                    File.Delete(transactionPath);
                }
            }
            catch
            {
            }
        }
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(DirectoryPath))
            {
                Directory.Delete(DirectoryPath, recursive: true);
            }
        }
        catch
        {
        }
    }
}

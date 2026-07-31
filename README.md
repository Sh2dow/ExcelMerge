# ExcelMerge

ExcelMerge compares and three-way merges `.xlsx`, `.csv`, and `.tsv` files. The rewrite is built on .NET 10, Avalonia 12, the Open XML SDK, and disk-backed indexes designed for multi-million-cell workbooks.

The legacy projects remain in this repository as references. New development uses `ExcelMerge.Rewrite.slnx` and the projects under `src/`.

## Features

- Two-way workbook comparison and BASE/LOCAL/REMOTE three-way merge
- Typed value, formula, style, row metadata, and structural conflict detection
- LOCAL-based transactional `.xlsx` output with atomic replacement
- REMOTE style dependency mapping, remote worksheet graph import, and BOTH row insertion
- Formula, defined-name, table, validation, comment, drawing, and worksheet-coordinate transforms
- RFC 4180 CSV and TSV reading/writing with UTF-8 and UTF-16 input support
- Virtual synchronized Desktop grid with search, change navigation, conflict resolution, copy, recent sessions, and key columns
- English, Japanese, and Simplified Chinese Desktop resources
- Headless CLI commands and a Git merge-driver contract
- Streaming worksheet/shared-string ingestion with fixed-width disk indexes and bounded caches

Legacy `.xls` workbooks are rejected without modification. Convert them to `.xlsx` before comparison or merge. Encrypted, signed, malformed, or unsupported packages fail closed.

## Requirements

- .NET 10 SDK
- A platform supported by Avalonia for framework-dependent Desktop builds
- Windows x64 for the Native AOT publish commands below

## Build And Test

```powershell
dotnet build ExcelMerge.Rewrite.slnx --configuration Release
dotnet test tests/ExcelMerge.Rewrite.Tests/ExcelMerge.Rewrite.Tests.csproj --configuration Release
```

## Desktop

Start the application:

```powershell
dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj
```

Open a comparison or merge directly:

```powershell
dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj -- diff local.xlsx remote.xlsx
dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj -- merge base.xlsx local.xlsx remote.xlsx
dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj -- merge base.xlsx local.xlsx remote.xlsx result.xlsx
```

The Desktop saves a merge only after every conflict is resolved. LOCAL is the visual and package baseline.

## CLI

```powershell
dotnet run --project src/ExcelMerge.Cli/ExcelMerge.Cli.csproj -- diff LOCAL REMOTE
dotnet run --project src/ExcelMerge.Cli/ExcelMerge.Cli.csproj -- merge BASE LOCAL REMOTE --output RESULT
dotnet run --project src/ExcelMerge.Cli/ExcelMerge.Cli.csproj -- merge-driver BASE OURS THEIRS MARKER_SIZE REPOSITORY_PATH
```

`diff` returns `0` for equal inputs and `1` for differences. `merge` writes only conflict-free automatic merges. Exit code `1` also represents unresolved conflicts, cancellation, or runtime failure; invalid command lines return `2`.

## Git Integration

Register the Desktop as a visual difftool:

```ini
[diff]
    tool = ExcelMerge

[difftool "ExcelMerge"]
    cmd = "C:/Tools/ExcelMerge/ExcelMerge.Desktop.exe" diff "$LOCAL" "$REMOTE"
```

Register the headless CLI merge driver:

```gitattributes
*.xlsx -text merge=excelmerge
*.csv text merge=excelmerge
*.tsv text merge=excelmerge
```

```powershell
git config merge.excelmerge.name "Excel workbook merge"
git config merge.excelmerge.driver '"C:/Tools/ExcelMerge/ExcelMerge.Cli.exe" merge-driver %O %A %B %L %P'
```

Git supplies BASE as `%O`, the required output/current version as `%A`, incoming content as `%B`, marker size as `%L`, and repository path as `%P`. A successful merge atomically replaces `%A`; unresolved conflicts leave it unchanged and return `1`.

Register Desktop as the conflict-resolution mergetool. The fifth argument is Git's `$MERGED` destination, so **Save result** writes directly back to the worktree file:

```powershell
git config merge.tool excelmerge
git config mergetool.excelmerge.cmd '\"C:/Tools/ExcelMerge/ExcelMerge.Desktop.exe\" merge \"$BASE\" \"$LOCAL\" \"$REMOTE\" \"$MERGED\"'
git config mergetool.excelmerge.trustExitCode true
```

After an automatic driver conflict, run `git mergetool --tool=excelmerge`. Resolving and saving returns `0`; closing without saving returns `1`, so Git leaves the file unresolved.

## Native AOT

```powershell
dotnet publish src/ExcelMerge.Cli/ExcelMerge.Cli.csproj -c Release -r win-x64 --self-contained true -p:PublishAot=true
dotnet publish src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj -c Release -r win-x64 --self-contained true -p:PublishAot=true
```

Desktop bindings are compiled so trimming and AOT do not depend on reflection binding.

## Benchmarks

The default BenchmarkDotNet run uses the 200,000-cell smoke corpus:

```powershell
dotnet run --project benchmarks/ExcelMerge.Rewrite.Benchmarks/ExcelMerge.Rewrite.Benchmarks.csproj -c Release -- --filter "*" --join
```

Select larger generated corpora with `EXCELMERGE_BENCHMARK_SCALES=standard`, `capacity`, or `all`. The capacity probe measures sampled live memory for 5,000,000 sparse cells with unique shared strings:

```powershell
$env:EXCELMERGE_BENCHMARK_SCALES = "capacity"
dotnet run --project benchmarks/ExcelMerge.Rewrite.Benchmarks/ExcelMerge.Rewrite.Benchmarks.csproj -c Release -- --filter "*OpenXmlIndexingBenchmarks*" --job Dry
dotnet run --project benchmarks/ExcelMerge.Rewrite.Benchmarks/ExcelMerge.Rewrite.Benchmarks.csproj -c Release -- --capacity-probe
```

## Project Layout

- `ExcelMerge.Domain`: immutable workbook, change, conflict, resolution, and merge-plan contracts
- `ExcelMerge.Engine`: alignment, two-way diff, and three-way merge planning
- `ExcelMerge.Storage`: workspaces and chunked row/text indexes
- `ExcelMerge.OpenXml`: streaming `.xlsx` adapter and transactional fidelity writer
- `ExcelMerge.Delimited`: CSV/TSV adapter
- `ExcelMerge.Application`: sessions, orchestration, progress, settings, and recent work
- `ExcelMerge.Desktop`: Avalonia application
- `ExcelMerge.Cli`: command-line and Git driver application
- `tests/ExcelMerge.Rewrite.Tests`: rewrite regression suite
- `benchmarks/ExcelMerge.Rewrite.Benchmarks`: generated performance corpora and benchmarks

External commands, the legacy PowerShell console, and log templates are intentionally deferred. See `docs/rewrite/status.md` for verified gates and remaining manual compatibility checks.

## License

MIT License. Copyright (c) 2017 skanmera.

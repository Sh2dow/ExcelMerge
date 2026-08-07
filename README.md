# ExcelMerge

Project source: [https://github.com/ksgfk/ExcelMerge](https://github.com/ksgfk/ExcelMerge)

ExcelMerge is a .NET desktop and command-line application for comparing and three-way merging spreadsheet files. It supports `.xlsx`, `.csv`, and `.tsv` inputs and uses disk-backed indexes plus a virtualized grid to keep workbook processing bounded.

## Features

- Compare LOCAL and REMOTE workbooks across worksheets.
- Merge BASE, LOCAL, and REMOTE workbooks, applying one-sided changes automatically.
- Resolve merge conflicts in the Desktop app with LOCAL, REMOTE, BOTH, or custom values where applicable.
- Compare typed cell values, formulas, cached formula values, display text, styles, and row metadata.
- Align rows with configurable key columns and verified row signatures.
- Search cells, hide unchanged rows, and navigate changes or conflicts in a synchronized two-pane grid.
- Preserve the LOCAL `.xlsx` package and styles as the merge baseline.
- Write results transactionally, validate generated `.xlsx` packages, and reject unsupported transformations before replacing the destination.
- Use the CLI for automation or as a Git merge driver.

## Documentation

- [AI installation guide for packaged releases](docs/AI-INSTALL.md)
- [AI/Agent Git merge-driver and conflict-resolution guide](docs/ai-git-merge-driver.md)
- [中文用户指南](docs/user-guide.zh-CN.md)

## Requirements

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- A writable temporary directory and destination directory
- A graphical environment when running the Desktop app

Microsoft Excel is not required.

## Build

The root solution uses the `.slnx` format:

```shell
dotnet build ExcelMerge.slnx --configuration Release
```

## Release Package

Create self-contained Native AOT CLI and Desktop binaries plus a ZIP archive containing the AI installation guide, AI Git guide, Chinese user guide, and package manifest:

    pwsh -File scripts/package-excelmerge.ps1

The default output is artifacts/ExcelMerge-win-x64.zip. Use -RuntimeIdentifier linux-x64 or another supported runtime identifier for a different target. When packaging for a runtime that cannot execute on the current machine, add -SkipBinaryCheck.

## Desktop App

Start the Avalonia application from the repository root:

```shell
dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj
```

Choose **Compare** for two files or **Merge** for BASE, LOCAL, and REMOTE files. All files in one operation must have the same format. You can also drop two files onto the window for Compare mode or three files for Merge mode.

The Desktop app provides:

- Worksheet, change, conflict, and search navigation
- LOCAL and REMOTE panes with a BASE/LOCAL/REMOTE/RESULT inspector
- Substring, exact, case-sensitive, and regular-expression search
- Conflict resolution and result saving
- English and Simplified Chinese interfaces with system, light, and dark themes
- Comparison settings for key columns, formulas, display text, styles, row metadata, and blank cells

Startup arguments are also supported:

```shell
dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj -- diff LOCAL REMOTE
dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj -- merge BASE LOCAL REMOTE RESULT
```

## CLI

Run CLI commands directly from source:

```shell
dotnet run --project src/ExcelMerge.Cli/ExcelMerge.Cli.csproj -- diff LOCAL REMOTE
dotnet run --project src/ExcelMerge.Cli/ExcelMerge.Cli.csproj -- diff --local LOCAL --remote REMOTE
dotnet run --project src/ExcelMerge.Cli/ExcelMerge.Cli.csproj -- merge BASE LOCAL REMOTE --output RESULT
dotnet run --project src/ExcelMerge.Cli/ExcelMerge.Cli.csproj -- --help
```

The `diff` command writes no textual report; its result is represented by the exit code. CLI merge succeeds only when the merge has no unresolved conflicts. Use the Desktop app when interactive conflict resolution is required.

| Exit code | Meaning |
| --- | --- |
| `0` | Clean diff, successful merge, help, or version output |
| `1` | Differences, unresolved conflicts, cancellation, or operational failure |
| `2` | Invalid command-line usage |

### Git Merge Driver

The merge-driver command follows Git's `%O %A %B %L %P` contract:

```text
excelmerge merge-driver %O %A %B %L %P
```

`%O` is BASE, `%A` is OURS, `%B` is THEIRS, `%L` is the marker size, and `%P` is the repository path. On success, the driver transactionally replaces OURS. A published CLI executable must be available as `excelmerge` before using this command in Git configuration.

## Format Behavior

| Format | Behavior |
| --- | --- |
| `.xlsx` | Reads standard, unencrypted Open XML workbooks. Merge output uses LOCAL as the package and visual-style baseline and is validated before commit. |
| `.csv` | Exposes one worksheet named `Data`; every field remains text. Output is UTF-8 without a BOM and uses CRLF line endings. |
| `.tsv` | Uses the same single-sheet text model as CSV, with tab delimiters. |

Legacy `.xls` files are not supported and must be converted to `.xlsx`. Merge output must use the same extension as its inputs. ExcelMerge compares stored formulas and cached values but does not calculate formulas.

Encrypted or corrupt Open XML packages are rejected. Merge also refuses digitally signed workbooks and unsupported package transformations rather than committing output with partial fidelity.

## File Safety

- Source files are fingerprinted and checked again while saving; a changed source aborts the operation.
- Writers create transaction files in the destination directory and clean them up after failure or cancellation.
- Existing destinations are preserved when validation or pre-commit checks fail.
- Replacement is atomic where the platform supports it, with an overwrite-move fallback elsewhere.

## Tests

Run the full MSTest suite directly rather than targeting the solution, which also contains the benchmark executable:

```shell
dotnet test tests/ExcelMerge.Tests/ExcelMerge.Tests.csproj --configuration Release
```

Run one test class:

```shell
dotnet test tests/ExcelMerge.Tests/ExcelMerge.Tests.csproj --configuration Release --filter "FullyQualifiedName~ExcelMerge.Tests.EngineTests"
```

Tests generate their CSV and XLSX fixtures under the system temporary directory. They require no Excel installation, display server, database, network connection, or other service.

## Benchmarks

BenchmarkDotNet uses a short `quick` job with memory diagnostics unless an explicit job is supplied:

```shell
dotnet run --project benchmarks/ExcelMerge.Benchmarks/ExcelMerge.Benchmarks.csproj --configuration Release -- --filter "*EngineBenchmarks*"
```

Open XML benchmarks default to the 200,000-cell smoke scale. Set `EXCELMERGE_BENCHMARK_SCALES` to `standard`, `capacity`, or `all` for larger runs. The separate `--capacity-probe` processes five million cells and is not intended for routine verification.

## Project Structure

| Project | Responsibility |
| --- | --- |
| `ExcelMerge.Domain` | Typed cells, workbook metadata, changes, conflicts, and resolutions |
| `ExcelMerge.Engine` | Row alignment, two-way diff, three-way merge, and merge planning |
| `ExcelMerge.Storage` | Workspace lifetime, disk-backed cell/text stores, caching, and cleanup |
| `ExcelMerge.OpenXml` | Streaming `.xlsx` reading and transactional package writing |
| `ExcelMerge.Delimited` | CSV/TSV parsing and writing |
| `ExcelMerge.Application` | Format-independent sessions and use-case orchestration |
| `ExcelMerge.Desktop` | Avalonia UI, view models, and virtual diff grid |
| `ExcelMerge.Cli` | Command-line and Git merge-driver entrypoints |

`Domain` is the dependency root. Format-specific and UI framework types must not leak into `Domain` or `Engine`.

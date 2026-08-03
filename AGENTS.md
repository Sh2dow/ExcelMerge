# ExcelMerge Agent Notes

## Build And Test

- Use the .NET 10 SDK; the root solution is `ExcelMerge.slnx`, not a `.sln` file.
- Build everything with `dotnet build ExcelMerge.slnx --configuration Release`.
- Run the suite directly with `dotnet test tests/ExcelMerge.Tests/ExcelMerge.Tests.csproj --configuration Release`; the solution also contains a BenchmarkDotNet executable.
- Run one test class with `dotnet test tests/ExcelMerge.Tests/ExcelMerge.Tests.csproj --configuration Release --filter "FullyQualifiedName~ExcelMerge.Tests.EngineTests"`.
- Run one method with `dotnet test tests/ExcelMerge.Tests/ExcelMerge.Tests.csproj --configuration Release --filter "FullyQualifiedName=ExcelMerge.Tests.EngineTests.Two_way_diff_reports_cell_changes_and_compact_column_runs"`.
- Tests generate CSV/XLSX fixtures under `%TEMP%\ExcelMerge.Tests`; they require writable temporary disk space but no Excel installation, display server, database, network, or other service.
- Run the CLI with `dotnet run --project src/ExcelMerge.Cli/ExcelMerge.Cli.csproj -- <args>` and the Avalonia app with `dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj`.

## Project Boundaries

- `Domain` is the dependency root. `Engine` and `Storage` depend on it; `OpenXml`/`Delimited` depend on `Domain` plus `Storage`; `Application` composes all backends; CLI/Desktop are the outer shells. Keep Avalonia and Open XML SDK types out of `Domain` and `Engine`.
- `ExcelMergeApplication.OpenCompareAsync` and `OpenMergeAsync` create workspace-backed sessions. A session owns its snapshots and disk stores; keep it alive while reading them and dispose it before replacing or abandoning a document.
- `.xlsx` output copies LOCAL as the package/style baseline, patches selected parts, validates, then atomically replaces the destination. Unsupported transformations must fail before commit.
- CSV/TSV inputs become one worksheet named `Data`, and every field remains text. All inputs in an operation must share a format; merge output must use that format's extension.
- Keep desktop behavior in view models/services. Code-behind is reserved for rendering, hit testing, native drag/drop, and window integration.

## Behavioral Invariants

- Internal row/column coordinates are zero-based; Open XML cell and row references are one-based.
- Snapshot rows and cells are sparse but strictly ordered, unique, and nonnegative. An absent cell is not automatically the same as a present blank cell.
- Comparisons use typed values. Formula text, typed cached values, and optional display text are distinct comparison inputs.
- `ThreeWayMergeResult.RowMappings` is the complete writer topology. `ViewRows` is presentation-filtered and must never drive output generation.
- `BOTH` is a row-level resolution only. Every conflict needs exactly one matching row- or cell-scope resolution, and unresolved plans must not reach a writer.
- Save paths recheck source fingerprints and available disk space. Writers use same-directory transaction files; preserve cleanup and atomic-replacement behavior on failure or cancellation.
- CLI exit code `1` means either differences or failure; usage errors return `2`. A clean diff and successful merge return `0`.

## Known Documentation Drift

- Treat executable source as authoritative over `docs/architecture.md`: worksheet selection currently happens after all workbook sheets are indexed.
- `ChunkedCellStore` uses append-only chunk files plus a fixed-width disk offset index; it is not memory-mapped.
- The Desktop path has no explicit worker scheduler or `Task.Run` boundary, so the documented claim that parsing/diffing/saving never runs on the UI thread is not currently guaranteed.

## Performance Work

- Benchmarks default to a custom short `quick` job with memory diagnostics: `dotnet run --project benchmarks/ExcelMerge.Benchmarks/ExcelMerge.Benchmarks.csproj --configuration Release -- --filter "*EngineBenchmarks*"`.
- `EXCELMERGE_BENCHMARK_SCALES` accepts `smoke`, `standard`/`1m`, `capacity`/`5m`, or `all` for Open XML benchmarks. `--capacity-probe` is a separate five-million-cell run; do not use it as routine verification.

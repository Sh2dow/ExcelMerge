# Rewrite Status

Last updated: 2026-07-31

## Verified Baseline

- `ExcelMerge.Rewrite.slnx` builds on .NET SDK 10.0.302 without compiler warnings.
- `tests/ExcelMerge.Rewrite.Tests` contains 74 passing tests.
- Desktop and CLI publish warning-free as self-contained `win-x64` Native AOT executables.
- Native CLI `--help` and native Desktop startup/responding-window smoke checks pass.
- The 200,000-cell BenchmarkDotNet suite executes 13 storage, Engine, reader, viewport, and writer cases.
- The 5,000,000-cell capacity corpus completes for dense/sparse and repeated/unique shared strings.
- Worst measured reader case: sparse unique shared strings, 54.25 seconds, 12.6 MB sampled managed peak, and 68.8 MB working-set peak.
- Validated capacity writer: 47.47 seconds, 1.06 GB sampled managed peak, and 946 MB working-set peak.

Measurements were recorded on Windows 10, an Intel Core i7-11700F, 16 logical cores, NVMe storage, and .NET 10.0.10. Short-run figures are baselines, not fixed release limits.

## Implemented

- `ExcelMerge.Domain`: typed cells, workbook/sheet topology, changes, automatic decisions, conflicts, cell/row resolutions, row overrides, and immutable merge plans.
- `ExcelMerge.Engine`: key-column signatures, patience/bounded alignment, exact hash verification, two-way differences, independent BASE alignment, structural conflicts, and complete plan assembly.
- `ExcelMerge.Storage`: owned/leased workspaces, stale cleanup, disk preflight, CRC row chunks, fixed-width row and UTF-8 text indexes, chunk rotation, and bounded LRU caches.
- `ExcelMerge.OpenXml` reader: package preflight, metadata-only discovery, disk-backed shared strings, styles, dates, typed cells, shared formulas, sparse rows, cancellation, and source-change detection.
- `ExcelMerge.OpenXml` writer: LOCAL-copy transaction, source verification, remote styles/worksheets, BOTH, coordinate/formula/name transforms, table remapping, unsupported dependency rejection, recalculation, bounded large-part validation, and atomic commit.
- `ExcelMerge.Delimited`: RFC 4180 CSV and TSV parsing, BOM/encoding handling, limits, typed snapshots, and transactional UTF-8 output.
- `ExcelMerge.Application`: format dispatch, compare/merge sessions, sheet pairing, progress, cancellation, workspace ownership, resolution mutation, save orchestration, recent descriptors, and atomic JSON settings.
- `ExcelMerge.Desktop`: Avalonia `WinExe`, direct startup arguments, virtual synchronized grid, bounded viewport/text caches, change map, search, navigation, inspection, conflict resolution, save, settings, recent sessions, diagnostics, accessibility name, and three locales.
- `ExcelMerge.Cli`: strict diff/merge/merge-driver parsing, aliases, extensionless Git staging, transactional `%A` replacement, cancellation, and `0/1/2` exit contracts.
- Benchmarks: 200K/1M/5M streaming corpora, metadata, dense/sparse and repeated/unique indexing, viewport fetch, storage, alignment, two-way, three-way, and validated package writing.

## Acceptance Coverage

- Automated tests cover reader failures, writer atomicity/fidelity, remote rich strings and styles, row insertion transforms, package relationships, formulas, names, delimited adapters, Application sessions/state, CLI parsing, and Git staging.
- Shared strings no longer retain complete managed arrays or SDK DOM collections; reader text and writer OOXML fragments use chunked fixed-width disk indexes.
- Large worksheet/shared-string validation streams individual schema elements and validates a stripped shadow package, keeping SDK validation below the 2 GB target.
- `.xls`, encrypted packages, signed packages, malformed relationships, unsupported metadata/pivot/query/external/threaded dependencies, and unsafe formula structures fail closed.

## Remaining Manual Checks

- Open the Windows fidelity fixture set in supported Microsoft Excel versions and confirm no repair prompt.
- Open the cross-platform fixture set in LibreOffice.
- Exercise Desktop keyboard, screen-reader, high-DPI, Japanese, and Simplified Chinese workflows on target systems.
- Establish repeatable CI hardware before enforcing the documented 10% throughput and 5% allocation regression thresholds.

External commands, the PowerShell console, and configurable log templates remain deferred as stated in `current-capabilities.md`.

## Next Milestone

Produce a release-candidate fixture bundle, complete the manual Excel/LibreOffice/accessibility matrix, and record signed release artifacts. No legacy implementation is required by the rewrite solution.

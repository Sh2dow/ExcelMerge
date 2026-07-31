# Rewrite Architecture

## Principles

- The rewrite does not reference the legacy core, NPOI, NetDiff, FastWpfGrid, or the WPF application.
- Domain and merge behavior do not depend on Avalonia or Open XML SDK types.
- Workbook parsing, indexing, diffing, and saving never run on the UI thread.
- Large worksheets are streamed into bounded, disk-backed indexes rather than materialized as UI objects.
- Unsupported transformations fail before output is committed.
- LOCAL is the package and visual-style baseline for merge output.

## Project Boundaries

| Project | Responsibility |
| --- | --- |
| `ExcelMerge.Domain` | Typed cells, coordinates, workbook metadata, changes, conflicts, and resolutions |
| `ExcelMerge.Engine` | Row/column alignment, two-way diff, three-way merge, and conflict planning |
| `ExcelMerge.OpenXml` | Streaming `.xlsx` parsing, package graph handling, style mapping, and transactional writing |
| `ExcelMerge.Storage` | Chunk files, row offsets, memory-mapped indexes, cache lifetime, and cleanup |
| `ExcelMerge.Application` | Sessions, use cases, progress, cancellation, settings, and recent work |
| `ExcelMerge.Desktop` | Avalonia shell, view models, virtual diff grid, navigation, inspector, and dialogs |
| `ExcelMerge.Cli` | Command-line diff/merge and Git driver contracts |

## Data Flow

1. Read workbook metadata and package relationships without loading worksheet XML.
2. Stream selected worksheets into a temporary indexed cell store.
3. Publish rows to the viewport as soon as the first chunks are indexed.
4. Calculate stable row signatures and exact equality data in the background.
5. Produce compact row mappings and change runs instead of one observable object per cell.
6. Store conflict resolutions independently from presentation state.
7. Build an immutable merge plan.
8. Copy the LOCAL package and stream-transform only affected package parts into a temporary output.
9. Validate package structure and relationships before atomically replacing the destination.

## Diff and Merge Rules

- Formula cells compare formula text and typed cached values separately.
- Literal cells compare type and normalized raw value, not culture-dependent display strings.
- User-selected key columns take priority for row identity.
- Unique row signatures provide fallback anchors.
- Patience alignment partitions large inputs; bounded Myers alignment handles small ambiguous gaps.
- Hash matches are always verified against exact cell sequences.
- Each side is aligned independently to BASE before changes are merged.
- Structural conflicts include add/add, delete/edit, incompatible row insertion, worksheet delete/edit, and ambiguous rename.
- BOTH is a row-level operation and is never represented as an isolated cell value.

## UI Composition

- `AppShellWindow`: application commands and comparison documents.
- `ComparisonWorkspaceView`: the complete comparison workflow.
- `WorkbookNavigatorView`: worksheet pairing, filters, and conflict status.
- `VirtualDiffGrid`: one custom-drawn control with shared geometry for both sides.
- `CellInspectorView`: BASE, LOCAL, REMOTE, RESULT, formula, type, and inline text details.
- `ConflictResolverView`: persistent keyboard-driven resolution controls.
- `SettingsWindow`: language, appearance, performance, key profiles, and shortcuts.
- `DiagnosticsWindow`: operation details and privacy-safe logs.

Code-behind is limited to rendering, hit testing, native drag/drop, and window integration. Application behavior belongs in view models and services.

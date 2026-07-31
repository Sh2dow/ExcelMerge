# Current Capability Baseline

This document records the user-visible behavior that exists before the rewrite. The legacy projects are references only and are not implementation dependencies for the new application.

## Current Avalonia Application

### Inputs and modes

- Open or drop LOCAL and REMOTE files for a two-way comparison.
- Open BASE, LOCAL, and REMOTE files for a three-way merge.
- Drop one file on a specific side or drop two files together.
- Select worksheet pairs independently in diff mode.
- Select matching worksheets and scan all worksheets in merge mode.
- Start from the desktop, command line, Git difftool, or Git merge driver.

### Comparison workspace

- Show synchronized LOCAL and REMOTE tables.
- Highlight added, removed, and modified cells.
- Synchronize scrolling, selection, column widths, and row heights.
- Hide unchanged rows.
- Navigate to the previous or next changed row.
- Select worksheets from a workbook overview.
- Use a draggable sheet map to inspect and navigate change locations.
- Change table font size, resize rows, and collapse the file input area.

### Merge workflow

- Detect value conflicts against BASE across all worksheets.
- Auto-merge changes made on only one side.
- Navigate conflicts across cells and worksheets.
- Track unresolved and resolved conflict counts.
- Resolve one cell or a complete row with LOCAL, REMOTE, BOTH, or a custom value.
- Preview character-level changes and the proposed result.
- Save RESULT after all conflicts are resolved.
- Return Git merge-driver exit codes that distinguish success, cancellation, and invalid arguments.

### Current output behavior

- Save `.xlsx` or `.xls` output.
- Use LOCAL as the preservation baseline where possible.
- Preserve formulas, selected styles, row heights, column widths, merged ranges, and untouched package parts in supported `.xlsx` paths.
- Duplicate a row for BOTH.

The current implementation does not reliably preserve complex formatting when BOTH forces the NPOI rewrite path. This behavior is a defect, not a compatibility promise.

## Legacy WPF Reference

The excluded WPF project also contains these user-facing features:

- Search with exact, case-sensitive, and regular-expression modes.
- Previous/next navigation for changed cells, changed rows, added rows, removed rows, and search results.
- Swap comparison sides.
- Copy selected cells as TSV or CSV.
- Configure row headers, column headers, and file-specific matching profiles.
- Store recent files and recent file pairs.
- Configure blank-row/column trimming, grid metrics, colors, and fonts.
- English, Japanese, and Simplified Chinese resources.
- Configurable external commands, log templates, and a PowerShell console.

The rewrite restores search, navigation, copy, recent sessions, header/key profiles, keyboard shortcuts, and localization in its first release. External commands, the PowerShell console, and log templates are deferred.

## First-Release Format Scope

- `.xlsx`: compare, three-way merge, and fidelity-preserving save.
- `.csv` and `.tsv`: compare and content merge through a delimited-text adapter.
- `.xls`: explicitly rejected with a localized explanation.
- Encrypted, digitally signed, or otherwise unsupported packages: detected before modification and rejected without producing an output file.

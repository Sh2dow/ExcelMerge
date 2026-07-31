# OOXML Fidelity and Compatibility

## Output Strategy

The writer starts from the LOCAL `.xlsx` package. It copies package parts without changing their uncompressed content unless a merge plan requires a modification.

Cell value selection and visual style selection are separate concerns:

- Existing destination cells retain LOCAL style unless the merge explicitly changes formatting.
- REMOTE-only cells import the required style dependency graph.
- BOTH keeps the LOCAL row first and inserts a REMOTE row immediately below it.
- Custom values retain the LOCAL style and are written as typed values when the requested type is known.

## Style Mapping

REMOTE cell styles are canonicalized and mapped into the LOCAL style table without renumbering existing LOCAL records. Mapping includes:

- Number formats
- Fonts
- Fills
- Borders
- Cell style XFs
- Cell XFs
- Alignment and protection data
- Row and column style references

Theme-based colors and fonts are resolved against the REMOTE theme before import when retaining the theme reference would change their appearance under the LOCAL theme.

## Row Insertion Transform

BOTH applies a coordinate transform to affected worksheet and workbook parts. The transform updates, where present:

- Row and cell addresses
- Worksheet dimensions
- A1 formula references
- Merged ranges
- Conditional-formatting and data-validation ranges
- Hyperlinks, comments, and VML anchors
- Drawing anchors
- Auto-filter and table ranges
- Print areas, print titles, and defined names
- Pane locations and row breaks

Calculation chains are removed after formula-affecting edits and the workbook is marked for full recalculation.

## Remote-Only Worksheets

A remote-only worksheet is imported as a package relationship graph rather than reconstructed cell by cell. Related tables, drawings, comments, hyperlinks, and worksheet relationships are copied with new relationship and part identifiers. Style and shared-string references are remapped into the LOCAL package.

## Fail-Closed Cases

The writer must reject the operation before committing output when it cannot safely transform a package. Initial fail-closed cases include:

- Encrypted packages
- Digitally signed packages
- Unsupported formula structures intersecting an inserted row
- Invalid or cyclic package relationships
- Unsupported cell metadata references intersecting a changed cell
- Insufficient temporary disk space
- Source files changing during the merge session

## Acceptance

- `OpenXmlValidator` reports no errors for every generated fixture.
- Microsoft Excel opens every Windows fidelity fixture without a repair prompt.
- LibreOffice opens every cross-platform smoke fixture successfully.
- Untouched part content hashes are identical to LOCAL.
- Changed cells, formulas, styles, row geometry, merged ranges, and relationships match the merge plan.
- Saving is atomic: cancellation or validation failure leaves the destination unchanged.

For large worksheet and shared-string parts, validation is bounded: each row or shared-string item is streamed through `OpenXmlValidator`, then the complete package graph and remaining content are validated in a shadow package whose already-streamed collections are empty. Ordinary packages continue through whole-package SDK validation directly.

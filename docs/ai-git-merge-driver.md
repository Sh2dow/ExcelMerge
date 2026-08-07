# AI/Agent Guide: Git Merge Driver and Conflict Resolution

This document is operational guidance for an AI agent or automation author integrating ExcelMerge with Git. Treat the executable source and tests as authoritative if this document and older project documentation disagree.

Project source: [https://github.com/ksgfk/ExcelMerge](https://github.com/ksgfk/ExcelMerge).

If installation, packaging, Git driver, parsing, or merge behavior is unclear, read the relevant source, tests, and project configuration from that repository. Use the source matching the packaged commit when possible; documentation is a guide, not a substitute for the current implementation.

## Scope and invariants

- Supported inputs are `.xlsx`, `.csv`, and `.tsv`. All three inputs in one operation must use the same extension/format.
- Git invokes the driver with five arguments: `%O %A %B %L %P`.
- `%O` is BASE, `%A` is OURS/LOCAL, `%B` is THEIRS/REMOTE, `%L` is Git's marker size, and `%P` is the repository-relative path.
- `merge-driver` writes the merged result back to `%A` only after a complete merge, successful validation, and a successful transactional commit.
- A non-zero driver result means Git must leave the path unresolved. Do not run `git add` until the result has been inspected or the conflict has been resolved.
- `BOTH` is a row-level decision. It is not a general cell-value choice.
- An unresolved merge must never be passed to a writer. The CLI therefore returns failure instead of producing a partial workbook.

## Build a stable CLI executable

Use a published CLI executable for Git. Do not point a long-lived Git configuration at a transient `dotnet run` build output unless the repository explicitly owns that path.

    dotnet build ExcelMerge.slnx --configuration Release
    dotnet publish src/ExcelMerge.Cli/ExcelMerge.Cli.csproj ^
      --configuration Release ^
      --output .artifacts/excelmerge-cli

Verify the published command before configuring Git. The exact filename can depend on the target platform and publish options; on Windows it is normally `ExcelMerge.Cli.exe`.

    & ./.artifacts/excelmerge-cli/ExcelMerge.Cli.exe --version

For a machine without the required .NET runtime, publish for the target runtime with `--runtime` and `--self-contained true`, then use the resulting executable. Keep the executable at a stable path and update the Git configuration whenever it moves.

## Configure the repository

Commit the file association in `.gitattributes` so every clone knows which files should use the driver:

    *.xlsx merge=excelmerge
    *.csv  merge=excelmerge
    *.tsv  merge=excelmerge

The attribute file is project configuration and belongs in source control. It does not install the executable or create the local Git driver definition; each developer or build agent still needs the local Git configuration below.

## Configure Git

The driver definition is machine-specific because it contains an executable path. Configure it in the repository with `--local` when the project has a fixed tool location, or globally with `--global` when the same executable is shared by many repositories.

The following command works in PowerShell and Git Bash when the executable path is written with forward slashes. Replace the path with the published CLI path:

    git config --local merge.excelmerge.name "ExcelMerge three-way workbook merge"
    git config --local merge.excelmerge.driver '"C:/Tools/ExcelMerge/ExcelMerge.Cli.exe" merge-driver %O %A %B %L %P'

Inspect the effective configuration and attribute resolution:

    git config --show-origin --get merge.excelmerge.driver
    git check-attr merge -- report.xlsx data.csv data.tsv

The expected attribute value is `excelmerge`. If the driver is configured globally, use `--global` in both `git config` commands. A checked-in `.gitattributes` entry without a matching local `merge.excelmerge.driver` is incomplete.

The driver command must preserve all five placeholders exactly. Do not remove `%P`: ExcelMerge uses the original repository path to determine whether the input is `.xlsx`, `.csv`, or `.tsv`. Do not hard-code BASE, LOCAL, or REMOTE paths.

## Direct driver contract

To exercise the driver without starting a Git merge, use copies of three same-format files. The command overwrites the OURS copy on success:

    & "C:/Tools/ExcelMerge/ExcelMerge.Cli.exe" merge-driver "C:/work/base.xlsx" "C:/work/ours.xlsx" "C:/work/theirs.xlsx" 7 "reports/report.xlsx"

Never use the only copy of a valuable workbook as the OURS test input. The driver stages all three inputs in a temporary workspace, saves there, and then transactionally replaces the OURS destination. Temporary staging is cleaned up after the operation.

The CLI exit codes are:

| Code | Meaning | Git/automation action |
| --- | --- | --- |
| `0` | Clean diff, successful merge, help, or version output | Accept the operation result. |
| `1` | Differences, unresolved conflicts, cancellation, or operational failure | Treat as not resolved; inspect diagnostics and Git status. |
| `2` | Invalid command-line usage | Fix the command or configuration before retrying. |

For `merge-driver`, code `1` is intentionally shared by unresolved conflicts and operational failures. Read stderr and inspect the source files before deciding whether to retry.

## Automatic merge behavior

The three-way merge compares BASE against LOCAL and REMOTE independently, then combines the changes.

- A change made only on one side is applied automatically.
- The same change on both sides is not a conflict.
- Incompatible edits to the same cell, row, worksheet, or worksheet metadata become conflicts.
- Formula text and a formula's typed cached value are separate comparison inputs. ExcelMerge does not calculate formulas.
- `.xlsx` output uses LOCAL as the package and style baseline. The generated package is validated before replacement.
- CSV and TSV inputs are modeled as one worksheet named `Data`, and every field remains text.

Do not infer a clean merge from a successful process start. A successful driver means the full result was saved and committed. A failed driver must not be followed by `git add` merely to silence Git.

## Resolve a failed Git merge

When the custom driver returns non-zero:

1. Preserve the unmerged state and run `git status --short`.
2. Identify the repository-relative path and obtain the three index stages: stage 1 is BASE, stage 2 is OURS, and stage 3 is THEIRS.
3. Extract those blobs to real files with a binary-safe Git operation. For Git Bash, the basic form is:

       git show ":1:path/to/report.xlsx" > /tmp/report-base.xlsx
       git show ":2:path/to/report.xlsx" > /tmp/report-local.xlsx
       git show ":3:path/to/report.xlsx" > /tmp/report-remote.xlsx

   For `.xlsx`, use a binary-safe file API or shell redirection appropriate to the host. Do not decode the blob as text or edit the workbook bytes directly.

4. Open the three extracted files in the Desktop app's Merge mode:

       dotnet run --project src/ExcelMerge.Desktop/ExcelMerge.Desktop.csproj -- merge BASE LOCAL REMOTE RESULT

   The final `RESULT` argument is optional in the Desktop app. When provided, it becomes the suggested save path.

5. Resolve every conflict in the conflict navigator. Use LOCAL, REMOTE, or BOTH for row-level conflicts; use a custom value only where the selected conflict is a cell conflict and the UI permits it.
6. Save the result. Validate that the result opens and that the intended sheets, rows, formulas, styles, and values are present.
7. Copy the resolved result over the working-tree path, then stage it:

       Copy-Item -LiteralPath resolved/report.xlsx -Destination path/to/report.xlsx
       git add -- path/to/report.xlsx
       git status

   Use a binary-safe copy operation after confirming the destination is the intended unmerged path.

Do not use `git checkout --ours` or `git checkout --theirs` as a substitute for opening a real three-way merge when the intent is to preserve changes from both sides. Those commands discard one side wholesale.

## Resolution semantics

For each conflict, the selected resolution must be explicit:

- **LOCAL** keeps the LOCAL value/row/worksheet choice.
- **REMOTE** keeps the REMOTE value/row/worksheet choice.
- **BOTH** keeps both rows when the conflict has both a LOCAL and a REMOTE row. It is only available for supported row-level situations.
- **Custom** creates a text cell value and is available for cell-level conflicts, not row-level conflicts.

Selecting a resolution in the presentation grid is not enough if an automation writes its own merge plan. The complete writer topology is represented by `ThreeWayMergeResult.RowMappings`; do not generate output from filtered `ViewRows`. Every conflict needs exactly one matching row- or cell-scope resolution, and an unresolved plan must be rejected.

## Verification checklist for an AI agent

Before changing a project configuration:

- Confirm the executable responds to `--version`.
- Confirm all three inputs have the same supported format.
- Confirm `.gitattributes` maps the target path to `excelmerge`.
- Confirm `git config --get merge.excelmerge.driver` contains all five placeholders.
- Confirm `%P` has the original file extension and repository-relative path.
- Use temporary copies for a direct driver test.

After a merge:

- Require driver exit code `0` before staging.
- Check Git status and the final file type/extension.
- For `.xlsx`, open or validate the package and check that the destination was not replaced after an error or cancellation.
- If exit code is `1`, preserve the three source versions, inspect stderr, and use the Desktop merge workflow for interactive resolution.
- Do not treat a normal textual conflict-marker search as sufficient for `.xlsx`; the workbook is a package and may be structurally invalid even if Git reports a file-level merge.

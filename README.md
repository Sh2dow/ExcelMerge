# ExcelMerge

ExcelMerge is a visual diff tool for Excel and delimited text files. The desktop UI is being migrated from WPF to [Avalonia](https://avaloniaui.net/) so the merge workflow can be developed on a modern, cross-platform foundation.

## Current Features

- Compare `.xlsx`, `.xls`, `.csv`, and `.tsv` files
- Select a worksheet independently on each side
- View cell changes in synchronized side-by-side grids
- Highlight added, removed, and modified cells
- Hide unchanged rows
- Navigate to the previous or next changed row
- Drop one file onto LOCAL or REMOTE, or drop two files together to open both sides
- Resolve three-way conflicts with Use LOCAL, Use REMOTE, or KEEP BOTH
- Track unresolved conflicts across worksheets and jump directly to each sheet
- Navigate conflict cells precisely across worksheets and track the remaining unresolved count
- Adjust table font size and collapse the file-input drawer to maximize the comparison area
- Generate and save a merged RESULT `.xlsx` or `.xls` workbook
- Preserve LOCAL cell styles, formulas, row heights, column widths, and unchanged workbook content when saving a merge
- Launch as a standalone application or a Git diff tool

Merge mode validates BASE, displays the LOCAL/REMOTE diff, auto-merges non-conflicting cells, and applies resolved values to a copy of the LOCAL workbook instead of rebuilding every sheet. Existing `.xlsx` sheets are patched at the OOXML cell level so untouched inline strings and package parts remain intact. REMOTE-only sheets copy their cell styles and basic sheet layout into RESULT. KEEP BOTH duplicates the conflicting row in RESULT.

## Requirements

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- Windows, macOS, or Linux supported by Avalonia

## Build

```powershell
dotnet build ExcelMerge.sln --configuration Release
```

## Run

Open the application and choose two files:

```powershell
dotnet run --project ExcelMerge.Avalonia
```

Open a comparison directly:

```powershell
dotnet run --project ExcelMerge.Avalonia -- diff -l local.xlsx -r remote.xlsx
```

Start with all three Git merge inputs:

```powershell
dotnet run --project ExcelMerge.Avalonia -- merge -b base.xlsx -l local.xlsx -r remote.xlsx
```

Long options and positional paths are also accepted:

```powershell
dotnet run --project ExcelMerge.Avalonia -- merge --base-path base.xlsx --local-path local.xlsx --remote-path remote.xlsx
dotnet run --project ExcelMerge.Avalonia -- merge base.xlsx local.xlsx remote.xlsx
```

The Git merge-driver entry point accepts Git's five standard placeholders:

```powershell
ExcelMerge.Avalonia.exe merge-driver <base> <ours-output> <theirs> <marker-size> <repository-path>
```

The equivalent named merge options are `--base`, `--ours`, `--theirs`, `--output`, `--marker-size`, and `--path`.

The previous diff aliases `-s/--src-path` and `-d/--dst-path` remain supported.

## Git Difftool

Build or publish the Avalonia application, then register its executable in `.gitconfig`:

```ini
[diff]
    tool = ExcelMerge

[difftool "ExcelMerge"]
    cmd = C:/path/to/ExcelMerge.Avalonia.exe diff -s "$LOCAL" -d "$REMOTE"
```

Run it with:

```powershell
git difftool -t ExcelMerge
```

## Git Merge Driver

Add the file types to the repository's `.gitattributes`:

```gitattributes
*.xlsx -text merge=excelmerge
*.xls -text merge=excelmerge
```

Publish the Avalonia application, then register the driver. This PowerShell command preserves the quoting required for paths containing spaces:

```powershell
git config merge.excelmerge.name "Excel workbook merge"
git config merge.excelmerge.driver '"C:/Tools/ExcelMerge/ExcelMerge.Avalonia.exe" merge-driver %O %A %B %L %P'
```

Git supplies BASE as `%O`, the current version and required output as `%A`, the incoming version as `%B`, the conflict marker size as `%L`, and the repository path as `%P`. ExcelMerge writes a successful result over `%A`. If there are no cell conflicts, review the automatic merge preview and select **Auto Merge**. Otherwise, resolve the conflicts in the UI and select **Complete Git merge**.

Merge-driver exit codes are:

- `0`: the merged workbook was written successfully
- `1`: the merge was cancelled, left unresolved, or failed at runtime
- `2`: the command-line arguments are invalid

## Project Layout

- `ExcelMerge.Avalonia`: current Avalonia desktop UI
- `ExcelMerge`: workbook reader and worksheet diff model
- `NetDiff`: sequence diff engine
- `ExcelMerge.Tests`: workbook and worksheet regression tests
- `NetDiff/NetDiff.Test`: diff engine regression tests

The legacy `ExcelMerge.GUI`, `FastWpfGrid`, `ExcelMerge.Installer`, and `ExcelMerge.ShellExtension` directories remain as migration references but are excluded from the default solution.

## License

MIT License

Copyright (c) 2017 skanmera

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

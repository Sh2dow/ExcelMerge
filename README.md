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
- Generate and save a merged RESULT `.xlsx` or `.xls` workbook
- Launch as a standalone application or a Git diff tool

Merge mode validates BASE, displays the LOCAL/REMOTE diff, auto-merges non-conflicting cells, and writes resolved cells and merged-cell ranges to a new RESULT workbook. KEEP BOTH duplicates the conflicting row in RESULT.

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

This is diff-tool integration only. RESULT can be saved from the GUI, but the executable should not be registered as a Git merge driver until a non-interactive output-path contract and merge exit codes are implemented.

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

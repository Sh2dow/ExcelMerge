# Git conflict workbook fixtures

Use these files as the three inputs for merge mode:

- `01-base.xlsx`: common ancestor (`BASE`)
- `02-local.xlsx`: current branch (`LOCAL`)
- `03-remote.xlsx`: incoming branch (`REMOTE`)

Run the application from the repository root:

```powershell
dotnet run --project ExcelMerge.Avalonia -- merge ExcelMerge.Tests/Fixtures/GitConflict/01-base.xlsx ExcelMerge.Tests/Fixtures/GitConflict/02-local.xlsx ExcelMerge.Tests/Fixtures/GitConflict/03-remote.xlsx
```

## Expected changes

### Inventory

| Row | Record | Expected behavior |
| --- | --- | --- |
| 3 | `P-100` | Conflict: LOCAL price is `629`, REMOTE price is `569`, BASE price is `599` |
| 4 | `P-200` | Auto-merge LOCAL-only stock change: `50` to `45` |
| 5 | `P-300` | Auto-merge REMOTE-only owner change: `Alice` to `Carol` |
| 6 | `P-400` | No conflict: both sides change status from `Active` to `On Hold` |
| 7 | `P-500` | Auto-merge LOCAL-only product rename |
| 8 | `P-600` | Auto-merge REMOTE-only price change |
| 9 | `P-700` | Auto-merge LOCAL-only added row |
| 10 | `P-800` | Auto-merge REMOTE-only added row |
| 11 | `P-900` | Conflict in multiple cells: LOCAL product is `Multi-Conflict Local` and price is `110`, REMOTE product is `Multi-Conflict Remote` and price is `120`, BASE product is `Multi-Conflict Item` and price is `100` |
| 12 | `P-901` | Conflict with long multiline product text for preview layout testing |

### Team

| Row | Record | Expected behavior |
| --- | --- | --- |
| 3 | `E-001` | Conflict: LOCAL role is `Tech Lead`, REMOTE role is `Architect`, BASE role is `Engineer` |
| 4 | `E-002` | Auto-merge LOCAL-only office change |
| 5 | `E-003` | Auto-merge REMOTE-only allocation change |
| 6 | `E-004` | No conflict: both sides change manager to `Grace` |

The title in each sheet is a horizontal merged-cell range. Resolving `P-100` or `E-001` with KEEP BOTH should duplicate that data row in RESULT without affecting the title range.

# Performance Budget

## Capacity Baseline

- One workbook: 500 MB compressed and 5,000,000 non-empty cells.
- Three-way merge: up to 1.5 GB compressed input and 15,000,000 non-empty cells.
- Baseline machine: 8 logical cores, 16 GB RAM, and an NVMe SSD.

## Memory and Storage

- Peak managed memory target: at most 2 GB for the three-way baseline.
- No worksheet XML part or complete workbook package may be loaded into one byte array or DOM.
- Shared strings and cells use append-only chunk stores with fixed-width offset indexes.
- Viewport data uses a bounded LRU cache.
- Temporary-space requirements are estimated before indexing and saving.
- Abandoned session data is cleaned on startup and normal shutdown.

## Responsiveness

- Application shell becomes interactive within one second on the baseline machine.
- Workbook metadata and worksheet names appear without waiting for cell indexing.
- Parsed rows are published progressively in bounded batches.
- Cancellation is observed within 200 ms at parser and engine checkpoints.
- No UI-thread operation is allowed to parse files, align rows, scan all cells, or save packages.
- Indexed-region scrolling targets 60 frames per second with a 33 ms p95 input-to-frame ceiling.

## Virtual Grid

- One control draws both panes and owns one row/column geometry model.
- Only visible cells plus bounded overscan are measured and drawn.
- Text layouts are cached by value, font, width, and culture.
- Row and column prefix indexes provide logarithmic hit testing and scroll-to-address.
- Change-map rendering uses compact runs and a cached bitmap, not one visual per marker.

## Regression Gates

- BenchmarkDotNet scenarios cover metadata load, shared strings, sparse rows, dense rows, row alignment, conflict scanning, viewport fetch, and package writing.
- Generated corpora include 200,000-cell smoke, 1,000,000-cell standard, and 5,000,000-cell capacity cases.
- Phase zero records elapsed-time baselines before hard time limits are fixed.
- Pull requests may not regress standard-case throughput by more than 10 percent or allocations by more than 5 percent without an approved architecture note.

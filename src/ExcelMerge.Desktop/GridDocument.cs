using System.Text.RegularExpressions;
using ExcelMerge.Application;
using ExcelMerge.Domain;
using ExcelMerge.Engine;

namespace ExcelMerge.Desktop;

public enum GridDocumentKind
{
    Compare,
    Merge,
}

public enum GridRowVisualState
{
    Unchanged,
    Changed,
    Conflict,
    Resolved,
}

public enum GridPane
{
    Local,
    Remote,
}

public readonly record struct GridRowDescriptor(
    int? BaseRowIndex,
    int? LocalRowIndex,
    int? RemoteRowIndex,
    GridRowVisualState State,
    int ConflictCount);

public sealed record GridLoadedRow(
    int ViewRowIndex,
    GridRowDescriptor Descriptor,
    RowRecord? BaseRow,
    RowRecord? LocalRow,
    RowRecord? RemoteRow);

public readonly record struct GridCellSelection(
    int ViewRowIndex,
    int ColumnIndex,
    GridPane Pane);

public readonly record struct GridSearchResult(
    int ViewRowIndex,
    int ColumnIndex,
    GridPane Pane,
    string Text);

public readonly record struct GridChangeRun(
    int StartRowIndex,
    int RowCount,
    GridRowVisualState State);

public sealed class GridDocument
{
    private readonly IWorksheetSnapshot? _base;
    private readonly IWorksheetSnapshot? _local;
    private readonly IWorksheetSnapshot? _remote;
    private readonly GridRowDescriptor[] _rows;

    private GridDocument(
        GridDocumentKind kind,
        string title,
        IWorksheetSnapshot? baseWorksheet,
        IWorksheetSnapshot? localWorksheet,
        IWorksheetSnapshot? remoteWorksheet,
        GridRowDescriptor[] rows)
    {
        Kind = kind;
        Title = title;
        _base = baseWorksheet;
        _local = localWorksheet;
        _remote = remoteWorksheet;
        _rows = rows;
        ChangeRuns = BuildChangeRuns(rows);
        MaxColumnIndex = new[]
        {
            baseWorksheet?.Metadata.LastColumnIndex ?? 0,
            localWorksheet?.Metadata.LastColumnIndex ?? 0,
            remoteWorksheet?.Metadata.LastColumnIndex ?? 0,
        }.Max();
    }

    public GridDocumentKind Kind { get; }

    public string Title { get; }

    public int RowCount => _rows.Length;

    public int MaxColumnIndex { get; }

    public IReadOnlyList<GridRowDescriptor> Rows => _rows;

    public IReadOnlyList<GridChangeRun> ChangeRuns { get; }

    public static GridDocument FromCompare(CompareSheetResult sheet, bool hideUnchanged = false)
    {
        var rows = sheet.Difference.ViewRows.ToArray()
            .Select(static row => new GridRowDescriptor(
                BaseRowIndex: null,
                row.BaseRowIndex,
                row.SourceRowIndex,
                row.Kind == ChangeKind.Unchanged
                    ? GridRowVisualState.Unchanged
                    : GridRowVisualState.Changed,
                ConflictCount: 0))
            .Where(row => !hideUnchanged || row.State != GridRowVisualState.Unchanged)
            .ToArray();
        var title = sheet.LocalWorksheet?.Metadata.Name ?? sheet.RemoteWorksheet?.Metadata.Name ?? sheet.Id;
        return new GridDocument(
            GridDocumentKind.Compare,
            title,
            baseWorksheet: null,
            sheet.LocalWorksheet,
            sheet.RemoteWorksheet,
            rows);
    }

    public static GridDocument FromMerge(
        MergeSheetResult sheet,
        IReadOnlySet<long> resolvedConflicts,
        bool hideUnchanged = false)
    {
        var conflicts = sheet.Merge.Conflicts.ToArray();
        var rows = sheet.Merge.ViewRows.ToArray()
            .Select(row =>
            {
                var resolved = row.ConflictCount != 0 && conflicts
                    .AsSpan(row.ConflictStartIndex, row.ConflictCount)
                    .ToArray()
                    .All(conflict => resolvedConflicts.Contains(conflict.Id));
                var state = row.State switch
                {
                    MergeViewRowState.Unchanged => GridRowVisualState.Unchanged,
                    MergeViewRowState.Automatic => GridRowVisualState.Changed,
                    MergeViewRowState.Conflict when resolved => GridRowVisualState.Resolved,
                    MergeViewRowState.Conflict => GridRowVisualState.Conflict,
                    _ => GridRowVisualState.Unchanged,
                };
                return new GridRowDescriptor(
                    row.BaseRowIndex,
                    row.LocalRowIndex,
                    row.RemoteRowIndex,
                    state,
                    row.ConflictCount);
            })
            .Where(row => !hideUnchanged || row.State != GridRowVisualState.Unchanged)
            .ToArray();
        var title = sheet.LocalWorksheet?.Metadata.Name ??
            sheet.RemoteWorksheet?.Metadata.Name ??
            sheet.BaseWorksheet?.Metadata.Name ??
            sheet.Id;
        return new GridDocument(
            GridDocumentKind.Merge,
            title,
            sheet.BaseWorksheet,
            sheet.LocalWorksheet,
            sheet.RemoteWorksheet,
            rows);
    }

    public async ValueTask<GridLoadedRow> LoadRowAsync(
        int viewRowIndex,
        CancellationToken cancellationToken = default)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(viewRowIndex);
        if (viewRowIndex >= _rows.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(viewRowIndex));
        }

        var descriptor = _rows[viewRowIndex];
        var baseRow = await ReadOptionalAsync(_base, descriptor.BaseRowIndex, cancellationToken)
            .ConfigureAwait(false);
        var localRow = await ReadOptionalAsync(_local, descriptor.LocalRowIndex, cancellationToken)
            .ConfigureAwait(false);
        var remoteRow = await ReadOptionalAsync(_remote, descriptor.RemoteRowIndex, cancellationToken)
            .ConfigureAwait(false);
        return new GridLoadedRow(viewRowIndex, descriptor, baseRow, localRow, remoteRow);
    }

    public int FindNextChangedRow(int currentRow, int direction)
    {
        if (_rows.Length == 0)
        {
            return -1;
        }

        direction = direction < 0 ? -1 : 1;
        for (var offset = 1; offset <= _rows.Length; offset++)
        {
            var index = (currentRow + (offset * direction)) % _rows.Length;
            if (index < 0)
            {
                index += _rows.Length;
            }

            if (_rows[index].State != GridRowVisualState.Unchanged)
            {
                return index;
            }
        }

        return -1;
    }

    public int FindViewRow(int? baseRowIndex, int? localRowIndex, int? remoteRowIndex)
    {
        for (var index = 0; index < _rows.Length; index++)
        {
            var row = _rows[index];
            if (row.BaseRowIndex == baseRowIndex &&
                row.LocalRowIndex == localRowIndex &&
                row.RemoteRowIndex == remoteRowIndex)
            {
                return index;
            }
        }

        return -1;
    }

    public async ValueTask<IReadOnlyList<GridSearchResult>> SearchAsync(
        string query,
        bool exact,
        bool caseSensitive,
        bool regularExpression,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrEmpty(query))
        {
            return Array.Empty<GridSearchResult>();
        }

        Regex? regex = null;
        if (regularExpression)
        {
            regex = new Regex(
                query,
                caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase,
                TimeSpan.FromMilliseconds(250));
        }

        var comparison = caseSensitive
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;
        var results = new List<GridSearchResult>();
        for (var rowIndex = 0; rowIndex < RowCount; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = await LoadRowAsync(rowIndex, cancellationToken).ConfigureAwait(false);
            SearchRow(row.LocalRow, GridPane.Local);
            SearchRow(row.RemoteRow, GridPane.Remote);

            void SearchRow(RowRecord? record, GridPane pane)
            {
                if (!record.HasValue)
                {
                    return;
                }

                foreach (var cell in record.Value.Cells.Span)
                {
                    var text = FormatValue(cell.Value);
                    var matches = regex?.IsMatch(text) ?? (exact
                        ? string.Equals(text, query, comparison)
                        : text.Contains(query, comparison));
                    if (matches)
                    {
                        results.Add(new GridSearchResult(
                            rowIndex,
                            cell.Address.ColumnIndex,
                            pane,
                            text));
                    }
                }
            }
        }

        return results;
    }

    public static CellRecord? FindCell(RowRecord? row, int columnIndex)
    {
        if (!row.HasValue)
        {
            return null;
        }

        foreach (var cell in row.Value.Cells.Span)
        {
            if (cell.Address.ColumnIndex == columnIndex)
            {
                return cell;
            }

            if (cell.Address.ColumnIndex > columnIndex)
            {
                break;
            }
        }

        return null;
    }

    public static string FormatValue(CellValue value)
    {
        if (value.DisplayText is not null)
        {
            return value.DisplayText;
        }

        return value.IsFormula
            ? "=" + (value.Formula ?? string.Empty)
            : value.Scalar.GetValueOrDefault().ToString();
    }

    private static async ValueTask<RowRecord?> ReadOptionalAsync(
        IWorksheetSnapshot? worksheet,
        int? rowIndex,
        CancellationToken cancellationToken) =>
        worksheet is null || !rowIndex.HasValue
            ? null
            : await worksheet.GetRowAsync(rowIndex.Value, cancellationToken).ConfigureAwait(false);

    private static GridChangeRun[] BuildChangeRuns(GridRowDescriptor[] rows)
    {
        var runs = new List<GridChangeRun>();
        var start = -1;
        var state = GridRowVisualState.Unchanged;
        for (var index = 0; index <= rows.Length; index++)
        {
            var next = index < rows.Length ? rows[index].State : GridRowVisualState.Unchanged;
            if (start >= 0 && (index == rows.Length || next != state))
            {
                runs.Add(new GridChangeRun(start, index - start, state));
                start = -1;
            }

            if (index < rows.Length && next != GridRowVisualState.Unchanged && start < 0)
            {
                start = index;
                state = next;
            }
        }

        return [.. runs];
    }
}

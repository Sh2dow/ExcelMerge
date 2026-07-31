using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Threading;
using ExcelMerge.Domain;

namespace ExcelMerge.Desktop;

public sealed class VirtualDiffGrid : Control
{
    public static readonly StyledProperty<GridDocument?> DocumentProperty =
        AvaloniaProperty.Register<VirtualDiffGrid, GridDocument?>(nameof(Document));

    private static readonly IBrush HeaderBackground = new SolidColorBrush(Color.Parse("#E9ECEF"));
    private static readonly IBrush GridLine = new SolidColorBrush(Color.Parse("#D5D9DD"));
    private static readonly IBrush TextBrush = new SolidColorBrush(Color.Parse("#202428"));
    private static readonly IBrush MutedTextBrush = new SolidColorBrush(Color.Parse("#68717B"));
    private static readonly IBrush LocalHeader = new SolidColorBrush(Color.Parse("#DDEFE5"));
    private static readonly IBrush RemoteHeader = new SolidColorBrush(Color.Parse("#F2DFE4"));
    private static readonly IBrush ChangedBackground = new SolidColorBrush(Color.Parse("#FFF4C7"));
    private static readonly IBrush ConflictBackground = new SolidColorBrush(Color.Parse("#FADCE2"));
    private static readonly IBrush ResolvedBackground = new SolidColorBrush(Color.Parse("#DDEFE5"));
    private static readonly IBrush SelectedBackground = new SolidColorBrush(Color.Parse("#DCEBFA"));
    private static readonly IBrush MapBackground = new SolidColorBrush(Color.Parse("#EEF0F1"));
    private static readonly Pen GridPen = new(GridLine, 1);
    private static readonly Pen SplitPen = new(new SolidColorBrush(Color.Parse("#8A939D")), 1.5);
    private static readonly Pen SelectionPen = new(new SolidColorBrush(Color.Parse("#1769AA")), 2);
    private static readonly Pen MapViewportPen = new(new SolidColorBrush(Color.Parse("#444C54")), 1);
    private readonly Dictionary<int, GridLoadedRow> _cache = [];
    private readonly LinkedList<int> _lru = [];
    private readonly HashSet<int> _pending = [];
    private readonly Dictionary<TextLayoutKey, TextLayoutEntry> _textCache = [];
    private readonly LinkedList<TextLayoutKey> _textLru = [];
    private CancellationTokenSource _loadCancellation = new();
    private int _firstRow;
    private int _firstColumn;
    private GridCellSelection? _selection;
    private int _generation;
    private bool _mapDragging;

    static VirtualDiffGrid()
    {
        AffectsRender<VirtualDiffGrid>(DocumentProperty);
    }

    public VirtualDiffGrid()
    {
        Focusable = true;
        ClipToBounds = true;
        GridSettings.Changed += GridSettingsChanged;
    }

    public event EventHandler<GridCellSelection>? CellSelected;

    public GridDocument? Document
    {
        get => GetValue(DocumentProperty);
        set => SetValue(DocumentProperty, value);
    }

    public int FirstVisibleRow => _firstRow;

    public override void Render(DrawingContext context)
    {
        base.Render(context);
        var bounds = new Rect(Bounds.Size);
        context.DrawRectangle(Brushes.White, null, bounds);
        if (Document is null || bounds.Width < 280 || bounds.Height < 80)
        {
            DrawText(
                context,
                LocalizationService.Get("NoSession"),
                new Point(18, 18),
                MutedTextBrush,
                13);
            return;
        }

        const double headerHeight = 30;
        const double rowHeaderWidth = 48;
        var paneWidth = bounds.Width / 2;
        var contentHeight = bounds.Height - headerHeight;
        var visibleRows = Math.Max(1, (int)Math.Ceiling(contentHeight / GridSettings.RowHeight));
        var visibleColumns = Math.Max(
            1,
            (int)Math.Ceiling((paneWidth - rowHeaderWidth) / GridSettings.ColumnWidth));
        ClampOffsets(visibleRows, visibleColumns);
        EnsureRowsLoaded(_firstRow, visibleRows + 2);

        DrawPaneHeader(context, 0, paneWidth, rowHeaderWidth, headerHeight, GridPane.Local);
        DrawPaneHeader(context, paneWidth, paneWidth, rowHeaderWidth, headerHeight, GridPane.Remote);
        context.DrawLine(SplitPen, new Point(paneWidth, 0), new Point(paneWidth, bounds.Height));

        for (var visibleRow = 0; visibleRow < visibleRows; visibleRow++)
        {
            var viewRow = _firstRow + visibleRow;
            if (viewRow >= Document.RowCount)
            {
                break;
            }

            var y = headerHeight + (visibleRow * GridSettings.RowHeight);
            var descriptor = Document.Rows[viewRow];
            var background = descriptor.State switch
            {
                GridRowVisualState.Changed => ChangedBackground,
                GridRowVisualState.Conflict => ConflictBackground,
                GridRowVisualState.Resolved => ResolvedBackground,
                _ => Brushes.White,
            };
            context.DrawRectangle(background, null, new Rect(0, y, bounds.Width, GridSettings.RowHeight));
            DrawRowHeader(context, 0, y, rowHeaderWidth, viewRow, descriptor.LocalRowIndex);
            DrawRowHeader(context, paneWidth, y, rowHeaderWidth, viewRow, descriptor.RemoteRowIndex);
            _cache.TryGetValue(viewRow, out var loaded);
            DrawCells(
                context,
                loaded?.LocalRow,
                GridPane.Local,
                0,
                paneWidth,
                rowHeaderWidth,
                y,
                visibleColumns,
                viewRow);
            DrawCells(
                context,
                loaded?.RemoteRow,
                GridPane.Remote,
                paneWidth,
                paneWidth,
                rowHeaderWidth,
                y,
                visibleColumns,
                viewRow);
            context.DrawLine(
                GridPen,
                new Point(0, y + GridSettings.RowHeight),
                new Point(bounds.Width, y + GridSettings.RowHeight));
        }

        DrawChangeMap(context, bounds, headerHeight, visibleRows);
    }

    public void ScrollTo(int viewRowIndex, int columnIndex = 0)
    {
        if (Document is null)
        {
            return;
        }

        _firstRow = Math.Clamp(viewRowIndex, 0, Math.Max(0, Document.RowCount - 1));
        _firstColumn = Math.Clamp(columnIndex, 0, Math.Max(0, Document.MaxColumnIndex));
        _selection = new GridCellSelection(viewRowIndex, columnIndex, GridPane.Local);
        InvalidateVisual();
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == DocumentProperty)
        {
            ResetDocument();
        }
    }

    protected override void OnPointerWheelChanged(PointerWheelEventArgs e)
    {
        base.OnPointerWheelChanged(e);
        if (Document is null)
        {
            return;
        }

        if (e.KeyModifiers.HasFlag(KeyModifiers.Shift))
        {
            _firstColumn = Math.Clamp(
                _firstColumn - Math.Sign(e.Delta.Y),
                0,
                Math.Max(0, Document.MaxColumnIndex));
        }
        else
        {
            var step = e.Delta.Y == 0 ? 0 : -Math.Sign(e.Delta.Y) * 3;
            _firstRow = Math.Clamp(
                _firstRow + step,
                0,
                Math.Max(0, Document.RowCount - 1));
        }

        InvalidateVisual();
        e.Handled = true;
    }

    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        Focus();
        if (Document is null || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            return;
        }

        const double headerHeight = 30;
        const double rowHeaderWidth = 48;
        var point = e.GetPosition(this);
        if (point.X >= Bounds.Width - 14)
        {
            _mapDragging = true;
            e.Pointer.Capture(this);
            ScrollFromMap(point.Y, headerHeight);
            e.Handled = true;
            return;
        }

        if (point.Y < headerHeight)
        {
            return;
        }

        var paneWidth = Bounds.Width / 2;
        var pane = point.X < paneWidth ? GridPane.Local : GridPane.Remote;
        var paneX = pane == GridPane.Local ? point.X : point.X - paneWidth;
        if (paneX < rowHeaderWidth)
        {
            paneX = rowHeaderWidth;
        }

        var row = _firstRow + (int)((point.Y - headerHeight) / GridSettings.RowHeight);
        var column = _firstColumn + (int)((paneX - rowHeaderWidth) / GridSettings.ColumnWidth);
        if (row < 0 || row >= Document.RowCount || column < 0)
        {
            return;
        }

        _selection = new GridCellSelection(row, column, pane);
        CellSelected?.Invoke(this, _selection.Value);
        InvalidateVisual();
        e.Handled = true;
    }

    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (_mapDragging)
        {
            ScrollFromMap(e.GetPosition(this).Y, 30);
            e.Handled = true;
        }
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (_mapDragging)
        {
            _mapDragging = false;
            e.Pointer.Capture(null);
            e.Handled = true;
        }
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (Document is null)
        {
            return;
        }

        switch (e.Key)
        {
            case Key.Up:
                _firstRow = Math.Max(0, _firstRow - 1);
                break;
            case Key.Down:
                _firstRow = Math.Min(Math.Max(0, Document.RowCount - 1), _firstRow + 1);
                break;
            case Key.PageUp:
                _firstRow = Math.Max(0, _firstRow - 20);
                break;
            case Key.PageDown:
                _firstRow = Math.Min(Math.Max(0, Document.RowCount - 1), _firstRow + 20);
                break;
            case Key.Left:
                _firstColumn = Math.Max(0, _firstColumn - 1);
                break;
            case Key.Right:
                _firstColumn = Math.Min(Document.MaxColumnIndex, _firstColumn + 1);
                break;
            case Key.Home:
                _firstRow = 0;
                _firstColumn = 0;
                break;
            default:
                return;
        }

        InvalidateVisual();
        e.Handled = true;
    }

    private void DrawPaneHeader(
        DrawingContext context,
        double paneLeft,
        double paneWidth,
        double rowHeaderWidth,
        double headerHeight,
        GridPane pane)
    {
        context.DrawRectangle(
            pane == GridPane.Local ? LocalHeader : RemoteHeader,
            null,
            new Rect(paneLeft, 0, paneWidth, headerHeight));
        DrawText(
            context,
            pane == GridPane.Local ? LocalizationService.Get("Local") : LocalizationService.Get("Remote"),
            new Point(paneLeft + 8, 7),
            TextBrush,
            11,
            FontWeight.SemiBold);
        var columns = Math.Max(1, (int)Math.Ceiling((paneWidth - rowHeaderWidth) / GridSettings.ColumnWidth));
        for (var index = 0; index < columns; index++)
        {
            var x = paneLeft + rowHeaderWidth + (index * GridSettings.ColumnWidth);
            context.DrawLine(GridPen, new Point(x, 0), new Point(x, Bounds.Height));
            using (context.PushClip(new Rect(x + 1, 0, GridSettings.ColumnWidth - 2, headerHeight)))
            {
                DrawText(
                    context,
                    ColumnName(_firstColumn + index),
                    new Point(x + 6, 7),
                    MutedTextBrush,
                    11,
                    FontWeight.SemiBold,
                    GridSettings.ColumnWidth - 12);
            }
        }

        context.DrawLine(
            GridPen,
            new Point(paneLeft, headerHeight),
            new Point(paneLeft + paneWidth, headerHeight));
    }

    private void DrawRowHeader(
        DrawingContext context,
        double paneLeft,
        double y,
        double width,
        int viewRow,
        int? sourceRow)
    {
        context.DrawRectangle(
            HeaderBackground,
            null,
            new Rect(paneLeft, y, width, GridSettings.RowHeight));
        DrawText(
            context,
            sourceRow.HasValue ? (sourceRow.Value + 1).ToString(CultureInfo.CurrentCulture) : "-",
            new Point(paneLeft + 6, y + 5),
            MutedTextBrush,
            11);
        context.DrawLine(
            GridPen,
            new Point(paneLeft + width, y),
            new Point(paneLeft + width, y + GridSettings.RowHeight));
    }

    private void DrawCells(
        DrawingContext context,
        RowRecord? row,
        GridPane pane,
        double paneLeft,
        double paneWidth,
        double rowHeaderWidth,
        double y,
        int visibleColumns,
        int viewRow)
    {
        for (var index = 0; index < visibleColumns; index++)
        {
            var column = _firstColumn + index;
            var x = paneLeft + rowHeaderWidth + (index * GridSettings.ColumnWidth);
            var cellBounds = new Rect(x, y, GridSettings.ColumnWidth, GridSettings.RowHeight);
            if (_selection is { } selected &&
                selected.ViewRowIndex == viewRow &&
                selected.ColumnIndex == column &&
                selected.Pane == pane)
            {
                context.DrawRectangle(SelectedBackground, SelectionPen, cellBounds.Deflate(1));
            }

            var cell = GridDocument.FindCell(row, column);
            if (cell.HasValue)
            {
                using (context.PushClip(cellBounds.Deflate(2)))
                {
                    DrawText(
                        context,
                        GridDocument.FormatValue(cell.Value.Value),
                        new Point(x + 6, y + 5),
                        TextBrush,
                        11,
                        maxWidth: GridSettings.ColumnWidth - 12);
                }
            }

            context.DrawLine(
                GridPen,
                new Point(x + GridSettings.ColumnWidth, y),
                new Point(x + GridSettings.ColumnWidth, y + GridSettings.RowHeight));
        }
    }

    private void EnsureRowsLoaded(int start, int count)
    {
        var document = Document;
        if (document is null)
        {
            return;
        }

        var end = Math.Min(document.RowCount, start + count);
        for (var row = start; row < end; row++)
        {
            if (_cache.ContainsKey(row) || !_pending.Add(row))
            {
                continue;
            }

            _ = LoadRowAsync(document, row, _generation, _loadCancellation.Token);
        }
    }

    private async Task LoadRowAsync(
        GridDocument document,
        int rowIndex,
        int generation,
        CancellationToken cancellationToken)
    {
        try
        {
            var row = await document.LoadRowAsync(rowIndex, cancellationToken).ConfigureAwait(false);
            Dispatcher.UIThread.Post(() =>
            {
                _pending.Remove(rowIndex);
                if (generation != _generation || !ReferenceEquals(document, Document))
                {
                    return;
                }

                _cache[rowIndex] = row;
                Touch(rowIndex);
                TrimCache();
                InvalidateVisual();
            });
        }
        catch (OperationCanceledException)
        {
            Dispatcher.UIThread.Post(() => _pending.Remove(rowIndex));
        }
        catch
        {
            Dispatcher.UIThread.Post(() => _pending.Remove(rowIndex));
        }
    }

    private void Touch(int rowIndex)
    {
        _lru.Remove(rowIndex);
        _lru.AddFirst(rowIndex);
    }

    private void TrimCache()
    {
        while (_cache.Count > GridSettings.CacheRows && _lru.Last is { } last)
        {
            _cache.Remove(last.Value);
            _lru.RemoveLast();
        }
    }

    private void ClampOffsets(int visibleRows, int visibleColumns)
    {
        if (Document is null)
        {
            return;
        }

        _firstRow = Math.Clamp(_firstRow, 0, Math.Max(0, Document.RowCount - visibleRows));
        _firstColumn = Math.Clamp(
            _firstColumn,
            0,
            Math.Max(0, Document.MaxColumnIndex - visibleColumns + 1));
    }

    private void ResetDocument()
    {
        _loadCancellation.Cancel();
        _loadCancellation.Dispose();
        _loadCancellation = new CancellationTokenSource();
        _generation++;
        _cache.Clear();
        _lru.Clear();
        _pending.Clear();
        _firstRow = 0;
        _firstColumn = 0;
        _selection = null;
        InvalidateVisual();
    }

    private void GridSettingsChanged(object? sender, EventArgs eventArgs)
    {
        TrimCache();
        InvalidateVisual();
    }

    private void DrawChangeMap(
        DrawingContext context,
        Rect bounds,
        double headerHeight,
        int visibleRows)
    {
        var document = Document!;
        const double mapWidth = 10;
        var x = bounds.Width - mapWidth - 2;
        var height = Math.Max(1, bounds.Height - headerHeight - 2);
        context.DrawRectangle(MapBackground, GridPen, new Rect(x, headerHeight + 1, mapWidth, height));
        foreach (var run in document.ChangeRuns)
        {
            var y = headerHeight + 1 + (run.StartRowIndex / (double)Math.Max(1, document.RowCount) * height);
            var runHeight = Math.Max(2, run.RowCount / (double)Math.Max(1, document.RowCount) * height);
            var brush = run.State switch
            {
                GridRowVisualState.Conflict => ConflictBackground,
                GridRowVisualState.Resolved => ResolvedBackground,
                _ => ChangedBackground,
            };
            context.DrawRectangle(brush, null, new Rect(x + 1, y, mapWidth - 2, runHeight));
        }

        var viewportY = headerHeight + 1 + (_firstRow / (double)Math.Max(1, document.RowCount) * height);
        var viewportHeight = Math.Max(3, visibleRows / (double)Math.Max(1, document.RowCount) * height);
        context.DrawRectangle(
            null,
            MapViewportPen,
            new Rect(x, viewportY, mapWidth, Math.Min(viewportHeight, height)));
    }

    private void ScrollFromMap(double pointerY, double headerHeight)
    {
        if (Document is null)
        {
            return;
        }

        var height = Math.Max(1, Bounds.Height - headerHeight);
        var ratio = Math.Clamp((pointerY - headerHeight) / height, 0, 1);
        _firstRow = Math.Clamp(
            (int)Math.Round(ratio * Math.Max(0, Document.RowCount - 1)),
            0,
            Math.Max(0, Document.RowCount - 1));
        InvalidateVisual();
    }

    private static string ColumnName(int columnIndex)
    {
        Span<char> buffer = stackalloc char[8];
        var cursor = buffer.Length;
        var value = columnIndex + 1;
        while (value > 0)
        {
            value--;
            buffer[--cursor] = (char)('A' + (value % 26));
            value /= 26;
        }

        return buffer[cursor..].ToString();
    }

    private void DrawText(
        DrawingContext context,
        string text,
        Point origin,
        IBrush brush,
        double fontSize,
        FontWeight? weight = null,
        double maxWidth = double.PositiveInfinity)
    {
        var key = new TextLayoutKey(
            text,
            fontSize,
            weight ?? FontWeight.Normal,
            maxWidth,
            CultureInfo.CurrentCulture.Name,
            brush);
        if (!_textCache.TryGetValue(key, out var entry))
        {
            var node = _textLru.AddFirst(key);
            entry = new TextLayoutEntry(
                new FormattedText(
                    text,
                    CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    new Typeface("Segoe UI", FontStyle.Normal, key.Weight),
                    fontSize,
                    brush),
                node);
            _textCache.Add(key, entry);
            while (_textCache.Count > 4_096 && _textLru.Last is { } last)
            {
                _textCache.Remove(last.Value);
                _textLru.RemoveLast();
            }
        }
        else
        {
            _textLru.Remove(entry.Node);
            _textLru.AddFirst(entry.Node);
        }

        context.DrawText(entry.Text, origin);
    }

    private readonly record struct TextLayoutKey(
        string Text,
        double FontSize,
        FontWeight Weight,
        double MaxWidth,
        string Culture,
        IBrush Brush);

    private sealed record TextLayoutEntry(
        FormattedText Text,
        LinkedListNode<TextLayoutKey> Node);
}

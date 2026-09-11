using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using ExcelMerge.Domain;

namespace ExcelMerge.Desktop;

public sealed class VirtualDiffGrid : Control
{
    private const double HeaderHeight = 30;
    private const double RowHeaderWidth = 48;
    private const double ResizeGripSize = 4;
    private const double MinimumColumnWidth = 32;
    private const double MaximumColumnWidth = 1_000;
    private const double MinimumRowHeight = 18;
    private const double MaximumRowHeight = 400;

    public static readonly StyledProperty<GridDocument?> DocumentProperty =
        AvaloniaProperty.Register<VirtualDiffGrid, GridDocument?>(nameof(Document));

    private GridPalette _palette = GridPalette.Light;
    private readonly Dictionary<int, GridLoadedRow> _cache = [];
    private readonly LinkedList<int> _lru = [];
    private readonly HashSet<int> _pending = [];
    private readonly Dictionary<TextLayoutKey, TextLayoutEntry> _textCache = [];
    private readonly LinkedList<TextLayoutKey> _textLru = [];
    private readonly Dictionary<int, double> _columnWidths = [];
    private readonly Dictionary<int, double> _rowHeights = [];
    private readonly Dictionary<(string Old, string New), TextSpan[]> _diffSpanCache = [];
    private readonly ContentWidthFitter _autoFit = new();
    private CancellationTokenSource _loadCancellation = new();
    private int _firstRow;
    private int _firstColumn;
    private int _horizontalScrollMaximum;
    private int _horizontalViewportSize = 1;
    private int _verticalScrollMaximum;
    private int _verticalViewportSize = 1;
    private GridCellSelection? _selection;
    private int _generation;
    private GridResizeDrag? _resizeDrag;
    private Cursor? _columnResizeCursor;
    private Cursor? _rowResizeCursor;

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

    public event EventHandler? HorizontalScrollChanged;

    public event EventHandler? VerticalScrollChanged;

    public GridDocument? Document
    {
        get => GetValue(DocumentProperty);
        set => SetValue(DocumentProperty, value);
    }

    public int FirstVisibleRow => _firstRow;

    public int FirstVisibleColumn => _firstColumn;

    public int HorizontalScrollMaximum => _horizontalScrollMaximum;

    public int HorizontalViewportSize => _horizontalViewportSize;

    public int VerticalScrollMaximum => _verticalScrollMaximum;

    public int VerticalViewportSize => _verticalViewportSize;

    public Rect GetPaneBounds(GridPane pane)
    {
        var paneWidth = Math.Max(0, Bounds.Width) / 2;
        return new Rect(
            pane == GridPane.Local ? 0 : paneWidth,
            HeaderHeight,
            paneWidth,
            Math.Max(0, Bounds.Height - HeaderHeight));
    }

    public override void Render(DrawingContext context)
    {
        base.Render(context);
        var bounds = new Rect(Bounds.Size);
        context.DrawRectangle(_palette.Background, null, bounds);
        if (Document is null || bounds.Width < 280 || bounds.Height < 80)
        {
            DrawText(
                context,
                LocalizationService.Get("NoSession"),
                new Point(18, 18),
                _palette.MutedText,
                13);
            return;
        }

        var paneWidth = bounds.Width / 2;
        var contentHeight = bounds.Height - HeaderHeight;
        ClampOffsets(contentHeight, bounds.Width);
        var visibleRows = CalculateVisibleRowCount(contentHeight);
        var visibleColumns = CalculateVisibleColumnCount(bounds.Width);
        EnsureRowsLoaded(_firstRow, visibleRows + 2);

        DrawPaneHeader(context, 0, paneWidth, RowHeaderWidth, HeaderHeight, GridPane.Local);
        DrawPaneHeader(context, paneWidth, paneWidth, RowHeaderWidth, HeaderHeight, GridPane.Remote);
        context.DrawLine(_palette.SplitPen, new Point(paneWidth, 0), new Point(paneWidth, bounds.Height));

        var y = HeaderHeight;
        for (var visibleRow = 0; visibleRow < visibleRows; visibleRow++)
        {
            var viewRow = _firstRow + visibleRow;
            if (viewRow >= Document.RowCount || y >= bounds.Height)
            {
                break;
            }

            var rowHeight = GetRowHeight(viewRow);
            var descriptor = Document.Rows[viewRow];
            var background = descriptor.State == GridRowVisualState.Changed
                ? _palette.ChangedBackground
                : _palette.Background;
            context.DrawRectangle(background, null, new Rect(0, y, bounds.Width, rowHeight));
            var rowHeaderState = Document.GetRowHeaderVisualState(viewRow);
            DrawRowHeader(
                context,
                0,
                y,
                RowHeaderWidth,
                rowHeight,
                descriptor.LocalRowIndex,
                rowHeaderState);
            DrawRowHeader(
                context,
                paneWidth,
                y,
                RowHeaderWidth,
                rowHeight,
                descriptor.RemoteRowIndex,
                rowHeaderState);
            _cache.TryGetValue(viewRow, out var loaded);
            DrawCells(
                context,
                loaded,
                GridPane.Local,
                y,
                rowHeight,
                visibleColumns,
                viewRow);
            DrawCells(
                context,
                loaded,
                GridPane.Remote,
                y,
                rowHeight,
                visibleColumns,
                viewRow);
            context.DrawLine(
                _palette.GridPen,
                new Point(0, y + rowHeight),
                new Point(bounds.Width, y + rowHeight));
            y += rowHeight;
        }
    }

    public void ScrollTo(int viewRowIndex, int columnIndex = 0)
    {
        if (Document is null)
        {
            return;
        }

        SetFirstRow(viewRowIndex);
        SetFirstColumn(columnIndex);
        _selection = new GridCellSelection(viewRowIndex, columnIndex, GridPane.Local);
        InvalidateVisual();
    }

    public void ScrollHorizontalTo(double value)
    {
        if (!double.IsFinite(value))
        {
            throw new ArgumentOutOfRangeException(nameof(value));
        }

        var clamped = Math.Clamp(value, 0, _horizontalScrollMaximum);
        SetFirstColumn((int)Math.Round(clamped));
    }

    public void ScrollVerticalTo(double value)
    {
        if (!double.IsFinite(value))
        {
            throw new ArgumentOutOfRangeException(nameof(value));
        }

        var clamped = Math.Clamp(value, 0, _verticalScrollMaximum);
        SetFirstRow((int)Math.Round(clamped));
    }

    public void CenterOnMapPosition(double horizontalRatio, double verticalRatio)
    {
        if (!double.IsFinite(horizontalRatio))
        {
            throw new ArgumentOutOfRangeException(nameof(horizontalRatio));
        }

        if (!double.IsFinite(verticalRatio))
        {
            throw new ArgumentOutOfRangeException(nameof(verticalRatio));
        }

        if (Document is not { RowCount: > 0 } document)
        {
            return;
        }

        var targetColumn = (int)Math.Round(
            Math.Clamp(horizontalRatio, 0, 1) * document.MaxColumnIndex);
        var targetRow = (int)Math.Round(
            Math.Clamp(verticalRatio, 0, 1) * (document.RowCount - 1));
        SetFirstColumn(CalculateCenteredFirstColumn(targetColumn));
        SetFirstRow(CalculateCenteredFirstRow(targetRow));
    }

    public double GetColumnWidth(int columnIndex)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
        if (_columnWidths.TryGetValue(columnIndex, out var width))
        {
            return width;
        }

        var fallback = Document?.SourceDefaultColumnWidth ?? GridSettings.ColumnWidth;
        return _autoFit.GetWidth(columnIndex) is { } fit
            ? Math.Max(fallback, fit)
            : fallback;
    }

    public double GetRowHeight(int viewRowIndex)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(viewRowIndex);
        return _rowHeights.GetValueOrDefault(viewRowIndex, GridSettings.RowHeight);
    }

    public void ResizeColumn(int columnIndex, double width)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
        if (!double.IsFinite(width))
        {
            throw new ArgumentOutOfRangeException(nameof(width));
        }

        _columnWidths[columnIndex] = Math.Clamp(width, MinimumColumnWidth, MaximumColumnWidth);
        UpdateHorizontalScrollMetrics(forceNotification: true);
        InvalidateVisual();
    }

    public void ResizeRow(int viewRowIndex, double height)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(viewRowIndex);
        if (!double.IsFinite(height))
        {
            throw new ArgumentOutOfRangeException(nameof(height));
        }

        _rowHeights[viewRowIndex] = Math.Clamp(height, MinimumRowHeight, MaximumRowHeight);
        UpdateVerticalScrollMetrics(forceNotification: true);
        InvalidateVisual();
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == DocumentProperty)
        {
            RemapCustomSizes(
                change.GetOldValue<GridDocument?>(),
                change.GetNewValue<GridDocument?>());
            ResetDocument();
        }
        else if (change.Property == BoundsProperty)
        {
            UpdateHorizontalScrollMetrics();
            UpdateVerticalScrollMetrics();
        }
    }

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        ActualThemeVariantChanged += OnActualThemeVariantChanged;
        UpdatePalette();
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        ActualThemeVariantChanged -= OnActualThemeVariantChanged;
        base.OnDetachedFromVisualTree(e);
    }

    private void OnActualThemeVariantChanged(object? sender, EventArgs e) => UpdatePalette();

    private void UpdatePalette()
    {
        var palette = GridPalette.ForVariant(ActualThemeVariant);
        if (!ReferenceEquals(palette, _palette))
        {
            _palette = palette;
            InvalidateVisual();
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
            SetFirstColumn(_firstColumn - Math.Sign(e.Delta.Y));
        }
        else
        {
            var step = e.Delta.Y == 0 ? 0 : -Math.Sign(e.Delta.Y) * 3;
            SetFirstRow(_firstRow + step);
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

        var point = e.GetPosition(this);
        if (TryHitColumnResize(point, out var columnToResize))
        {
            BeginResize(e.Pointer, GridResizeKind.Column, columnToResize, point.X);
            e.Handled = true;
            return;
        }

        if (TryHitRowResize(point, out var rowToResize))
        {
            BeginResize(e.Pointer, GridResizeKind.Row, rowToResize, point.Y);
            e.Handled = true;
            return;
        }

        if (point.Y < HeaderHeight)
        {
            return;
        }

        var paneWidth = Bounds.Width / 2;
        var pane = point.X < paneWidth ? GridPane.Local : GridPane.Remote;
        var paneX = pane == GridPane.Local ? point.X : point.X - paneWidth;
        if (paneX < RowHeaderWidth)
        {
            paneX = RowHeaderWidth;
        }

        var row = HitTestRow(point.Y);
        var column = HitTestColumn(paneX);
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
        var point = e.GetPosition(this);
        if (_resizeDrag is { } resize)
        {
            var position = resize.Kind == GridResizeKind.Column ? point.X : point.Y;
            var size = resize.InitialSize + position - resize.InitialPointerPosition;
            if (resize.Kind == GridResizeKind.Column)
            {
                ResizeColumn(resize.Index, size);
            }
            else
            {
                ResizeRow(resize.Index, size);
            }

            e.Handled = true;
            return;
        }

        UpdateResizeCursor(point);
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (_resizeDrag is not null)
        {
            _resizeDrag = null;
            e.Pointer.Capture(null);
            UpdateResizeCursor(e.GetPosition(this));
            e.Handled = true;
        }
    }

    protected override void OnPointerExited(PointerEventArgs e)
    {
        base.OnPointerExited(e);
        if (_resizeDrag is null)
        {
            Cursor = null;
        }
    }

    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
    {
        base.OnPointerCaptureLost(e);
        _resizeDrag = null;
        Cursor = null;
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
                SetFirstRow(_firstRow - 1);
                break;
            case Key.Down:
                SetFirstRow(_firstRow + 1);
                break;
            case Key.PageUp:
                SetFirstRow(_firstRow - Math.Max(1, _verticalViewportSize));
                break;
            case Key.PageDown:
                SetFirstRow(_firstRow + Math.Max(1, _verticalViewportSize));
                break;
            case Key.Left:
                SetFirstColumn(_firstColumn - 1);
                break;
            case Key.Right:
                SetFirstColumn(_firstColumn + 1);
                break;
            case Key.Home:
                SetFirstRow(0);
                SetFirstColumn(0);
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
        using (context.PushClip(new Rect(paneLeft, 0, paneWidth, headerHeight)))
        {
            context.DrawRectangle(
                pane == GridPane.Local ? _palette.LocalHeader : _palette.RemoteHeader,
                null,
                new Rect(paneLeft, 0, paneWidth, headerHeight));
            DrawText(
                context,
                pane == GridPane.Local ? LocalizationService.Get("Local") : LocalizationService.Get("Remote"),
                new Point(paneLeft + 8, 7),
                _palette.Text,
                11,
                FontWeight.SemiBold);
            var columns = CalculateVisibleColumnCount(Bounds.Width);
            var x = paneLeft + rowHeaderWidth;
            for (var index = 0; index < columns; index++)
            {
                var column = _firstColumn + index;
                var columnWidth = GetColumnWidth(column);
                context.DrawLine(_palette.GridPen, new Point(x, 0), new Point(x, headerHeight));
                using (context.PushClip(new Rect(x + 1, 0, columnWidth - 2, headerHeight)))
                {
                    DrawText(
                        context,
                        ColumnName(column),
                        new Point(x + 6, 7),
                        _palette.MutedText,
                        11,
                        FontWeight.SemiBold,
                        columnWidth - 12);
                }

                x += columnWidth;
            }

            context.DrawLine(
                _palette.GridPen,
                new Point(paneLeft, headerHeight),
                new Point(paneLeft + paneWidth, headerHeight));
        }
    }

    private void DrawRowHeader(
        DrawingContext context,
        double paneLeft,
        double y,
        double width,
        double rowHeight,
        int? sourceRow,
        GridCellVisualState visualState)
    {
        var bounds = new Rect(paneLeft, y, width, rowHeight);
        using (context.PushClip(bounds))
        {
            var background = visualState switch
            {
                GridCellVisualState.Conflict => _palette.ConflictBackground,
                GridCellVisualState.Resolved => _palette.ResolvedBackground,
                _ => _palette.HeaderBackground,
            };
            context.DrawRectangle(background, null, bounds);
            DrawText(
                context,
                sourceRow.HasValue ? (sourceRow.Value + 1).ToString(CultureInfo.CurrentCulture) : "-",
                new Point(paneLeft + 6, y + 5),
                _palette.MutedText,
                11);
            context.DrawLine(
                _palette.GridPen,
                new Point(paneLeft + width, y),
                new Point(paneLeft + width, y + rowHeight));
        }
    }

    private void DrawCells(
        DrawingContext context,
        GridLoadedRow? loaded,
        GridPane pane,
        double y,
        double rowHeight,
        int visibleColumns,
        int viewRow)
    {
        var row = loaded is null
            ? null
            : pane == GridPane.Local ? loaded.LocalRow : loaded.RemoteRow;
        var paneBounds = GetPaneBounds(pane);
        using (context.PushClip(new Rect(paneBounds.X, y, paneBounds.Width, rowHeight)))
        {
            var x = paneBounds.X + RowHeaderWidth;
            for (var index = 0; index < visibleColumns; index++)
            {
                var column = _firstColumn + index;
                var columnWidth = GetColumnWidth(column);
                var cellBounds = new Rect(x, y, columnWidth, rowHeight);
                var visualState = Document!.GetCellVisualState(viewRow, column);
                var selected = _selection is { } selection &&
                    selection.ViewRowIndex == viewRow &&
                    selection.ColumnIndex == column &&
                    selection.Pane == pane;
                if (visualState != GridCellVisualState.None)
                {
                    context.DrawRectangle(
                        visualState == GridCellVisualState.Conflict
                            ? _palette.ConflictBackground
                            : _palette.ResolvedBackground,
                        visualState == GridCellVisualState.Conflict ? _palette.ConflictPen : _palette.ResolvedPen,
                        cellBounds.Deflate(1));
                }
                else if (selected)
                {
                    context.DrawRectangle(_palette.SelectedBackground, null, cellBounds.Deflate(1));
                }

                if (selected)
                {
                    context.DrawRectangle(null, _palette.SelectionPen, cellBounds.Deflate(1));
                }

                var cell = GridDocument.FindCell(row, column);
                if (cell.HasValue)
                {
                    var text = GridDocument.FormatValue(cell.Value.Value);
                    var origin = new Point(x + 6, y + 5);
                    var diff = GetDiffHighlight(loaded, pane, viewRow, column, text);
                    if (diff is not null &&
                        GetTextLayout(text, 11, null, _palette.Text).Width > columnWidth - 12)
                    {
                        diff = null;
                    }
                    using (context.PushClip(cellBounds.Deflate(2)))
                    {
                        if (diff is { } highlight)
                        {
                            DrawHighlightedText(
                                context,
                                text,
                                highlight.Spans,
                                origin,
                                _palette.Text,
                                11,
                                highlight.Brush);
                        }
                        else
                        {
                            DrawText(
                                context,
                                text,
                                origin,
                                _palette.Text,
                                11,
                                maxWidth: columnWidth - 12);
                        }
                    }
                }

                context.DrawLine(
                    _palette.GridPen,
                    new Point(x + columnWidth, y),
                    new Point(x + columnWidth, y + rowHeight));
                x += columnWidth;
            }
        }
    }

    private (TextSpan[] Spans, IBrush Brush)? GetDiffHighlight(
        GridLoadedRow? loaded,
        GridPane pane,
        int viewRow,
        int column,
        string text)
    {
        if (loaded is null || Document is null || text.Length == 0 || text.IndexOf('\n') >= 0)
        {
            return null;
        }

        RowRecord? counterpart;
        IBrush brush;
        if (Document.Kind == GridDocumentKind.Compare)
        {
            if (loaded.Descriptor.State != GridRowVisualState.Changed)
            {
                return null;
            }

            counterpart = pane == GridPane.Local ? loaded.RemoteRow : loaded.LocalRow;
            brush = _palette.ChangedHighlight;
        }
        else
        {
            if (Document.GetCellVisualState(viewRow, column) != GridCellVisualState.Conflict)
            {
                return null;
            }

            counterpart = loaded.BaseRow;
            brush = _palette.ConflictHighlight;
        }

        var counterpartCell = GridDocument.FindCell(counterpart, column);
        var counterpartText = counterpartCell.HasValue
            ? GridDocument.FormatValue(counterpartCell.Value.Value)
            : string.Empty;
        if (counterpartText == text)
        {
            return null;
        }

        var key = (counterpartText, text);
        if (!_diffSpanCache.TryGetValue(key, out var spans))
        {
            spans = TextDiffer.GetChangedSpans(counterpartText, text);
            if (_diffSpanCache.Count > 4_096)
            {
                _diffSpanCache.Clear();
            }

            _diffSpanCache[key] = spans;
        }

        return spans.Length == 0 ? null : (spans, brush);
    }

    private void DrawHighlightedText(
        DrawingContext context,
        string text,
        TextSpan[] spans,
        Point origin,
        IBrush brush,
        double fontSize,
        IBrush highlight)
    {
        var x = origin.X;
        var cursor = 0;
        foreach (var span in spans)
        {
            if (span.Start > cursor)
            {
                var before = GetTextLayout(text[cursor..span.Start], fontSize, null, brush);
                context.DrawText(before, new Point(x, origin.Y));
                x += before.Width;
            }

            var changed = GetTextLayout(text.Substring(span.Start, span.Length), fontSize, null, brush);
            context.DrawRectangle(
                highlight,
                null,
                new Rect(x, origin.Y - 2, changed.Width, changed.Height + 3));
            context.DrawText(changed, new Point(x, origin.Y));
            x += changed.Width;
            cursor = span.Start + span.Length;
        }

        if (cursor < text.Length)
        {
            context.DrawText(GetTextLayout(text[cursor..], fontSize, null, brush), new Point(x, origin.Y));
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
                if (GridDocument.GetSourceRowHeight(row) is { } sourceHeight)
                {
                    if (_rowHeights.TryAdd(rowIndex, sourceHeight))
                    {
                        UpdateVerticalScrollMetrics();
                    }
                }

                if (ApplyContentAutoFit(row))
                {
                    UpdateHorizontalScrollMetrics();
                }

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

    private bool ApplyContentAutoFit(GridLoadedRow row)
    {
        var changed = false;
        Measure(row.LocalRow);
        Measure(row.RemoteRow);
        return changed;

        void Measure(RowRecord? record)
        {
            if (!record.HasValue)
            {
                return;
            }

            foreach (var cell in record.Value.Cells.Span)
            {
                var column = cell.Address.ColumnIndex;
                if (_columnWidths.ContainsKey(column))
                {
                    continue;
                }

                var text = GridDocument.FormatValue(cell.Value);
                if (text.Length == 0)
                {
                    continue;
                }

                changed |= _autoFit.TryGrow(
                    column,
                    GetTextLayout(text, 11, null, _palette.Text).Width);
            }
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

    private void BeginResize(
        IPointer pointer,
        GridResizeKind kind,
        int index,
        double pointerPosition)
    {
        var initialSize = kind == GridResizeKind.Column
            ? GetColumnWidth(index)
            : GetRowHeight(index);
        _resizeDrag = new GridResizeDrag(kind, index, pointerPosition, initialSize);
        Cursor = GetResizeCursor(kind);
        pointer.Capture(this);
    }

    private bool TryHitColumnResize(Point point, out int columnIndex)
    {
        columnIndex = -1;
        if (point.Y < 0 || point.Y > HeaderHeight || point.X < 0 || point.X > Bounds.Width)
        {
            return false;
        }

        var paneWidth = Bounds.Width / 2;
        var paneX = point.X < paneWidth ? point.X : point.X - paneWidth;
        if (paneX < RowHeaderWidth - ResizeGripSize)
        {
            return false;
        }

        var boundary = RowHeaderWidth;
        var visibleColumns = CalculateVisibleColumnCount(Bounds.Width);
        for (var index = 0; index < visibleColumns; index++)
        {
            var column = _firstColumn + index;
            boundary += GetColumnWidth(column);
            if (Math.Abs(paneX - boundary) <= ResizeGripSize && boundary <= paneWidth + ResizeGripSize)
            {
                columnIndex = column;
                return true;
            }

            if (boundary > paneWidth + ResizeGripSize)
            {
                break;
            }
        }

        return false;
    }

    private bool TryHitRowResize(Point point, out int viewRowIndex)
    {
        viewRowIndex = -1;
        if (Document is null || point.Y < HeaderHeight || point.Y > Bounds.Height)
        {
            return false;
        }

        var paneWidth = Bounds.Width / 2;
        var paneX = point.X < paneWidth ? point.X : point.X - paneWidth;
        if (paneX < 0 || paneX > RowHeaderWidth)
        {
            return false;
        }

        var boundary = HeaderHeight;
        var visibleRows = CalculateVisibleRowCount(Math.Max(0, Bounds.Height - HeaderHeight));
        for (var index = 0; index < visibleRows; index++)
        {
            var row = _firstRow + index;
            if (row >= Document.RowCount)
            {
                break;
            }

            boundary += GetRowHeight(row);
            if (Math.Abs(point.Y - boundary) <= ResizeGripSize)
            {
                viewRowIndex = row;
                return true;
            }
        }

        return false;
    }

    private void UpdateResizeCursor(Point point)
    {
        if (TryHitColumnResize(point, out _))
        {
            Cursor = GetResizeCursor(GridResizeKind.Column);
        }
        else if (TryHitRowResize(point, out _))
        {
            Cursor = GetResizeCursor(GridResizeKind.Row);
        }
        else
        {
            Cursor = null;
        }
    }

    private Cursor GetResizeCursor(GridResizeKind kind)
    {
        if (kind == GridResizeKind.Column)
        {
            return _columnResizeCursor ??= new Cursor(StandardCursorType.SizeWestEast);
        }

        return _rowResizeCursor ??= new Cursor(StandardCursorType.SizeNorthSouth);
    }

    private int HitTestColumn(double paneX)
    {
        var offset = Math.Max(0, paneX - RowHeaderWidth);
        var column = _firstColumn;
        var visibleColumns = CalculateVisibleColumnCount(Bounds.Width);
        for (var index = 0; index < visibleColumns; index++, column++)
        {
            var width = GetColumnWidth(column);
            if (offset < width)
            {
                return column;
            }

            offset -= width;
        }

        return column - 1;
    }

    private int HitTestRow(double y)
    {
        if (Document is null)
        {
            return -1;
        }

        var offset = y - HeaderHeight;
        var row = _firstRow;
        while (row < Document.RowCount)
        {
            var height = GetRowHeight(row);
            if (offset < height)
            {
                return row;
            }

            offset -= height;
            row++;
        }

        return -1;
    }

    private void ClampOffsets(double contentHeight, double width)
    {
        if (Document is null)
        {
            return;
        }

        _firstRow = Math.Clamp(_firstRow, 0, CalculateLastFirstRow(contentHeight));
        _firstColumn = Math.Clamp(_firstColumn, 0, CalculateLastFirstColumn(width));
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
        _diffSpanCache.Clear();
        _firstRow = 0;
        _firstColumn = 0;
        _selection = null;
        if (Document?.SourceColumnWidths is { } sourceColumnWidths)
        {
            foreach (var (columnIndex, width) in sourceColumnWidths)
            {
                _columnWidths.TryAdd(columnIndex, width);
            }
        }

        UpdateHorizontalScrollMetrics(forceNotification: true);
        UpdateVerticalScrollMetrics(forceNotification: true);
        InvalidateVisual();
    }

    private void RemapCustomSizes(GridDocument? oldDocument, GridDocument? newDocument)
    {
        if (oldDocument is null ||
            newDocument is null ||
            oldDocument.Kind != newDocument.Kind ||
            !string.Equals(oldDocument.Title, newDocument.Title, StringComparison.Ordinal))
        {
            _columnWidths.Clear();
            _rowHeights.Clear();
            _autoFit.Clear();
            return;
        }

        var heightsByRow = new Dictionary<GridRowIdentity, double>();
        foreach (var (viewRowIndex, height) in _rowHeights)
        {
            if (viewRowIndex < oldDocument.RowCount)
            {
                heightsByRow[GridRowIdentity.From(oldDocument.Rows[viewRowIndex])] = height;
            }
        }

        _rowHeights.Clear();
        for (var viewRowIndex = 0; viewRowIndex < newDocument.RowCount; viewRowIndex++)
        {
            if (heightsByRow.TryGetValue(
                GridRowIdentity.From(newDocument.Rows[viewRowIndex]),
                out var height))
            {
                _rowHeights[viewRowIndex] = height;
            }
        }
    }

    private void GridSettingsChanged(object? sender, EventArgs eventArgs)
    {
        TrimCache();
        UpdateHorizontalScrollMetrics(forceNotification: true);
        UpdateVerticalScrollMetrics(forceNotification: true);
        InvalidateVisual();
    }

    private void SetFirstRow(int row)
    {
        var clamped = Math.Clamp(row, 0, _verticalScrollMaximum);
        if (_firstRow == clamped)
        {
            return;
        }

        _firstRow = clamped;
        UpdateVerticalScrollMetrics(forceNotification: true);
        InvalidateVisual();
    }

    private void SetFirstColumn(int column)
    {
        var clamped = Math.Clamp(column, 0, _horizontalScrollMaximum);
        if (_firstColumn == clamped)
        {
            return;
        }

        _firstColumn = clamped;
        UpdateHorizontalScrollMetrics(forceNotification: true);
        InvalidateVisual();
    }

    private void UpdateHorizontalScrollMetrics(bool forceNotification = false)
    {
        var maximum = CalculateLastFirstColumn(Bounds.Width);
        var firstColumn = Math.Clamp(_firstColumn, 0, maximum);
        var viewportSize = CalculateVisibleColumnCount(Bounds.Width, firstColumn);
        var changed = viewportSize != _horizontalViewportSize ||
            maximum != _horizontalScrollMaximum ||
            firstColumn != _firstColumn;
        _horizontalViewportSize = viewportSize;
        _horizontalScrollMaximum = maximum;
        _firstColumn = firstColumn;
        if (changed || forceNotification)
        {
            HorizontalScrollChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private void UpdateVerticalScrollMetrics(bool forceNotification = false)
    {
        var contentHeight = Math.Max(0, Bounds.Height - HeaderHeight);
        var maximum = CalculateLastFirstRow(contentHeight);
        var firstRow = Math.Clamp(_firstRow, 0, maximum);
        var viewportSize = CalculateVisibleRowCount(contentHeight, firstRow);
        var changed = viewportSize != _verticalViewportSize ||
            maximum != _verticalScrollMaximum ||
            firstRow != _firstRow;
        _verticalViewportSize = viewportSize;
        _verticalScrollMaximum = maximum;
        _firstRow = firstRow;
        if (changed || forceNotification)
        {
            VerticalScrollChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    private int CalculateVisibleColumnCount(double width, int? firstColumn = null)
    {
        var remainingWidth = Math.Max(0, width) / 2 - RowHeaderWidth;
        var column = firstColumn ?? _firstColumn;
        var count = 0;
        do
        {
            remainingWidth -= GetColumnWidth(column++);
            count++;
        }
        while (remainingWidth > 0);

        return count;
    }

    private int CalculateVisibleRowCount(double contentHeight, int? firstRow = null)
    {
        var remainingHeight = Math.Max(0, contentHeight);
        var row = firstRow ?? _firstRow;
        var count = 0;
        do
        {
            remainingHeight -= GetRowHeight(row++);
            count++;
        }
        while (remainingHeight > 0 && row < (Document?.RowCount ?? 0));

        return count;
    }

    private int CalculateLastFirstColumn(double width)
    {
        if (Document is null)
        {
            return 0;
        }

        var remainingWidth = Math.Max(0, width) / 2 - RowHeaderWidth;
        var firstColumn = Document.MaxColumnIndex;
        var occupiedBeforeLast = 0d;
        var occupied = GetColumnWidth(firstColumn);
        while (firstColumn > 0)
        {
            var candidateWidth = GetColumnWidth(firstColumn - 1);
            if (occupiedBeforeLast + candidateWidth >= remainingWidth)
            {
                break;
            }

            firstColumn--;
            occupiedBeforeLast += candidateWidth;
            occupied += candidateWidth;
            if (occupied >= remainingWidth)
            {
                break;
            }
        }

        return firstColumn;
    }

    private int CalculateCenteredFirstColumn(int targetColumn)
    {
        var center = Math.Max(0, (Bounds.Width / 2) - RowHeaderWidth) / 2;
        var firstColumn = targetColumn;
        var targetOffset = GetColumnWidth(targetColumn) / 2;
        while (firstColumn > 0)
        {
            var candidateOffset = targetOffset + GetColumnWidth(firstColumn - 1);
            if (Math.Abs(candidateOffset - center) > Math.Abs(targetOffset - center))
            {
                break;
            }

            firstColumn--;
            targetOffset = candidateOffset;
        }

        return firstColumn;
    }

    private int CalculateLastFirstRow(double contentHeight)
    {
        if (Document is null || Document.RowCount == 0)
        {
            return 0;
        }

        var remainingHeight = Math.Max(0, contentHeight);
        var firstRow = Document.RowCount - 1;
        var occupiedBeforeLast = 0d;
        var occupied = GetRowHeight(firstRow);
        while (firstRow > 0)
        {
            var candidateHeight = GetRowHeight(firstRow - 1);
            if (occupiedBeforeLast + candidateHeight >= remainingHeight)
            {
                break;
            }

            firstRow--;
            occupiedBeforeLast += candidateHeight;
            occupied += candidateHeight;
            if (occupied >= remainingHeight)
            {
                break;
            }
        }

        return firstRow;
    }

    private int CalculateCenteredFirstRow(int targetRow)
    {
        var center = Math.Max(0, Bounds.Height - HeaderHeight) / 2;
        var firstRow = targetRow;
        var targetOffset = GetRowHeight(targetRow) / 2;
        while (firstRow > 0)
        {
            var candidateOffset = targetOffset + GetRowHeight(firstRow - 1);
            if (Math.Abs(candidateOffset - center) > Math.Abs(targetOffset - center))
            {
                break;
            }

            firstRow--;
            targetOffset = candidateOffset;
        }

        return firstRow;
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
        context.DrawText(GetTextLayout(text, fontSize, weight, brush, maxWidth), origin);
    }

    private FormattedText GetTextLayout(
        string text,
        double fontSize,
        FontWeight? weight,
        IBrush brush,
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
            var formattedText = new FormattedText(
                text,
                CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface("Segoe UI", FontStyle.Normal, key.Weight),
                fontSize,
                brush);
            if (double.IsFinite(maxWidth))
            {
                formattedText.MaxTextWidth = Math.Max(0, maxWidth);
            }

            entry = new TextLayoutEntry(
                formattedText,
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

        return entry.Text;
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

    private enum GridResizeKind
    {
        Column,
        Row,
    }

    private readonly record struct GridResizeDrag(
        GridResizeKind Kind,
        int Index,
        double InitialPointerPosition,
        double InitialSize);

    private readonly record struct GridRowIdentity(
        int? BaseRowIndex,
        int? LocalRowIndex,
        int? RemoteRowIndex)
    {
        public static GridRowIdentity From(GridRowDescriptor row) =>
            new(row.BaseRowIndex, row.LocalRowIndex, row.RemoteRowIndex);
    }
}

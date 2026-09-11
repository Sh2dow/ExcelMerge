using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Styling;

namespace ExcelMerge.Desktop;

public sealed class DiffChangeMap : Control
{
    public static readonly StyledProperty<GridDocument?> DocumentProperty =
        AvaloniaProperty.Register<DiffChangeMap, GridDocument?>(nameof(Document));

    public static readonly StyledProperty<int> FirstVisibleRowProperty =
        AvaloniaProperty.Register<DiffChangeMap, int>(nameof(FirstVisibleRow));

    public static readonly StyledProperty<int> VisibleRowCountProperty =
        AvaloniaProperty.Register<DiffChangeMap, int>(nameof(VisibleRowCount), 1);

    public static readonly StyledProperty<int> FirstVisibleColumnProperty =
        AvaloniaProperty.Register<DiffChangeMap, int>(nameof(FirstVisibleColumn));

    public static readonly StyledProperty<int> VisibleColumnCountProperty =
        AvaloniaProperty.Register<DiffChangeMap, int>(nameof(VisibleColumnCount), 1);

    private GridPalette _palette = GridPalette.Light;
    private bool _dragging;

    static DiffChangeMap()
    {
        AffectsRender<DiffChangeMap>(
            DocumentProperty,
            FirstVisibleRowProperty,
            VisibleRowCountProperty,
            FirstVisibleColumnProperty,
            VisibleColumnCountProperty);
    }

    public DiffChangeMap()
    {
        ClipToBounds = true;
        Cursor = new Cursor(StandardCursorType.Hand);
    }

    public event Action<double, double>? PositionRequested;

    public GridDocument? Document
    {
        get => GetValue(DocumentProperty);
        set => SetValue(DocumentProperty, value);
    }

    public int FirstVisibleRow
    {
        get => GetValue(FirstVisibleRowProperty);
        set => SetValue(FirstVisibleRowProperty, value);
    }

    public int VisibleRowCount
    {
        get => GetValue(VisibleRowCountProperty);
        set => SetValue(VisibleRowCountProperty, value);
    }

    public int FirstVisibleColumn
    {
        get => GetValue(FirstVisibleColumnProperty);
        set => SetValue(FirstVisibleColumnProperty, value);
    }

    public int VisibleColumnCount
    {
        get => GetValue(VisibleColumnCountProperty);
        set => SetValue(VisibleColumnCountProperty, value);
    }

    public override void Render(DrawingContext context)
    {
        base.Render(context);
        var bounds = new Rect(Bounds.Size);
        context.DrawRectangle(_palette.MapBackground, _palette.MapBorderPen, bounds.Deflate(0.5));
        if (Document is not { RowCount: > 0 } document || bounds.Height <= 2 || bounds.Width <= 2)
        {
            return;
        }

        var content = bounds.Deflate(1);
        var columnCount = Math.Max(1, document.MaxColumnIndex + 1);
        foreach (var run in document.ChangeRuns)
        {
            if (run.State != GridRowVisualState.Changed)
            {
                continue;
            }

            var y = content.Y + (run.StartRowIndex / (double)document.RowCount * content.Height);
            var height = Math.Max(2, run.RowCount / (double)document.RowCount * content.Height);
            context.DrawRectangle(
                _palette.MapChanged,
                null,
                new Rect(content.X, y, content.Width, Math.Min(height, content.Bottom - y)));
        }


        var rowMarkerHeight = Math.Min(
            content.Height,
            Math.Max(2, content.Height / document.RowCount));
        foreach (var marker in document.RowMarkers)
        {
            var centerY = content.Y +
                ((marker.ViewRowIndex + 0.5) / document.RowCount * content.Height);
            var y = Math.Clamp(
                centerY - (rowMarkerHeight / 2),
                content.Y,
                content.Bottom - rowMarkerHeight);
            context.DrawRectangle(
                MarkerBrush(marker.State),
                null,
                new Rect(content.X, y, content.Width, rowMarkerHeight));
        }

        var cellMarkerWidth = Math.Min(content.Width, Math.Max(2, content.Width / columnCount));
        foreach (var marker in document.CellMarkers)
        {
            var centerX = content.X +
                ((marker.ColumnIndex + 0.5) / columnCount * content.Width);
            var centerY = content.Y +
                ((marker.ViewRowIndex + 0.5) / document.RowCount * content.Height);
            var x = Math.Clamp(
                centerX - (cellMarkerWidth / 2),
                content.X,
                content.Right - cellMarkerWidth);
            var y = Math.Clamp(
                centerY - (rowMarkerHeight / 2),
                content.Y,
                content.Bottom - rowMarkerHeight);
            context.DrawRectangle(
                MarkerBrush(marker.State),
                null,
                new Rect(x, y, cellMarkerWidth, rowMarkerHeight));
        }

        var viewportHeight = Math.Clamp(
            Math.Max(4, VisibleRowCount / (double)document.RowCount * content.Height),
            0,
            content.Height);
        var viewportY = content.Y +
            (Math.Clamp(FirstVisibleRow, 0, document.RowCount - 1) /
                (double)document.RowCount * content.Height);
        viewportY = Math.Min(viewportY, content.Bottom - viewportHeight);
        var viewportWidth = Math.Clamp(
            Math.Max(4, VisibleColumnCount / (double)columnCount * content.Width),
            0,
            content.Width);
        var viewportX = content.X +
            (Math.Clamp(FirstVisibleColumn, 0, columnCount - 1) /
                (double)columnCount * content.Width);
        viewportX = Math.Min(viewportX, content.Right - viewportWidth);
        context.DrawRectangle(
            _palette.MapViewportFill,
            _palette.MapViewportPen,
            new Rect(viewportX, viewportY, viewportWidth, viewportHeight));
    }

    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        if (Document is null || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            return;
        }

        _dragging = true;
        e.Pointer.Capture(this);
        RequestPosition(e.GetPosition(this));
        e.Handled = true;
    }

    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (_dragging)
        {
            RequestPosition(e.GetPosition(this));
            e.Handled = true;
        }
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (_dragging)
        {
            _dragging = false;
            e.Pointer.Capture(null);
            e.Handled = true;
        }
    }

    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
    {
        base.OnPointerCaptureLost(e);
        _dragging = false;
    }

    private void RequestPosition(Point point)
    {
        var horizontalRatio = Math.Clamp(
            (point.X - 1) / Math.Max(1, Bounds.Width - 2),
            0,
            1);
        var verticalRatio = Math.Clamp(
            (point.Y - 1) / Math.Max(1, Bounds.Height - 2),
            0,
            1);
        PositionRequested?.Invoke(horizontalRatio, verticalRatio);
    }

    private IBrush MarkerBrush(GridCellVisualState state) =>
        state == GridCellVisualState.Conflict ? _palette.MapConflict : _palette.MapResolved;

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        ActualThemeVariantChanged += OnActualThemeVariantChanged;
        _palette = GridPalette.ForVariant(ActualThemeVariant);
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        ActualThemeVariantChanged -= OnActualThemeVariantChanged;
        base.OnDetachedFromVisualTree(e);
    }

    private void OnActualThemeVariantChanged(object? sender, EventArgs e)
    {
        var palette = GridPalette.ForVariant(ActualThemeVariant);
        if (!ReferenceEquals(palette, _palette))
        {
            _palette = palette;
            InvalidateVisual();
        }
    }
}

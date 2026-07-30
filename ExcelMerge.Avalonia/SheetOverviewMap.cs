using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using ExcelMerge;

namespace ExcelMerge.Avalonia;

internal enum SheetOverviewMarkerKind : byte
{
    None,
    Changed,
    ResolvedConflict,
    UnresolvedConflict,
}

public sealed class SheetOverviewPositionRequestedEventArgs(Point position) : EventArgs
{
    public Point Position { get; } = position;
}

public sealed class SheetOverviewMap : Control
{
    private static readonly IBrush BackgroundBrush = new SolidColorBrush(Color.Parse("#F8FAFC"));
    private static readonly IBrush ChangedBrush = new SolidColorBrush(Color.Parse("#EAB308"));
    private static readonly IBrush UnresolvedConflictBrush = new SolidColorBrush(Color.Parse("#D6455D"));
    private static readonly IBrush ResolvedConflictBrush = new SolidColorBrush(Color.Parse("#22A06B"));
    private static readonly IBrush ViewportBrush = new SolidColorBrush(Color.Parse("#385B8DEF"));
    private static readonly Pen BorderPen = new(new SolidColorBrush(Color.Parse("#CBD5E1")), 1);
    private static readonly Pen ViewportPen = new(new SolidColorBrush(Color.Parse("#2563EB")), 1.5);

    private readonly List<Marker> _markers = new();
    private IReadOnlyList<RenderedMarker> _renderedMarkers = Array.Empty<RenderedMarker>();
    private Rect _viewport = new(0, 0, 1, 1);
    private int _rowCount;
    private int _columnCount;
    private int _renderWidth = -1;
    private int _renderHeight = -1;
    private bool _isDragging;

    public event EventHandler<SheetOverviewPositionRequestedEventArgs>? PositionRequested;

    public SheetOverviewMap()
    {
        Cursor = new Cursor(StandardCursorType.SizeAll);
        ClipToBounds = true;
    }

    public void SetRows(IReadOnlyList<DiffRow> rows)
    {
        _markers.Clear();
        _rowCount = rows.Count;
        _columnCount = rows.SelectMany(row => row.Cells)
            .Select(cell => cell.ColumnIndex + 1)
            .DefaultIfEmpty(0)
            .Max();
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            foreach (var cell in rows[rowIndex].Cells)
            {
                var kind = Classify(cell.Status, cell.IsConflict, cell.Resolution);
                if (kind != SheetOverviewMarkerKind.None)
                    _markers.Add(new Marker(rowIndex, cell.ColumnIndex, kind));
            }
        }

        _renderWidth = -1;
        _renderHeight = -1;
        InvalidateVisual();
    }

    public void SetViewport(Rect viewport)
    {
        var left = Math.Clamp(viewport.Left, 0, 1);
        var top = Math.Clamp(viewport.Top, 0, 1);
        var right = Math.Clamp(viewport.Right, left, 1);
        var bottom = Math.Clamp(viewport.Bottom, top, 1);
        var next = new Rect(left, top, right - left, bottom - top);
        if (_viewport == next)
            return;

        _viewport = next;
        InvalidateVisual();
    }

    public override void Render(DrawingContext context)
    {
        base.Render(context);
        var bounds = new Rect(Bounds.Size);
        context.DrawRectangle(BackgroundBrush, BorderPen, bounds);
        if (bounds.Width <= 4 || bounds.Height <= 4)
            return;

        EnsureRenderedMarkers((int)Math.Floor(bounds.Width - 4), (int)Math.Floor(bounds.Height - 4));
        foreach (var marker in _renderedMarkers)
            context.DrawRectangle(BrushFor(marker.Kind), null, marker.Bounds.Translate(new Vector(2, 2)));

        var viewport = ViewportBounds(bounds.Width, bounds.Height);
        context.DrawRectangle(ViewportBrush, ViewportPen, viewport, 2, 2);
    }

    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            return;

        _isDragging = true;
        e.Pointer.Capture(this);
        RequestPosition(e.GetPosition(this));
        e.Handled = true;
    }

    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (!_isDragging || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            return;

        RequestPosition(e.GetPosition(this));
        e.Handled = true;
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (!_isDragging)
            return;

        _isDragging = false;
        e.Pointer.Capture(null);
        e.Handled = true;
    }

    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
    {
        base.OnPointerCaptureLost(e);
        _isDragging = false;
    }

    internal static SheetOverviewMarkerKind Classify(
        ExcelCellStatus status,
        bool isConflict,
        MergeResolution resolution)
    {
        if (isConflict)
        {
            return resolution == MergeResolution.Unresolved
                ? SheetOverviewMarkerKind.UnresolvedConflict
                : SheetOverviewMarkerKind.ResolvedConflict;
        }

        return status == ExcelCellStatus.None
            ? SheetOverviewMarkerKind.None
            : SheetOverviewMarkerKind.Changed;
    }

    private void RequestPosition(Point point)
    {
        if (Bounds.Width <= 0 || Bounds.Height <= 0 || _rowCount == 0 || _columnCount == 0)
            return;

        PositionRequested?.Invoke(this, new SheetOverviewPositionRequestedEventArgs(new Point(
            Math.Clamp(point.X / Bounds.Width, 0, 1),
            Math.Clamp(point.Y / Bounds.Height, 0, 1))));
    }

    private Rect ViewportBounds(double width, double height)
    {
        const double inset = 2;
        const double minimumWidth = 10;
        const double minimumHeight = 14;
        var contentWidth = Math.Max(0, width - inset * 2);
        var contentHeight = Math.Max(0, height - inset * 2);
        var viewportWidth = Math.Min(contentWidth, Math.Max(minimumWidth, _viewport.Width * contentWidth));
        var viewportHeight = Math.Min(contentHeight, Math.Max(minimumHeight, _viewport.Height * contentHeight));
        var centerX = inset + (_viewport.Left + _viewport.Width / 2) * contentWidth;
        var centerY = inset + (_viewport.Top + _viewport.Height / 2) * contentHeight;
        var x = Math.Clamp(centerX - viewportWidth / 2, inset, width - inset - viewportWidth);
        var y = Math.Clamp(centerY - viewportHeight / 2, inset, height - inset - viewportHeight);
        return new Rect(x, y, viewportWidth, viewportHeight);
    }

    private void EnsureRenderedMarkers(int width, int height)
    {
        if (_renderWidth == width && _renderHeight == height)
            return;

        _renderWidth = width;
        _renderHeight = height;
        if (width <= 0 || height <= 0 || _rowCount == 0 || _columnCount == 0)
        {
            _renderedMarkers = Array.Empty<RenderedMarker>();
            return;
        }

        var pixels = new SheetOverviewMarkerKind[width * height];
        foreach (var marker in _markers)
        {
            var x = ScaleIndex(marker.ColumnIndex, _columnCount, width);
            var y = ScaleIndex(marker.RowIndex, _rowCount, height);
            for (var pixelY = Math.Max(0, y - 1); pixelY <= Math.Min(height - 1, y + 1); pixelY++)
            {
                for (var pixelX = Math.Max(0, x - 1); pixelX <= Math.Min(width - 1, x + 1); pixelX++)
                {
                    var pixelIndex = pixelY * width + pixelX;
                    if (marker.Kind > pixels[pixelIndex])
                        pixels[pixelIndex] = marker.Kind;
                }
            }
        }

        var rendered = new List<RenderedMarker>();
        for (var y = 0; y < height; y++)
        {
            var x = 0;
            while (x < width)
            {
                var kind = pixels[y * width + x];
                if (kind == SheetOverviewMarkerKind.None)
                {
                    x++;
                    continue;
                }

                var start = x++;
                while (x < width && pixels[y * width + x] == kind)
                    x++;
                rendered.Add(new RenderedMarker(new Rect(start, y, x - start, 1), kind));
            }
        }
        _renderedMarkers = rendered;
    }

    private static int ScaleIndex(int index, int count, int size)
    {
        if (count <= 1 || size <= 1)
            return 0;
        return (int)Math.Round((double)index / (count - 1) * (size - 1));
    }

    private static IBrush BrushFor(SheetOverviewMarkerKind kind) => kind switch
    {
        SheetOverviewMarkerKind.Changed => ChangedBrush,
        SheetOverviewMarkerKind.ResolvedConflict => ResolvedConflictBrush,
        SheetOverviewMarkerKind.UnresolvedConflict => UnresolvedConflictBrush,
        _ => Brushes.Transparent,
    };

    private readonly record struct Marker(int RowIndex, int ColumnIndex, SheetOverviewMarkerKind Kind);
    private readonly record struct RenderedMarker(Rect Bounds, SheetOverviewMarkerKind Kind);
}

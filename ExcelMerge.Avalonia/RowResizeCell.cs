using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;

namespace ExcelMerge.Avalonia;

internal sealed class RowResizeCell : Grid
{
    private const double MinimumHeight = 24;
    private const double MaximumHeight = 300;

    private readonly DiffRow _row;
    private readonly Action<DiffRow, double> _resize;
    private readonly Border _grip;
    private IPointer? _pointer;
    private Visual? _positionRoot;
    private double _startPointerY;
    private double _startHeight;

    public RowResizeCell(DiffRow row, Action<DiffRow, double> resize, Action<DiffRow>? click = null)
    {
        _row = row;
        _resize = resize;
        Background = Brushes.Transparent;

        Children.Add(new TextBlock
        {
            Text = row.DisplayIndex,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            IsHitTestVisible = false,
        });

        _grip = new Border
        {
            Height = 6,
            HorizontalAlignment = HorizontalAlignment.Stretch,
            VerticalAlignment = VerticalAlignment.Bottom,
            Background = Brushes.Transparent,
            Cursor = new Cursor(StandardCursorType.SizeNorthSouth),
            Child = new Border
            {
                Height = 1,
                VerticalAlignment = VerticalAlignment.Bottom,
                Background = new SolidColorBrush(Color.Parse("#66707A")),
                IsHitTestVisible = false,
            },
        };
        _grip.PointerPressed += GripPointerPressed;
        _grip.PointerMoved += GripPointerMoved;
        _grip.PointerReleased += GripPointerReleased;
        _grip.PointerCaptureLost += GripPointerCaptureLost;
        Children.Add(_grip);

        if (click != null && row.HasConflict)
        {
            Cursor = new Cursor(StandardCursorType.Hand);
            ToolTip.SetTip(this, "Click to resolve all conflicts in this row");
            AddHandler(
                PointerPressedEvent,
                (_, e) =>
                {
                    if (_grip.IsPointerOver || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
                        return;
                    click(row);
                    e.Handled = true;
                },
                RoutingStrategies.Bubble,
                handledEventsToo: true);
        }
    }

    private void GripPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (!e.GetCurrentPoint(_grip).Properties.IsLeftButtonPressed)
            return;

        _positionRoot = TopLevel.GetTopLevel(this);
        _pointer = e.Pointer;
        _startPointerY = e.GetPosition(_positionRoot).Y;
        _startHeight = _row.RowHeight;
        e.Pointer.Capture(_grip);
        e.PreventGestureRecognition();
        e.Handled = true;
    }

    private void GripPointerMoved(object? sender, PointerEventArgs e)
    {
        if (e.Pointer != _pointer || _positionRoot == null)
            return;

        var delta = e.GetPosition(_positionRoot).Y - _startPointerY;
        _resize(_row, Math.Clamp(_startHeight + delta, MinimumHeight, MaximumHeight));
        e.Handled = true;
    }

    private void GripPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (e.Pointer != _pointer)
            return;

        if (e.Pointer.Captured == _grip)
            e.Pointer.Capture(null);
        ResetPointer();
        e.Handled = true;
    }

    private void GripPointerCaptureLost(object? sender, PointerCaptureLostEventArgs e)
    {
        if (e.Pointer == _pointer)
            ResetPointer();
    }

    private void ResetPointer()
    {
        _pointer = null;
        _positionRoot = null;
    }
}

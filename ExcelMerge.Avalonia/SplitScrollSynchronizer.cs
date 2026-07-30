using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Threading;

namespace ExcelMerge.Avalonia;

internal sealed class SplitScrollSynchronizer : IDisposable
{
    private const double Epsilon = 0.01;

    private readonly DataGrid _localGrid;
    private readonly DataGrid _remoteGrid;
    private ScrollBar? _localVertical;
    private ScrollBar? _remoteVertical;
    private ScrollBar? _localHorizontal;
    private ScrollBar? _remoteHorizontal;
    private ScrollBar? _verticalFollower;
    private ScrollBar? _horizontalFollower;
    private long _verticalGeneration;
    private long _horizontalGeneration;
    private bool _applyingVertical;
    private bool _applyingHorizontal;
    private bool _disposed;

    public event EventHandler? ViewportChanged;

    public NormalizedViewport VerticalViewport => GetViewport(_localVertical ?? _remoteVertical);
    public NormalizedViewport HorizontalViewport => GetViewport(_localHorizontal ?? _remoteHorizontal);

    public SplitScrollSynchronizer(DataGrid localGrid, DataGrid remoteGrid)
    {
        _localGrid = localGrid;
        _remoteGrid = remoteGrid;
        _localGrid.TemplateApplied += GridTemplateApplied;
        _remoteGrid.TemplateApplied += GridTemplateApplied;
        _localGrid.VerticalScroll += LocalVerticalScroll;
        _remoteGrid.VerticalScroll += RemoteVerticalScroll;
        _localGrid.HorizontalScroll += LocalHorizontalScroll;
        _remoteGrid.HorizontalScroll += RemoteHorizontalScroll;

        _localGrid.ApplyTemplate();
        _remoteGrid.ApplyTemplate();
        CaptureScrollBars();
    }

    private void GridTemplateApplied(object? sender, TemplateAppliedEventArgs e)
    {
        Dispatcher.UIThread.Post(CaptureScrollBars, DispatcherPriority.Loaded);
    }

    private void CaptureScrollBars()
    {
        if (_disposed)
            return;

        DetachBarHandlers();
        _localVertical = FindScrollBar(_localGrid, "PART_VerticalScrollbar");
        _remoteVertical = FindScrollBar(_remoteGrid, "PART_VerticalScrollbar");
        _localHorizontal = FindScrollBar(_localGrid, "PART_HorizontalScrollbar");
        _remoteHorizontal = FindScrollBar(_remoteGrid, "PART_HorizontalScrollbar");
        AttachBarHandlers();
    }

    private static ScrollBar? FindScrollBar(DataGrid grid, string name)
    {
        return grid.GetTemplateDescendants()
            .OfType<ScrollBar>()
            .FirstOrDefault(scrollBar => scrollBar.Name == name);
    }

    private void AttachBarHandlers()
    {
        if (_localVertical != null)
            _localVertical.ValueChanged += LocalVerticalValueChanged;
        if (_remoteVertical != null)
            _remoteVertical.ValueChanged += RemoteVerticalValueChanged;
        if (_localHorizontal != null)
            _localHorizontal.ValueChanged += LocalHorizontalValueChanged;
        if (_remoteHorizontal != null)
            _remoteHorizontal.ValueChanged += RemoteHorizontalValueChanged;
    }

    private void DetachBarHandlers()
    {
        if (_localVertical != null)
            _localVertical.ValueChanged -= LocalVerticalValueChanged;
        if (_remoteVertical != null)
            _remoteVertical.ValueChanged -= RemoteVerticalValueChanged;
        if (_localHorizontal != null)
            _localHorizontal.ValueChanged -= LocalHorizontalValueChanged;
        if (_remoteHorizontal != null)
            _remoteHorizontal.ValueChanged -= RemoteHorizontalValueChanged;
    }

    private void LocalVerticalValueChanged(object? sender, RangeBaseValueChangedEventArgs e)
    {
        if (_applyingVertical)
            return;
        SynchronizeVertical(_localVertical, _remoteVertical);
        ViewportChanged?.Invoke(this, EventArgs.Empty);
    }

    private void RemoteVerticalValueChanged(object? sender, RangeBaseValueChangedEventArgs e)
    {
        if (_applyingVertical)
            return;
        SynchronizeVertical(_remoteVertical, _localVertical);
        ViewportChanged?.Invoke(this, EventArgs.Empty);
    }

    private void LocalHorizontalValueChanged(object? sender, RangeBaseValueChangedEventArgs e)
    {
        if (_applyingHorizontal)
            return;
        SynchronizeHorizontal(_localHorizontal, _remoteHorizontal);
        ViewportChanged?.Invoke(this, EventArgs.Empty);
    }

    private void RemoteHorizontalValueChanged(object? sender, RangeBaseValueChangedEventArgs e)
    {
        if (_applyingHorizontal)
            return;
        SynchronizeHorizontal(_remoteHorizontal, _localHorizontal);
        ViewportChanged?.Invoke(this, EventArgs.Empty);
    }

    private void LocalVerticalScroll(object? sender, ScrollEventArgs e) => TakeVerticalLeadership(_localVertical);
    private void RemoteVerticalScroll(object? sender, ScrollEventArgs e) => TakeVerticalLeadership(_remoteVertical);
    private void LocalHorizontalScroll(object? sender, ScrollEventArgs e) => TakeHorizontalLeadership(_localHorizontal);
    private void RemoteHorizontalScroll(object? sender, ScrollEventArgs e) => TakeHorizontalLeadership(_remoteHorizontal);

    private void TakeVerticalLeadership(ScrollBar? leader)
    {
        if (_applyingVertical || leader == null)
            return;

        _verticalFollower = null;
        _verticalGeneration++;
    }

    private void TakeHorizontalLeadership(ScrollBar? leader)
    {
        if (_applyingHorizontal || leader == null)
            return;

        _horizontalFollower = null;
        _horizontalGeneration++;
    }

    private void SynchronizeVertical(ScrollBar? source, ScrollBar? target)
    {
        if (_applyingVertical || source == null || target == null || ReferenceEquals(source, _verticalFollower))
            return;

        var desired = MapAlignedOffset(
            source.Value, source.Minimum, source.Maximum,
            target.Minimum, target.Maximum);
        if (Math.Abs(target.Value - desired) <= Epsilon)
            return;

        _applyingVertical = true;
        _verticalFollower = target;
        var generation = ++_verticalGeneration;
        try
        {
            MoveScrollBar(target, desired, true);
        }
        finally
        {
            _applyingVertical = false;
        }
        ReleaseFollowerAfterLayout(true, generation);
    }

    private void SynchronizeHorizontal(ScrollBar? source, ScrollBar? target)
    {
        if (_applyingHorizontal || source == null || target == null || ReferenceEquals(source, _horizontalFollower))
            return;

        var desired = MapAlignedOffset(
            source.Value, source.Minimum, source.Maximum,
            target.Minimum, target.Maximum);
        if (Math.Abs(target.Value - desired) <= Epsilon)
            return;

        _applyingHorizontal = true;
        _horizontalFollower = target;
        var generation = ++_horizontalGeneration;
        try
        {
            MoveScrollBar(target, desired, false);
        }
        finally
        {
            _applyingHorizontal = false;
        }
        ReleaseFollowerAfterLayout(false, generation);
    }

    private void ReleaseFollowerAfterLayout(bool vertical, long generation)
    {
        Dispatcher.UIThread.Post(() =>
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (_disposed)
                    return;

                if (vertical && generation == _verticalGeneration)
                    _verticalFollower = null;
                else if (!vertical && generation == _horizontalGeneration)
                    _horizontalFollower = null;
            }, DispatcherPriority.Loaded);
        }, DispatcherPriority.Loaded);
    }

    private static void MoveScrollBar(ScrollBar scrollBar, double desired, bool vertical)
    {
        desired = Math.Clamp(desired, scrollBar.Minimum, scrollBar.Maximum);
        var delta = desired - scrollBar.Value;
        if (Math.Abs(delta) <= Epsilon)
            return;

        var previousLargeChange = scrollBar.LargeChange;
        try
        {
            scrollBar.SetCurrentValue(RangeBase.LargeChangeProperty, Math.Abs(delta));
            if (vertical)
            {
                if (delta > 0)
                    scrollBar.PageDown();
                else
                    scrollBar.PageUp();
            }
            else
            {
                if (delta > 0)
                    scrollBar.PageRight();
                else
                    scrollBar.PageLeft();
            }
        }
        finally
        {
            scrollBar.SetCurrentValue(RangeBase.LargeChangeProperty, previousLargeChange);
        }
    }

    public void ScrollTo(double horizontalPosition, double verticalPosition)
    {
        if (_disposed)
            return;
        if (_localVertical == null || _localHorizontal == null)
            CaptureScrollBars();

        if (_localVertical != null)
            MoveScrollBar(_localVertical, ContentPositionToOffset(
                verticalPosition,
                _localVertical.Minimum,
                _localVertical.Maximum,
                _localVertical.ViewportSize), true);
        if (_localHorizontal != null)
            MoveScrollBar(_localHorizontal, ContentPositionToOffset(
                horizontalPosition,
                _localHorizontal.Minimum,
                _localHorizontal.Maximum,
                _localHorizontal.ViewportSize), false);
    }

    internal static double ContentPositionToOffset(
        double position,
        double minimum,
        double maximum,
        double viewportSize)
    {
        var range = Math.Max(0, maximum - minimum);
        var viewport = double.IsFinite(viewportSize) && viewportSize > 0 ? viewportSize : 0;
        var extent = range + viewport;
        var desired = minimum + Math.Clamp(position, 0, 1) * extent - viewport / 2;
        return Math.Clamp(desired, minimum, maximum);
    }

    private static NormalizedViewport GetViewport(ScrollBar? scrollBar)
    {
        if (scrollBar == null)
            return new NormalizedViewport(0, 1);

        var range = Math.Max(0, scrollBar.Maximum - scrollBar.Minimum);
        var viewport = double.IsFinite(scrollBar.ViewportSize) && scrollBar.ViewportSize > 0
            ? scrollBar.ViewportSize
            : 0;
        var extent = range + viewport;
        if (extent <= Epsilon)
            return new NormalizedViewport(0, 1);

        var start = Math.Clamp((scrollBar.Value - scrollBar.Minimum) / extent, 0, 1);
        var end = Math.Clamp((scrollBar.Value - scrollBar.Minimum + viewport) / extent, start, 1);
        return new NormalizedViewport(start, end);
    }

    internal static double MapAlignedOffset(
        double sourceValue,
        double sourceMinimum,
        double sourceMaximum,
        double targetMinimum,
        double targetMaximum)
    {
        if (sourceMaximum - sourceMinimum <= Epsilon)
            return targetMinimum;
        if (sourceValue <= sourceMinimum + Epsilon)
            return targetMinimum;
        if (sourceValue >= sourceMaximum - Epsilon)
            return targetMaximum;

        var alignedValue = targetMinimum + sourceValue - sourceMinimum;
        return Math.Clamp(alignedValue, targetMinimum, targetMaximum);
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _disposed = true;
        _localGrid.TemplateApplied -= GridTemplateApplied;
        _remoteGrid.TemplateApplied -= GridTemplateApplied;
        _localGrid.VerticalScroll -= LocalVerticalScroll;
        _remoteGrid.VerticalScroll -= RemoteVerticalScroll;
        _localGrid.HorizontalScroll -= LocalHorizontalScroll;
        _remoteGrid.HorizontalScroll -= RemoteHorizontalScroll;
        DetachBarHandlers();
        _localVertical = null;
        _remoteVertical = null;
        _localHorizontal = null;
        _remoteHorizontal = null;
    }
}

internal readonly record struct NormalizedViewport(double Start, double End);

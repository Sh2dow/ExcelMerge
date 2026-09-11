namespace ExcelMerge.Desktop;

/// <summary>
/// Grow-only per-column content width estimates from measured cell text. Never shrinks a column;
/// the caller skips columns that have a manual or source-provided width.
/// </summary>
public sealed class ContentWidthFitter
{
    public const double MinimumWidth = 64;
    public const double MaximumWidth = 480;
    public const double CellPadding = 16;

    private readonly Dictionary<int, double> _widths = [];

    public bool TryGrow(int columnIndex, double measuredTextWidth)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);
        if (!double.IsFinite(measuredTextWidth) || measuredTextWidth < 0)
        {
            return false;
        }

        var width = Math.Clamp(
            Math.Ceiling(measuredTextWidth + CellPadding),
            MinimumWidth,
            MaximumWidth);
        if (width <= _widths.GetValueOrDefault(columnIndex))
        {
            return false;
        }

        _widths[columnIndex] = width;
        return true;
    }

    public double? GetWidth(int columnIndex) =>
        _widths.TryGetValue(columnIndex, out var width) ? width : null;

    public void Clear() => _widths.Clear();
}

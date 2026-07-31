namespace ExcelMerge.Desktop;

public static class GridSettings
{
    private static double _rowHeight = 25;
    private static double _columnWidth = 118;
    private static int _cacheRows = 256;

    public static event EventHandler? Changed;

    public static double RowHeight
    {
        get => _rowHeight;
        set
        {
            var normalized = Math.Clamp(value, 18, 64);
            if (Math.Abs(_rowHeight - normalized) < 0.01)
                return;
            _rowHeight = normalized;
            Changed?.Invoke(null, EventArgs.Empty);
        }
    }

    public static double ColumnWidth
    {
        get => _columnWidth;
        set
        {
            var normalized = Math.Clamp(value, 64, 320);
            if (Math.Abs(_columnWidth - normalized) < 0.01)
                return;
            _columnWidth = normalized;
            Changed?.Invoke(null, EventArgs.Empty);
        }
    }

    public static int CacheRows
    {
        get => _cacheRows;
        set
        {
            var normalized = Math.Clamp(value, 32, 2048);
            if (_cacheRows == normalized)
                return;
            _cacheRows = normalized;
            Changed?.Invoke(null, EventArgs.Empty);
        }
    }
}

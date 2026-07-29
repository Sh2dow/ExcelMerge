using Avalonia;
using Avalonia.Controls;

namespace ExcelMerge.Avalonia;

internal sealed class ColumnWidthSynchronizer : IDisposable
{
    private readonly IReadOnlyList<DataGridColumn> _localColumns;
    private readonly IReadOnlyList<DataGridColumn> _remoteColumns;
    private bool _synchronizing;

    public ColumnWidthSynchronizer(DataGrid localGrid, DataGrid remoteGrid)
    {
        if (localGrid.Columns.Count != remoteGrid.Columns.Count)
            throw new ArgumentException("Both grids must have the same number of columns.");

        _localColumns = localGrid.Columns.ToList();
        _remoteColumns = remoteGrid.Columns.ToList();
        foreach (var column in _localColumns)
            column.PropertyChanged += ColumnPropertyChanged;
        foreach (var column in _remoteColumns)
            column.PropertyChanged += ColumnPropertyChanged;
    }

    private void ColumnPropertyChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (_synchronizing || e.Property != DataGridColumn.WidthProperty || sender is not DataGridColumn source)
            return;

        var localIndex = IndexOf(_localColumns, source);
        var target = localIndex >= 0
            ? _remoteColumns[localIndex]
            : _localColumns[IndexOf(_remoteColumns, source)];

        _synchronizing = true;
        try
        {
            target.Width = source.Width;
        }
        finally
        {
            _synchronizing = false;
        }
    }

    private static int IndexOf(IReadOnlyList<DataGridColumn> columns, DataGridColumn column)
    {
        for (var index = 0; index < columns.Count; index++)
        {
            if (ReferenceEquals(columns[index], column))
                return index;
        }

        return -1;
    }

    public void Dispose()
    {
        foreach (var column in _localColumns)
            column.PropertyChanged -= ColumnPropertyChanged;
        foreach (var column in _remoteColumns)
            column.PropertyChanged -= ColumnPropertyChanged;
    }
}

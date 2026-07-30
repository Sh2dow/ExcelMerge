using Avalonia.Media;
using ExcelMerge;
using System.ComponentModel;

namespace ExcelMerge.Avalonia;

public sealed class DiffRow : INotifyPropertyChanged
{
    private readonly IReadOnlyDictionary<int, DiffCell> _cellsByColumn;
    private readonly DiffCell[] _conflictCells;
    private readonly bool _isChanged;

    public int Index { get; }
    public IReadOnlyList<DiffCell> Cells { get; }
    public bool IsChanged => _isChanged;
    public bool HasConflict => _conflictCells.Length > 0;
    public IReadOnlyList<DiffCell> ConflictCells => _conflictCells;
    public int ConflictRowIndex => _conflictCells.Length > 0 ? _conflictCells[0].OriginalRowIndex : Index - 1;
    public bool IsResolved => _conflictCells.All(cell => cell.Resolution != MergeResolution.Unresolved);
    public string DisplayIndex => _conflictCells.Any(cell => cell.Resolution == MergeResolution.Both) ? $"{Index} x2" : Index.ToString();
    public double RowHeight { get; private set; } = 28;
    public bool HasCustomRowHeight { get; private set; }

    public event PropertyChangedEventHandler? PropertyChanged;

    public DiffRow(ExcelRowDiff row, IReadOnlyDictionary<(int Row, int Column), string>? baseValues = null)
    {
        Index = row.Index + 1;
        Cells = row.Cells.Values.Select(cell => new DiffCell(this, cell)).ToList();
        if (baseValues != null)
        {
            foreach (var cell in Cells)
                cell.ClassifyConflict(baseValues);
        }
        _isChanged = Cells.Any(cell => cell.Status != ExcelCellStatus.None);
        _conflictCells = Cells.Where(cell => cell.IsConflict).ToArray();
        _cellsByColumn = Cells.ToDictionary(cell => cell.ColumnIndex);
    }

    internal DiffCell? GetCell(int columnIndex) =>
        _cellsByColumn.TryGetValue(columnIndex, out var cell) ? cell : null;

    internal void NotifyResolutionChanged()
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsResolved)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayIndex)));
    }

    public void Resolve(MergeResolution resolution)
    {
        if (HasConflict)
        {
            foreach (var cell in _conflictCells)
                cell.Resolve(resolution);
        }
    }

    public void SetDefaultRowHeight(double height)
    {
        if (!HasCustomRowHeight)
            RowHeight = height;
    }

    public void SetCustomRowHeight(double height)
    {
        RowHeight = height;
        HasCustomRowHeight = true;
    }
}

public sealed class DiffCell : INotifyPropertyChanged
{
    private static readonly IBrush AddedRightBrush = new SolidColorBrush(Color.Parse("#DDF6E7"));
    private static readonly IBrush AddedLeftBrush = new SolidColorBrush(Color.Parse("#ECF8F1"));
    private static readonly IBrush RemovedRightBrush = new SolidColorBrush(Color.Parse("#FFF0F2"));
    private static readonly IBrush RemovedLeftBrush = new SolidColorBrush(Color.Parse("#FBDDE2"));
    private static readonly IBrush ModifiedBrush = new SolidColorBrush(Color.Parse("#FFF1BF"));
    private static readonly IBrush ConflictBrush = new SolidColorBrush(Color.Parse("#F7CBD2"));
    private static readonly IBrush AcceptedBrush = new SolidColorBrush(Color.Parse("#D9F2E3"));
    private static readonly IBrush RejectedBrush = new SolidColorBrush(Color.Parse("#F3E5E8"));
    private static readonly IBrush BothBrush = new SolidColorBrush(Color.Parse("#DCEBFA"));

    private readonly ExcelCellDiff _cell;
    private readonly DiffRow _owner;
    private InlineDiffResult? _inlineDiff;
    public int ColumnIndex => _cell.ColumnIndex;
    public int OriginalColumnIndex => Status == ExcelCellStatus.Added
        ? _cell.DstCell.OriginalColumnIndex
        : _cell.SrcCell.OriginalColumnIndex;
    public int OriginalRowIndex => Status == ExcelCellStatus.Added
        ? _cell.DstCell.OriginalRowIndex
        : _cell.SrcCell.OriginalRowIndex;
    public ExcelCellStatus Status => _cell.Status;
    public string LocalValue => _cell.SrcCell.Value;
    public string RemoteValue => _cell.DstCell.Value;
    public string BaseValue { get; private set; } = string.Empty;
    public bool IsConflict { get; private set; }
    public MergeResolution Resolution { get; private set; }
    public string? CustomValue { get; private set; }
    public string LocalDisplayValue => Value(false);
    public string RemoteDisplayValue => Value(true);
    public IBrush LocalBackground => Brush(false);
    public IBrush RemoteBackground => Brush(true);
    public string ResolutionToolTip => Resolution == MergeResolution.Unresolved
        ? "Click to resolve this conflict"
        : $"Resolved: {Resolution}. Click to change.";
    internal InlineDiffResult InlineDiff => _inlineDiff ??= InlineTextDiff.Create(LocalValue, RemoteValue);

    public event PropertyChangedEventHandler? PropertyChanged;

    public DiffCell(DiffRow owner, ExcelCellDiff cell)
    {
        _owner = owner;
        _cell = cell;
    }

    internal void ClassifyConflict(IReadOnlyDictionary<(int Row, int Column), string> baseValues)
    {
        var coordinate = Status == ExcelCellStatus.Added
            ? (_cell.DstCell.OriginalRowIndex, _cell.DstCell.OriginalColumnIndex)
            : (_cell.SrcCell.OriginalRowIndex, _cell.SrcCell.OriginalColumnIndex);
        baseValues.TryGetValue(coordinate, out var baseValue);
        BaseValue = baseValue ?? string.Empty;
        IsConflict = !string.Equals(LocalValue, RemoteValue, StringComparison.Ordinal)
            && !string.Equals(LocalValue, BaseValue, StringComparison.Ordinal)
            && !string.Equals(RemoteValue, BaseValue, StringComparison.Ordinal);
    }

    public string Value(bool remote)
    {
        return Resolution == MergeResolution.Custom
            ? CustomValue ?? string.Empty
            : remote ? RemoteValue : LocalValue;
    }

    public void Resolve(MergeResolution resolution, string? customValue = null)
    {
        if (!IsConflict)
            return;
        var nextCustomValue = resolution == MergeResolution.Custom ? customValue ?? string.Empty : null;
        if (Resolution == resolution && string.Equals(CustomValue, nextCustomValue, StringComparison.Ordinal))
            return;
        Resolution = resolution;
        CustomValue = nextCustomValue;
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Resolution)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CustomValue)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LocalDisplayValue)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RemoteDisplayValue)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(LocalBackground)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(RemoteBackground)));
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(ResolutionToolTip)));
        _owner.NotifyResolutionChanged();
    }

    public IBrush Brush(bool remote)
    {
        if (IsConflict)
        {
            return Resolution switch
            {
                MergeResolution.Local => remote ? RejectedBrush : AcceptedBrush,
                MergeResolution.Remote => remote ? AcceptedBrush : RejectedBrush,
                MergeResolution.Both => BothBrush,
                MergeResolution.Custom => BothBrush,
                _ => ConflictBrush,
            };
        }

        return Status switch
        {
            ExcelCellStatus.Added => remote ? AddedRightBrush : AddedLeftBrush,
            ExcelCellStatus.Removed => remote ? RemovedRightBrush : RemovedLeftBrush,
            ExcelCellStatus.Modified => ModifiedBrush,
            _ => Brushes.Transparent,
        };
    }
}

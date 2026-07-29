using Avalonia.Media;
using ExcelMerge;

namespace ExcelMerge.Avalonia;

public sealed class DiffRow
{
    public int Index { get; }
    public IReadOnlyList<DiffCell> Cells { get; }
    public bool IsChanged => Cells.Any(c => c.Status != ExcelCellStatus.None);
    public bool HasConflict => Cells.Any(cell => cell.IsConflict);
    public IReadOnlyList<DiffCell> ConflictCells => Cells.Where(cell => cell.IsConflict).ToList();
    public int ConflictRowIndex => ConflictCells.FirstOrDefault()?.OriginalRowIndex ?? Index - 1;
    public MergeResolution Resolution { get; private set; }
    public string DisplayIndex => Resolution == MergeResolution.Both ? $"{Index} x2" : Index.ToString();
    public double RowHeight { get; set; } = 32;

    public DiffRow(ExcelRowDiff row, IReadOnlyDictionary<(int Row, int Column), string>? baseValues = null)
    {
        Index = row.Index + 1;
        Cells = row.Cells.Values.Select(cell => new DiffCell(this, cell)).ToList();
        if (baseValues != null)
        {
            foreach (var cell in Cells)
                cell.ClassifyConflict(baseValues);
        }
    }

    public void Resolve(MergeResolution resolution)
    {
        if (HasConflict)
            Resolution = resolution;
    }
}

public sealed class DiffCell
{
    private static readonly IBrush AddedRightBrush = new SolidColorBrush(Color.Parse("#285A3A"));
    private static readonly IBrush AddedLeftBrush = new SolidColorBrush(Color.Parse("#1D3327"));
    private static readonly IBrush RemovedRightBrush = new SolidColorBrush(Color.Parse("#3A2528"));
    private static readonly IBrush RemovedLeftBrush = new SolidColorBrush(Color.Parse("#6B3038"));
    private static readonly IBrush ModifiedBrush = new SolidColorBrush(Color.Parse("#66551D"));
    private static readonly IBrush ConflictBrush = new SolidColorBrush(Color.Parse("#7A2E38"));
    private static readonly IBrush AcceptedBrush = new SolidColorBrush(Color.Parse("#285A3A"));
    private static readonly IBrush RejectedBrush = new SolidColorBrush(Color.Parse("#35282B"));
    private static readonly IBrush BothBrush = new SolidColorBrush(Color.Parse("#294D70"));

    private readonly DiffRow _owner;
    private readonly ExcelCellDiff _cell;
    public int ColumnIndex => _cell.ColumnIndex;
    public int OriginalRowIndex => Status == ExcelCellStatus.Added
        ? _cell.DstCell.OriginalRowIndex
        : _cell.SrcCell.OriginalRowIndex;
    public ExcelCellStatus Status => _cell.Status;
    public string LocalValue => _cell.SrcCell.Value;
    public string RemoteValue => _cell.DstCell.Value;
    public string BaseValue { get; private set; } = string.Empty;
    public bool IsConflict { get; private set; }

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

    public string Value(bool remote) => remote ? RemoteValue : LocalValue;

    public IBrush Brush(bool remote)
    {
        if (IsConflict)
        {
            return _owner.Resolution switch
            {
                MergeResolution.Local => remote ? RejectedBrush : AcceptedBrush,
                MergeResolution.Remote => remote ? AcceptedBrush : RejectedBrush,
                MergeResolution.Both => BothBrush,
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

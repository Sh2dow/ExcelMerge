using System.Globalization;

namespace ExcelMerge.Domain;

/// <summary>A zero-based worksheet coordinate.</summary>
public readonly record struct CellAddress : IComparable<CellAddress>
{
    public CellAddress(int rowIndex, int columnIndex)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
        ArgumentOutOfRangeException.ThrowIfNegative(columnIndex);

        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
    }

    public int RowIndex { get; }

    public int ColumnIndex { get; }

    public int CompareTo(CellAddress other)
    {
        var rowComparison = RowIndex.CompareTo(other.RowIndex);
        return rowComparison != 0 ? rowComparison : ColumnIndex.CompareTo(other.ColumnIndex);
    }

    public void Deconstruct(out int rowIndex, out int columnIndex)
    {
        rowIndex = RowIndex;
        columnIndex = ColumnIndex;
    }

    public override string ToString() =>
        string.Create(CultureInfo.InvariantCulture, $"R{RowIndex + 1}C{ColumnIndex + 1}");
}

public enum CellKind
{
    Blank,
    Text,
    Number,
    Boolean,
    DateTime,
    Error,
    Formula,
}

/// <summary>
/// An allocation-free scalar cell value. Formula text and cached formula values are represented by
/// <see cref="CellValue"/> rather than by this type.
/// </summary>
public readonly record struct CellScalar
{
    private readonly long _payload;
    private readonly string? _text;

    private CellScalar(CellKind kind, long payload, string? text)
    {
        Kind = kind;
        _payload = payload;
        _text = text;
    }

    public CellKind Kind { get; }

    public string? TextValue => Kind is CellKind.Text or CellKind.Error ? _text : null;

    public double? NumberValue =>
        Kind == CellKind.Number ? BitConverter.Int64BitsToDouble(_payload) : null;

    public bool? BooleanValue => Kind == CellKind.Boolean ? _payload != 0 : null;

    public DateTime? DateTimeValue =>
        Kind == CellKind.DateTime ? DateTime.FromBinary(_payload) : null;

    public static CellScalar Blank => default;

    public static CellScalar FromText(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new CellScalar(CellKind.Text, 0, value);
    }

    public static CellScalar FromNumber(double value) =>
        new(CellKind.Number, BitConverter.DoubleToInt64Bits(value), null);

    public static CellScalar FromBoolean(bool value) =>
        new(CellKind.Boolean, value ? 1 : 0, null);

    public static CellScalar FromDateTime(DateTime value) =>
        new(CellKind.DateTime, value.ToBinary(), null);

    public static CellScalar FromError(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new CellScalar(CellKind.Error, 0, value);
    }

    public override string ToString() => Kind switch
    {
        CellKind.Blank => string.Empty,
        CellKind.Text or CellKind.Error => _text ?? string.Empty,
        CellKind.Number => BitConverter.Int64BitsToDouble(_payload).ToString("R", CultureInfo.InvariantCulture),
        CellKind.Boolean => _payload == 0 ? bool.FalseString : bool.TrueString,
        CellKind.DateTime => DateTime.FromBinary(_payload).ToString("O", CultureInfo.InvariantCulture),
        _ => string.Empty,
    };
}

/// <summary>
/// A literal scalar or a formula with an optional typed cached result. Display text is retained for
/// presentation but is not a substitute for the typed value.
/// </summary>
public readonly record struct CellValue
{
    private readonly CellScalar _value;
    private readonly bool _hasCachedValue;

    public CellValue(CellScalar scalar, string? displayText = null)
        : this(scalar, false, null, displayText)
    {
    }

    private CellValue(
        CellScalar value,
        bool hasCachedValue,
        string? formula,
        string? displayText)
    {
        _value = value;
        _hasCachedValue = hasCachedValue;
        Formula = formula;
        DisplayText = displayText;
    }

    public CellKind Kind => IsFormula ? CellKind.Formula : _value.Kind;

    public bool IsFormula => Formula is not null;

    public CellScalar? Scalar => IsFormula ? null : _value;

    public string? Formula { get; }

    public CellScalar? CachedValue => IsFormula && _hasCachedValue ? _value : null;

    public string? DisplayText { get; }

    public static CellValue Blank => default;

    public static CellValue FromText(string value, string? displayText = null) =>
        new(CellScalar.FromText(value), displayText);

    public static CellValue FromNumber(double value, string? displayText = null) =>
        new(CellScalar.FromNumber(value), displayText);

    public static CellValue FromBoolean(bool value, string? displayText = null) =>
        new(CellScalar.FromBoolean(value), displayText);

    public static CellValue FromDateTime(DateTime value, string? displayText = null) =>
        new(CellScalar.FromDateTime(value), displayText);

    public static CellValue FromError(string value, string? displayText = null) =>
        new(CellScalar.FromError(value), displayText);

    public static CellValue FromFormula(
        string formula,
        CellScalar? cachedValue = null,
        string? displayText = null)
    {
        ArgumentNullException.ThrowIfNull(formula);
        return new CellValue(cachedValue.GetValueOrDefault(), cachedValue.HasValue, formula, displayText);
    }

    public CellValue WithDisplayText(string? displayText) =>
        new(_value, _hasCachedValue, Formula, displayText);

    public override string ToString() =>
        DisplayText ?? (IsFormula ? Formula ?? string.Empty : _value.ToString());
}

public readonly record struct CellRecord(
    CellAddress Address,
    CellValue Value,
    int? StyleIndex = null);

/// <summary>Cells must be ordered by column and belong to <see cref="RowIndex"/>.</summary>
public readonly record struct RowRecord
{
    public RowRecord(
        int rowIndex,
        ReadOnlyMemory<CellRecord> cells,
        double? height = null,
        bool isHidden = false,
        int? styleIndex = null)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(rowIndex);
        if (styleIndex < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(styleIndex),
                styleIndex,
                "A style index cannot be negative.");
        }

        RowIndex = rowIndex;
        Cells = cells;
        Height = height;
        IsHidden = isHidden;
        StyleIndex = styleIndex;
    }

    public int RowIndex { get; }

    public ReadOnlyMemory<CellRecord> Cells { get; }

    public double? Height { get; }

    public bool IsHidden { get; }

    public int? StyleIndex { get; }

    public int CellCount => Cells.Length;
}

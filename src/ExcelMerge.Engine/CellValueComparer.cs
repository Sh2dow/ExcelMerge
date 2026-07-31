using ExcelMerge.Domain;

namespace ExcelMerge.Engine;

/// <summary>
/// Compares CellValue's typed representation exactly. Display text is excluded unless explicitly
/// requested because formatting is not a semantic value for normal worksheet comparison.
/// </summary>
public sealed class CellValueComparer : IEqualityComparer<CellValue>
{
    public static CellValueComparer Exact { get; } = new(new CellComparisonOptions());

    public CellValueComparer(CellComparisonOptions? options = null)
    {
        Options = options ?? new CellComparisonOptions();
    }

    public CellComparisonOptions Options { get; }

    public bool Equals(CellValue x, CellValue y) => Equals(x, y, CancellationToken.None);

    public bool Equals(CellValue x, CellValue y, CancellationToken cancellationToken)
    {
        if (x.IsFormula != y.IsFormula)
            return false;

        if (x.IsFormula)
        {
            if (!OrdinalEquals(x.Formula, y.Formula, cancellationToken))
                return false;

            if (Options.CompareFormulaCachedValues &&
                !EqualsCachedValue(x, y, cancellationToken))
            {
                return false;
            }
        }
        else if (!EqualsScalar(x.Scalar.GetValueOrDefault(), y.Scalar.GetValueOrDefault(), cancellationToken))
        {
            return false;
        }

        return !Options.CompareDisplayText ||
            OrdinalEquals(x.DisplayText, y.DisplayText, cancellationToken);
    }

    public int GetHashCode(CellValue value) => unchecked((int)GetStableHash(value));

    internal ulong GetStableHash(
        CellValue value,
        StableHashLane lane = StableHashLane.Primary,
        CancellationToken cancellationToken = default)
    {
        var hash = new StableHashBuilder(lane);
        AddToHash(ref hash, value, Options, cancellationToken);
        return hash.Value;
    }

    internal static bool EqualsScalar(
        CellScalar left,
        CellScalar right,
        CancellationToken cancellationToken)
    {
        if (left.Kind != right.Kind)
            return false;

        return left.Kind switch
        {
            CellKind.Blank => true,
            CellKind.Text or CellKind.Error => OrdinalEquals(
                left.TextValue,
                right.TextValue,
                cancellationToken),
            CellKind.Number => DoubleBits(left.NumberValue.GetValueOrDefault()) ==
                DoubleBits(right.NumberValue.GetValueOrDefault()),
            CellKind.Boolean => left.BooleanValue == right.BooleanValue,
            CellKind.DateTime => DateTimeBits(left.DateTimeValue.GetValueOrDefault()) ==
                DateTimeBits(right.DateTimeValue.GetValueOrDefault()),
            _ => false,
        };
    }

    internal static void AddScalarToHash(
        ref StableHashBuilder hash,
        CellScalar scalar,
        CancellationToken cancellationToken = default)
    {
        hash.AddByte((byte)scalar.Kind);
        switch (scalar.Kind)
        {
            case CellKind.Blank:
                break;
            case CellKind.Text:
            case CellKind.Error:
                hash.AddString(scalar.TextValue, cancellationToken);
                break;
            case CellKind.Number:
                hash.AddUInt64(DoubleBits(scalar.NumberValue.GetValueOrDefault()));
                break;
            case CellKind.Boolean:
                hash.AddByte(scalar.BooleanValue.GetValueOrDefault() ? (byte)1 : (byte)0);
                break;
            case CellKind.DateTime:
                hash.AddInt64(DateTimeBits(scalar.DateTimeValue.GetValueOrDefault()));
                break;
        }
    }

    private static bool EqualsCachedValue(
        CellValue left,
        CellValue right,
        CancellationToken cancellationToken)
    {
        if (left.CachedValue.HasValue != right.CachedValue.HasValue)
            return false;
        return !left.CachedValue.HasValue || EqualsScalar(
            left.CachedValue.GetValueOrDefault(),
            right.CachedValue.GetValueOrDefault(),
            cancellationToken);
    }

    private static void AddToHash(
        ref StableHashBuilder hash,
        CellValue value,
        CellComparisonOptions options,
        CancellationToken cancellationToken)
    {
        if (value.IsFormula)
        {
            hash.AddByte(0xF0);
            hash.AddString(value.Formula, cancellationToken);
            if (options.CompareFormulaCachedValues)
            {
                hash.AddByte(value.CachedValue.HasValue ? (byte)1 : (byte)0);
                if (value.CachedValue.HasValue)
                    AddScalarToHash(ref hash, value.CachedValue.Value, cancellationToken);
            }
        }
        else
        {
            hash.AddByte(0x0F);
            AddScalarToHash(ref hash, value.Scalar.GetValueOrDefault(), cancellationToken);
        }

        if (options.CompareDisplayText)
        {
            hash.AddByte(value.DisplayText is null ? (byte)0 : (byte)1);
            hash.AddString(value.DisplayText, cancellationToken);
        }
    }

    private static bool OrdinalEquals(
        string? left,
        string? right,
        CancellationToken cancellationToken)
    {
        if (ReferenceEquals(left, right))
            return true;
        if (left is null || right is null || left.Length != right.Length)
            return false;

        for (var index = 0; index < left.Length; index++)
        {
            if ((index & 0xFFF) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            if (left[index] != right[index])
                return false;
        }

        return true;
    }

    private static long DateTimeBits(DateTime value) => value.ToBinary();

    private static ulong DoubleBits(double value) =>
        unchecked((ulong)BitConverter.DoubleToInt64Bits(value));
}

internal enum StableHashLane
{
    Primary,
    Secondary,
}

/// <summary>A small deterministic hash state; it deliberately does not use process-randomized hashing.</summary>
internal struct StableHashBuilder
{
    private ulong _value;

    public StableHashBuilder(StableHashLane lane)
    {
        _value = lane == StableHashLane.Primary
            ? 14695981039346656037UL
            : 1099511628211UL ^ 0x9E3779B97F4A7C15UL;
    }

    public ulong Value => _value;

    public void AddByte(byte value)
    {
        unchecked
        {
            _value ^= value;
            _value *= 1099511628211UL;
            _value ^= _value >> 29;
        }
    }

    public void AddUInt64(ulong value)
    {
        AddByte((byte)value);
        AddByte((byte)(value >> 8));
        AddByte((byte)(value >> 16));
        AddByte((byte)(value >> 24));
        AddByte((byte)(value >> 32));
        AddByte((byte)(value >> 40));
        AddByte((byte)(value >> 48));
        AddByte((byte)(value >> 56));
    }

    public void AddInt64(long value) => AddUInt64(unchecked((ulong)value));

    public void AddInt32(int value) => AddUInt64(unchecked((ulong)(uint)value));

    public void AddString(string? value, CancellationToken cancellationToken = default)
    {
        if (value is null)
        {
            AddByte(0);
            return;
        }

        AddByte(1);
        AddInt32(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if ((index & 0xFFF) == 0)
                cancellationToken.ThrowIfCancellationRequested();
            var character = value[index];
            AddByte((byte)character);
            AddByte((byte)(character >> 8));
        }
    }
}

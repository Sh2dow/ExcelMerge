using System.Buffers.Binary;
using System.Text;
using ExcelMerge.Domain;

namespace ExcelMerge.Storage;

internal static class RowCodec
{
    private const int RowFixedSize = (sizeof(int) * 2) + sizeof(byte);
    private const int CellFixedSize = (sizeof(int) * 3) + sizeof(byte);
    private const byte HasHeightFlag = 1 << 0;
    private const byte IsHiddenFlag = 1 << 1;
    private const byte HasStyleFlag = 1 << 2;
    private static readonly UTF8Encoding Utf8 = new(false, true);

    public static int GetEncodedLength(
        RowRecord row,
        int maximumRowSizeBytes,
        CancellationToken cancellationToken)
    {
        ValidateRow(row);

        long length = RowFixedSize;
        if (row.Height.HasValue)
        {
            length += sizeof(double);
        }

        if (row.StyleIndex.HasValue)
        {
            length += sizeof(int);
        }

        EnsureWithinLimit(length, maximumRowSizeBytes, nameof(row));

        var cells = row.Cells.Span;
        for (var i = 0; i < cells.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = cells[i];

            length += CellFixedSize;
            length += GetByteCount(cell.Value.DisplayText);
            length += GetValueLength(cell.Value);
            EnsureWithinLimit(length, maximumRowSizeBytes, nameof(row));
        }

        return checked((int)length);
    }

    public static void Encode(RowRecord row, Span<byte> destination, CancellationToken cancellationToken)
    {
        var cursor = 0;
        WriteInt32(destination, ref cursor, row.RowIndex);
        WriteInt32(destination, ref cursor, row.CellCount);

        byte flags = 0;
        if (row.Height.HasValue)
        {
            flags |= HasHeightFlag;
        }

        if (row.IsHidden)
        {
            flags |= IsHiddenFlag;
        }

        if (row.StyleIndex.HasValue)
        {
            flags |= HasStyleFlag;
        }

        destination[cursor++] = flags;

        if (row.Height is { } height)
        {
            WriteInt64(destination, ref cursor, BitConverter.DoubleToInt64Bits(height));
        }

        if (row.StyleIndex is { } rowStyleIndex)
        {
            WriteInt32(destination, ref cursor, rowStyleIndex);
        }

        var cells = row.Cells.Span;
        for (var i = 0; i < cells.Length; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = cells[i];

            WriteInt32(destination, ref cursor, cell.Address.ColumnIndex);
            WriteInt32(destination, ref cursor, cell.StyleIndex ?? -1);
            destination[cursor++] = (byte)cell.Value.Kind;
            WriteString(destination, ref cursor, cell.Value.DisplayText);
            WriteValue(destination, ref cursor, cell.Value);
        }

        if (cursor != destination.Length)
        {
            throw new InvalidOperationException("The row encoder produced an unexpected payload length.");
        }
    }

    public static RowRecord Decode(
        ReadOnlySpan<byte> source,
        int expectedCellCount,
        int expectedRowIndex,
        CancellationToken cancellationToken)
    {
        try
        {
            var cursor = 0;
            var rowIndex = ReadInt32(source, ref cursor);
            var cellCount = ReadInt32(source, ref cursor);
            var flags = ReadByte(source, ref cursor);

            if (rowIndex != expectedRowIndex || rowIndex < 0)
            {
                throw Corrupt("The row index does not match its index entry.");
            }

            if (cellCount != expectedCellCount || cellCount < 0)
            {
                throw Corrupt("The cell count does not match its index entry.");
            }

            if ((flags & ~(HasHeightFlag | IsHiddenFlag | HasStyleFlag)) != 0)
            {
                throw Corrupt("The row contains unsupported flags.");
            }

            var height = (flags & HasHeightFlag) != 0
                ? BitConverter.Int64BitsToDouble(ReadInt64(source, ref cursor))
                : (double?)null;
            var rowStyleIndex = (flags & HasStyleFlag) != 0
                ? ReadInt32(source, ref cursor)
                : (int?)null;

            if (rowStyleIndex < 0)
            {
                throw Corrupt("The row contains an invalid style index.");
            }

            if (cellCount > (source.Length - cursor) / CellFixedSize)
            {
                throw Corrupt("The row cell count exceeds the available payload.");
            }

            var cells = new CellRecord[cellCount];
            var previousColumnIndex = -1;
            for (var i = 0; i < cells.Length; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var columnIndex = ReadInt32(source, ref cursor);
                var styleIndex = ReadInt32(source, ref cursor);
                var kind = (CellKind)ReadByte(source, ref cursor);
                var displayText = ReadString(source, ref cursor);

                if (columnIndex <= previousColumnIndex || styleIndex < -1 || !IsDefined(kind))
                {
                    throw Corrupt("The row contains invalid cell metadata.");
                }

                var value = ReadValue(source, ref cursor, kind, displayText);
                cells[i] = new CellRecord(
                    new CellAddress(rowIndex, columnIndex),
                    value,
                    styleIndex == -1 ? null : styleIndex);
                previousColumnIndex = columnIndex;
            }

            if (cursor != source.Length)
            {
                throw Corrupt("The row contains trailing data.");
            }

            return new RowRecord(
                rowIndex,
                cells,
                height,
                (flags & IsHiddenFlag) != 0,
                rowStyleIndex);
        }
        catch (StorageCorruptionException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (
            exception is ArgumentException or DecoderFallbackException or OverflowException)
        {
            throw new StorageCorruptionException("The row payload is invalid.", exception);
        }
    }

    private static void ValidateRow(RowRecord row)
    {
        if (row.RowIndex < 0 || row.StyleIndex < 0)
        {
            throw new ArgumentException("The row contains invalid metadata.", nameof(row));
        }

        var cells = row.Cells.Span;
        var previousColumnIndex = -1;
        for (var i = 0; i < cells.Length; i++)
        {
            var cell = cells[i];
            if (cell.Address.RowIndex != row.RowIndex ||
                cell.Address.ColumnIndex <= previousColumnIndex ||
                cell.StyleIndex < 0 ||
                !IsDefined(cell.Value.Kind))
            {
                throw new ArgumentException(
                    "Cells must belong to the row, be ordered by unique column, and contain valid metadata.",
                    nameof(row));
            }

            ValidateValue(cell.Value, nameof(row));
            previousColumnIndex = cell.Address.ColumnIndex;
        }
    }

    private static void ValidateValue(CellValue value, string parameterName)
    {
        if (value.IsFormula)
        {
            if (value.Formula is null)
            {
                throw new ArgumentException("A formula cell requires formula text.", parameterName);
            }

            if (value.CachedValue is { } cachedValue)
            {
                ValidateScalar(cachedValue, parameterName);
            }

            return;
        }

        if (value.Scalar is not { } scalar)
        {
            throw new ArgumentException("A literal cell requires a scalar value.", parameterName);
        }

        ValidateScalar(scalar, parameterName);
    }

    private static void ValidateScalar(CellScalar scalar, string parameterName)
    {
        if (!IsScalarKind(scalar.Kind) ||
            scalar.Kind is CellKind.Text or CellKind.Error && scalar.TextValue is null)
        {
            throw new ArgumentException("The cell contains an invalid scalar value.", parameterName);
        }
    }

    private static long GetValueLength(CellValue value)
    {
        if (value.IsFormula)
        {
            long length = GetRequiredStringLength(value.Formula) + sizeof(byte);
            if (value.CachedValue is { } cachedValue)
            {
                length += sizeof(byte) + GetScalarLength(cachedValue);
            }

            return length;
        }

        return GetScalarLength(value.Scalar.GetValueOrDefault());
    }

    private static long GetScalarLength(CellScalar scalar) => scalar.Kind switch
    {
        CellKind.Blank => 0,
        CellKind.Text or CellKind.Error => GetRequiredStringLength(scalar.TextValue),
        CellKind.Number or CellKind.DateTime => sizeof(long),
        CellKind.Boolean => sizeof(byte),
        _ => throw new ArgumentException("The cell contains an invalid scalar value."),
    };

    private static void WriteValue(Span<byte> destination, ref int cursor, CellValue value)
    {
        if (value.IsFormula)
        {
            WriteRequiredString(destination, ref cursor, value.Formula);
            if (value.CachedValue is { } cachedValue)
            {
                destination[cursor++] = 1;
                destination[cursor++] = (byte)cachedValue.Kind;
                WriteScalar(destination, ref cursor, cachedValue);
            }
            else
            {
                destination[cursor++] = 0;
            }

            return;
        }

        WriteScalar(destination, ref cursor, value.Scalar.GetValueOrDefault());
    }

    private static void WriteScalar(Span<byte> destination, ref int cursor, CellScalar scalar)
    {
        switch (scalar.Kind)
        {
            case CellKind.Blank:
                return;
            case CellKind.Text:
            case CellKind.Error:
                WriteRequiredString(destination, ref cursor, scalar.TextValue);
                return;
            case CellKind.Number:
                WriteInt64(
                    destination,
                    ref cursor,
                    BitConverter.DoubleToInt64Bits(scalar.NumberValue.GetValueOrDefault()));
                return;
            case CellKind.Boolean:
                destination[cursor++] = scalar.BooleanValue.GetValueOrDefault() ? (byte)1 : (byte)0;
                return;
            case CellKind.DateTime:
                WriteInt64(destination, ref cursor, scalar.DateTimeValue.GetValueOrDefault().ToBinary());
                return;
            default:
                throw new InvalidOperationException("The cell contains an invalid scalar value.");
        }
    }

    private static CellValue ReadValue(
        ReadOnlySpan<byte> source,
        ref int cursor,
        CellKind kind,
        string? displayText)
    {
        if (kind == CellKind.Formula)
        {
            var formula = ReadRequiredString(source, ref cursor);
            var hasCachedValue = ReadByte(source, ref cursor);
            if (hasCachedValue > 1)
            {
                throw Corrupt("The formula cached-value flag is invalid.");
            }

            CellScalar? cachedValue = null;
            if (hasCachedValue == 1)
            {
                var cachedKind = (CellKind)ReadByte(source, ref cursor);
                if (!IsScalarKind(cachedKind))
                {
                    throw Corrupt("The formula cached-value kind is invalid.");
                }

                cachedValue = ReadScalar(source, ref cursor, cachedKind);
            }

            return CellValue.FromFormula(formula, cachedValue, displayText);
        }

        return new CellValue(ReadScalar(source, ref cursor, kind), displayText);
    }

    private static CellScalar ReadScalar(ReadOnlySpan<byte> source, ref int cursor, CellKind kind) => kind switch
    {
        CellKind.Blank => CellScalar.Blank,
        CellKind.Text => CellScalar.FromText(ReadRequiredString(source, ref cursor)),
        CellKind.Number => CellScalar.FromNumber(BitConverter.Int64BitsToDouble(ReadInt64(source, ref cursor))),
        CellKind.Boolean => CellScalar.FromBoolean(ReadBoolean(source, ref cursor)),
        CellKind.DateTime => CellScalar.FromDateTime(DateTime.FromBinary(ReadInt64(source, ref cursor))),
        CellKind.Error => CellScalar.FromError(ReadRequiredString(source, ref cursor)),
        _ => throw Corrupt("The cell scalar kind is invalid."),
    };

    private static bool ReadBoolean(ReadOnlySpan<byte> source, ref int cursor)
    {
        var value = ReadByte(source, ref cursor);
        return value switch
        {
            0 => false,
            1 => true,
            _ => throw Corrupt("A Boolean cell has an invalid payload."),
        };
    }

    private static int GetByteCount(string? value) => value is null ? 0 : Utf8.GetByteCount(value);

    private static int GetRequiredStringLength(string? value)
    {
        if (value is null)
        {
            throw new ArgumentException("A required cell string is missing.");
        }

        return checked(sizeof(int) + Utf8.GetByteCount(value));
    }

    private static void EnsureWithinLimit(long length, int maximumRowSizeBytes, string parameterName)
    {
        if (length > maximumRowSizeBytes)
        {
            throw new ArgumentException(
                $"The serialized row exceeds the configured {maximumRowSizeBytes:N0}-byte limit.",
                parameterName);
        }
    }

    private static void WriteInt32(Span<byte> destination, ref int cursor, int value)
    {
        BinaryPrimitives.WriteInt32LittleEndian(destination.Slice(cursor, sizeof(int)), value);
        cursor += sizeof(int);
    }

    private static void WriteInt64(Span<byte> destination, ref int cursor, long value)
    {
        BinaryPrimitives.WriteInt64LittleEndian(destination.Slice(cursor, sizeof(long)), value);
        cursor += sizeof(long);
    }

    private static void WriteString(Span<byte> destination, ref int cursor, string? value)
    {
        if (value is null)
        {
            WriteInt32(destination, ref cursor, -1);
            return;
        }

        WriteRequiredString(destination, ref cursor, value);
    }

    private static void WriteRequiredString(Span<byte> destination, ref int cursor, string? value)
    {
        if (value is null)
        {
            throw new InvalidOperationException("A required cell string is missing.");
        }

        var length = Utf8.GetByteCount(value);
        WriteInt32(destination, ref cursor, length);
        var bytesWritten = Utf8.GetBytes(value.AsSpan(), destination.Slice(cursor, length));
        cursor += bytesWritten;
    }

    private static int ReadInt32(ReadOnlySpan<byte> source, ref int cursor)
    {
        EnsureAvailable(source, cursor, sizeof(int));
        var value = BinaryPrimitives.ReadInt32LittleEndian(source.Slice(cursor, sizeof(int)));
        cursor += sizeof(int);
        return value;
    }

    private static long ReadInt64(ReadOnlySpan<byte> source, ref int cursor)
    {
        EnsureAvailable(source, cursor, sizeof(long));
        var value = BinaryPrimitives.ReadInt64LittleEndian(source.Slice(cursor, sizeof(long)));
        cursor += sizeof(long);
        return value;
    }

    private static byte ReadByte(ReadOnlySpan<byte> source, ref int cursor)
    {
        EnsureAvailable(source, cursor, sizeof(byte));
        return source[cursor++];
    }

    private static string? ReadString(ReadOnlySpan<byte> source, ref int cursor)
    {
        var length = ReadInt32(source, ref cursor);
        if (length == -1)
        {
            return null;
        }

        return ReadStringContent(source, ref cursor, length);
    }

    private static string ReadRequiredString(ReadOnlySpan<byte> source, ref int cursor)
    {
        var length = ReadInt32(source, ref cursor);
        if (length < 0)
        {
            throw Corrupt("A required string has an invalid byte length.");
        }

        return ReadStringContent(source, ref cursor, length);
    }

    private static string ReadStringContent(ReadOnlySpan<byte> source, ref int cursor, int length)
    {
        if (length < 0)
        {
            throw Corrupt("A string has an invalid byte length.");
        }

        EnsureAvailable(source, cursor, length);
        var value = Utf8.GetString(source.Slice(cursor, length));
        cursor += length;
        return value;
    }

    private static void EnsureAvailable(ReadOnlySpan<byte> source, int cursor, int length)
    {
        if (cursor < 0 || length < 0 || cursor > source.Length - length)
        {
            throw Corrupt("The row payload ended unexpectedly.");
        }
    }

    private static bool IsDefined(CellKind kind) => kind is
        CellKind.Blank or
        CellKind.Text or
        CellKind.Number or
        CellKind.Boolean or
        CellKind.DateTime or
        CellKind.Error or
        CellKind.Formula;

    private static bool IsScalarKind(CellKind kind) => kind is
        CellKind.Blank or
        CellKind.Text or
        CellKind.Number or
        CellKind.Boolean or
        CellKind.DateTime or
        CellKind.Error;

    private static StorageCorruptionException Corrupt(string message) => new(message);
}

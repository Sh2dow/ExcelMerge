using System.Buffers;
using System.Text;

namespace ExcelMerge.Delimited;

internal sealed class DelimitedParser
{
    private readonly TextReader _reader;
    private readonly char _delimiter;
    private readonly string _sourcePath;
    private readonly int _maximumFieldCharacters;
    private readonly int _maximumColumns;
    private readonly Func<string[], ValueTask> _consumeRecord;
    private readonly StringBuilder _field = new();
    private readonly List<string> _fields = [];

    private ParserState _state;
    private long _characterOffset;
    private long _recordsRead;
    private bool _recordStarted;
    private bool _skipLineFeed;

    private DelimitedParser(
        TextReader reader,
        char delimiter,
        string sourcePath,
        int maximumFieldCharacters,
        int maximumColumns,
        Func<string[], ValueTask> consumeRecord)
    {
        _reader = reader;
        _delimiter = delimiter;
        _sourcePath = sourcePath;
        _maximumFieldCharacters = maximumFieldCharacters;
        _maximumColumns = maximumColumns;
        _consumeRecord = consumeRecord;
    }

    public static ValueTask ParseAsync(
        TextReader reader,
        char delimiter,
        string sourcePath,
        int maximumFieldCharacters,
        int maximumColumns,
        Func<string[], ValueTask> consumeRecord,
        CancellationToken cancellationToken)
    {
        var parser = new DelimitedParser(
            reader,
            delimiter,
            sourcePath,
            maximumFieldCharacters,
            maximumColumns,
            consumeRecord);
        return parser.ParseAsync(cancellationToken);
    }

    private async ValueTask ParseAsync(CancellationToken cancellationToken)
    {
        var buffer = ArrayPool<char>.Shared.Rent(64 * 1024);
        try
        {
            while (true)
            {
                int count;
                try
                {
                    count = await _reader.ReadAsync(
                        buffer.AsMemory(0, buffer.Length),
                        cancellationToken).ConfigureAwait(false);
                }
                catch (DecoderFallbackException exception)
                {
                    throw Invalid(
                        DelimitedReaderError.InvalidEncoding,
                        "The input contains a byte sequence that is invalid for its detected encoding.",
                        exception);
                }

                if (count == 0)
                {
                    break;
                }

                for (var index = 0; index < count; index++)
                {
                    if ((index & 1023) == 0)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    await ProcessCharacterAsync(buffer[index]).ConfigureAwait(false);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (_state == ParserState.InQuotedField)
            {
                throw Invalid(
                    DelimitedReaderError.UnterminatedQuotedField,
                    "A quoted field reaches the end of the input without a closing quote.");
            }

            if (_state != ParserState.StartField || _recordStarted || _fields.Count > 0)
            {
                await CompleteRecordAsync().ConfigureAwait(false);
            }
        }
        finally
        {
            ArrayPool<char>.Shared.Return(buffer, clearArray: true);
        }
    }

    private async ValueTask ProcessCharacterAsync(char value)
    {
        _characterOffset++;
        if (_skipLineFeed)
        {
            _skipLineFeed = false;
            if (value == '\n')
            {
                return;
            }
        }

        switch (_state)
        {
            case ParserState.StartField:
                await ProcessStartFieldAsync(value).ConfigureAwait(false);
                break;

            case ParserState.InUnquotedField:
                await ProcessUnquotedFieldAsync(value).ConfigureAwait(false);
                break;

            case ParserState.InQuotedField:
                if (value == '"')
                {
                    _state = ParserState.AfterClosingQuote;
                }
                else
                {
                    Append(value);
                }

                break;

            case ParserState.AfterClosingQuote:
                await ProcessAfterClosingQuoteAsync(value).ConfigureAwait(false);
                break;

            default:
                throw new InvalidOperationException("The delimited parser entered an unknown state.");
        }
    }

    private async ValueTask ProcessStartFieldAsync(char value)
    {
        if (value == _delimiter)
        {
            CompleteField();
            _recordStarted = true;
            return;
        }

        if (value == '"')
        {
            _state = ParserState.InQuotedField;
            _recordStarted = true;
            return;
        }

        if (IsRecordSeparator(value))
        {
            _skipLineFeed = value == '\r';
            await CompleteRecordAsync().ConfigureAwait(false);
            return;
        }

        Append(value);
        _state = ParserState.InUnquotedField;
        _recordStarted = true;
    }

    private async ValueTask ProcessUnquotedFieldAsync(char value)
    {
        if (value == _delimiter)
        {
            CompleteField();
            _state = ParserState.StartField;
            return;
        }

        if (IsRecordSeparator(value))
        {
            _skipLineFeed = value == '\r';
            await CompleteRecordAsync().ConfigureAwait(false);
            return;
        }

        if (value == '"')
        {
            throw Invalid(
                DelimitedReaderError.UnexpectedQuote,
                "A quote is only valid at the start of a field or as an escaped quote.");
        }

        Append(value);
    }

    private async ValueTask ProcessAfterClosingQuoteAsync(char value)
    {
        if (value == '"')
        {
            Append('"');
            _state = ParserState.InQuotedField;
            return;
        }

        if (value == _delimiter)
        {
            CompleteField();
            _state = ParserState.StartField;
            return;
        }

        if (IsRecordSeparator(value))
        {
            _skipLineFeed = value == '\r';
            await CompleteRecordAsync().ConfigureAwait(false);
            return;
        }

        throw Invalid(
            DelimitedReaderError.UnexpectedCharacterAfterQuote,
            "Only a delimiter or record separator may follow a closing quote.");
    }

    private void Append(char value)
    {
        if (_maximumFieldCharacters > 0 && _field.Length >= _maximumFieldCharacters)
        {
            throw Invalid(
                DelimitedReaderError.FieldLimitExceeded,
                "A field exceeds the configured character limit.");
        }

        _field.Append(value);
    }

    private void CompleteField()
    {
        if (_maximumColumns > 0 && _fields.Count >= _maximumColumns)
        {
            throw Invalid(
                DelimitedReaderError.ColumnLimitExceeded,
                "A record exceeds the configured field limit.");
        }

        _fields.Add(_field.Length == 0 ? string.Empty : _field.ToString());
        _field.Clear();
    }

    private async ValueTask CompleteRecordAsync()
    {
        CompleteField();
        var record = _fields.ToArray();
        _fields.Clear();
        _field.Clear();
        _state = ParserState.StartField;
        _recordStarted = false;

        await _consumeRecord(record).ConfigureAwait(false);
        _recordsRead++;
    }

    private DelimitedReaderException Invalid(
        DelimitedReaderError error,
        string message,
        Exception? innerException = null) =>
        new(
            error,
            message,
            _sourcePath,
            _recordsRead + 1,
            _fields.Count + 1,
            _characterOffset == 0 ? null : _characterOffset,
            innerException);

    private static bool IsRecordSeparator(char value) => value is '\r' or '\n';

    private enum ParserState
    {
        StartField,
        InUnquotedField,
        InQuotedField,
        AfterClosingQuote,
    }
}

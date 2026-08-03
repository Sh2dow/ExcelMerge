using System.Globalization;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelMerge.Domain;
using DomainCellValue = ExcelMerge.Domain.CellValue;

namespace ExcelMerge.OpenXml;

internal sealed class OpenXmlReaderStyleCatalog
{
    private readonly ReaderCellStyle[] _styles;
    private readonly CultureInfo _displayCulture;

    private OpenXmlReaderStyleCatalog(ReaderCellStyle[] styles, CultureInfo displayCulture)
    {
        _styles = styles;
        _displayCulture = displayCulture;
    }

    public static OpenXmlReaderStyleCatalog Load(
        WorkbookStylesPart? part,
        string sourcePath,
        CultureInfo displayCulture,
        CancellationToken cancellationToken)
    {
        if (part is null)
        {
            return new OpenXmlReaderStyleCatalog(
                [new ReaderCellStyle(0, "General", false, false)],
                displayCulture);
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var stylesheet = part.Stylesheet ?? throw InvalidStyles(
                sourcePath,
                "The workbook style part does not contain a stylesheet root element.");
            var customFormats = new Dictionary<uint, string>();
            if (stylesheet.NumberingFormats is { } numberingFormats)
            {
                foreach (var format in numberingFormats.Elements<NumberingFormat>())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var numberFormatId = format.NumberFormatId?.Value;
                    var formatCode = format.FormatCode?.Value;
                    if (numberFormatId is null || formatCode is null ||
                        !customFormats.TryAdd(numberFormatId.Value, formatCode))
                    {
                        throw InvalidStyles(
                            sourcePath,
                            "The style table contains an invalid or duplicate number format.");
                    }
                }
            }

            var cellFormats = stylesheet.CellFormats?.Elements<CellFormat>().ToArray();
            if (cellFormats is not { Length: > 0 })
            {
                throw InvalidStyles(sourcePath, "The style table does not contain cell formats.");
            }

            var styles = new ReaderCellStyle[cellFormats.Length];
            for (var index = 0; index < cellFormats.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var numberFormatId = cellFormats[index].NumberFormatId?.Value ?? 0;
                var formatCode = customFormats.TryGetValue(numberFormatId, out var customCode)
                    ? customCode
                    : GetBuiltInFormatCode(numberFormatId);
                var isDuration = IsDurationFormat(formatCode);
                var isDate = IsBuiltInDateFormat(numberFormatId) ||
                    IsDateFormat(formatCode);
                styles[index] = new ReaderCellStyle(
                    numberFormatId,
                    formatCode,
                    isDate,
                    isDuration);
            }

            return new OpenXmlReaderStyleCatalog(styles, displayCulture);
        }
        catch (OpenXmlReaderException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (
            exception is InvalidDataException or FormatException or OverflowException or XmlException)
        {
            throw InvalidStyles(sourcePath, "The workbook style table is invalid.", exception);
        }
    }

    public bool IsValidStyleIndex(int styleIndex) => (uint)styleIndex < (uint)_styles.Length;

    public CellScalar ConvertNumber(
        double value,
        int styleIndex,
        bool uses1904DateSystem)
    {
        var style = _styles[styleIndex];
        if (style.IsDate &&
            !style.IsDuration &&
            TryConvertSerialDate(value, uses1904DateSystem, out var dateTime))
        {
            return CellScalar.FromDateTime(dateTime);
        }

        return CellScalar.FromNumber(value);
    }

    public string? Format(
        CellScalar scalar,
        int styleIndex,
        bool uses1904DateSystem,
        bool retainText)
    {
        var style = _styles[styleIndex];
        return scalar.Kind switch
        {
            CellKind.Blank => string.Empty,
            CellKind.Text => retainText ? scalar.TextValue : null,
            CellKind.Boolean => scalar.BooleanValue.GetValueOrDefault() ? "TRUE" : "FALSE",
            CellKind.Error => retainText ? scalar.TextValue : null,
            CellKind.DateTime => FormatDateTime(
                scalar.DateTimeValue.GetValueOrDefault(),
                style),
            CellKind.Number => FormatNumber(
                scalar.NumberValue.GetValueOrDefault(),
                style,
                uses1904DateSystem),
            _ => null,
        };
    }

    private string FormatNumber(double value, ReaderCellStyle style, bool uses1904DateSystem)
    {
        if (style.IsDuration)
        {
            return FormatElapsed(value, style.FormatCode);
        }

        if (style.IsDate && TryConvertSerialDate(value, uses1904DateSystem, out var dateTime))
        {
            return FormatDateTime(dateTime, style);
        }

        if (style.IsDate)
        {
            return FormatGeneral(value);
        }

        var formatCode = style.FormatCode;
        if (string.IsNullOrEmpty(formatCode) || formatCode.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            return FormatGeneral(value);
        }

        var numericFormat = SanitizeNumericFormat(formatCode);
        if (numericFormat is null)
        {
            return FormatGeneral(value);
        }

        try
        {
            return value.ToString(numericFormat, _displayCulture);
        }
        catch (FormatException)
        {
            return FormatGeneral(value);
        }
    }

    private string FormatDateTime(DateTime value, ReaderCellStyle style)
    {
        if (!style.IsDate || string.IsNullOrEmpty(style.FormatCode))
        {
            return value.ToString("G", _displayCulture);
        }

        var dateFormat = TranslateDateFormat(style.FormatCode);
        if (string.IsNullOrEmpty(dateFormat))
        {
            return value.ToString("G", _displayCulture);
        }

        try
        {
            return value.ToString(dateFormat, _displayCulture);
        }
        catch (FormatException)
        {
            return value.ToString("G", _displayCulture);
        }
    }

    private string FormatElapsed(double serialValue, string? formatCode)
    {
        if (!double.IsFinite(serialValue) || string.IsNullOrEmpty(formatCode))
        {
            return FormatGeneral(serialValue);
        }

        var absoluteSeconds = Math.Abs(serialValue) * TimeSpan.SecondsPerDay;
        if (absoluteSeconds > long.MaxValue)
        {
            return FormatGeneral(serialValue);
        }

        var totalSeconds = (long)Math.Floor(absoluteSeconds);
        var sign = serialValue < 0 ? "-" : string.Empty;
        if (formatCode.Contains("[h]", StringComparison.OrdinalIgnoreCase) ||
            formatCode.Contains("[hh]", StringComparison.OrdinalIgnoreCase))
        {
            var hours = totalSeconds / TimeSpan.SecondsPerHour;
            var minutes = (totalSeconds / TimeSpan.SecondsPerMinute) % 60;
            var seconds = totalSeconds % 60;
            return formatCode.Contains('s', StringComparison.OrdinalIgnoreCase)
                ? string.Create(_displayCulture, $"{sign}{hours}:{minutes:00}:{seconds:00}")
                : string.Create(_displayCulture, $"{sign}{hours}:{minutes:00}");
        }

        if (formatCode.Contains("[m]", StringComparison.OrdinalIgnoreCase) ||
            formatCode.Contains("[mm]", StringComparison.OrdinalIgnoreCase))
        {
            var minutes = totalSeconds / TimeSpan.SecondsPerMinute;
            var seconds = totalSeconds % 60;
            return string.Create(_displayCulture, $"{sign}{minutes}:{seconds:00}");
        }

        if (formatCode.Contains("[s]", StringComparison.OrdinalIgnoreCase) ||
            formatCode.Contains("[ss]", StringComparison.OrdinalIgnoreCase))
        {
            return sign + totalSeconds.ToString(_displayCulture);
        }

        return FormatGeneral(serialValue);
    }

    private string FormatGeneral(double value) => value.ToString("G15", _displayCulture);

    private static bool TryConvertSerialDate(
        double serialValue,
        bool uses1904DateSystem,
        out DateTime dateTime)
    {
        dateTime = default;
        if (!double.IsFinite(serialValue))
        {
            return false;
        }

        double adjustedSerial;
        DateTime epoch;
        if (uses1904DateSystem)
        {
            epoch = new DateTime(1904, 1, 1);
            adjustedSerial = serialValue;
        }
        else
        {
            if (Math.Floor(serialValue) == 60)
            {
                // Excel's fictitious 1900-02-29 cannot be represented by System.DateTime.
                return false;
            }

            epoch = new DateTime(1899, 12, 31);
            adjustedSerial = serialValue >= 60 ? serialValue - 1 : serialValue;
        }

        try
        {
            dateTime = epoch.AddDays(adjustedSerial);
            return true;
        }
        catch (ArgumentOutOfRangeException)
        {
            return false;
        }
    }

    private static bool IsBuiltInDateFormat(uint numberFormatId) =>
        numberFormatId is >= 14 and <= 22 or
            >= 27 and <= 36 or
            >= 45 and <= 47 or
            >= 50 and <= 58;

    private static bool IsDateFormat(string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode))
        {
            return false;
        }

        var section = GetFirstFormatSection(formatCode);
        for (var index = 0; index < section.Length; index++)
        {
            var current = section[index];
            if (current == '"')
            {
                SkipQuoted(section, ref index);
                continue;
            }

            if (current == '\\' || current is '_' or '*')
            {
                index++;
                continue;
            }

            if (current == '[')
            {
                var close = section.IndexOf(']', index + 1);
                if (close < 0)
                {
                    return false;
                }

                var bracketed = section.AsSpan(index + 1, close - index - 1);
                if (bracketed.Equals("h", StringComparison.OrdinalIgnoreCase) ||
                    bracketed.Equals("hh", StringComparison.OrdinalIgnoreCase) ||
                    bracketed.Equals("m", StringComparison.OrdinalIgnoreCase) ||
                    bracketed.Equals("mm", StringComparison.OrdinalIgnoreCase) ||
                    bracketed.Equals("s", StringComparison.OrdinalIgnoreCase) ||
                    bracketed.Equals("ss", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                index = close;
                continue;
            }

            if (current is 'y' or 'Y' or 'm' or 'M' or 'd' or 'D' or 'h' or 'H' or 's' or 'S')
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsDurationFormat(string? formatCode)
    {
        if (string.IsNullOrEmpty(formatCode))
        {
            return false;
        }

        return formatCode.Contains("[h]", StringComparison.OrdinalIgnoreCase) ||
            formatCode.Contains("[hh]", StringComparison.OrdinalIgnoreCase) ||
            formatCode.Contains("[m]", StringComparison.OrdinalIgnoreCase) ||
            formatCode.Contains("[mm]", StringComparison.OrdinalIgnoreCase) ||
            formatCode.Contains("[s]", StringComparison.OrdinalIgnoreCase) ||
            formatCode.Contains("[ss]", StringComparison.OrdinalIgnoreCase);
    }

    private static string? SanitizeNumericFormat(string formatCode)
    {
        if (formatCode.Contains('?') || formatCode.Contains('@'))
        {
            return null;
        }

        var result = new StringBuilder(formatCode.Length);
        for (var index = 0; index < formatCode.Length; index++)
        {
            var current = formatCode[index];
            if (current == '[')
            {
                var close = formatCode.IndexOf(']', index + 1);
                if (close < 0)
                {
                    return null;
                }

                var content = formatCode.AsSpan(index + 1, close - index - 1);
                if (content.Length > 1 && content[0] == '$')
                {
                    var localeSeparator = content.IndexOf('-');
                    var symbol = localeSeparator < 0 ? content[1..] : content[1..localeSeparator];
                    foreach (var value in symbol)
                    {
                        result.Append('\\').Append(value);
                    }
                }

                index = close;
                continue;
            }

            if (current is '_' or '*')
            {
                index++;
                continue;
            }

            result.Append(current);
        }

        return result.Length == 0 ? null : result.ToString();
    }

    private static string? TranslateDateFormat(string formatCode)
    {
        var section = GetFirstFormatSection(formatCode);
        var result = new StringBuilder(section.Length + 8);
        var hasAmPm = section.Contains("AM/PM", StringComparison.OrdinalIgnoreCase) ||
            section.Contains("A/P", StringComparison.OrdinalIgnoreCase);
        var previousToken = '\0';

        for (var index = 0; index < section.Length;)
        {
            if (section.AsSpan(index).StartsWith("AM/PM", StringComparison.OrdinalIgnoreCase))
            {
                result.Append("tt");
                index += 5;
                previousToken = 't';
                continue;
            }

            if (section.AsSpan(index).StartsWith("A/P", StringComparison.OrdinalIgnoreCase))
            {
                result.Append('t');
                index += 3;
                previousToken = 't';
                continue;
            }

            var current = section[index];
            if (current == '"')
            {
                var start = index++;
                while (index < section.Length)
                {
                    if (section[index++] == '"')
                    {
                        break;
                    }
                }

                result.Append(section, start, index - start);
                continue;
            }

            if (current == '\\')
            {
                result.Append(current);
                if (++index < section.Length)
                {
                    result.Append(section[index++]);
                }

                continue;
            }

            if (current is '_' or '*')
            {
                index = Math.Min(index + 2, section.Length);
                continue;
            }

            if (current == '[')
            {
                var close = section.IndexOf(']', index + 1);
                if (close < 0)
                {
                    return null;
                }

                index = close + 1;
                continue;
            }

            if (current is 'y' or 'Y' or 'd' or 'D' or 'h' or 'H' or 'm' or 'M' or 's' or 'S')
            {
                var token = char.ToLowerInvariant(current);
                var start = index++;
                while (index < section.Length &&
                    char.ToLowerInvariant(section[index]) == token)
                {
                    index++;
                }

                var length = index - start;
                switch (token)
                {
                    case 'y':
                        result.Append('y', Math.Min(length, 4));
                        break;
                    case 'd':
                        result.Append('d', Math.Min(length, 4));
                        break;
                    case 'h':
                        result.Append(hasAmPm ? 'h' : 'H', Math.Min(length, 2));
                        break;
                    case 's':
                        result.Append('s', Math.Min(length, 2));
                        break;
                    case 'm':
                        var isMinute = previousToken == 'h' || NextDateTokenIsSecond(section, index);
                        result.Append(isMinute ? 'm' : 'M', Math.Min(length, 4));
                        break;
                }

                previousToken = token;
                continue;
            }

            if (current == '0' && previousToken == 's')
            {
                var start = index++;
                while (index < section.Length && section[index] == '0')
                {
                    index++;
                }

                result.Append('f', Math.Min(index - start, 7));
                continue;
            }

            if (char.IsAsciiLetter(current))
            {
                result.Append('\\').Append(current);
            }
            else
            {
                result.Append(current);
            }

            index++;
        }

        return result.ToString();
    }

    private static bool NextDateTokenIsSecond(string formatCode, int index)
    {
        while (index < formatCode.Length)
        {
            var current = formatCode[index];
            if (current == '"')
            {
                SkipQuoted(formatCode, ref index);
            }
            else if (current == '\\' || current is '_' or '*')
            {
                index += 2;
            }
            else if (char.IsAsciiLetter(current))
            {
                return current is 's' or 'S';
            }
            else if (current is '/' or ';')
            {
                return false;
            }
            else
            {
                index++;
            }
        }

        return false;
    }

    private static string GetFirstFormatSection(string formatCode)
    {
        var quoted = false;
        for (var index = 0; index < formatCode.Length; index++)
        {
            if (formatCode[index] == '"')
            {
                quoted = !quoted;
            }
            else if (formatCode[index] == '\\')
            {
                index++;
            }
            else if (!quoted && formatCode[index] == ';')
            {
                return formatCode[..index];
            }
        }

        return formatCode;
    }

    private static void SkipQuoted(string value, ref int index)
    {
        index++;
        while (index < value.Length && value[index] != '"')
        {
            index++;
        }
    }

    private static string? GetBuiltInFormatCode(uint numberFormatId) => numberFormatId switch
    {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yy h:mm",
        37 => "#,##0;(#,##0)",
        38 => "#,##0;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => null,
    };

    private static OpenXmlReaderException InvalidStyles(
        string sourcePath,
        string message,
        Exception? innerException = null) =>
        new(
            OpenXmlReaderError.InvalidStyles,
            message,
            sourcePath,
            innerException: innerException);

    private readonly record struct ReaderCellStyle(
        uint NumberFormatId,
        string? FormatCode,
        bool IsDate,
        bool IsDuration);
}

internal static class OpenXmlReaderCellConverter
{
    public static DomainCellValue Convert(
        Cell cell,
        CellAddress address,
        int effectiveStyleIndex,
        OpenXmlReaderSharedFormulaResolver sharedFormulas,
        OpenXmlReaderSharedStrings sharedStrings,
        OpenXmlReaderStyleCatalog styles,
        bool uses1904DateSystem,
        bool includeDisplayText,
        SheetMetadata sheet,
        string sourcePath,
        CancellationToken cancellationToken)
    {
        var formula = cell.CellFormula;
        if (formula is not null)
        {
            if (cell.InlineString is not null)
            {
                throw InvalidCell(sheet, sourcePath, address, "A formula cell contains an inline string.");
            }

            var formulaText = sharedFormulas.Resolve(formula, address);
            CellScalar? cachedValue = null;
            string? displayText = null;
            if (cell.CellValue is not null)
            {
                cachedValue = ReadScalar(
                    cell,
                    address,
                    effectiveStyleIndex,
                    sharedStrings,
                    styles,
                    uses1904DateSystem,
                    sheet,
                    sourcePath,
                    hasExplicitValue: true,
                    cancellationToken);
                if (includeDisplayText)
                {
                    displayText = styles.Format(
                        cachedValue.Value,
                        effectiveStyleIndex,
                        uses1904DateSystem,
                        retainText: true);
                }
            }

            return DomainCellValue.FromFormula(formulaText, cachedValue, displayText);
        }

        var scalar = ReadScalar(
            cell,
            address,
            effectiveStyleIndex,
            sharedStrings,
            styles,
            uses1904DateSystem,
            sheet,
            sourcePath,
            hasExplicitValue: cell.CellValue is not null || cell.InlineString is not null,
            cancellationToken);
        var display = includeDisplayText
            ? styles.Format(
                scalar,
                effectiveStyleIndex,
                uses1904DateSystem,
                retainText: false)
            : null;
        return new DomainCellValue(scalar, display);
    }

    private static CellScalar ReadScalar(
        Cell cell,
        CellAddress address,
        int effectiveStyleIndex,
        OpenXmlReaderSharedStrings sharedStrings,
        OpenXmlReaderStyleCatalog styles,
        bool uses1904DateSystem,
        SheetMetadata sheet,
        string sourcePath,
        bool hasExplicitValue,
        CancellationToken cancellationToken)
    {
        var dataType = cell.DataType?.Value;
        var rawValue = cell.CellValue?.Text;

        if (dataType is null || dataType == CellValues.Number)
        {
            if (!hasExplicitValue || string.IsNullOrEmpty(rawValue))
            {
                return CellScalar.Blank;
            }

            if (cell.InlineString is not null ||
                !double.TryParse(
                    rawValue,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out var number) ||
                !double.IsFinite(number))
            {
                throw InvalidCell(sheet, sourcePath, address, "The numeric value is invalid.");
            }

            return styles.ConvertNumber(number, effectiveStyleIndex, uses1904DateSystem);
        }

        if (dataType == CellValues.SharedString)
        {
            if (!int.TryParse(
                    rawValue,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out var sharedStringIndex) ||
                sharedStringIndex < 0)
            {
                throw InvalidCell(sheet, sourcePath, address, "The shared-string index is invalid.");
            }

            return CellScalar.FromText(sharedStrings.Get(sharedStringIndex, sourcePath, sheet.Id));
        }

        if (dataType == CellValues.InlineString)
        {
            if (cell.CellValue is not null)
            {
                throw InvalidCell(sheet, sourcePath, address, "An inline-string cell also contains a value element.");
            }

            return CellScalar.FromText(
                cell.InlineString is null
                    ? string.Empty
                    : OpenXmlReaderSharedStrings.ExtractText(
                        cell.InlineString,
                        cancellationToken));
        }

        if (dataType == CellValues.String)
        {
            if (cell.InlineString is not null)
            {
                throw InvalidCell(sheet, sourcePath, address, "A string cell contains incompatible inline content.");
            }

            return CellScalar.FromText(rawValue ?? string.Empty);
        }

        if (dataType == CellValues.Boolean)
        {
            var boolean = rawValue switch
            {
                "0" => false,
                "1" => true,
                "false" => false,
                "true" => true,
                _ => (bool?)null,
            };
            if (boolean is null)
            {
                throw InvalidCell(sheet, sourcePath, address, "The Boolean value is invalid.");
            }

            return CellScalar.FromBoolean(boolean.Value);
        }

        if (dataType == CellValues.Error)
        {
            if (rawValue is null)
            {
                throw InvalidCell(sheet, sourcePath, address, "The error value is missing.");
            }

            return CellScalar.FromError(rawValue);
        }

        if (dataType == CellValues.Date)
        {
            if (rawValue is null ||
                !DateTime.TryParse(
                    rawValue,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind,
                    out var dateTime))
            {
                throw InvalidCell(sheet, sourcePath, address, "The ISO date value is invalid.");
            }

            return CellScalar.FromDateTime(dateTime);
        }

        throw InvalidCell(sheet, sourcePath, address, "The cell data type is not supported.");
    }

    private static OpenXmlReaderException InvalidCell(
        SheetMetadata sheet,
        string sourcePath,
        CellAddress address,
        string message) =>
        new(
            OpenXmlReaderError.InvalidWorksheet,
            $"Worksheet '{sheet.Name}', cell {address}: {message}",
            sourcePath,
            sheet.Id);
}

using System.Globalization;
using System.Text;

namespace ExcelMerge.OpenXml;

internal static class OpenXmlFormulaRowRewriter
{
    public static string Rewrite(
        string formula,
        Func<int, int?> rowMapper,
        string? sourceSheetName,
        bool rewriteUnqualifiedReferences = true)
    {
        ArgumentNullException.ThrowIfNull(formula);
        ArgumentNullException.ThrowIfNull(rowMapper);
        if (formula.Length == 0)
        {
            return formula;
        }

        if (ContainsThreeDimensionalReference(formula))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "Three-dimensional formula references cannot participate in a row transform.");
        }

        var result = new StringBuilder(formula.Length);
        for (var index = 0; index < formula.Length;)
        {
            var current = formula[index];
            if (current == '"')
            {
                CopyQuoted(formula, result, ref index, '"');
                continue;
            }

            if (current == '[')
            {
                CopyBracketed(formula, result, ref index);
                continue;
            }

            if (TryReadReference(formula, index, out var reference) &&
                ShouldRewriteReference(
                    formula,
                    index,
                    sourceSheetName,
                    rewriteUnqualifiedReferences))
            {
                var mappedRow = rowMapper(reference.RowIndex);
                if (!mappedRow.HasValue || mappedRow.Value is < 0 or >= 1_048_576)
                {
                    result.Append("#REF!");
                }
                else
                {
                    result.Append(formula, index, reference.DigitsStart - index);
                    result.Append(mappedRow.Value + 1);
                }

                index += reference.Length;
                continue;
            }

            result.Append(current);
            index++;
        }

        return RewriteRowRanges(
            result.ToString(),
            rowMapper,
            sourceSheetName,
            rewriteUnqualifiedReferences);
    }

    public static string RewriteDefinedName(
        string formula,
        Func<int, int?> rowMapper,
        string? sourceSheetName,
        bool rewriteUnqualifiedReferences)
    {
        return Rewrite(
            formula,
            rowMapper,
            sourceSheetName,
            rewriteUnqualifiedReferences);
    }

    public static bool ContainsForeignSheetReference(string formula, string allowedSheetName)
    {
        ArgumentNullException.ThrowIfNull(formula);
        ArgumentException.ThrowIfNullOrWhiteSpace(allowedSheetName);
        var inString = false;
        for (var index = 0; index < formula.Length; index++)
        {
            if (formula[index] == '"')
            {
                if (inString && index + 1 < formula.Length && formula[index + 1] == '"')
                {
                    index++;
                    continue;
                }

                inString = !inString;
                continue;
            }

            if (inString || formula[index] != '!')
            {
                continue;
            }

            var qualifierEnd = index;
            string qualifier;
            if (index > 0 && formula[index - 1] == '\'')
            {
                var cursor = index - 2;
                while (cursor >= 0)
                {
                    if (formula[cursor] != '\'')
                    {
                        cursor--;
                        continue;
                    }

                    if (cursor > 0 && formula[cursor - 1] == '\'')
                    {
                        cursor -= 2;
                        continue;
                    }

                    break;
                }

                if (cursor < 0 || cursor > 0 && formula[cursor - 1] == ']')
                {
                    return true;
                }

                qualifier = formula[(cursor + 1)..(qualifierEnd - 1)]
                    .Replace("''", "'", StringComparison.Ordinal);
            }
            else
            {
                var cursor = index - 1;
                while (cursor >= 0 && IsSheetNameCharacter(formula[cursor]))
                {
                    cursor--;
                }

                if (cursor >= 0 && formula[cursor] is ']' or ':')
                {
                    return true;
                }

                qualifier = formula[(cursor + 1)..qualifierEnd];
            }

            if (!string.Equals(qualifier, allowedSheetName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    public static string RenameSheetReferences(
        string formula,
        string sourceSheetName,
        string resultSheetName)
    {
        ArgumentNullException.ThrowIfNull(formula);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceSheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(resultSheetName);
        if (string.Equals(sourceSheetName, resultSheetName, StringComparison.Ordinal))
        {
            return formula;
        }

        var result = new StringBuilder(formula.Length);
        for (var index = 0; index < formula.Length;)
        {
            if (formula[index] == '"')
            {
                CopyQuoted(formula, result, ref index, '"');
                continue;
            }

            if (formula[index] == '[')
            {
                CopyBracketed(formula, result, ref index);
                continue;
            }

            if (formula[index] == '\'' &&
                TryReadQuotedSheetQualifier(formula, index, out var qualifierEnd, out var quotedName))
            {
                if (string.Equals(quotedName, sourceSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    AppendSheetQualifier(result, resultSheetName, forceQuotes: true);
                }
                else
                {
                    result.Append(formula, index, qualifierEnd - index);
                }

                index = qualifierEnd;
                continue;
            }

            if (IsSheetNameCharacter(formula[index]) &&
                (index == 0 ||
                 !IsSheetNameCharacter(formula[index - 1]) && formula[index - 1] is not (']' or ':')))
            {
                var cursor = index;
                while (cursor < formula.Length && IsSheetNameCharacter(formula[cursor]))
                {
                    cursor++;
                }

                if (cursor < formula.Length && formula[cursor] == '!' &&
                    string.Equals(
                        formula[index..cursor],
                        sourceSheetName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    AppendSheetQualifier(result, resultSheetName, forceQuotes: false);
                    index = cursor + 1;
                    continue;
                }
            }

            result.Append(formula[index++]);
        }

        return result.ToString();
    }

    private static bool ShouldRewriteReference(
        string formula,
        int referenceStart,
        string? sourceSheetName,
        bool rewriteUnqualified)
    {
        if (referenceStart == 0 || formula[referenceStart - 1] != '!')
        {
            return rewriteUnqualified ||
                referenceStart > 0 &&
                formula[referenceStart - 1] == ':' &&
                sourceSheetName is not null &&
                RangeStartUsesSheet(formula, referenceStart - 1, sourceSheetName);
        }

        if (sourceSheetName is null)
        {
            return false;
        }

        var qualifierEnd = referenceStart - 1;
        if (qualifierEnd == 0)
        {
            return false;
        }

        string qualifier;
        if (formula[qualifierEnd - 1] == '\'')
        {
            var cursor = qualifierEnd - 2;
            while (cursor >= 0)
            {
                if (formula[cursor] != '\'')
                {
                    cursor--;
                    continue;
                }

                if (cursor > 0 && formula[cursor - 1] == '\'')
                {
                    cursor -= 2;
                    continue;
                }

                break;
            }

            if (cursor < 0)
            {
                return false;
            }

            qualifier = formula[(cursor + 1)..(qualifierEnd - 1)].Replace("''", "'", StringComparison.Ordinal);
            if (cursor > 0 && formula[cursor - 1] == ']')
            {
                return false;
            }
        }
        else
        {
            var cursor = qualifierEnd - 1;
            while (cursor >= 0 && IsSheetNameCharacter(formula[cursor]))
            {
                cursor--;
            }

            if (cursor >= 0 && formula[cursor] is ']' or ':')
            {
                return false;
            }

            qualifier = formula[(cursor + 1)..qualifierEnd];
        }

        return string.Equals(qualifier, sourceSheetName, StringComparison.OrdinalIgnoreCase);
    }

    private static bool RangeStartUsesSheet(
        string formula,
        int colonIndex,
        string sourceSheetName)
    {
        var qualifierEnd = formula.LastIndexOf('!', colonIndex - 1);
        if (qualifierEnd < 0 ||
            !TryReadReference(formula, qualifierEnd + 1, out var rangeStart) ||
            qualifierEnd + 1 + rangeStart.Length != colonIndex)
        {
            return false;
        }

        return ShouldRewriteReference(
            formula,
            qualifierEnd + 1,
            sourceSheetName,
            rewriteUnqualified: false);
    }

    private static string RewriteRowRanges(
        string formula,
        Func<int, int?> rowMapper,
        string? sourceSheetName,
        bool rewriteUnqualifiedReferences)
    {
        var result = new StringBuilder(formula.Length);
        for (var index = 0; index < formula.Length;)
        {
            if (formula[index] == '"')
            {
                CopyQuoted(formula, result, ref index, '"');
                continue;
            }

            if (formula[index] == '[')
            {
                CopyBracketed(formula, result, ref index);
                continue;
            }

            if (TryReadRowRange(formula, index, out var range) &&
                (range.SheetName is null
                    ? rewriteUnqualifiedReferences
                    : sourceSheetName is not null && string.Equals(
                        range.SheetName,
                        sourceSheetName,
                        StringComparison.OrdinalIgnoreCase)))
            {
                var firstRow = rowMapper(range.FirstRowIndex);
                var lastRow = rowMapper(range.LastRowIndex);
                if (!firstRow.HasValue || !lastRow.HasValue ||
                    firstRow.Value is < 0 or >= 1_048_576 ||
                    lastRow.Value is < 0 or >= 1_048_576)
                {
                    result.Append("#REF!");
                }
                else
                {
                    result.Append(formula, index, range.FirstDigitsStart - index);
                    result.Append(firstRow.Value + 1);
                    result.Append(
                        formula,
                        range.FirstDigitsEnd,
                        range.LastDigitsStart - range.FirstDigitsEnd);
                    result.Append(lastRow.Value + 1);
                }

                index = range.End;
                continue;
            }

            result.Append(formula[index++]);
        }

        return result.ToString();
    }

    private static bool TryReadRowRange(
        string formula,
        int start,
        out FormulaRowRange range)
    {
        range = default;
        if (start > 0 && IsIdentifierCharacter(formula[start - 1]))
        {
            return false;
        }

        var cursor = start;
        string? sheetName = null;
        if (cursor < formula.Length && formula[cursor] == '\'')
        {
            if (!TryReadQuotedSheetQualifier(formula, cursor, out cursor, out sheetName))
            {
                return false;
            }
        }
        else if (cursor < formula.Length && IsSheetNameCharacter(formula[cursor]))
        {
            var qualifierStart = cursor;
            while (cursor < formula.Length && IsSheetNameCharacter(formula[cursor]))
            {
                cursor++;
            }

            if (cursor >= formula.Length || formula[cursor] != '!')
            {
                cursor = start;
            }
            else
            {
                if (start > 0 && formula[start - 1] is ']' or ':')
                {
                    return false;
                }

                sheetName = formula[qualifierStart..cursor];
                cursor++;
            }
        }

        if (!TryReadRowNumber(
                formula,
                ref cursor,
                out var firstDigitsStart,
                out var firstDigitsEnd,
                out var firstRowIndex) ||
            cursor >= formula.Length || formula[cursor++] != ':' ||
            !TryReadRowNumber(
                formula,
                ref cursor,
                out var lastDigitsStart,
                out _,
                out var lastRowIndex) ||
            cursor < formula.Length && IsIdentifierCharacter(formula[cursor]))
        {
            return false;
        }

        range = new FormulaRowRange(
            sheetName,
            firstRowIndex,
            lastRowIndex,
            firstDigitsStart,
            firstDigitsEnd,
            lastDigitsStart,
            cursor);
        return true;
    }

    private static bool TryReadRowNumber(
        string formula,
        ref int cursor,
        out int digitsStart,
        out int digitsEnd,
        out int rowIndex)
    {
        if (cursor < formula.Length && formula[cursor] == '$')
        {
            cursor++;
        }

        digitsStart = cursor;
        while (cursor < formula.Length && char.IsAsciiDigit(formula[cursor]))
        {
            cursor++;
        }

        digitsEnd = cursor;
        if (digitsStart == digitsEnd ||
            !int.TryParse(
                formula.AsSpan(digitsStart, digitsEnd - digitsStart),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out var oneBasedRow) ||
            oneBasedRow is < 1 or > 1_048_576)
        {
            rowIndex = default;
            return false;
        }

        rowIndex = oneBasedRow - 1;
        return true;
    }

    private static bool TryReadQuotedSheetQualifier(
        string formula,
        int start,
        out int end,
        out string sheetName)
    {
        var name = new StringBuilder();
        var cursor = start + 1;
        while (cursor < formula.Length)
        {
            if (formula[cursor] != '\'')
            {
                name.Append(formula[cursor++]);
                continue;
            }

            if (cursor + 1 < formula.Length && formula[cursor + 1] == '\'')
            {
                name.Append('\'');
                cursor += 2;
                continue;
            }

            if (cursor + 1 < formula.Length && formula[cursor + 1] == '!')
            {
                end = cursor + 2;
                sheetName = name.ToString();
                return true;
            }

            break;
        }

        end = default;
        sheetName = string.Empty;
        return false;
    }

    private static void AppendSheetQualifier(
        StringBuilder result,
        string sheetName,
        bool forceQuotes)
    {
        var requiresQuotes = forceQuotes ||
            sheetName.Length == 0 ||
            sheetName.Any(static character => !IsSheetNameCharacter(character)) ||
            OpenXmlReaderAddress.TryParseCell(sheetName, out _);
        if (requiresQuotes)
        {
            result.Append('\'');
            result.Append(sheetName.Replace("'", "''", StringComparison.Ordinal));
            result.Append("'!");
        }
        else
        {
            result.Append(sheetName);
            result.Append('!');
        }
    }

    private static bool TryReadReference(
        string formula,
        int start,
        out FormulaReference reference)
    {
        reference = default;
        if (start > 0 && IsIdentifierCharacter(formula[start - 1]))
        {
            return false;
        }

        var cursor = start;
        if (cursor < formula.Length && formula[cursor] == '$')
        {
            cursor++;
        }

        var lettersStart = cursor;
        while (cursor < formula.Length &&
            char.IsAsciiLetter(formula[cursor]) &&
            cursor - lettersStart < 3)
        {
            cursor++;
        }

        if (cursor == lettersStart ||
            cursor < formula.Length && char.IsAsciiLetter(formula[cursor]))
        {
            return false;
        }

        if (cursor < formula.Length && formula[cursor] == '$')
        {
            cursor++;
        }

        var digitsStart = cursor;
        while (cursor < formula.Length && char.IsAsciiDigit(formula[cursor]))
        {
            cursor++;
        }

        if (cursor == digitsStart ||
            cursor < formula.Length &&
            (IsIdentifierCharacter(formula[cursor]) || formula[cursor] is '!' or '('))
        {
            return false;
        }

        var oneBasedColumn = 0;
        var lettersEnd = digitsStart;
        if (lettersEnd > lettersStart && formula[lettersEnd - 1] == '$')
        {
            lettersEnd--;
        }

        for (var index = lettersStart; index < lettersEnd; index++)
        {
            var letter = char.ToUpperInvariant(formula[index]);
            oneBasedColumn = checked((oneBasedColumn * 26) + (letter - 'A' + 1));
        }

        if (oneBasedColumn is < 1 or > 16_384 ||
            !int.TryParse(
                formula.AsSpan(digitsStart, cursor - digitsStart),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out var oneBasedRow) ||
            oneBasedRow is < 1 or > 1_048_576)
        {
            return false;
        }

        reference = new FormulaReference(
            oneBasedRow - 1,
            digitsStart,
            cursor - start);
        return true;
    }

    private static void CopyQuoted(
        string source,
        StringBuilder destination,
        ref int index,
        char quote)
    {
        destination.Append(source[index++]);
        while (index < source.Length)
        {
            var current = source[index++];
            destination.Append(current);
            if (current != quote)
            {
                continue;
            }

            if (index < source.Length && source[index] == quote)
            {
                destination.Append(source[index++]);
                continue;
            }

            break;
        }
    }

    private static void CopyBracketed(string source, StringBuilder destination, ref int index)
    {
        while (index < source.Length)
        {
            var current = source[index++];
            destination.Append(current);
            if (current == ']')
            {
                break;
            }
        }
    }

    private static bool ContainsThreeDimensionalReference(string formula)
    {
        var inString = false;
        for (var index = 0; index < formula.Length; index++)
        {
            if (formula[index] == '"')
            {
                if (inString && index + 1 < formula.Length && formula[index + 1] == '"')
                {
                    index++;
                    continue;
                }

                inString = !inString;
                continue;
            }

            if (inString || formula[index] != '!')
            {
                continue;
            }

            for (var cursor = index - 1; cursor >= 0; cursor--)
            {
                var value = formula[cursor];
                if (value == ':')
                {
                    return true;
                }

                if (value is ']' or '(' or ')' or ',' or ';' or '+' or '-' or '*' or '/' or
                    '^' or '&' or '=' or '<' or '>' or '{' or '}')
                {
                    break;
                }
            }
        }

        return false;
    }

    private static bool IsIdentifierCharacter(char value) =>
        char.IsAsciiLetterOrDigit(value) || value is '_' or '.' or '\\';

    private static bool IsSheetNameCharacter(char value) =>
        char.IsAsciiLetterOrDigit(value) || value is '_' or '.';

    private readonly record struct FormulaReference(
        int RowIndex,
        int DigitsStart,
        int Length);

    private readonly record struct FormulaRowRange(
        string? SheetName,
        int FirstRowIndex,
        int LastRowIndex,
        int FirstDigitsStart,
        int FirstDigitsEnd,
        int LastDigitsStart,
        int End);
}

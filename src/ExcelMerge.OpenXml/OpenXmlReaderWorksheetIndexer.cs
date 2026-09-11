using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelMerge.Domain;
using ExcelMerge.Storage;

namespace ExcelMerge.OpenXml;

internal sealed class OpenXmlReaderSharedStrings : IDisposable
{
    private readonly ChunkedTextStore? _store;

    private OpenXmlReaderSharedStrings(ChunkedTextStore? store)
    {
        _store = store;
    }

    public long Count => _store?.Count ?? 0;

    public static OpenXmlReaderSharedStrings Load(
        SharedStringTablePart? part,
        Workspace workspace,
        string sourcePath,
        Action<long>? reportProgress,
        CancellationToken cancellationToken)
    {
        if (part is null)
        {
            return new OpenXmlReaderSharedStrings(null);
        }

        var store = workspace.CreateTextStore();
        try
        {
            var sawRoot = false;
            try
            {
                using var reader = DocumentFormat.OpenXml.OpenXmlReader.Create(part);
                while (reader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!reader.IsStartElement)
                    {
                        continue;
                    }

                    if (reader.Depth == 0)
                    {
                        sawRoot = reader.ElementType == typeof(SharedStringTable);
                        if (!sawRoot)
                        {
                            throw InvalidTable(sourcePath, "The shared-string part has an invalid root element.");
                        }
                    }

                    if (reader.ElementType != typeof(SharedStringItem))
                    {
                        continue;
                    }

                    if (reader.LoadCurrentElement() is not SharedStringItem item)
                    {
                        throw InvalidTable(sourcePath, "A shared-string item could not be read.");
                    }

                    store.Append(ExtractText(item, cancellationToken), cancellationToken);
                    if ((store.Count & 4095) == 0)
                    {
                        reportProgress?.Invoke(store.Count);
                    }
                }
            }
            catch (OpenXmlReaderException)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (exception is InvalidDataException or FormatException)
            {
                throw InvalidTable(sourcePath, "The shared-string table is invalid.", exception);
            }

            if (!sawRoot)
            {
                throw InvalidTable(sourcePath, "The shared-string part is empty.");
            }

            store.Flush();
            reportProgress?.Invoke(store.Count);
            return new OpenXmlReaderSharedStrings(store);
        }
        catch
        {
            store.Dispose();
            throw;
        }
    }

    public string Get(int index, string sourcePath, string sheetId)
    {
        if (index < 0 || index >= Count || _store is null)
        {
            throw new OpenXmlReaderException(
                OpenXmlReaderError.InvalidSharedStringTable,
                $"A cell refers to missing shared-string index {index}.",
                sourcePath,
                sheetId);
        }

        return _store.Read(index);
    }

    public void Dispose() => _store?.Dispose();

    public static string ExtractText(
        OpenXmlCompositeElement container,
        CancellationToken cancellationToken = default)
    {
        StringBuilder? builder = null;
        string? singleText = null;

        foreach (var child in container.ChildElements)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string? text = child switch
            {
                DocumentFormat.OpenXml.Spreadsheet.Text value => value.Text,
                Run run => string.Concat(
                    run.Elements<DocumentFormat.OpenXml.Spreadsheet.Text>()
                        .Select(static value => value.Text)),
                _ => null,
            };

            if (text is null)
            {
                continue;
            }

            if (singleText is null && builder is null)
            {
                singleText = text;
                continue;
            }

            builder ??= new StringBuilder(singleText);
            builder.Append(text);
        }

        return builder?.ToString() ?? singleText ?? string.Empty;
    }

    private static OpenXmlReaderException InvalidTable(
        string sourcePath,
        string message,
        Exception? innerException = null) =>
        new(
            OpenXmlReaderError.InvalidSharedStringTable,
            message,
            sourcePath,
            innerException: innerException);
}

internal static class OpenXmlReaderWorksheetIndexer
{
    private const int MaximumRowCount = 1_048_576;
    private const int MaximumColumnCount = 16_384;

    public static async ValueTask<IndexedWorksheet> IndexAsync(
        OpenXmlWorkbookReader.DiscoveredSheet sheet,
        ChunkedCellStore store,
        OpenXmlReaderSharedStrings sharedStrings,
        OpenXmlReaderStyleCatalog styles,
        bool uses1904DateSystem,
        OpenXmlWorkbookReader.ReaderSettings settings,
        Action<long, long>? reportProgress,
        CancellationToken cancellationToken)
    {
        var storedRowIndices = new List<int>();
        var columnStyles = new int?[MaximumColumnCount];
        Dictionary<int, double>? columnWidths = null;
        double? defaultColumnWidth = null;
        var sharedFormulas = new OpenXmlReaderSharedFormulaResolver(
            sheet.Metadata,
            sheet.SourcePath);
        var nextRowIndex = 0;
        int? firstRowIndex = null;
        int? lastRowIndex = null;
        int? firstColumnIndex = null;
        int? lastColumnIndex = null;
        long cellCount = 0;
        long nonEmptyCellCount = 0;
        var sawWorksheet = false;
        var sawSheetData = false;
        var insideSheetData = false;

        try
        {
            using var reader = DocumentFormat.OpenXml.OpenXmlReader.Create(sheet.WorksheetPart);
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();

                if (reader.Depth == 0 && reader.IsStartElement)
                {
                    sawWorksheet = reader.ElementType == typeof(Worksheet);
                    if (!sawWorksheet)
                    {
                        throw InvalidWorksheet(sheet, "The worksheet part has an invalid root element.");
                    }
                }

                if (reader.IsStartElement && reader.ElementType == typeof(SheetFormatProperties))
                {
                    if (sawSheetData || reader.LoadCurrentElement() is not SheetFormatProperties format)
                    {
                        throw InvalidWorksheet(sheet, "The worksheet has invalid sheet format properties.");
                    }

                    defaultColumnWidth = ReadDefaultColumnWidth(format, sheet);
                    continue;
                }

                if (reader.IsStartElement && reader.ElementType == typeof(Columns))
                {
                    if (sawSheetData || reader.LoadCurrentElement() is not Columns columns)
                    {
                        throw InvalidWorksheet(sheet, "The worksheet has invalid column metadata ordering.");
                    }

                    ReadColumnStyles(
                        columns,
                        columnStyles,
                        ref columnWidths,
                        styles,
                        sheet,
                        cancellationToken);
                    continue;
                }

                if (reader.ElementType == typeof(SheetData))
                {
                    if (reader.IsStartElement)
                    {
                        if (sawSheetData)
                        {
                            throw InvalidWorksheet(sheet, "The worksheet contains multiple sheetData elements.");
                        }

                        sawSheetData = true;
                        insideSheetData = true;
                    }
                    else if (reader.IsEndElement)
                    {
                        insideSheetData = false;
                    }

                    continue;
                }

                if (!reader.IsStartElement || reader.ElementType != typeof(Row))
                {
                    continue;
                }

                if (!insideSheetData || reader.LoadCurrentElement() is not Row row)
                {
                    throw InvalidWorksheet(sheet, "A worksheet row appears outside sheetData.");
                }

                var rowRecord = ReadRow(
                    row,
                    ref nextRowIndex,
                    columnStyles,
                    sharedFormulas,
                    sharedStrings,
                    styles,
                    uses1904DateSystem,
                    settings,
                    sheet,
                    cancellationToken);
                await store.AppendRowAsync(rowRecord, cancellationToken).ConfigureAwait(false);
                storedRowIndices.Add(rowRecord.RowIndex);

                firstRowIndex ??= rowRecord.RowIndex;
                lastRowIndex = rowRecord.RowIndex;
                cellCount = checked(cellCount + rowRecord.CellCount);
                foreach (var cell in rowRecord.Cells.Span)
                {
                    firstColumnIndex = firstColumnIndex.HasValue
                        ? Math.Min(firstColumnIndex.Value, cell.Address.ColumnIndex)
                        : cell.Address.ColumnIndex;
                    lastColumnIndex = lastColumnIndex.HasValue
                        ? Math.Max(lastColumnIndex.Value, cell.Address.ColumnIndex)
                        : cell.Address.ColumnIndex;
                    if (cell.Value.Kind != CellKind.Blank)
                    {
                        nonEmptyCellCount++;
                    }
                }

                if (storedRowIndices.Count % settings.ProgressIntervalRows == 0)
                {
                    reportProgress?.Invoke(storedRowIndices.Count, cellCount);
                }
            }
        }
        catch (OpenXmlReaderException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is InvalidDataException or FormatException)
        {
            throw InvalidWorksheet(sheet, "The worksheet XML is invalid.", exception);
        }

        if (!sawWorksheet || !sawSheetData || insideSheetData)
        {
            throw InvalidWorksheet(sheet, "The worksheet is incomplete or missing sheetData.");
        }

        reportProgress?.Invoke(storedRowIndices.Count, cellCount);
        var metadata = sheet.Metadata with
        {
            FirstRowIndex = firstRowIndex,
            LastRowIndex = lastRowIndex,
            FirstColumnIndex = firstColumnIndex,
            LastColumnIndex = lastColumnIndex,
            NonEmptyCellCount = nonEmptyCellCount,
            ColumnWidths = BuildColumnWidthList(columnWidths),
            DefaultColumnWidth = defaultColumnWidth,
        };
        return new IndexedWorksheet(metadata, storedRowIndices.ToArray(), cellCount);
    }

    private static RowRecord ReadRow(
        Row row,
        ref int nextRowIndex,
        int?[] columnStyles,
        OpenXmlReaderSharedFormulaResolver sharedFormulas,
        OpenXmlReaderSharedStrings sharedStrings,
        OpenXmlReaderStyleCatalog styles,
        bool uses1904DateSystem,
        OpenXmlWorkbookReader.ReaderSettings settings,
        OpenXmlWorkbookReader.DiscoveredSheet sheet,
        CancellationToken cancellationToken)
    {
        var firstCell = row.Elements<Cell>().FirstOrDefault();
        var rowIndex = GetRowIndex(row, firstCell, nextRowIndex, sheet);
        if (rowIndex < nextRowIndex)
        {
            throw InvalidWorksheet(sheet, "Worksheet rows are duplicated or out of order.");
        }

        nextRowIndex = checked(rowIndex + 1);
        var rowStyleIndex = GetStyleIndex(row.StyleIndex, styles, sheet, "row");
        var height = row.Height?.Value;
        if (height is { } rowHeight && (!double.IsFinite(rowHeight) || rowHeight < 0))
        {
            throw InvalidWorksheet(sheet, "A worksheet row has an invalid height.");
        }

        var cells = new List<CellRecord>();
        var previousColumnIndex = -1;
        foreach (var cell in row.Elements<Cell>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var columnIndex = GetColumnIndex(cell, rowIndex, previousColumnIndex + 1, sheet);
            if (columnIndex <= previousColumnIndex)
            {
                throw InvalidWorksheet(sheet, "Cells in a worksheet row are duplicated or out of order.");
            }

            var styleIndex = GetStyleIndex(cell.StyleIndex, styles, sheet, "cell");
            var effectiveStyleIndex = styleIndex ??
                rowStyleIndex ??
                columnStyles[columnIndex] ??
                0;
            var address = new CellAddress(rowIndex, columnIndex);
            var value = OpenXmlReaderCellConverter.Convert(
                cell,
                address,
                effectiveStyleIndex,
                sharedFormulas,
                sharedStrings,
                styles,
                uses1904DateSystem,
                settings.IncludeDisplayText,
                sheet.Metadata,
                sheet.SourcePath,
                cancellationToken);
            cells.Add(new CellRecord(address, value, styleIndex));
            previousColumnIndex = columnIndex;
        }

        return new RowRecord(
            rowIndex,
            cells.ToArray(),
            height,
            row.Hidden?.Value ?? false,
            rowStyleIndex);
    }

    private static int GetRowIndex(
        Row row,
        Cell? firstCell,
        int inferredRowIndex,
        OpenXmlWorkbookReader.DiscoveredSheet sheet)
    {
        if (row.RowIndex?.Value is { } oneBasedRowIndex)
        {
            if (oneBasedRowIndex is 0 or > MaximumRowCount)
            {
                throw InvalidWorksheet(sheet, "A row index is outside Excel's worksheet limits.");
            }

            return checked((int)oneBasedRowIndex - 1);
        }

        if (firstCell?.CellReference?.Value is { Length: > 0 } reference)
        {
            if (!OpenXmlReaderAddress.TryParseCell(reference, out var address))
            {
                throw InvalidWorksheet(sheet, $"Cell reference '{reference}' is invalid.");
            }

            return address.RowIndex;
        }

        if (inferredRowIndex >= MaximumRowCount)
        {
            throw InvalidWorksheet(sheet, "An inferred row index is outside Excel's worksheet limits.");
        }

        return inferredRowIndex;
    }

    private static int GetColumnIndex(
        Cell cell,
        int rowIndex,
        int inferredColumnIndex,
        OpenXmlWorkbookReader.DiscoveredSheet sheet)
    {
        var reference = cell.CellReference?.Value;
        if (!string.IsNullOrEmpty(reference))
        {
            if (!OpenXmlReaderAddress.TryParseCell(reference, out var address) ||
                address.RowIndex != rowIndex)
            {
                throw InvalidWorksheet(sheet, $"Cell reference '{reference}' does not belong to its row.");
            }

            return address.ColumnIndex;
        }

        if (inferredColumnIndex >= MaximumColumnCount)
        {
            throw InvalidWorksheet(sheet, "An inferred column index is outside Excel's worksheet limits.");
        }

        return inferredColumnIndex;
    }

    private static int? GetStyleIndex(
        UInt32Value? value,
        OpenXmlReaderStyleCatalog styles,
        OpenXmlWorkbookReader.DiscoveredSheet sheet,
        string owner)
    {
        if (value is null)
        {
            return null;
        }

        if (value.Value > int.MaxValue)
        {
            throw InvalidWorksheet(sheet, $"A {owner} style index is too large.");
        }

        var styleIndex = checked((int)value.Value);
        if (!styles.IsValidStyleIndex(styleIndex))
        {
            throw InvalidWorksheet(sheet, $"A {owner} refers to missing style index {styleIndex}.");
        }

        return styleIndex;
    }

    private static void ReadColumnStyles(
        Columns columns,
        int?[] destination,
        ref Dictionary<int, double>? columnWidths,
        OpenXmlReaderStyleCatalog styles,
        OpenXmlWorkbookReader.DiscoveredSheet sheet,
        CancellationToken cancellationToken)
    {
        foreach (var column in columns.Elements<Column>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var minimum = column.Min?.Value;
            var maximum = column.Max?.Value;
            if (minimum is null or 0 ||
                maximum is null or 0 ||
                minimum > maximum ||
                maximum > MaximumColumnCount)
            {
                throw InvalidWorksheet(sheet, "The worksheet has an invalid column range.");
            }

            double? customWidth = null;
            if (column.CustomWidth?.Value == true && column.Width is { } width)
            {
                if (!double.IsFinite(width.Value) || width.Value < 0)
                {
                    throw InvalidWorksheet(sheet, "The worksheet has an invalid column width.");
                }

                customWidth = width.Value;
            }

            var styleIndex = GetStyleIndex(column.Style, styles, sheet, "column");
            if (styleIndex is null && customWidth is null)
            {
                continue;
            }

            for (var columnIndex = checked((int)minimum.Value - 1);
                 columnIndex < maximum.Value;
                 columnIndex++)
            {
                if ((columnIndex & 1023) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }

                if (styleIndex is { } style)
                {
                    destination[columnIndex] = style;
                }

                if (customWidth is { } columnWidth)
                {
                    (columnWidths ??= [])[columnIndex] = columnWidth;
                }
            }
        }
    }

    private static double? ReadDefaultColumnWidth(
        SheetFormatProperties format,
        OpenXmlWorkbookReader.DiscoveredSheet sheet)
    {
        var width = format.DefaultColumnWidth?.Value ?? format.BaseColumnWidth?.Value;
        if (width is null)
        {
            return null;
        }

        if (!double.IsFinite(width.Value) || width.Value < 0)
        {
            throw InvalidWorksheet(sheet, "The worksheet has an invalid default column width.");
        }

        return width.Value;
    }

    private static double?[]? BuildColumnWidthList(Dictionary<int, double>? columnWidths)
    {
        if (columnWidths is null || columnWidths.Count == 0)
        {
            return null;
        }

        var widths = new double?[columnWidths.Keys.Max() + 1];
        foreach (var (columnIndex, width) in columnWidths)
        {
            widths[columnIndex] = width;
        }

        return widths;
    }

    private static OpenXmlReaderException InvalidWorksheet(
        OpenXmlWorkbookReader.DiscoveredSheet sheet,
        string message,
        Exception? innerException = null) =>
        new(
            OpenXmlReaderError.InvalidWorksheet,
            $"Worksheet '{sheet.Metadata.Name}' is invalid: {message}",
            sheet.SourcePath,
            sheet.Metadata.Id,
            innerException);

    internal readonly record struct IndexedWorksheet(
        SheetMetadata Metadata,
        int[] StoredRowIndices,
        long CellCount);
}

internal sealed class OpenXmlReaderSharedFormulaResolver
{
    private readonly Dictionary<uint, SharedFormulaAnchor> _anchors = new();
    private readonly SheetMetadata _sheet;
    private readonly string _partPath;

    public OpenXmlReaderSharedFormulaResolver(SheetMetadata sheet, string partPath)
    {
        _sheet = sheet;
        _partPath = partPath;
    }

    public string Resolve(CellFormula formula, CellAddress address)
    {
        if (formula.FormulaType?.Value != CellFormulaValues.Shared)
        {
            return formula.Text ?? string.Empty;
        }

        var sharedIndex = formula.SharedIndex?.Value;
        if (sharedIndex is null)
        {
            throw Invalid("A shared formula does not declare a shared index.");
        }

        var formulaText = formula.Text ?? string.Empty;
        if (formulaText.Length > 0)
        {
            OpenXmlReaderRange? range = null;
            if (formula.Reference?.Value is { Length: > 0 } reference)
            {
                if (!OpenXmlReaderAddress.TryParseRange(reference, out var parsedRange))
                {
                    throw Invalid($"Shared formula range '{reference}' is invalid.");
                }

                range = parsedRange;
                if (!parsedRange.Contains(address))
                {
                    throw Invalid("A shared-formula anchor is outside its declared range.");
                }
            }

            if (!_anchors.TryAdd(
                sharedIndex.Value,
                new SharedFormulaAnchor(formulaText, address, range)))
            {
                throw Invalid($"Shared formula index {sharedIndex.Value} has multiple anchors.");
            }

            return formulaText;
        }

        if (!_anchors.TryGetValue(sharedIndex.Value, out var anchor))
        {
            throw Invalid($"Shared formula index {sharedIndex.Value} has no preceding anchor.");
        }

        if (anchor.Range is { } declaredRange && !declaredRange.Contains(address))
        {
            throw Invalid("A shared-formula cell is outside its declared range.");
        }

        return OpenXmlReaderFormulaTranslator.Translate(
            anchor.Formula,
            address.RowIndex - anchor.Address.RowIndex,
            address.ColumnIndex - anchor.Address.ColumnIndex);
    }

    private OpenXmlReaderException Invalid(string message) =>
        new(
            OpenXmlReaderError.InvalidWorksheet,
            $"Worksheet '{_sheet.Name}' is invalid: {message}",
            _partPath,
            _sheet.Id);

    private readonly record struct SharedFormulaAnchor(
        string Formula,
        CellAddress Address,
        OpenXmlReaderRange? Range);
}

internal static class OpenXmlReaderFormulaTranslator
{
    public static string Translate(string formula, int rowDelta, int columnDelta)
    {
        if (rowDelta == 0 && columnDelta == 0)
        {
            return formula;
        }

        var result = new StringBuilder(formula.Length);
        for (var index = 0; index < formula.Length;)
        {
            var current = formula[index];
            if (current == '"')
            {
                CopyQuoted(formula, result, ref index, '"', doubledEscape: true);
                continue;
            }

            if (current == '\'')
            {
                CopyQuoted(formula, result, ref index, '\'', doubledEscape: true);
                continue;
            }

            if (current == '[')
            {
                CopyBracketed(formula, result, ref index);
                continue;
            }

            if (TryReadReference(formula, index, out var reference))
            {
                var row = reference.AbsoluteRow
                    ? reference.RowIndex
                    : reference.RowIndex + rowDelta;
                var column = reference.AbsoluteColumn
                    ? reference.ColumnIndex
                    : reference.ColumnIndex + columnDelta;
                if (row is < 0 or >= 1_048_576 || column is < 0 or >= 16_384)
                {
                    result.Append("#REF!");
                }
                else
                {
                    if (reference.AbsoluteColumn)
                    {
                        result.Append('$');
                    }

                    OpenXmlReaderAddress.AppendColumnName(result, column);
                    if (reference.AbsoluteRow)
                    {
                        result.Append('$');
                    }

                    result.Append(row + 1);
                }

                index += reference.Length;
                continue;
            }

            result.Append(current);
            index++;
        }

        return result.ToString();
    }

    private static bool TryReadReference(string formula, int start, out FormulaReference reference)
    {
        reference = default;
        if (start > 0 && IsIdentifierCharacter(formula[start - 1]))
        {
            return false;
        }

        var cursor = start;
        var absoluteColumn = cursor < formula.Length && formula[cursor] == '$';
        if (absoluteColumn)
        {
            cursor++;
        }

        var lettersStart = cursor;
        while (cursor < formula.Length && char.IsAsciiLetter(formula[cursor]) && cursor - lettersStart < 3)
        {
            cursor++;
        }

        if (cursor == lettersStart ||
            cursor < formula.Length && char.IsAsciiLetter(formula[cursor]))
        {
            return false;
        }

        var absoluteRow = cursor < formula.Length && formula[cursor] == '$';
        if (absoluteRow)
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

        var columnIndex = 0;
        for (var index = lettersStart; index < (absoluteRow ? digitsStart - 1 : digitsStart); index++)
        {
            var letter = char.ToUpperInvariant(formula[index]);
            columnIndex = checked((columnIndex * 26) + (letter - 'A' + 1));
        }

        if (columnIndex is < 1 or > 16_384 ||
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
            columnIndex - 1,
            absoluteRow,
            absoluteColumn,
            cursor - start);
        return true;
    }

    private static void CopyQuoted(
        string source,
        StringBuilder destination,
        ref int index,
        char quote,
        bool doubledEscape)
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

            if (doubledEscape && index < source.Length && source[index] == quote)
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

    private static bool IsIdentifierCharacter(char value) =>
        char.IsAsciiLetterOrDigit(value) || value is '_' or '.' or '\\';

    private readonly record struct FormulaReference(
        int RowIndex,
        int ColumnIndex,
        bool AbsoluteRow,
        bool AbsoluteColumn,
        int Length);
}

internal readonly record struct OpenXmlReaderRange(CellAddress Start, CellAddress End)
{
    public bool Contains(CellAddress address) =>
        address.RowIndex >= Start.RowIndex &&
        address.RowIndex <= End.RowIndex &&
        address.ColumnIndex >= Start.ColumnIndex &&
        address.ColumnIndex <= End.ColumnIndex;
}

internal static class OpenXmlReaderAddress
{
    public static bool TryParseCell(string value, out CellAddress address)
    {
        address = default;
        if (string.IsNullOrEmpty(value))
        {
            return false;
        }

        var index = 0;
        if (value[index] == '$')
        {
            index++;
        }

        var columnStart = index;
        while (index < value.Length && char.IsAsciiLetter(value[index]))
        {
            index++;
        }

        var columnLength = index - columnStart;
        if (columnLength is < 1 or > 3)
        {
            return false;
        }

        if (index < value.Length && value[index] == '$')
        {
            index++;
        }

        var rowStart = index;
        while (index < value.Length && char.IsAsciiDigit(value[index]))
        {
            index++;
        }

        if (index != value.Length || rowStart == index)
        {
            return false;
        }

        var oneBasedColumn = 0;
        for (var letterIndex = columnStart; letterIndex < columnStart + columnLength; letterIndex++)
        {
            var letter = char.ToUpperInvariant(value[letterIndex]);
            oneBasedColumn = checked((oneBasedColumn * 26) + (letter - 'A' + 1));
        }

        if (oneBasedColumn is < 1 or > 16_384 ||
            !int.TryParse(
                value.AsSpan(rowStart),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out var oneBasedRow) ||
            oneBasedRow is < 1 or > 1_048_576)
        {
            return false;
        }

        address = new CellAddress(oneBasedRow - 1, oneBasedColumn - 1);
        return true;
    }

    public static bool TryParseRange(string value, out OpenXmlReaderRange range)
    {
        range = default;
        var separator = value.IndexOf(':');
        if (separator < 0)
        {
            if (!TryParseCell(value, out var single))
            {
                return false;
            }

            range = new OpenXmlReaderRange(single, single);
            return true;
        }

        if (value.LastIndexOf(':') != separator ||
            !TryParseCell(value[..separator], out var start) ||
            !TryParseCell(value[(separator + 1)..], out var end) ||
            start.RowIndex > end.RowIndex ||
            start.ColumnIndex > end.ColumnIndex)
        {
            return false;
        }

        range = new OpenXmlReaderRange(start, end);
        return true;
    }

    public static void AppendColumnName(StringBuilder destination, int columnIndex)
    {
        Span<char> buffer = stackalloc char[3];
        var cursor = buffer.Length;
        var value = columnIndex + 1;
        while (value > 0)
        {
            value--;
            buffer[--cursor] = (char)('A' + (value % 26));
            value /= 26;
        }

        destination.Append(buffer[cursor..]);
    }
}

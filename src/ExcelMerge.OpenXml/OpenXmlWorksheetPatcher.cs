using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelMerge.Domain;
using DomainCellValue = ExcelMerge.Domain.CellValue;
using SpreadsheetCellValue = DocumentFormat.OpenXml.Spreadsheet.CellValue;

namespace ExcelMerge.OpenXml;

internal static class OpenXmlWorksheetPatcher
{
    public static void Apply(
        WorkbookPart workbookPart,
        WorkbookPart? remoteWorkbookPart,
        OpenXmlWriterSharedStrings? remoteSharedStrings,
        OpenXmlStyleMapper? styleMapper,
        CompiledWorksheetPatch patch,
        string scratchPath,
        bool uses1904DateSystem,
        CancellationToken cancellationToken)
    {
        if (patch.LocalRelationshipId is null ||
            !workbookPart.TryGetPartById(patch.LocalRelationshipId, out var part) ||
            part is not WorksheetPart worksheetPart)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPlan,
                $"LOCAL worksheet relationship '{patch.LocalRelationshipId}' is missing.");
        }

        var remoteRows = LoadRemoteRows(
            remoteWorkbookPart,
            remoteSharedStrings,
            styleMapper,
            patch,
            cancellationToken);
        var rowTransform = new OpenXmlRowIndexTransform(patch.RowTransform);
        var remoteRowTransform = new SourceRowIndexTransform(patch, rowTransform);
        var seenRows = new HashSet<int>();
        var seenStructuralRows = new HashSet<int>();
        var nextRowIndex = 0;
        var insertionGroups = patch.RowTransform.Insertions.OrderBy(static pair => pair.Key).ToArray();
        var nextInsertionGroup = 0;
        var insideSheetData = false;
        var rewriteDimension = patch.RowTransform.HasChanges || patch.Rows.Count != 0;
        var hasSheetDimension = false;
        var extent = new WorksheetExtent();
        using (var reader = DocumentFormat.OpenXml.OpenXmlReader.Create(
            worksheetPart,
            readMiscNodes: true))
        using (var stream = new FileStream(
            scratchPath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 128 * 1024,
            FileOptions.SequentialScan))
        using (var writer = DocumentFormat.OpenXml.OpenXmlWriter.Create(stream, Encoding.UTF8))
        {
            writer.WriteStartDocument();
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.IsStartElement && reader.ElementType == typeof(SheetData))
                {
                    insideSheetData = true;
                    writer.WriteStartElement(reader);
                }
                else if (reader.IsStartElement && reader.ElementType == typeof(SheetDimension) &&
                    rewriteDimension)
                {
                    var dimension = reader.LoadCurrentElement() ?? throw InvalidWorksheet(
                        patch,
                        "The worksheet dimension could not be loaded.");
                    hasSheetDimension = true;
                    writer.WriteElement(dimension);
                }
                else if (reader.IsStartElement && reader.ElementType == typeof(Row))
                {
                    if (reader.LoadCurrentElement() is not Row row)
                    {
                        throw InvalidWorksheet(patch, "A row could not be loaded.");
                    }

                    var rowIndex = GetRowIndex(row, nextRowIndex, patch);
                    nextRowIndex = checked(rowIndex + 1);
                    EmitInsertions(
                        writer,
                        insertionGroups,
                        ref nextInsertionGroup,
                        rowIndex,
                        rowTransform,
                        remoteRows,
                        remoteRowTransform,
                        extent,
                        patch,
                        uses1904DateSystem,
                        cancellationToken);

                    RewriteRowFormulas(
                        row,
                        rowTransform.MapReferenceRow,
                        patch.LocalName);
                    if (patch.RowTransform.RemoteMetadataRows.TryGetValue(
                        rowIndex,
                        out var remoteMetadataRowIndex))
                    {
                        ApplyRemoteRowMetadata(row, remoteRows[remoteMetadataRowIndex]);
                    }
                    if (patch.Rows.TryGetValue(rowIndex, out var cellPatches))
                    {
                        ApplyRowPatches(
                            row,
                            rowIndex,
                            cellPatches,
                            remoteRowTransform,
                            patch,
                            uses1904DateSystem,
                            cancellationToken);
                        seenRows.Add(rowIndex);
                    }

                    if (patch.RowTransform.DeletedLocalRows.Contains(rowIndex))
                    {
                        seenStructuralRows.Add(rowIndex);
                        continue;
                    }

                    Row resultRow;
                    if (patch.RowTransform.Replacements.TryGetValue(rowIndex, out var replacement))
                    {
                        resultRow = CreateSourceRow(
                            replacement,
                            remoteRows,
                            remoteRowTransform,
                            patch,
                            rowTransform.MapLocalRow(rowIndex),
                            uses1904DateSystem,
                            cancellationToken);
                        seenStructuralRows.Add(rowIndex);
                    }
                    else
                    {
                        resultRow = row;
                        SetRowIndex(resultRow, rowTransform.MapLocalRow(rowIndex), patch);
                    }

                    extent.Include(resultRow, patch);
                    writer.WriteElement(resultRow);
                }
                else if (reader.IsEndElement && reader.ElementType == typeof(SheetData))
                {
                    EmitInsertions(
                        writer,
                        insertionGroups,
                        ref nextInsertionGroup,
                        int.MaxValue,
                        rowTransform,
                        remoteRows,
                        remoteRowTransform,
                        extent,
                        patch,
                        uses1904DateSystem,
                        cancellationToken);
                    insideSheetData = false;
                    writer.WriteEndElement();
                }
                else if (reader.IsStartElement && patch.RowTransform.HasChanges &&
                    OpenXmlCoordinateTransformer.IsInlineCarrier(reader.ElementType))
                {
                    var carrier = reader.LoadCurrentElement() ?? throw InvalidWorksheet(
                        patch,
                        $"Coordinate element '{reader.LocalName}' could not be loaded.");
                    writer.WriteElement(OpenXmlCoordinateTransformer.TransformInline(
                        carrier,
                        rowTransform.MapReferenceRow,
                        rowTransform.MapLocalRow,
                        patch.LocalName));
                }
                else if (reader.IsStartElement)
                {
                    writer.WriteStartElement(reader);
                }
                else if (reader.IsEndElement)
                {
                    writer.WriteEndElement();
                }
                else if (reader.IsMiscNode)
                {
                    writer.WriteString(reader.GetText());
                }
            }
        }

        if (insideSheetData || nextInsertionGroup != insertionGroups.Length)
        {
            TryDelete(scratchPath);
            throw InvalidWorksheet(patch, "The worksheet has incomplete sheet data.");
        }

        var missingRow = patch.Rows.Keys.FirstOrDefault(rowIndex => !seenRows.Contains(rowIndex), -1);
        if (missingRow >= 0)
        {
            TryDelete(scratchPath);
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                $"Worksheet group '{patch.SheetGroupId}' requires creating LOCAL row {missingRow}.");
        }


        var missingStructuralRow = patch.RowTransform.DeletedLocalRows
            .Concat(patch.RowTransform.Replacements.Keys)
            .FirstOrDefault(rowIndex => !seenStructuralRows.Contains(rowIndex), -1);
        if (missingStructuralRow >= 0)
        {
            TryDelete(scratchPath);
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPlan,
                $"Worksheet group '{patch.SheetGroupId}' refers to missing LOCAL row {missingStructuralRow}.");
        }

        FeedWorksheetData(worksheetPart, scratchPath);
        if (hasSheetDimension)
        {
            RewriteSheetDimension(
                worksheetPart,
                scratchPath,
                extent.Reference,
                patch,
                cancellationToken);
        }

        if (patch.RowTransform.HasChanges)
        {
            OpenXmlCoordinateTransformer.TransformRelatedParts(
                worksheetPart,
                rowTransform.MapReferenceRow,
                rowTransform.MapLocalRow,
                patch.LocalName);
        }
    }

    private static void ApplyRowPatches(
        Row row,
        int rowIndex,
        IReadOnlyDictionary<int, CompiledCellPatch> patches,
        SourceRowIndexTransform remoteRowTransform,
        CompiledWorksheetPatch worksheetPatch,
        bool uses1904DateSystem,
        CancellationToken cancellationToken)
    {
        var cellsByColumn = new Dictionary<int, Cell>();
        var previousColumnIndex = -1;
        foreach (var cell in row.Elements<Cell>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var columnIndex = GetColumnIndex(cell, rowIndex, previousColumnIndex + 1, worksheetPatch);
            if (columnIndex <= previousColumnIndex || !cellsByColumn.TryAdd(columnIndex, cell))
            {
                throw InvalidWorksheet(worksheetPatch, "Cells are duplicated or out of order.");
            }

            previousColumnIndex = columnIndex;
        }

        foreach (var pair in patches.OrderBy(static pair => pair.Key))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var columnIndex = pair.Key;
            if (columnIndex is < 0 or >= 16_384)
            {
                throw InvalidWorksheet(worksheetPatch, "A result cell is outside worksheet limits.");
            }

            if (!cellsByColumn.TryGetValue(columnIndex, out var cell))
            {
                cell = new Cell
                {
                    CellReference = ToCellReference(rowIndex, columnIndex),
                };
                InsertCell(row, cell, columnIndex, cellsByColumn);
                cellsByColumn.Add(columnIndex, cell);
            }

            if (cell.CellMetaIndex is not null || cell.ValueMetaIndex is not null)
            {
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.UnsupportedOperation,
                    $"Cell {cell.CellReference?.Value} has unsupported metadata references.");
            }

            WriteValue(cell, pair.Value.Value, uses1904DateSystem);
            if (pair.Value.Source == CompiledCellSourceKind.Remote &&
                cell.CellFormula is { } formula)
            {
                RewriteFormula(
                    formula,
                    remoteRowTransform.Map,
                    worksheetPatch.RemoteName);
            }
        }
    }

    internal static void WriteValue(
        Cell cell,
        DomainCellValue? requestedValue,
        bool uses1904DateSystem)
    {
        cell.RemoveAllChildren<CellFormula>();
        cell.RemoveAllChildren<SpreadsheetCellValue>();
        cell.RemoveAllChildren<InlineString>();
        cell.DataType = null;

        if (!requestedValue.HasValue || requestedValue.Value.Kind == CellKind.Blank)
        {
            return;
        }

        var value = requestedValue.Value;
        if (value.IsFormula)
        {
            cell.CellFormula = new CellFormula(value.Formula ?? string.Empty);
            if (value.CachedValue is { } cachedValue)
            {
                WriteScalar(cell, cachedValue, uses1904DateSystem, formulaCache: true);
            }

            return;
        }

        WriteScalar(cell, value.Scalar.GetValueOrDefault(), uses1904DateSystem, formulaCache: false);
    }

    private static IReadOnlyDictionary<int, Row> LoadRemoteRows(
        WorkbookPart? remoteWorkbookPart,
        OpenXmlWriterSharedStrings? remoteSharedStrings,
        OpenXmlStyleMapper? styleMapper,
        CompiledWorksheetPatch patch,
        CancellationToken cancellationToken)
    {
        if (!patch.RowTransform.RequiresRemoteRows)
        {
            return new Dictionary<int, Row>();
        }

        if (remoteWorkbookPart is null || remoteSharedStrings is null || styleMapper is null || patch.RemoteRelationshipId is null ||
            !remoteWorkbookPart.TryGetPartById(patch.RemoteRelationshipId, out var part) ||
            part is not WorksheetPart worksheetPart)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPlan,
                $"Worksheet group '{patch.SheetGroupId}' has no readable REMOTE worksheet.");
        }

        var requested = patch.RowTransform.Replacements.Values
            .Concat(patch.RowTransform.Insertions.Values.SelectMany(static rows => rows))
            .Where(static source => source.Kind == CompiledRowSourceKind.Remote)
            .Select(static source => source.RemoteRowIndex!.Value)
            .Concat(patch.RowTransform.RemoteMetadataRows.Values)
            .ToHashSet();
        var rows = new Dictionary<int, Row>();
        var nextRowIndex = 0;
        using var reader = DocumentFormat.OpenXml.OpenXmlReader.Create(worksheetPart);
        while (reader.Read() && rows.Count < requested.Count)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!reader.IsStartElement || reader.ElementType != typeof(Row))
            {
                continue;
            }

            if (reader.LoadCurrentElement() is not Row row)
            {
                throw InvalidWorksheet(patch, "A REMOTE row could not be loaded.");
            }

            var rowIndex = GetRowIndex(row, nextRowIndex, patch);
            nextRowIndex = checked(rowIndex + 1);
            if (!requested.Contains(rowIndex))
            {
                continue;
            }

            OpenXmlImportedWorksheetTransformer.TransformRow(
                row,
                remoteSharedStrings,
                styleMapper,
                cancellationToken);
            rows.Add(rowIndex, row);
        }

        var missing = requested.FirstOrDefault(rowIndex => !rows.ContainsKey(rowIndex), -1);
        if (missing >= 0)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPlan,
                $"Worksheet group '{patch.SheetGroupId}' refers to missing REMOTE row {missing}.");
        }

        return rows;
    }

    private static void ApplyRemoteRowMetadata(Row localRow, Row remoteRow)
    {
        localRow.Height = remoteRow.Height?.Value;
        localRow.Hidden = remoteRow.Hidden?.Value;
        localRow.StyleIndex = remoteRow.StyleIndex?.Value;
        localRow.CustomFormat = remoteRow.StyleIndex is not null;
        localRow.CustomHeight = remoteRow.Height is not null;
        localRow.OutlineLevel = remoteRow.OutlineLevel?.Value;
        localRow.Collapsed = remoteRow.Collapsed?.Value;
        localRow.ThickTop = remoteRow.ThickTop?.Value;
        localRow.ThickBot = remoteRow.ThickBot?.Value;
        localRow.ShowPhonetic = remoteRow.ShowPhonetic?.Value;
    }

    private static void EmitInsertions(
        DocumentFormat.OpenXml.OpenXmlWriter writer,
        KeyValuePair<int, IReadOnlyList<CompiledRowSource>>[] insertionGroups,
        ref int nextGroup,
        int maximumBoundary,
        OpenXmlRowIndexTransform transform,
        IReadOnlyDictionary<int, Row> remoteRows,
        SourceRowIndexTransform remoteRowTransform,
        WorksheetExtent extent,
        CompiledWorksheetPatch patch,
        bool uses1904DateSystem,
        CancellationToken cancellationToken)
    {
        while (nextGroup < insertionGroups.Length && insertionGroups[nextGroup].Key <= maximumBoundary)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var group = insertionGroups[nextGroup++];
            for (var index = 0; index < group.Value.Count; index++)
            {
                var resultRowIndex = transform.MapInsertion(group.Key, index);
                var row = CreateSourceRow(
                    group.Value[index],
                    remoteRows,
                    remoteRowTransform,
                    patch,
                    resultRowIndex,
                    uses1904DateSystem,
                    cancellationToken);
                extent.Include(row, patch);
                writer.WriteElement(row);
            }
        }
    }

    private static Row CreateSourceRow(
        CompiledRowSource source,
        IReadOnlyDictionary<int, Row> remoteRows,
        SourceRowIndexTransform remoteRowTransform,
        CompiledWorksheetPatch patch,
        int resultRowIndex,
        bool uses1904DateSystem,
        CancellationToken cancellationToken)
    {
        Row row;
        if (source.Kind == CompiledRowSourceKind.Remote)
        {
            row = (Row)remoteRows[source.RemoteRowIndex!.Value].CloneNode(deep: true);
            RewriteRowFormulas(row, remoteRowTransform.Map, patch.RemoteName);
        }
        else
        {
            var custom = source.CustomRow ?? throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPlan,
                "A custom row instruction has no row value.");
            row = new Row
            {
                Height = custom.Height,
                Hidden = custom.IsHidden,
                StyleIndex = custom.StyleIndex is { } styleIndex ? (uint)styleIndex : null,
                CustomFormat = custom.StyleIndex.HasValue,
                CustomHeight = custom.Height.HasValue,
            };
            foreach (var sourceCell in custom.Cells.Span)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var cell = new Cell
                {
                    StyleIndex = sourceCell.StyleIndex is { } cellStyle ? (uint)cellStyle : null,
                };
                WriteValue(cell, sourceCell.Value, uses1904DateSystem);
                row.Append(cell);
            }
        }

        SetRowIndex(row, resultRowIndex, patch: null);
        return row;
    }

    private static void RewriteRowFormulas(
        Row row,
        Func<int, int?> rowMapper,
        string? sourceSheetName)
    {
        foreach (var formula in row.Descendants<CellFormula>())
        {
            RewriteFormula(formula, rowMapper, sourceSheetName);
        }
    }

    private static void RewriteFormula(
        CellFormula formula,
        Func<int, int?> rowMapper,
        string? sourceSheetName)
    {
        if (!string.IsNullOrEmpty(formula.Text))
        {
            formula.Text = OpenXmlFormulaRowRewriter.Rewrite(
                formula.Text,
                rowMapper,
                sourceSheetName);
        }

        if (formula.Reference?.Value is { Length: > 0 } reference)
        {
            formula.Reference = OpenXmlFormulaRowRewriter.Rewrite(
                reference,
                rowMapper,
                sourceSheetName);
        }
    }

    private static void SetRowIndex(Row row, int rowIndex, CompiledWorksheetPatch? patch)
    {
        if (rowIndex is < 0 or >= 1_048_576)
        {
            throw patch is null
                ? new OpenXmlWriterException(
                    OpenXmlWriterError.UnsupportedOperation,
                    "A structural merge exceeds the worksheet row limit.")
                : InvalidWorksheet(patch, "A structural merge exceeds the worksheet row limit.");
        }

        row.RowIndex = checked((uint)rowIndex + 1);
        var inferredColumn = 0;
        foreach (var cell in row.Elements<Cell>())
        {
            var columnIndex = inferredColumn;
            if (cell.CellReference?.Value is { Length: > 0 } reference &&
                OpenXmlReaderAddress.TryParseCell(reference, out var address))
            {
                columnIndex = address.ColumnIndex;
            }

            cell.CellReference = ToCellReference(rowIndex, columnIndex);
            inferredColumn = checked(columnIndex + 1);
        }
    }

    private sealed class SourceRowIndexTransform
    {
        private readonly Dictionary<int, int> _mappedRows = [];
        private readonly HashSet<int> _omittedRows = [];
        private readonly int[] _anchors;

        public SourceRowIndexTransform(
            CompiledWorksheetPatch patch,
            OpenXmlRowIndexTransform localTransform)
        {
            foreach (var mapping in patch.RowMappings.Span)
            {
                if (mapping.RemoteRowIndex is not { } remoteRowIndex)
                {
                    continue;
                }

                if (mapping.LocalRowIndex is { } localRowIndex &&
                    !patch.RowTransform.DeletedLocalRows.Contains(localRowIndex))
                {
                    _mappedRows[remoteRowIndex] = localTransform.MapLocalRow(localRowIndex);
                }
                else
                {
                    _omittedRows.Add(remoteRowIndex);
                }
            }

            foreach (var replacement in patch.RowTransform.Replacements)
            {
                if (replacement.Value is
                    {
                        Kind: CompiledRowSourceKind.Remote,
                        RemoteRowIndex: { } remoteRowIndex,
                    })
                {
                    _mappedRows[remoteRowIndex] = localTransform.MapLocalRow(replacement.Key);
                    _omittedRows.Remove(remoteRowIndex);
                }
            }

            foreach (var insertionGroup in patch.RowTransform.Insertions)
            {
                for (var index = 0; index < insertionGroup.Value.Count; index++)
                {
                    if (insertionGroup.Value[index] is
                        {
                            Kind: CompiledRowSourceKind.Remote,
                            RemoteRowIndex: { } remoteRowIndex,
                        })
                    {
                        _mappedRows[remoteRowIndex] = localTransform.MapInsertion(
                            insertionGroup.Key,
                            index);
                        _omittedRows.Remove(remoteRowIndex);
                    }
                }
            }

            _anchors = _mappedRows.Keys.Order().ToArray();
        }

        public int? Map(int rowIndex)
        {
            if (_mappedRows.TryGetValue(rowIndex, out var mapped))
            {
                return mapped;
            }

            if (_omittedRows.Contains(rowIndex) || _anchors.Length == 0)
            {
                return null;
            }

            var insertion = Array.BinarySearch(_anchors, rowIndex);
            if (insertion >= 0)
            {
                return _mappedRows[_anchors[insertion]];
            }

            insertion = ~insertion;
            var anchor = insertion > 0
                ? _anchors[insertion - 1]
                : _anchors[0];
            return checked(rowIndex + (_mappedRows[anchor] - anchor));
        }
    }

    private static void WriteScalar(
        Cell cell,
        CellScalar scalar,
        bool uses1904DateSystem,
        bool formulaCache)
    {
        switch (scalar.Kind)
        {
            case CellKind.Blank:
                return;
            case CellKind.Text:
                if (formulaCache)
                {
                    cell.DataType = CellValues.String;
                    cell.CellValue = new SpreadsheetCellValue(scalar.TextValue ?? string.Empty);
                }
                else
                {
                    cell.DataType = CellValues.InlineString;
                    cell.InlineString = new InlineString(CreateText(scalar.TextValue ?? string.Empty));
                }

                return;
            case CellKind.Number:
                cell.DataType = CellValues.Number;
                cell.CellValue = new SpreadsheetCellValue(
                    scalar.NumberValue.GetValueOrDefault().ToString("R", CultureInfo.InvariantCulture));
                return;
            case CellKind.Boolean:
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new SpreadsheetCellValue(
                    scalar.BooleanValue.GetValueOrDefault() ? "1" : "0");
                return;
            case CellKind.DateTime:
                cell.DataType = CellValues.Number;
                cell.CellValue = new SpreadsheetCellValue(
                    ToSerialDate(scalar.DateTimeValue.GetValueOrDefault(), uses1904DateSystem)
                        .ToString("R", CultureInfo.InvariantCulture));
                return;
            case CellKind.Error:
                cell.DataType = CellValues.Error;
                cell.CellValue = new SpreadsheetCellValue(scalar.TextValue ?? "#VALUE!");
                return;
            default:
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.InvalidPlan,
                    $"Cell kind '{scalar.Kind}' cannot be written as a scalar value.");
        }
    }

    private static Text CreateText(string value)
    {
        var text = new Text(value);
        if (value.Length != value.Trim().Length || value.Contains("  ", StringComparison.Ordinal))
        {
            text.Space = SpaceProcessingModeValues.Preserve;
        }

        return text;
    }

    private static double ToSerialDate(DateTime value, bool uses1904DateSystem) =>
        uses1904DateSystem
            ? (value - new DateTime(1904, 1, 1)).TotalDays
            : value.ToOADate();

    private static int GetRowIndex(Row row, int inferredRowIndex, CompiledWorksheetPatch patch)
    {
        if (row.RowIndex?.Value is { } oneBasedRowIndex)
        {
            if (oneBasedRowIndex is 0 or > 1_048_576)
            {
                throw InvalidWorksheet(patch, "A row index is outside worksheet limits.");
            }

            return checked((int)oneBasedRowIndex - 1);
        }

        var firstReference = row.Elements<Cell>().FirstOrDefault()?.CellReference?.Value;
        if (!string.IsNullOrEmpty(firstReference) &&
            OpenXmlReaderAddress.TryParseCell(firstReference, out var address))
        {
            return address.RowIndex;
        }

        return inferredRowIndex;
    }

    private static int GetColumnIndex(
        Cell cell,
        int rowIndex,
        int inferredColumnIndex,
        CompiledWorksheetPatch patch)
    {
        var reference = cell.CellReference?.Value;
        if (string.IsNullOrEmpty(reference))
        {
            return inferredColumnIndex;
        }

        if (!OpenXmlReaderAddress.TryParseCell(reference, out var address) || address.RowIndex != rowIndex)
        {
            throw InvalidWorksheet(patch, $"Cell reference '{reference}' is invalid for its row.");
        }

        return address.ColumnIndex;
    }

    private static void InsertCell(
        Row row,
        Cell cell,
        int columnIndex,
        IReadOnlyDictionary<int, Cell> cellsByColumn)
    {
        var following = cellsByColumn
            .Where(pair => pair.Key > columnIndex)
            .OrderBy(static pair => pair.Key)
            .Select(static pair => pair.Value)
            .FirstOrDefault();
        if (following is null)
        {
            row.Append(cell);
        }
        else
        {
            row.InsertBefore(cell, following);
        }
    }

    private static string ToCellReference(int rowIndex, int columnIndex)
    {
        var builder = new StringBuilder(10);
        OpenXmlReaderAddress.AppendColumnName(builder, columnIndex);
        builder.Append(rowIndex + 1);
        return builder.ToString();
    }

    private static void FeedWorksheetData(WorksheetPart worksheetPart, string scratchPath)
    {
        using var input = new FileStream(
            scratchPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 128 * 1024,
            FileOptions.SequentialScan);
        worksheetPart.FeedData(input);
    }

    private static void RewriteSheetDimension(
        WorksheetPart worksheetPart,
        string scratchPath,
        string reference,
        CompiledWorksheetPatch patch,
        CancellationToken cancellationToken)
    {
        var found = false;
        using (var reader = DocumentFormat.OpenXml.OpenXmlReader.Create(
            worksheetPart,
            readMiscNodes: true))
        using (var stream = new FileStream(
            scratchPath,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 128 * 1024,
            FileOptions.SequentialScan))
        using (var writer = DocumentFormat.OpenXml.OpenXmlWriter.Create(stream, Encoding.UTF8))
        {
            writer.WriteStartDocument();
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.IsStartElement && reader.ElementType == typeof(SheetDimension))
                {
                    var dimension = reader.LoadCurrentElement() as SheetDimension ?? throw InvalidWorksheet(
                        patch,
                        "The worksheet dimension could not be reloaded.");
                    dimension.Reference = reference;
                    writer.WriteElement(dimension);
                    found = true;
                }
                else if (reader.IsStartElement &&
                    typeof(OpenXmlLeafTextElement).IsAssignableFrom(reader.ElementType))
                {
                    var leaf = reader.LoadCurrentElement() ?? throw InvalidWorksheet(
                        patch,
                        $"Leaf element '{reader.LocalName}' could not be reloaded.");
                    writer.WriteElement(leaf);
                }
                else if (reader.IsStartElement)
                {
                    writer.WriteStartElement(reader);
                }
                else if (reader.IsEndElement)
                {
                    writer.WriteEndElement();
                }
                else if (reader.IsMiscNode)
                {
                    writer.WriteString(reader.GetText());
                }
            }
        }

        if (!found)
        {
            throw InvalidWorksheet(patch, "The worksheet dimension disappeared during transformation.");
        }

        FeedWorksheetData(worksheetPart, scratchPath);
    }

    private sealed class WorksheetExtent
    {
        private int _firstRow = int.MaxValue;
        private int _lastRow = -1;
        private int _firstColumn = int.MaxValue;
        private int _lastColumn = -1;

        public string Reference
        {
            get
            {
                if (_lastRow < 0)
                {
                    return "A1";
                }

                var first = ToCellReference(_firstRow, _firstColumn);
                var last = ToCellReference(_lastRow, _lastColumn);
                return string.Equals(first, last, StringComparison.Ordinal)
                    ? first
                    : $"{first}:{last}";
            }
        }

        public void Include(Row row, CompiledWorksheetPatch patch)
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var reference = cell.CellReference?.Value;
                if (string.IsNullOrEmpty(reference) ||
                    !OpenXmlReaderAddress.TryParseCell(reference, out var address))
                {
                    throw InvalidWorksheet(
                        patch,
                        $"Cell reference '{cell.CellReference?.Value}' is invalid after transformation.");
                }

                _firstRow = Math.Min(_firstRow, address.RowIndex);
                _lastRow = Math.Max(_lastRow, address.RowIndex);
                _firstColumn = Math.Min(_firstColumn, address.ColumnIndex);
                _lastColumn = Math.Max(_lastColumn, address.ColumnIndex);
            }
        }
    }

    private static OpenXmlWriterException InvalidWorksheet(
        CompiledWorksheetPatch patch,
        string message) =>
        new(
            OpenXmlWriterError.InvalidPackage,
            $"LOCAL worksheet group '{patch.SheetGroupId}' is invalid: {message}");

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch
        {
            // The operation cleanup retries after the package has been closed.
        }
    }
}

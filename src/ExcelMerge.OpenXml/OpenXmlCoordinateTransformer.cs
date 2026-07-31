using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using ThreadedComments = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace ExcelMerge.OpenXml;

internal static class OpenXmlCoordinateTransformer
{
    public static bool IsInlineCarrier(Type elementType) =>
        elementType == typeof(MergeCells) ||
        elementType == typeof(ConditionalFormatting) ||
        elementType == typeof(DataValidations) ||
        elementType == typeof(Hyperlinks) ||
        elementType == typeof(AutoFilter) ||
        elementType == typeof(SheetViews) ||
        elementType == typeof(RowBreaks) ||
        elementType == typeof(ProtectedRanges) ||
        elementType == typeof(IgnoredErrors) ||
        elementType == typeof(CellWatches);

    public static OpenXmlElement TransformInline(
        OpenXmlElement element,
        Func<int, int?> referenceRowMapper,
        Func<int, int> anchorRowMapper,
        string? sourceSheetName)
    {
        switch (element)
        {
            case MergeCells mergeCells:
                foreach (var mergeCell in mergeCells.Elements<MergeCell>())
                {
                    mergeCell.Reference = TransformReference(
                        mergeCell.Reference?.Value,
                        referenceRowMapper);
                }

                break;
            case ConditionalFormatting conditionalFormatting:
                TransformListValue(conditionalFormatting.SequenceOfReferences, referenceRowMapper);
                RewriteFormulaText(
                    conditionalFormatting.Descendants<Formula>(),
                    referenceRowMapper,
                    sourceSheetName);
                break;
            case DataValidations dataValidations:
                foreach (var validation in dataValidations.Elements<DataValidation>())
                {
                    TransformListValue(validation.SequenceOfReferences, referenceRowMapper);
                    RewriteFormulaText(
                        validation.Elements<Formula1>().Cast<OpenXmlLeafTextElement>()
                            .Concat(validation.Elements<Formula2>()),
                        referenceRowMapper,
                        sourceSheetName);
                }

                break;
            case Hyperlinks hyperlinks:
                foreach (var hyperlink in hyperlinks.Elements<Hyperlink>())
                {
                    hyperlink.Reference = TransformReference(
                        hyperlink.Reference?.Value,
                        referenceRowMapper);
                    if (!string.IsNullOrEmpty(hyperlink.Location?.Value))
                    {
                        hyperlink.Location = OpenXmlFormulaRowRewriter.Rewrite(
                            hyperlink.Location.Value,
                            referenceRowMapper,
                            sourceSheetName);
                    }
                }

                break;
            case AutoFilter autoFilter:
                autoFilter.Reference = TransformReference(
                    autoFilter.Reference?.Value,
                    referenceRowMapper);
                TransformSortStates(autoFilter, referenceRowMapper);
                break;
            case SheetViews sheetViews:
                foreach (var pane in sheetViews.Descendants<Pane>())
                {
                    pane.TopLeftCell = TransformReference(
                        pane.TopLeftCell?.Value,
                        anchorRow => anchorRowMapper(anchorRow));
                }

                foreach (var selection in sheetViews.Descendants<Selection>())
                {
                    selection.ActiveCell = TransformReference(
                        selection.ActiveCell?.Value,
                        anchorRow => anchorRowMapper(anchorRow));
                    TransformListValue(
                        selection.SequenceOfReferences,
                        anchorRow => anchorRowMapper(anchorRow));
                }

                break;
            case RowBreaks rowBreaks:
                foreach (var pageBreak in rowBreaks.Elements<Break>())
                {
                    if (pageBreak.Id?.Value is { } oneBasedRow && oneBasedRow > 0)
                    {
                        pageBreak.Id = checked((uint)anchorRowMapper(checked((int)oneBasedRow - 1)) + 1U);
                    }
                }

                break;
            case ProtectedRanges protectedRanges:
                foreach (var range in protectedRanges.Elements<ProtectedRange>())
                {
                    TransformListValue(range.SequenceOfReferences, referenceRowMapper);
                }

                break;
            case IgnoredErrors ignoredErrors:
                foreach (var error in ignoredErrors.Elements<IgnoredError>())
                {
                    TransformListValue(error.SequenceOfReferences, referenceRowMapper);
                }

                break;
            case CellWatches cellWatches:
                foreach (var watch in cellWatches.Elements<CellWatch>())
                {
                    watch.CellReference = TransformReference(
                        watch.CellReference?.Value,
                        referenceRowMapper);
                }

                break;
        }

        TransformSortStates(element, referenceRowMapper);
        return element;
    }

    public static void TransformRelatedParts(
        WorksheetPart worksheetPart,
        Func<int, int?> referenceRowMapper,
        Func<int, int> anchorRowMapper,
        string? sourceSheetName)
    {
        if (worksheetPart.PivotTableParts.Any() ||
            worksheetPart.QueryTableParts.Any() ||
            worksheetPart.SingleCellTablePart is not null)
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                "Pivot, query, and single-cell tables cannot yet participate in a row transform.");
        }

        foreach (var tablePart in worksheetPart.TableDefinitionParts)
        {
            var table = tablePart.Table;
            table.Reference = TransformReference(table.Reference?.Value, referenceRowMapper);
            if (table.AutoFilter is { } autoFilter)
            {
                autoFilter.Reference = TransformReference(
                    autoFilter.Reference?.Value,
                    referenceRowMapper);
                TransformSortStates(autoFilter, referenceRowMapper);
            }

            TransformSortStates(table, referenceRowMapper);
            table.Save();
        }

        if (worksheetPart.WorksheetCommentsPart?.Comments is { } comments)
        {
            foreach (var comment in comments.CommentList?.Elements<Comment>().ToArray() ?? [])
            {
                var transformed = TransformReferenceOrNull(
                    comment.Reference?.Value,
                    referenceRowMapper);
                if (transformed is null)
                    comment.Remove();
                else
                    comment.Reference = transformed;
            }

            comments.Save();
        }

        foreach (var threadedPart in worksheetPart.WorksheetThreadedCommentsParts)
        {
            var root = threadedPart.RootElement;
            if (root is null)
                continue;
            foreach (var comment in root.Descendants<ThreadedComments.ThreadedComment>().ToArray())
            {
                var transformed = TransformReferenceOrNull(comment.Ref?.Value, referenceRowMapper);
                if (transformed is null)
                    comment.Remove();
                else
                    comment.Ref = transformed;
            }

            root.Save();
        }

        if (worksheetPart.DrawingsPart?.WorksheetDrawing is { } drawing)
        {
            foreach (var rowId in drawing.Descendants<DrawingSpreadsheet.RowId>())
            {
                if (int.TryParse(
                    rowId.Text,
                    NumberStyles.None,
                    CultureInfo.InvariantCulture,
                    out var rowIndex))
                {
                    rowId.Text = anchorRowMapper(rowIndex).ToString(CultureInfo.InvariantCulture);
                }
            }

            drawing.Save();
        }

        foreach (var vmlPart in worksheetPart.VmlDrawingParts)
        {
            TransformVml(vmlPart, anchorRowMapper);
        }
    }

    private static void TransformVml(VmlDrawingPart part, Func<int, int> anchorRowMapper)
    {
        XDocument document;
        using (var input = part.GetStream(FileMode.Open, FileAccess.Read))
        {
            document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        }

        foreach (var anchor in document.Descendants().Where(static element => element.Name.LocalName == "Anchor"))
        {
            var values = (anchor.Value ?? string.Empty).Split(',');
            if (values.Length < 8 ||
                !int.TryParse(values[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var fromRow) ||
                !int.TryParse(values[6].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var toRow))
            {
                throw new OpenXmlWriterException(
                    OpenXmlWriterError.UnsupportedOperation,
                    "A VML drawing contains an unsupported anchor.");
            }

            values[2] = anchorRowMapper(fromRow).ToString(CultureInfo.InvariantCulture);
            values[6] = anchorRowMapper(toRow).ToString(CultureInfo.InvariantCulture);
            anchor.Value = string.Join(',', values);
        }

        foreach (var row in document.Descendants().Where(static element =>
            element.Name.LocalName == "Row" &&
            element.Name.NamespaceName.Contains("excel", StringComparison.OrdinalIgnoreCase)))
        {
            if (int.TryParse(row.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex))
            {
                row.Value = anchorRowMapper(rowIndex).ToString(CultureInfo.InvariantCulture);
            }
        }

        using var output = part.GetStream(FileMode.Create, FileAccess.Write);
        document.Save(output, SaveOptions.DisableFormatting);
    }

    private static void TransformSortStates(
        OpenXmlElement element,
        Func<int, int?> rowMapper)
    {
        foreach (var sortState in element.Descendants<SortState>())
        {
            sortState.Reference = TransformReference(sortState.Reference?.Value, rowMapper);
        }
    }

    private static void RewriteFormulaText(
        IEnumerable<OpenXmlLeafTextElement> formulas,
        Func<int, int?> rowMapper,
        string? sourceSheetName)
    {
        foreach (var formula in formulas)
        {
            if (!string.IsNullOrEmpty(formula.Text))
            {
                formula.Text = OpenXmlFormulaRowRewriter.Rewrite(
                    formula.Text,
                    rowMapper,
                    sourceSheetName);
            }
        }
    }

    private static void TransformListValue(
        ListValue<StringValue>? references,
        Func<int, int?> rowMapper)
    {
        if (references is null)
        {
            return;
        }

        references.InnerText = string.Join(
            ' ',
            references.Items.Select(reference => TransformReference(reference.Value, rowMapper)));
    }

    private static string? TransformReferenceOrNull(
        string? reference,
        Func<int, int?> rowMapper)
    {
        if (string.IsNullOrWhiteSpace(reference) ||
            !OpenXmlReaderAddress.TryParseCell(reference, out var address))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                $"Coordinate reference '{reference}' is invalid.");
        }

        return rowMapper(address.RowIndex) is { } mappedRow
            ? BuildReference(address.ColumnIndex, mappedRow)
            : null;
    }

    private static string TransformReference(
        string? reference,
        Func<int, int?> rowMapper)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                "A coordinate-bearing element has no reference.");
        }

        var transformed = OpenXmlFormulaRowRewriter.Rewrite(
            reference,
            rowMapper,
            sourceSheetName: null);
        if (transformed.Contains("#REF!", StringComparison.Ordinal))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.UnsupportedOperation,
                $"Row deletion invalidates coordinate range '{reference}'.");
        }

        return transformed;
    }

    private static string BuildReference(int columnIndex, int rowIndex)
    {
        var builder = new System.Text.StringBuilder(10);
        OpenXmlReaderAddress.AppendColumnName(builder, columnIndex);
        builder.Append(rowIndex + 1);
        return builder.ToString();
    }
}

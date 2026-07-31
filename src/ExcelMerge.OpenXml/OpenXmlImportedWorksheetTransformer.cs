using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelMerge.OpenXml;

internal static class OpenXmlImportedWorksheetTransformer
{
    public static void Transform(
        WorksheetPart worksheetPart,
        OpenXmlWriterSharedStrings sharedStrings,
        OpenXmlStyleMapper styleMapper,
        string scratchPath,
        CancellationToken cancellationToken)
    {
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
                if (reader.IsStartElement && reader.ElementType == typeof(Columns))
                {
                    if (reader.LoadCurrentElement() is not Columns columns)
                    {
                        throw InvalidWorksheet("A column collection could not be loaded.");
                    }
                    foreach (var column in columns.Elements<Column>())
                    {
                        if (column.Style?.Value is { } styleIndex)
                        {
                            column.Style = styleMapper.MapCellFormat(styleIndex);
                        }
                    }

                    writer.WriteElement(columns);
                }
                else if (reader.IsStartElement && reader.ElementType == typeof(Row))
                {
                    if (reader.LoadCurrentElement() is not Row row)
                    {
                        throw InvalidWorksheet("A row could not be loaded.");
                    }
                    TransformRow(row, sharedStrings, styleMapper, cancellationToken);

                    writer.WriteElement(row);
                }
                else if (reader.IsStartElement && reader.ElementType == typeof(ConditionalFormatting))
                {
                    if (reader.LoadCurrentElement() is not ConditionalFormatting conditionalFormatting)
                    {
                        throw InvalidWorksheet("Conditional formatting could not be loaded.");
                    }
                    foreach (var rule in conditionalFormatting.Elements<ConditionalFormattingRule>())
                    {
                        if (rule.FormatId?.Value is { } formatId)
                        {
                            rule.FormatId = styleMapper.MapDifferentialFormat(formatId);
                        }
                    }

                    writer.WriteElement(conditionalFormatting);
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

        using var input = new FileStream(
            scratchPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 128 * 1024,
            FileOptions.SequentialScan);
        worksheetPart.FeedData(input);
        styleMapper.Save();
    }

    internal static void TransformRow(
        Row row,
        OpenXmlWriterSharedStrings sharedStrings,
        OpenXmlStyleMapper styleMapper,
        CancellationToken cancellationToken)
    {
        if (row.StyleIndex?.Value is { } rowStyleIndex)
        {
            row.StyleIndex = styleMapper.MapCellFormat(rowStyleIndex);
        }

        foreach (var cell in row.Elements<Cell>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (cell.StyleIndex?.Value is { } cellStyleIndex)
            {
                cell.StyleIndex = styleMapper.MapCellFormat(cellStyleIndex);
            }

            ConvertSharedString(cell, sharedStrings);
        }
    }

    private static void ConvertSharedString(Cell cell, OpenXmlWriterSharedStrings sharedStrings)
    {
        if (cell.DataType?.Value != CellValues.SharedString)
        {
            return;
        }

        if (cell.CellFormula is not null ||
            !int.TryParse(
                cell.CellValue?.Text,
                NumberStyles.None,
                CultureInfo.InvariantCulture,
            out var index))
        {
            throw new OpenXmlWriterException(
                OpenXmlWriterError.InvalidPackage,
                $"REMOTE cell '{cell.CellReference?.Value}' has an invalid shared-string reference.");
        }

        var item = sharedStrings.Read(index, cell.CellReference?.Value);
        var inlineString = new InlineString();
        foreach (var child in item.ChildElements)
        {
            inlineString.Append(child.CloneNode(deep: true));
        }

        cell.RemoveAllChildren<DocumentFormat.OpenXml.Spreadsheet.CellValue>();
        cell.DataType = CellValues.InlineString;
        cell.InlineString = inlineString;
    }

    private static OpenXmlWriterException InvalidWorksheet(string message) =>
        new(OpenXmlWriterError.InvalidPackage, $"The imported REMOTE worksheet is invalid: {message}");
}

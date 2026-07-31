using System.IO.Compression;
using System.Text;
using System.Xml;

namespace ExcelMerge.Rewrite.Benchmarks;

public enum XlsxCorpusLayout
{
    Dense,
    Sparse,
}

public enum XlsxSharedStringDistribution
{
    Repeated,
    Unique,
}

internal sealed class XlsxBenchmarkCorpus : IDisposable
{
    private const string ContentTypesNamespace =
        "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string PackageRelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string OfficeRelationshipsNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string XmlnsNamespace = "http://www.w3.org/2000/xmlns/";
    private const int RepeatedSharedStringCount = 256;

    private static readonly int[] DenseColumnIndices =
        [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19];

    private static readonly int[] SparseColumnIndices =
        [0, 8, 31, 63, 127, 255, 511, 1023, 2047, 4095];

    private string? _directoryPath;

    private XlsxBenchmarkCorpus(
        string directoryPath,
        string filePath,
        int cellCount,
        int storedRowCount,
        int worksheetCount)
    {
        _directoryPath = directoryPath;
        FilePath = filePath;
        CellCount = cellCount;
        StoredRowCount = storedRowCount;
        WorksheetCount = worksheetCount;
    }

    public string DirectoryPath => _directoryPath ?? throw new ObjectDisposedException(nameof(XlsxBenchmarkCorpus));

    public string FilePath { get; }

    public int CellCount { get; }

    public int StoredRowCount { get; }

    public int WorksheetCount { get; }

    public static XlsxBenchmarkCorpus Create(
        int cellCount,
        XlsxCorpusLayout layout,
        XlsxSharedStringDistribution distribution,
        int worksheetCount = 1)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(cellCount);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(worksheetCount);
        if (!Enum.IsDefined(layout))
        {
            throw new ArgumentOutOfRangeException(nameof(layout));
        }

        if (!Enum.IsDefined(distribution))
        {
            throw new ArgumentOutOfRangeException(nameof(distribution));
        }

        var geometry = CorpusGeometry.Create(layout);
        var storedRowCount = checked((cellCount + geometry.CellsPerRow - 1) / geometry.CellsPerRow);
        var lastRowNumber = checked(((storedRowCount - 1) * geometry.RowStride) + 1);
        if (lastRowNumber > 1_048_576)
        {
            throw new ArgumentOutOfRangeException(
                nameof(cellCount),
                "The requested sparse corpus exceeds the XLSX worksheet row limit.");
        }

        var directoryPath = BenchmarkFileSystem.CreateTemporaryDirectory("excelmerge-xlsx-corpus-");
        var filePath = Path.Combine(directoryPath, "corpus.xlsx");
        try
        {
            WritePackage(
                filePath,
                cellCount,
                geometry,
                distribution,
                worksheetCount,
                storedRowCount,
                lastRowNumber);
            return new XlsxBenchmarkCorpus(
                directoryPath,
                filePath,
                cellCount,
                storedRowCount,
                worksheetCount);
        }
        catch
        {
            try
            {
                BenchmarkFileSystem.DeleteDirectory(directoryPath);
            }
            catch
            {
                // Preserve the corpus-generation failure.
            }

            throw;
        }
    }

    public void Dispose()
    {
        var directoryPath = _directoryPath;
        if (directoryPath is null)
        {
            return;
        }

        BenchmarkFileSystem.DeleteDirectory(directoryPath);
        _directoryPath = null;
    }

    private static void WritePackage(
        string filePath,
        int cellCount,
        CorpusGeometry geometry,
        XlsxSharedStringDistribution distribution,
        int worksheetCount,
        int storedRowCount,
        int lastRowNumber)
    {
        using var stream = new FileStream(
            filePath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 1024 * 1024,
            FileOptions.SequentialScan);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false);

        WriteContentTypes(archive, worksheetCount);
        WriteRootRelationships(archive);
        WriteWorkbook(archive, worksheetCount);
        WriteWorkbookRelationships(archive, worksheetCount);
        WriteStyles(archive);
        WriteSharedStrings(archive, cellCount, distribution);
        WriteDataWorksheet(
            archive,
            cellCount,
            geometry,
            distribution,
            storedRowCount,
            lastRowNumber);

        for (var sheetNumber = 2; sheetNumber <= worksheetCount; sheetNumber++)
        {
            WriteEmptyWorksheet(archive, sheetNumber);
        }
    }

    private static void WriteContentTypes(ZipArchive archive, int worksheetCount) =>
        WriteXmlEntry(archive, "[Content_Types].xml", writer =>
        {
            writer.WriteStartElement("Types", ContentTypesNamespace);
            WriteContentTypeDefault(
                writer,
                "rels",
                "application/vnd.openxmlformats-package.relationships+xml");
            WriteContentTypeDefault(writer, "xml", "application/xml");
            WriteContentTypeOverride(
                writer,
                "/xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            for (var sheetNumber = 1; sheetNumber <= worksheetCount; sheetNumber++)
            {
                WriteContentTypeOverride(
                    writer,
                    $"/xl/worksheets/sheet{sheetNumber}.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            }

            WriteContentTypeOverride(
                writer,
                "/xl/sharedStrings.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
            WriteContentTypeOverride(
                writer,
                "/xl/styles.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            writer.WriteEndElement();
        });

    private static void WriteRootRelationships(ZipArchive archive) =>
        WriteXmlEntry(archive, "_rels/.rels", writer =>
        {
            writer.WriteStartElement("Relationships", PackageRelationshipsNamespace);
            WriteRelationship(
                writer,
                "rId1",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
                "xl/workbook.xml");
            writer.WriteEndElement();
        });

    private static void WriteWorkbook(ZipArchive archive, int worksheetCount) =>
        WriteXmlEntry(archive, "xl/workbook.xml", writer =>
        {
            writer.WriteStartElement("workbook", SpreadsheetNamespace);
            writer.WriteAttributeString("xmlns", "r", XmlnsNamespace, OfficeRelationshipsNamespace);
            writer.WriteStartElement("workbookPr", SpreadsheetNamespace);
            writer.WriteAttributeString("date1904", "0");
            writer.WriteEndElement();
            writer.WriteStartElement("bookViews", SpreadsheetNamespace);
            writer.WriteStartElement("workbookView", SpreadsheetNamespace);
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteStartElement("sheets", SpreadsheetNamespace);
            for (var sheetNumber = 1; sheetNumber <= worksheetCount; sheetNumber++)
            {
                writer.WriteStartElement("sheet", SpreadsheetNamespace);
                writer.WriteAttributeString(
                    "name",
                    sheetNumber == 1 ? "Data" : $"Metadata {sheetNumber:D2}");
                WriteIntegerAttribute(writer, "sheetId", sheetNumber);
                writer.WriteAttributeString(
                    "r",
                    "id",
                    OfficeRelationshipsNamespace,
                    $"rId{sheetNumber}");
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteStartElement("calcPr", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "calcId", 191029);
            writer.WriteEndElement();
            writer.WriteEndElement();
        });

    private static void WriteWorkbookRelationships(ZipArchive archive, int worksheetCount) =>
        WriteXmlEntry(archive, "xl/_rels/workbook.xml.rels", writer =>
        {
            writer.WriteStartElement("Relationships", PackageRelationshipsNamespace);
            for (var sheetNumber = 1; sheetNumber <= worksheetCount; sheetNumber++)
            {
                WriteRelationship(
                    writer,
                    $"rId{sheetNumber}",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    $"worksheets/sheet{sheetNumber}.xml");
            }

            WriteRelationship(
                writer,
                $"rId{worksheetCount + 1}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                "sharedStrings.xml");
            WriteRelationship(
                writer,
                $"rId{worksheetCount + 2}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                "styles.xml");
            writer.WriteEndElement();
        });

    private static void WriteStyles(ZipArchive archive) =>
        WriteXmlEntry(archive, "xl/styles.xml", writer =>
        {
            writer.WriteStartElement("styleSheet", SpreadsheetNamespace);

            writer.WriteStartElement("fonts", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "count", 1);
            writer.WriteStartElement("font", SpreadsheetNamespace);
            WriteValueElement(writer, "sz", "val", 11);
            writer.WriteStartElement("color", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "theme", 1);
            writer.WriteEndElement();
            writer.WriteStartElement("name", SpreadsheetNamespace);
            writer.WriteAttributeString("val", "Calibri");
            writer.WriteEndElement();
            WriteValueElement(writer, "family", "val", 2);
            writer.WriteStartElement("scheme", SpreadsheetNamespace);
            writer.WriteAttributeString("val", "minor");
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("fills", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "count", 2);
            WritePatternFill(writer, "none");
            WritePatternFill(writer, "gray125");
            writer.WriteEndElement();

            writer.WriteStartElement("borders", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "count", 1);
            writer.WriteStartElement("border", SpreadsheetNamespace);
            foreach (var edge in new[] { "left", "right", "top", "bottom", "diagonal" })
            {
                writer.WriteStartElement(edge, SpreadsheetNamespace);
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("cellStyleXfs", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "count", 1);
            WriteXf(writer, includeXfId: false);
            writer.WriteEndElement();
            writer.WriteStartElement("cellXfs", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "count", 1);
            WriteXf(writer, includeXfId: true);
            writer.WriteEndElement();

            writer.WriteStartElement("cellStyles", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "count", 1);
            writer.WriteStartElement("cellStyle", SpreadsheetNamespace);
            writer.WriteAttributeString("name", "Normal");
            WriteIntegerAttribute(writer, "xfId", 0);
            WriteIntegerAttribute(writer, "builtinId", 0);
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.WriteStartElement("dxfs", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "count", 0);
            writer.WriteEndElement();
            writer.WriteStartElement("tableStyles", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "count", 0);
            writer.WriteAttributeString("defaultTableStyle", "TableStyleMedium2");
            writer.WriteAttributeString("defaultPivotStyle", "PivotStyleLight16");
            writer.WriteEndElement();
            writer.WriteEndElement();
        });

    private static void WriteSharedStrings(
        ZipArchive archive,
        int cellCount,
        XlsxSharedStringDistribution distribution)
    {
        var uniqueCount = distribution == XlsxSharedStringDistribution.Unique
            ? cellCount
            : Math.Min(cellCount, RepeatedSharedStringCount);
        WriteXmlEntry(archive, "xl/sharedStrings.xml", writer =>
        {
            writer.WriteStartElement("sst", SpreadsheetNamespace);
            WriteIntegerAttribute(writer, "count", cellCount);
            WriteIntegerAttribute(writer, "uniqueCount", uniqueCount);
            var prefix = distribution == XlsxSharedStringDistribution.Unique
                ? "unique-"
                : "repeated-";
            for (var index = 0; index < uniqueCount; index++)
            {
                writer.WriteStartElement("si", SpreadsheetNamespace);
                writer.WriteStartElement("t", SpreadsheetNamespace);
                writer.WriteString(prefix);
                writer.WriteValue(index);
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            writer.WriteEndElement();
        });
    }

    private static void WriteDataWorksheet(
        ZipArchive archive,
        int cellCount,
        CorpusGeometry geometry,
        XlsxSharedStringDistribution distribution,
        int storedRowCount,
        int lastRowNumber)
    {
        var columnNames = geometry.ColumnIndices.Select(ToColumnName).ToArray();
        var widestColumnName = columnNames[Math.Min(cellCount, geometry.CellsPerRow) - 1];
        WriteXmlEntry(archive, "xl/worksheets/sheet1.xml", writer =>
        {
            writer.WriteStartElement("worksheet", SpreadsheetNamespace);
            writer.WriteStartElement("dimension", SpreadsheetNamespace);
            writer.WriteStartAttribute(null, "ref", null);
            writer.WriteString("A1:");
            writer.WriteString(widestColumnName);
            writer.WriteValue(lastRowNumber);
            writer.WriteEndAttribute();
            writer.WriteEndElement();
            WriteSheetPreamble(writer);
            writer.WriteStartElement("sheetData", SpreadsheetNamespace);

            var remainingCellCount = cellCount;
            var cellOrdinal = 0;
            for (var storedRowIndex = 0; storedRowIndex < storedRowCount; storedRowIndex++)
            {
                var rowNumber = checked((storedRowIndex * geometry.RowStride) + 1);
                var cellsInRow = Math.Min(geometry.CellsPerRow, remainingCellCount);
                writer.WriteStartElement("row", SpreadsheetNamespace);
                WriteIntegerAttribute(writer, "r", rowNumber);
                for (var cellIndex = 0; cellIndex < cellsInRow; cellIndex++)
                {
                    writer.WriteStartElement("c", SpreadsheetNamespace);
                    writer.WriteStartAttribute(null, "r", null);
                    writer.WriteString(columnNames[cellIndex]);
                    writer.WriteValue(rowNumber);
                    writer.WriteEndAttribute();
                    writer.WriteAttributeString("t", "s");
                    writer.WriteStartElement("v", SpreadsheetNamespace);
                    writer.WriteValue(distribution == XlsxSharedStringDistribution.Unique
                        ? cellOrdinal
                        : cellOrdinal % RepeatedSharedStringCount);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    cellOrdinal++;
                }

                writer.WriteEndElement();
                remainingCellCount -= cellsInRow;
            }

            writer.WriteEndElement();
            writer.WriteEndElement();
        });
    }

    private static void WriteEmptyWorksheet(ZipArchive archive, int sheetNumber) =>
        WriteXmlEntry(archive, $"xl/worksheets/sheet{sheetNumber}.xml", writer =>
        {
            writer.WriteStartElement("worksheet", SpreadsheetNamespace);
            WriteSheetPreamble(writer);
            writer.WriteStartElement("sheetData", SpreadsheetNamespace);
            writer.WriteEndElement();
            writer.WriteEndElement();
        });

    private static void WriteSheetPreamble(XmlWriter writer)
    {
        writer.WriteStartElement("sheetViews", SpreadsheetNamespace);
        writer.WriteStartElement("sheetView", SpreadsheetNamespace);
        WriteIntegerAttribute(writer, "workbookViewId", 0);
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteStartElement("sheetFormatPr", SpreadsheetNamespace);
        writer.WriteAttributeString("defaultRowHeight", "15");
        writer.WriteEndElement();
    }

    private static void WriteContentTypeDefault(
        XmlWriter writer,
        string extension,
        string contentType)
    {
        writer.WriteStartElement("Default", ContentTypesNamespace);
        writer.WriteAttributeString("Extension", extension);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteContentTypeOverride(
        XmlWriter writer,
        string partName,
        string contentType)
    {
        writer.WriteStartElement("Override", ContentTypesNamespace);
        writer.WriteAttributeString("PartName", partName);
        writer.WriteAttributeString("ContentType", contentType);
        writer.WriteEndElement();
    }

    private static void WriteRelationship(
        XmlWriter writer,
        string id,
        string type,
        string target)
    {
        writer.WriteStartElement("Relationship", PackageRelationshipsNamespace);
        writer.WriteAttributeString("Id", id);
        writer.WriteAttributeString("Type", type);
        writer.WriteAttributeString("Target", target);
        writer.WriteEndElement();
    }

    private static void WritePatternFill(XmlWriter writer, string patternType)
    {
        writer.WriteStartElement("fill", SpreadsheetNamespace);
        writer.WriteStartElement("patternFill", SpreadsheetNamespace);
        writer.WriteAttributeString("patternType", patternType);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteXf(XmlWriter writer, bool includeXfId)
    {
        writer.WriteStartElement("xf", SpreadsheetNamespace);
        WriteIntegerAttribute(writer, "numFmtId", 0);
        WriteIntegerAttribute(writer, "fontId", 0);
        WriteIntegerAttribute(writer, "fillId", 0);
        WriteIntegerAttribute(writer, "borderId", 0);
        if (includeXfId)
        {
            WriteIntegerAttribute(writer, "xfId", 0);
        }

        writer.WriteEndElement();
    }

    private static void WriteValueElement(
        XmlWriter writer,
        string elementName,
        string attributeName,
        int value)
    {
        writer.WriteStartElement(elementName, SpreadsheetNamespace);
        WriteIntegerAttribute(writer, attributeName, value);
        writer.WriteEndElement();
    }

    private static void WriteIntegerAttribute(XmlWriter writer, string name, int value)
    {
        writer.WriteStartAttribute(null, name, null);
        writer.WriteValue(value);
        writer.WriteEndAttribute();
    }

    private static void WriteXmlEntry(
        ZipArchive archive,
        string entryName,
        Action<XmlWriter> writeContent)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            CheckCharacters = true,
            CloseOutput = false,
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
        });
        writer.WriteStartDocument();
        writeContent(writer);
        writer.WriteEndDocument();
    }

    private static string ToColumnName(int columnIndex)
    {
        Span<char> buffer = stackalloc char[3];
        var position = buffer.Length;
        var remaining = checked(columnIndex + 1);
        do
        {
            remaining--;
            buffer[--position] = (char)('A' + (remaining % 26));
            remaining /= 26;
        }
        while (remaining > 0);

        return new string(buffer[position..]);
    }

    private readonly record struct CorpusGeometry(
        int CellsPerRow,
        int RowStride,
        int[] ColumnIndices)
    {
        public static CorpusGeometry Create(XlsxCorpusLayout layout) => layout switch
        {
            XlsxCorpusLayout.Dense => new CorpusGeometry(20, 1, DenseColumnIndices),
            XlsxCorpusLayout.Sparse => new CorpusGeometry(10, 2, SparseColumnIndices),
            _ => throw new ArgumentOutOfRangeException(nameof(layout)),
        };
    }
}

internal static class BenchmarkFileSystem
{
    private const int CleanupRetryCount = 5;

    public static string CreateTemporaryDirectory(string prefix)
    {
        var baseDirectory = Path.Combine(Path.GetTempPath(), "ExcelMerge.Rewrite.Benchmarks");
        Directory.CreateDirectory(baseDirectory);
        for (var attempt = 0; attempt < 10; attempt++)
        {
            var path = Path.Combine(baseDirectory, prefix + Guid.NewGuid().ToString("N"));
            if (Directory.Exists(path) || File.Exists(path))
            {
                continue;
            }

            Directory.CreateDirectory(path);
            return path;
        }

        throw new IOException("A unique benchmark directory could not be allocated.");
    }

    public static void DeleteDirectory(string path)
    {
        for (var attempt = 0; ; attempt++)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    Directory.Delete(path, recursive: true);
                }

                return;
            }
            catch (Exception exception) when (
                (exception is IOException or UnauthorizedAccessException) &&
                attempt < CleanupRetryCount)
            {
                Thread.Sleep(TimeSpan.FromMilliseconds(50 * (attempt + 1)));
            }
        }
    }
}

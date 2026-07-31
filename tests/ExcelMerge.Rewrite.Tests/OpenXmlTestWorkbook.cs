using System.IO.Compression;
using System.Text;

namespace ExcelMerge.Rewrite.Tests;

internal static class OpenXmlTestWorkbook
{
    public const string EmptyWorksheet =
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
        "<sheetData/>" +
        "</worksheet>";

    public static string Create(
        TestDirectory directory,
        string worksheetXml,
        string? sharedStringsXml = null,
        string? stylesXml = null,
        string fileName = "fixture.xlsx",
        string sheetName = "Data",
        string? sheetState = null,
        bool uses1904DateSystem = false)
    {
        var path = directory.GetPath(fileName);
        using var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create);

        var sharedStringContentType = sharedStringsXml is null
            ? string.Empty
            : "<Override PartName=\"/xl/sharedStrings.xml\" " +
                "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
        var stylesContentType = stylesXml is null
            ? string.Empty
            : "<Override PartName=\"/xl/styles.xml\" " +
                "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
        AddEntry(
            archive,
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/xl/workbook.xml\" " +
                "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
            "<Override PartName=\"/xl/worksheets/sheet1.xml\" " +
                "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
            sharedStringContentType +
            stylesContentType +
            "</Types>");
        AddEntry(
            archive,
            "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" " +
                "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" " +
                "Target=\"xl/workbook.xml\"/>" +
            "</Relationships>");

        var stateAttribute = sheetState is null ? string.Empty : $" state=\"{sheetState}\"";
        AddEntry(
            archive,
            "xl/workbook.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            $"<workbookPr date1904=\"{(uses1904DateSystem ? 1 : 0)}\"/>" +
            "<sheets>" +
            $"<sheet name=\"{sheetName}\" sheetId=\"1\" r:id=\"rId1\"{stateAttribute}/>" +
            "</sheets>" +
            "</workbook>");

        var sharedStringRelationship = sharedStringsXml is null
            ? string.Empty
            : "<Relationship Id=\"rId2\" " +
                "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" " +
                "Target=\"sharedStrings.xml\"/>";
        var stylesRelationship = stylesXml is null
            ? string.Empty
            : "<Relationship Id=\"rId3\" " +
                "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" " +
                "Target=\"styles.xml\"/>";
        AddEntry(
            archive,
            "xl/_rels/workbook.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" " +
                "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" " +
                "Target=\"worksheets/sheet1.xml\"/>" +
            sharedStringRelationship +
            stylesRelationship +
            "</Relationships>");
        AddEntry(archive, "xl/worksheets/sheet1.xml", worksheetXml);

        if (sharedStringsXml is not null)
        {
            AddEntry(archive, "xl/sharedStrings.xml", sharedStringsXml);
        }

        if (stylesXml is not null)
        {
            AddEntry(archive, "xl/styles.xml", stylesXml);
        }

        return path;
    }

    private static void AddEntry(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
        using var writer = new StreamWriter(
            entry.Open(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }
}

internal sealed class SynchronousProgress<T>(Action<T> callback) : IProgress<T>
{
    public void Report(T value) => callback(value);
}

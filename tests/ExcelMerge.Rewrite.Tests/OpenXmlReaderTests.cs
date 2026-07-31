using ExcelMerge.Domain;
using ExcelMerge.Engine;
using ExcelMerge.OpenXml;
using ExcelMerge.Storage;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Rewrite.Tests;

[TestClass]
public sealed class OpenXmlReaderTests
{
    [TestMethod]
    public async Task Reader_discovers_metadata_and_indexes_typed_sparse_rows()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData>" +
            "<row r=\"1\">" +
            "<c r=\"A1\" t=\"s\"><v>0</v></c>" +
            "<c r=\"B1\"><v>42.5</v></c>" +
            "<c r=\"C1\" t=\"b\"><v>1</v></c>" +
            "<c r=\"D1\" t=\"inlineStr\"><is><r><t>rich </t></r><r><t>inline</t></r></is></c>" +
            "<c r=\"E1\"><f>B1+1</f><v>43.5</v></c>" +
            "</row>" +
            "<row r=\"4\" hidden=\"1\" ht=\"24\">" +
            "<c r=\"C4\" t=\"e\"><v>#N/A</v></c>" +
            "</row>" +
            "</sheetData>" +
            "</worksheet>",
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                "count=\"1\" uniqueCount=\"1\"><si><t>shared</t></si></sst>",
            stylesXml: null,
            sheetState: "hidden",
            uses1904DateSystem: true);
        var reader = new OpenXmlWorkbookReader();
        var progressReports = new List<OpenXmlReaderProgress>();

        var metadata = await reader.ReadMetadataAsync(path);

        Assert.AreEqual("fixture.xlsx", metadata.DisplayName);
        Assert.AreEqual(WorkbookFormat.Xlsx, metadata.Format);
        Assert.AreEqual(Path.GetFullPath(path), metadata.SourcePath);
        Assert.AreEqual(new FileInfo(path).Length, metadata.ContentLength);
        Assert.IsTrue(metadata.Uses1904DateSystem);
        Assert.AreEqual(SheetVisibility.Hidden, metadata.Sheets.Single().Visibility);

        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
            RowCacheCapacity = 0,
            RowCacheByteLimit = 0,
        });
        var result = await reader.IndexAsync(
            path,
            workspace,
            progress: new SynchronousProgress<OpenXmlReaderProgress>(progressReports.Add));
        var worksheet = result.Worksheets.Single();

        Assert.AreEqual(2, worksheet.StoredRowCount);
        CollectionAssert.AreEqual(new[] { 0, 3 }, worksheet.StoredRowIndices.ToArray());
        Assert.AreEqual(0, worksheet.Metadata.FirstRowIndex);
        Assert.AreEqual(3, worksheet.Metadata.LastRowIndex);
        Assert.AreEqual(0, worksheet.Metadata.FirstColumnIndex);
        Assert.AreEqual(4, worksheet.Metadata.LastColumnIndex);
        Assert.AreEqual(6L, worksheet.Metadata.NonEmptyCellCount);

        var firstRow = await worksheet.GetRowAsync(0);
        var fourthRow = await worksheet.GetRowAsync(3);
        Assert.IsTrue(firstRow.HasValue);
        Assert.IsTrue(fourthRow.HasValue);
        Assert.IsNull(await worksheet.GetRowAsync(1));

        var firstCells = firstRow.Value.Cells.ToArray();
        Assert.AreEqual("shared", firstCells[0].Value.Scalar?.TextValue);
        Assert.AreEqual(42.5, firstCells[1].Value.Scalar?.NumberValue);
        Assert.AreEqual(true, firstCells[2].Value.Scalar?.BooleanValue);
        Assert.AreEqual("rich inline", firstCells[3].Value.Scalar?.TextValue);
        Assert.IsTrue(firstCells[4].Value.IsFormula);
        Assert.AreEqual("B1+1", firstCells[4].Value.Formula);
        Assert.AreEqual(43.5, firstCells[4].Value.CachedValue?.NumberValue);
        Assert.AreEqual("#N/A", fourthRow.Value.Cells.Span[0].Value.Scalar?.TextValue);
        Assert.IsTrue(fourthRow.Value.IsHidden);
        Assert.AreEqual(24d, fourthRow.Value.Height);

        var batch = await worksheet.ReadRowsAsync(0, 4);
        Assert.AreEqual(4, batch.RowCount);
        Assert.AreEqual(2, batch.Rows.Length);
        Assert.IsTrue(progressReports.Any(report => report.Stage == OpenXmlReaderStage.ReadingSharedStrings));
        Assert.AreEqual(OpenXmlReaderStage.Completed, progressReports[^1].Stage);
    }

    [TestMethod]
    public async Task Reader_translates_shared_formulas_with_absolute_references()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData>" +
            "<row r=\"1\">" +
            "<c r=\"A1\"><v>1</v></c>" +
            "<c r=\"B1\"><f t=\"shared\" si=\"0\" ref=\"B1:B2\">" +
                "A$1+$A1+$A$1+&quot;A1&quot;</f><v>4</v></c>" +
            "</row>" +
            "<row r=\"2\"><c r=\"B2\"><f t=\"shared\" si=\"0\"/><v>5</v></c></row>" +
            "</sheetData>" +
            "</worksheet>");
        var reader = new OpenXmlWorkbookReader();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });

        var result = await reader.IndexAsync(path, workspace);
        var anchor = await result.Worksheets[0].GetRowAsync(0);
        var translated = await result.Worksheets[0].GetRowAsync(1);

        Assert.IsTrue(anchor.HasValue);
        Assert.IsTrue(translated.HasValue);
        Assert.AreEqual("A$1+$A1+$A$1+\"A1\"", anchor.Value.Cells.Span[1].Value.Formula);
        Assert.AreEqual("A$1+$A2+$A$1+\"A1\"", translated.Value.Cells.Span[0].Value.Formula);
    }

    [TestMethod]
    public async Task Reader_applies_sheet_selection_and_can_omit_display_text()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>value</v></c></row></sheetData>" +
            "</worksheet>");
        var reader = new OpenXmlWorkbookReader();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });

        var result = await reader.IndexAsync(
            path,
            workspace,
            new OpenXmlReaderOptions
            {
                SheetNames = ["data"],
                IncludeDisplayText = false,
            });
        var row = await result.Worksheets.Single().GetRowAsync(0);

        Assert.IsTrue(row.HasValue);
        Assert.AreEqual("value", row.Value.Cells.Span[0].Value.Scalar?.TextValue);
        Assert.IsNull(row.Value.Cells.Span[0].Value.DisplayText);

        var exception = await Assert.ThrowsExceptionAsync<OpenXmlReaderException>(() =>
            reader.IndexAsync(
                path,
                workspace,
                new OpenXmlReaderOptions { SheetNames = ["missing"] }).AsTask());
        Assert.AreEqual(OpenXmlReaderError.SheetNotFound, exception.Error);
    }

    [TestMethod]
    public async Task Reader_rejects_unsupported_and_non_zip_inputs()
    {
        using var temporaryDirectory = new TestDirectory();
        var reader = new OpenXmlWorkbookReader();
        var fakeXlsx = temporaryDirectory.GetPath("fake.xlsx");
        await File.WriteAllTextAsync(fakeXlsx, "not a zip package");

        var legacy = await Assert.ThrowsExceptionAsync<OpenXmlReaderException>(() =>
            reader.ReadMetadataAsync(temporaryDirectory.GetPath("legacy.xls")).AsTask());
        var delimited = await Assert.ThrowsExceptionAsync<OpenXmlReaderException>(() =>
            reader.ReadMetadataAsync(temporaryDirectory.GetPath("data.csv")).AsTask());
        var notZip = await Assert.ThrowsExceptionAsync<OpenXmlReaderException>(() =>
            reader.ReadMetadataAsync(fakeXlsx).AsTask());

        Assert.AreEqual(OpenXmlReaderError.LegacyBinaryWorkbook, legacy.Error);
        Assert.AreEqual(OpenXmlReaderError.UnsupportedFileExtension, delimited.Error);
        Assert.AreEqual(OpenXmlReaderError.NotZipPackage, notZip.Error);
    }

    [TestMethod]
    public async Task Reader_removes_partial_stores_when_worksheet_indexing_fails()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>" +
            "</worksheet>");
        var reader = new OpenXmlWorkbookReader();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });

        var exception = await Assert.ThrowsExceptionAsync<OpenXmlReaderException>(() =>
            reader.IndexAsync(path, workspace).AsTask());

        Assert.AreEqual(OpenXmlReaderError.InvalidWorksheet, exception.Error);
        Assert.AreEqual(0, Directory.GetDirectories(workspace.DirectoryPath).Length);
    }

    [TestMethod]
    public async Task Reader_detects_source_timestamp_changes_and_cleans_up()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            OpenXmlTestWorkbook.EmptyWorksheet);
        var originalTimestamp = File.GetLastWriteTimeUtc(path);
        var changed = false;
        var progress = new SynchronousProgress<OpenXmlReaderProgress>(report =>
        {
            if (!changed && report.Stage == OpenXmlReaderStage.DiscoveringWorkbook)
            {
                File.SetLastWriteTimeUtc(path, originalTimestamp.AddMinutes(5));
                changed = true;
            }
        });
        var reader = new OpenXmlWorkbookReader();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });

        var exception = await Assert.ThrowsExceptionAsync<OpenXmlReaderException>(() =>
            reader.IndexAsync(path, workspace, progress: progress).AsTask());

        Assert.IsTrue(changed);
        Assert.AreEqual(OpenXmlReaderError.SourceChanged, exception.Error);
        Assert.AreEqual(0, Directory.GetDirectories(workspace.DirectoryPath).Length);
    }

    [TestMethod]
    public async Task Indexed_worksheets_flow_directly_into_the_diff_engine()
    {
        using var temporaryDirectory = new TestDirectory();
        var localPath = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>local</v></c></row></sheetData>" +
            "</worksheet>",
            fileName: "local.xlsx");
        var remotePath = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>remote</v></c></row></sheetData>" +
            "</worksheet>",
            fileName: "remote.xlsx");
        var reader = new OpenXmlWorkbookReader();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });

        var local = await reader.IndexAsync(localPath, workspace);
        var remote = await reader.IndexAsync(remotePath, workspace);
        var diff = await new TwoWayDiffEngine().CompareAsync(
            local.Worksheets.Single(),
            remote.Worksheets.Single(),
            SourceSide.Remote);

        Assert.AreEqual(ChangeKind.Modified, diff.SheetChange.Kind);
        Assert.AreEqual(SourceSide.Remote, diff.SheetChange.Source);
        Assert.AreEqual(1L, diff.Statistics.ChangedCellCount);
        Assert.AreEqual("local", diff.SheetChange.RowChanges.Span[0].CellChanges.Span[0].BaseValue?.Scalar?.TextValue);
        Assert.AreEqual("remote", diff.SheetChange.RowChanges.Span[0].CellChanges.Span[0].SourceValue?.Scalar?.TextValue);
    }
}

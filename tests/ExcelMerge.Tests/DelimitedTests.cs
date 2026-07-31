using System.Text;
using ExcelMerge.Delimited;
using ExcelMerge.Domain;
using ExcelMerge.Engine;
using ExcelMerge.Storage;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class DelimitedTests
{
    [TestMethod]
    public async Task Reader_parses_rfc4180_records_and_indexes_contiguous_rows()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = temporaryDirectory.GetPath("sample.csv");
        await File.WriteAllTextAsync(
            path,
            "alpha,\"b,b\",\"line1\r\nline2\",\"say \"\"hi\"\"\",\r\n\r\nlast",
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        var reader = new DelimitedWorkbookReader();

        var metadata = await reader.ReadMetadataAsync(path);

        Assert.AreEqual(WorkbookFormat.Csv, metadata.Format);
        Assert.AreEqual("sample.csv", metadata.DisplayName);
        Assert.AreEqual("sample", metadata.Sheets.Single().Name);
        Assert.AreEqual(0, metadata.Sheets[0].FirstRowIndex);
        Assert.AreEqual(2, metadata.Sheets[0].LastRowIndex);
        Assert.AreEqual(0, metadata.Sheets[0].FirstColumnIndex);
        Assert.AreEqual(4, metadata.Sheets[0].LastColumnIndex);
        Assert.AreEqual(5L, metadata.Sheets[0].NonEmptyCellCount);

        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var reports = new List<DelimitedReaderProgress>();
        var result = await reader.IndexAsync(
            path,
            workspace,
            new DelimitedReaderOptions { ProgressIntervalRows = 1 },
            new SynchronousProgress<DelimitedReaderProgress>(reports.Add));

        Assert.AreEqual(DelimitedFileFormat.Csv, result.Format);
        Assert.AreEqual(DelimitedTextEncoding.Utf8, result.TextEncoding);
        Assert.IsFalse(result.HasByteOrderMark);
        Assert.AreEqual(1, result.Worksheets.Count);
        Assert.AreEqual(3, result.Worksheet.StoredRowCount);
        Assert.AreEqual(DelimitedReaderStage.Detecting, reports[0].Stage);
        Assert.AreEqual(DelimitedReaderStage.Completed, reports[^1].Stage);
        Assert.AreEqual(3L, reports[^1].RowsRead);
        Assert.AreEqual(7L, reports[^1].CellsRead);

        var first = await result.Worksheet.GetRowAsync(0);
        var blank = await result.Worksheet.GetRowAsync(1);
        var last = await result.Worksheet.GetRowAsync(2);
        Assert.IsTrue(first.HasValue);
        Assert.IsTrue(blank.HasValue);
        Assert.IsTrue(last.HasValue);
        CollectionAssert.AreEqual(
            new[] { "alpha", "b,b", "line1\r\nline2", "say \"hi\"", string.Empty },
            GetTexts(first.Value));
        CollectionAssert.AreEqual(new[] { string.Empty }, GetTexts(blank.Value));
        CollectionAssert.AreEqual(new[] { "last" }, GetTexts(last.Value));
    }

    [TestMethod]
    public async Task Reader_detects_supported_byte_order_marks_and_tsv_format()
    {
        using var temporaryDirectory = new TestDirectory();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var reader = new DelimitedWorkbookReader();
        var cases = new (string Name, Encoding Encoding, DelimitedTextEncoding Expected)[]
        {
            ("utf8.tsv", new UTF8Encoding(true, true), DelimitedTextEncoding.Utf8),
            ("utf16-le.tsv", new UnicodeEncoding(false, true, true), DelimitedTextEncoding.Utf16LittleEndian),
            ("utf16-be.tsv", new UnicodeEncoding(true, true, true), DelimitedTextEncoding.Utf16BigEndian),
        };

        foreach (var item in cases)
        {
            var path = temporaryDirectory.GetPath(item.Name);
            await File.WriteAllTextAsync(path, "left\tright\n", item.Encoding);

            var result = await reader.IndexAsync(path, workspace);
            var row = await result.Worksheet.GetRowAsync(0);

            Assert.AreEqual(DelimitedFileFormat.Tsv, result.Format, item.Name);
            Assert.AreEqual(item.Expected, result.TextEncoding, item.Name);
            Assert.IsTrue(result.HasByteOrderMark, item.Name);
            Assert.IsTrue(row.HasValue, item.Name);
            CollectionAssert.AreEqual(new[] { "left", "right" }, GetTexts(row.Value), item.Name);
            Assert.AreEqual(1, result.Worksheet.StoredRowCount, item.Name);
        }
    }

    [TestMethod]
    public async Task Reader_requires_explicit_format_for_unknown_extensions()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = temporaryDirectory.GetPath("data.txt");
        await File.WriteAllTextAsync(path, "a,b");
        var reader = new DelimitedWorkbookReader();

        var exception = await Assert.ThrowsExceptionAsync<DelimitedReaderException>(() =>
            reader.ReadMetadataAsync(path).AsTask());
        var metadata = await reader.ReadMetadataAsync(
            path,
            new DelimitedReaderOptions
            {
                Format = DelimitedFileFormat.Csv,
                WorksheetName = "Imported",
            });

        Assert.AreEqual(DelimitedReaderError.UnsupportedFileExtension, exception.Error);
        Assert.AreEqual(WorkbookFormat.Csv, metadata.Format);
        Assert.AreEqual("Imported", metadata.Sheets.Single().Name);
    }

    [TestMethod]
    public async Task Reader_rejects_invalid_syntax_and_removes_partial_store()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = temporaryDirectory.GetPath("bad.csv");
        await File.WriteAllTextAsync(path, "valid\r\n\"unterminated");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });

        var exception = await Assert.ThrowsExceptionAsync<DelimitedReaderException>(() =>
            new DelimitedWorkbookReader().IndexAsync(path, workspace).AsTask());

        Assert.AreEqual(DelimitedReaderError.UnterminatedQuotedField, exception.Error);
        Assert.AreEqual(2L, exception.RecordNumber);
        Assert.AreEqual(0, Directory.GetDirectories(workspace.DirectoryPath).Length);
    }

    [TestMethod]
    public async Task Reader_enforces_limits_and_encoding_contract()
    {
        using var temporaryDirectory = new TestDirectory();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var reader = new DelimitedWorkbookReader();
        var columnsPath = temporaryDirectory.GetPath("columns.csv");
        var rowsPath = temporaryDirectory.GetPath("rows.csv");
        var fieldPath = temporaryDirectory.GetPath("field.csv");
        var invalidUtf8Path = temporaryDirectory.GetPath("invalid.csv");
        var utf32Path = temporaryDirectory.GetPath("utf32.csv");
        await File.WriteAllTextAsync(columnsPath, "a,b");
        await File.WriteAllTextAsync(rowsPath, "a\nb");
        await File.WriteAllTextAsync(fieldPath, "four");
        await File.WriteAllBytesAsync(invalidUtf8Path, [0xC3, 0x28]);
        await File.WriteAllBytesAsync(utf32Path, [0xFF, 0xFE, 0x00, 0x00, 0x61, 0x00, 0x00, 0x00]);

        var columns = await Assert.ThrowsExceptionAsync<DelimitedReaderException>(() =>
            reader.IndexAsync(
                columnsPath,
                workspace,
                new DelimitedReaderOptions { MaximumColumns = 1 }).AsTask());
        var rows = await Assert.ThrowsExceptionAsync<DelimitedReaderException>(() =>
            reader.IndexAsync(
                rowsPath,
                workspace,
                new DelimitedReaderOptions { MaximumRows = 1 }).AsTask());
        var field = await Assert.ThrowsExceptionAsync<DelimitedReaderException>(() =>
            reader.IndexAsync(
                fieldPath,
                workspace,
                new DelimitedReaderOptions { MaximumFieldCharacters = 3 }).AsTask());
        var invalidUtf8 = await Assert.ThrowsExceptionAsync<DelimitedReaderException>(() =>
            reader.IndexAsync(invalidUtf8Path, workspace).AsTask());
        var utf32 = await Assert.ThrowsExceptionAsync<DelimitedReaderException>(() =>
            reader.IndexAsync(utf32Path, workspace).AsTask());

        Assert.AreEqual(DelimitedReaderError.ColumnLimitExceeded, columns.Error);
        Assert.AreEqual(DelimitedReaderError.RowLimitExceeded, rows.Error);
        Assert.AreEqual(DelimitedReaderError.FieldLimitExceeded, field.Error);
        Assert.AreEqual(DelimitedReaderError.InvalidEncoding, invalidUtf8.Error);
        Assert.AreEqual(DelimitedReaderError.UnsupportedEncoding, utf32.Error);
        Assert.AreEqual(0, Directory.GetDirectories(workspace.DirectoryPath).Length);
    }

    [TestMethod]
    public async Task Writer_emits_transactional_utf8_and_round_trips_content()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = temporaryDirectory.GetPath("result.csv");
        var rows = new[]
        {
            new RowRecord(
                0,
                new CellRecord[]
                {
                    new(new CellAddress(0, 0), CellValue.FromText("a,b")),
                    new(new CellAddress(0, 1), CellValue.FromText("say \"hi\"")),
                    new(new CellAddress(0, 2), CellValue.FromText("line1\nline2")),
                    new(new CellAddress(0, 4), CellValue.FromNumber(12.5)),
                }),
            new RowRecord(
                2,
                new CellRecord[]
                {
                    new(new CellAddress(2, 0), CellValue.FromBoolean(true)),
                    new(new CellAddress(2, 1), CellValue.FromFormula("A1")),
                }),
        };
        var reports = new List<DelimitedWriterProgress>();

        var result = await new DelimitedContentWriter().WriteAsync(
            path,
            rows,
            new DelimitedWriterOptions
            {
                EmitUtf8ByteOrderMark = true,
                ProgressIntervalRows = 1,
            },
            new SynchronousProgress<DelimitedWriterProgress>(reports.Add));

        Assert.AreEqual(DelimitedFileFormat.Csv, result.Format);
        Assert.AreEqual(2L, result.SourceRowsProcessed);
        Assert.AreEqual(3L, result.RecordsWritten);
        Assert.AreEqual(6L, result.CellsWritten);
        Assert.IsTrue(result.HasByteOrderMark);
        Assert.AreEqual(DelimitedWriterStage.Completed, reports[^1].Stage);
        var bytes = await File.ReadAllBytesAsync(path);
        CollectionAssert.AreEqual(new byte[] { 0xEF, 0xBB, 0xBF }, bytes[..3]);
        Assert.AreEqual(
            "\"a,b\",\"say \"\"hi\"\"\",\"line1\nline2\",,12.5\r\n\r\nTRUE,=A1\r\n",
            await File.ReadAllTextAsync(path));

        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var indexed = await new DelimitedWorkbookReader().IndexAsync(path, workspace);
        Assert.IsTrue(indexed.HasByteOrderMark);
        Assert.AreEqual(3, indexed.Worksheet.StoredRowCount);
        CollectionAssert.AreEqual(
            new[] { "a,b", "say \"hi\"", "line1\nline2", string.Empty, "12.5" },
            GetTexts((await indexed.Worksheet.GetRowAsync(0)).GetValueOrDefault()));
        CollectionAssert.AreEqual(
            new[] { "TRUE", "=A1" },
            GetTexts((await indexed.Worksheet.GetRowAsync(2)).GetValueOrDefault()));
    }

    [TestMethod]
    public async Task Writer_preserves_destination_on_validation_failure_and_cancellation()
    {
        using var temporaryDirectory = new TestDirectory();
        var path = temporaryDirectory.GetPath("result.tsv");
        await File.WriteAllTextAsync(path, "original");
        var writer = new DelimitedContentWriter();
        var invalidRows = new[]
        {
            new RowRecord(
                0,
                new CellRecord[]
                {
                    new(new CellAddress(0, 1), CellValue.FromText("second")),
                    new(new CellAddress(0, 0), CellValue.FromText("first")),
                }),
        };

        var invalid = await Assert.ThrowsExceptionAsync<DelimitedWriterException>(() =>
            writer.WriteAsync(path, invalidRows).AsTask());
        var exists = await Assert.ThrowsExceptionAsync<DelimitedWriterException>(() =>
            writer.WriteAsync(
                path,
                new[] { TestData.Row(0, (0, TestData.Text("replacement"))) },
                new DelimitedWriterOptions { Overwrite = false }).AsTask());

        using var cancellation = new CancellationTokenSource();
        var progress = new SynchronousProgress<DelimitedWriterProgress>(report =>
        {
            if (report.Stage == DelimitedWriterStage.Writing)
            {
                cancellation.Cancel();
            }
        });
        await Assert.ThrowsExceptionAsync<OperationCanceledException>(() =>
            writer.WriteAsync(
                path,
                new[] { TestData.Row(0, (0, TestData.Text("replacement"))) },
                progress: progress,
                cancellationToken: cancellation.Token).AsTask());

        Assert.AreEqual(DelimitedWriterError.InvalidCellOrder, invalid.Error);
        Assert.AreEqual(DelimitedWriterError.DestinationExists, exists.Error);
        Assert.AreEqual("original", await File.ReadAllTextAsync(path));
        Assert.AreEqual(0, Directory.GetFiles(temporaryDirectory.Path, "*.tmp").Length);
    }

    [TestMethod]
    public async Task Indexed_delimited_worksheets_flow_into_the_diff_engine()
    {
        using var temporaryDirectory = new TestDirectory();
        var localPath = temporaryDirectory.GetPath("local.csv");
        var remotePath = temporaryDirectory.GetPath("remote.csv");
        await File.WriteAllTextAsync(localPath, "id,value\r\n1,local");
        await File.WriteAllTextAsync(remotePath, "id,value\r\n1,remote");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var reader = new DelimitedWorkbookReader();
        var local = await reader.IndexAsync(localPath, workspace);
        var remote = await reader.IndexAsync(remotePath, workspace);

        var diff = await new TwoWayDiffEngine().CompareAsync(
            local.Worksheet,
            remote.Worksheet,
            SourceSide.Remote);

        Assert.AreEqual(ChangeKind.Modified, diff.SheetChange.Kind);
        Assert.AreEqual(1L, diff.Statistics.ChangedCellCount);
        Assert.AreEqual(
            "remote",
            diff.SheetChange.RowChanges.Span[0].CellChanges.Span[0].SourceValue?.Scalar?.TextValue);
    }

    private static string[] GetTexts(in RowRecord row) =>
        row.Cells.Span.ToArray()
            .Select(static cell => cell.Value.Scalar?.TextValue ?? string.Empty)
            .ToArray();
}

using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using ExcelMerge.Domain;
using ExcelMerge.Engine;
using ExcelMerge.OpenXml;
using ExcelMerge.Storage;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Rewrite.Tests;

[TestClass]
public sealed class OpenXmlWriterTests
{
    [TestMethod]
    public async Task Writer_normalizes_noncanonical_font_child_order_before_validation()
    {
        using var temporaryDirectory = new TestDirectory();
        const string styles =
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<fonts count=\"1\"><font><name val=\"Calibri\"/><charset val=\"1\"/>" +
            "<family val=\"2\"/><color rgb=\"FF112233\"/><sz val=\"11\"/>" +
            "<scheme val=\"minor\"/></font></fonts>" +
            "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>" +
            "<fill><patternFill patternType=\"gray125\"/></fill></fills>" +
            "<borders count=\"1\"><border/></borders>" +
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
            "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>" +
            "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
            "<dxfs count=\"0\"/></styleSheet>";
        var basePath = CreateStyledTextWorkbook(temporaryDirectory, "base.xlsx", "base", styles);
        var localPath = CreateStyledTextWorkbook(temporaryDirectory, "local.xlsx", "base", styles);
        var remotePath = CreateStyledTextWorkbook(temporaryDirectory, "remote.xlsx", "remote", styles);
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        var localHash = await HashFileAsync(localPath);
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (plan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        CollectionAssert.AreEqual(localHash, await HashFileAsync(localPath));
        using var document = SpreadsheetDocument.Open(destinationPath, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Microsoft365)
            .Validate(document)
            .ToArray();
        Assert.AreEqual(
            0,
            errors.Length,
            string.Join(Environment.NewLine, errors.Select(static error => error.Description)));
        CollectionAssert.AreEqual(
            new[] { "sz", "color", "name", "family", "charset", "scheme" },
            document.WorkbookPart!.WorkbookStylesPart!.Stylesheet.Fonts!
                .Elements<Font>()
                .Single()
                .ChildElements
                .Select(static child => child.LocalName)
                .ToArray());
    }

    [TestMethod]
    public async Task Writer_applies_remote_only_cell_change_and_preserves_sources()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateTextWorkbook(temporaryDirectory, "base.xlsx", "base");
        var localPath = CreateTextWorkbook(temporaryDirectory, "local.xlsx", "base");
        var remotePath = CreateTextWorkbook(temporaryDirectory, "remote.xlsx", "remote");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        var localHashBefore = await HashFileAsync(localPath);
        var progress = new List<OpenXmlWriterProgress>();
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });

        var (plan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);
        var result = await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 },
            new SynchronousProgress<OpenXmlWriterProgress>(progress.Add));

        Assert.AreEqual(destinationPath, result.DestinationPath);
        Assert.IsTrue(File.Exists(destinationPath));
        CollectionAssert.AreEqual(localHashBefore, await HashFileAsync(localPath));
        Assert.AreEqual(OpenXmlWriterStage.Completed, progress[^1].Stage);

        var indexedResult = await new OpenXmlWorkbookReader().IndexAsync(destinationPath, workspace);
        var row = await indexedResult.Worksheets.Single().GetRowAsync(0);
        Assert.IsTrue(row.HasValue);
        Assert.AreEqual("remote", row.Value.Cells.Span[0].Value.Scalar?.TextValue);
    }

    [TestMethod]
    public async Task Writer_applies_remote_conflict_resolution_over_existing_destination()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateTextWorkbook(temporaryDirectory, "base.xlsx", "base");
        var localPath = CreateTextWorkbook(temporaryDirectory, "local.xlsx", "local");
        var remotePath = CreateTextWorkbook(temporaryDirectory, "remote.xlsx", "remote");
        var destinationPath = CreateTextWorkbook(temporaryDirectory, "result.xlsx", "old result");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (unresolvedPlan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var conflictId = unresolvedPlan.Conflicts.Span[0].Id;
        var plan = unresolvedPlan with
        {
            CellResolutions = new[] { CellResolution.UseRemote(conflictId) },
        };
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        var indexedResult = await new OpenXmlWorkbookReader().IndexAsync(destinationPath, workspace);
        var row = await indexedResult.Worksheets.Single().GetRowAsync(0);
        Assert.IsTrue(row.HasValue);
        Assert.AreEqual("remote", row.Value.Cells.Span[0].Value.Scalar?.TextValue);
    }

    [TestMethod]
    public async Task Writer_rejects_unresolved_plan_without_changing_destination()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateTextWorkbook(temporaryDirectory, "base.xlsx", "base");
        var localPath = CreateTextWorkbook(temporaryDirectory, "local.xlsx", "local");
        var remotePath = CreateTextWorkbook(temporaryDirectory, "remote.xlsx", "remote");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await File.WriteAllTextAsync(destinationPath, "existing destination");
        var destinationHash = await HashFileAsync(destinationPath);
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (plan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        var exception = await Assert.ThrowsExceptionAsync<OpenXmlWriterException>(() =>
            new OpenXmlWorkbookWriter().WriteAsync(
                request,
                new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 }).AsTask());

        Assert.AreEqual(OpenXmlWriterError.UnresolvedConflicts, exception.Error);
        CollectionAssert.AreEqual(destinationHash, await HashFileAsync(destinationPath));
        AssertNoTransactionFiles(temporaryDirectory.Path);
    }

    [TestMethod]
    public async Task Writer_cancellation_removes_transaction_and_preserves_destination()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateTextWorkbook(temporaryDirectory, "base.xlsx", "base");
        var localPath = CreateTextWorkbook(temporaryDirectory, "local.xlsx", "base");
        var remotePath = CreateTextWorkbook(temporaryDirectory, "remote.xlsx", "remote");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await File.WriteAllTextAsync(destinationPath, "existing destination");
        var destinationHash = await HashFileAsync(destinationPath);
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (plan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);
        using var cancellation = new CancellationTokenSource();
        var progress = new SynchronousProgress<OpenXmlWriterProgress>(report =>
        {
            if (report.Stage == OpenXmlWriterStage.CopyingLocal)
            {
                cancellation.Cancel();
            }
        });

        try
        {
            await new OpenXmlWorkbookWriter().WriteAsync(
                request,
                new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 },
                progress,
                cancellation.Token);
            Assert.Fail("The writer should have observed cancellation.");
        }
        catch (OperationCanceledException)
        {
        }

        CollectionAssert.AreEqual(destinationHash, await HashFileAsync(destinationPath));
        AssertNoTransactionFiles(temporaryDirectory.Path);
    }

    [TestMethod]
    public async Task Writer_detects_source_change_before_creating_output()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateTextWorkbook(temporaryDirectory, "base.xlsx", "base");
        var localPath = CreateTextWorkbook(temporaryDirectory, "local.xlsx", "base");
        var remotePath = CreateTextWorkbook(temporaryDirectory, "remote.xlsx", "remote");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (plan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);
        File.SetLastWriteTimeUtc(remotePath, File.GetLastWriteTimeUtc(remotePath).AddMinutes(1));

        var exception = await Assert.ThrowsExceptionAsync<OpenXmlWriterException>(() =>
            new OpenXmlWorkbookWriter().WriteAsync(
                request,
                new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 }).AsTask());

        Assert.AreEqual(OpenXmlWriterError.SourceChanged, exception.Error);
        Assert.IsFalse(File.Exists(destinationPath));
        AssertNoTransactionFiles(temporaryDirectory.Path);
    }

    [TestMethod]
    public async Task Writer_imports_a_remote_selected_worksheet_relationship_graph()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateTextWorkbook(temporaryDirectory, "base.xlsx", "base");
        var localPath = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>keep</v></c></row></sheetData>" +
            "</worksheet>",
            fileName: "local.xlsx",
            sheetName: "Other");
        var remotePath = CreateTextWorkbook(temporaryDirectory, "remote.xlsx", "remote");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var reader = new OpenXmlWorkbookReader();
        var baseResult = await reader.IndexAsync(basePath, workspace);
        var localResult = await reader.IndexAsync(localPath, workspace);
        var remoteResult = await reader.IndexAsync(remotePath, workspace);
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseResult.Worksheets.Single(),
            localSheet: null,
            remoteResult.Worksheets.Single(),
            new ThreeWayMergeOptions { SheetGroupId = "data-pair" });
        var unresolvedPlan = MergePlanFactory.Create(
            baseResult.Metadata,
            localResult.Metadata,
            remoteResult.Metadata,
            [merge]);
        var conflictId = unresolvedPlan.Conflicts.Span[0].Id;
        var plan = unresolvedPlan with
        {
            RowResolutions = new[] { RowResolution.UseRemote(conflictId) },
        };
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        var output = await reader.IndexAsync(destinationPath, workspace);
        CollectionAssert.AreEqual(
            new[] { "Other", "Data" },
            output.Metadata.Sheets.Select(static sheet => sheet.Name).ToArray());
        var imported = output.Worksheets.Single(sheet => sheet.Metadata.Name == "Data");
        var row = await imported.GetRowAsync(0);
        Assert.IsTrue(row.HasValue);
        Assert.AreEqual("remote", row.Value.Cells.Span[0].Value.Scalar?.TextValue);
    }

    [TestMethod]
    public async Task Writer_remaps_remote_styles_and_converts_shared_strings_when_importing_sheet()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateTextWorkbook(temporaryDirectory, "base.xlsx", "base");
        var localPath = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>",
            fileName: "local.xlsx",
            sheetName: "Other");
        var remotePath = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\" s=\"1\"><v>0</v></c></row></sheetData>" +
            "</worksheet>",
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">" +
            "<si><r><rPr><b/></rPr><t xml:space=\"preserve\">rich </t></r><r><t>text</t></r></si></sst>",
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"0.00\"/></numFmts>" +
            "<fonts count=\"2\"><font/><font><b/><name val=\"Calibri\"/></font></fonts>" +
            "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>" +
            "<fill><patternFill patternType=\"gray125\"/></fill></fills>" +
            "<borders count=\"1\"><border/></borders>" +
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
            "<cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
            "<xf numFmtId=\"164\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/></cellXfs>" +
            "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>" +
            "</styleSheet>",
            fileName: "remote.xlsx");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var reader = new OpenXmlWorkbookReader();
        var baseResult = await reader.IndexAsync(basePath, workspace);
        var localResult = await reader.IndexAsync(localPath, workspace);
        var remoteResult = await reader.IndexAsync(remotePath, workspace);
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet: null,
            localSheet: null,
            remoteResult.Worksheets.Single(),
            new ThreeWayMergeOptions { SheetGroupId = "styled-remote" });
        var plan = MergePlanFactory.Create(
            baseResult.Metadata,
            localResult.Metadata,
            remoteResult.Metadata,
            [merge]);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        using (var document = SpreadsheetDocument.Open(destinationPath, isEditable: false))
        {
            var workbookPart = document.WorkbookPart!;
            var sheet = workbookPart.Workbook.Sheets!.Elements<Sheet>()
                .Single(candidate => candidate.Name?.Value == "Data");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
            var cell = worksheetPart.Worksheet.Descendants<Cell>().Single();
            var runs = cell.InlineString!.Elements<Run>().ToArray();
            Assert.AreEqual(CellValues.InlineString, cell.DataType!.Value);
            Assert.AreEqual(2, runs.Length);
            Assert.IsNotNull(runs[0].RunProperties?.GetFirstChild<Bold>());
            Assert.AreEqual(SpaceProcessingModeValues.Preserve, runs[0].Text!.Space!.Value);
        }

        var output = await reader.IndexAsync(destinationPath, workspace);
        var imported = output.Worksheets.Single(sheet => sheet.Metadata.Name == "Data");
        var row = await imported.GetRowAsync(0);
        Assert.IsTrue(row.HasValue);
        Assert.AreEqual("rich text", row.Value.Cells.Span[0].Value.Scalar?.TextValue);
        Assert.IsTrue(row.Value.Cells.Span[0].StyleIndex > 0);
    }

    [TestMethod]
    public async Task Writer_inserts_a_remote_only_row_and_shifts_following_local_rows()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateRowsWorkbook(temporaryDirectory, "base.xlsx", "A", "B");
        var localPath = CreateRowsWorkbook(temporaryDirectory, "local.xlsx", "A", "B");
        var remotePath = CreateRowsWorkbook(temporaryDirectory, "remote.xlsx", "A", "inserted", "B");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (plan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        CollectionAssert.AreEqual(
            new[] { "A", "inserted", "B" },
            await ReadFirstColumnAsync(destinationPath, workspace));
    }

    [TestMethod]
    public async Task Writer_applies_both_by_placing_remote_row_immediately_after_local_row()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateRowsWorkbook(temporaryDirectory, "base.xlsx", "A", "B");
        var localPath = CreateRowsWorkbook(temporaryDirectory, "local.xlsx", "A", "LOCAL", "B");
        var remotePath = CreateRowsWorkbook(temporaryDirectory, "remote.xlsx", "A", "REMOTE", "B");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (unresolvedPlan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var conflict = unresolvedPlan.Conflicts.ToArray().Single();
        Assert.AreEqual(ConflictKind.RowAddAdd, conflict.Kind);
        var plan = unresolvedPlan with
        {
            RowResolutions = new[] { RowResolution.UseBoth(conflict.Id) },
        };
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        CollectionAssert.AreEqual(
            new[] { "A", "LOCAL", "REMOTE", "B" },
            await ReadFirstColumnAsync(destinationPath, workspace));
    }

    [TestMethod]
    public async Task Writer_rewrites_local_and_remote_formula_rows_without_double_translation()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateFormulaWorkbook(temporaryDirectory, "base.xlsx", insertRow: false);
        var localPath = CreateFormulaWorkbook(temporaryDirectory, "local.xlsx", insertRow: false);
        var remotePath = CreateFormulaWorkbook(temporaryDirectory, "remote.xlsx", insertRow: true);
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (plan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        var output = await new OpenXmlWorkbookReader().IndexAsync(destinationPath, workspace);
        var firstRow = await output.Worksheets.Single().GetRowAsync(0);
        Assert.IsTrue(firstRow.HasValue);
        Assert.AreEqual(
            "IF(A3=\"A2\",$A$3,A3)",
            firstRow.Value.Cells.Span[1].Value.Formula);
    }

    [TestMethod]
    public async Task Writer_transforms_all_supported_inline_coordinate_carriers_together()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateCoordinateWorkbook(temporaryDirectory, "base.xlsx", insertRow: false);
        var localPath = CreateCoordinateWorkbook(temporaryDirectory, "local.xlsx", insertRow: false);
        var remotePath = CreateCoordinateWorkbook(temporaryDirectory, "remote.xlsx", insertRow: true);
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (plan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        using var document = SpreadsheetDocument.Open(destinationPath, isEditable: false);
        var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet;
        Assert.AreEqual("A3:B3", worksheet.Descendants<MergeCell>().Single().Reference?.Value);
        Assert.AreEqual(
            "A3",
            worksheet.Descendants<ConditionalFormatting>().Single().SequenceOfReferences?.InnerText);
        Assert.AreEqual("A3", worksheet.Descendants<Formula>().Single().Text);
        Assert.AreEqual(
            "A3",
            worksheet.Descendants<DataValidation>().Single().SequenceOfReferences?.InnerText);
        Assert.AreEqual("A3", worksheet.Descendants<Formula1>().Single().Text);
        Assert.AreEqual("A1:A3", worksheet.Descendants<AutoFilter>().Single().Reference?.Value);
        Assert.AreEqual("A3", worksheet.Descendants<Hyperlink>().Single().Reference?.Value);
        Assert.AreEqual("A3", worksheet.Descendants<Hyperlink>().Single().Location?.Value);
        Assert.AreEqual(3U, worksheet.Descendants<Break>().Single().Id?.Value);
    }

    [TestMethod]
    public async Task Writer_applies_remote_row_metadata_without_replacing_cell_values()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateRowMetadataWorkbook(temporaryDirectory, "base.xlsx", 20, hidden: false);
        var localPath = CreateRowMetadataWorkbook(temporaryDirectory, "local.xlsx", 21, hidden: false);
        var remotePath = CreateRowMetadataWorkbook(temporaryDirectory, "remote.xlsx", 22, hidden: true);
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var reader = new OpenXmlWorkbookReader();
        var baseResult = await reader.IndexAsync(basePath, workspace);
        var localResult = await reader.IndexAsync(localPath, workspace);
        var remoteResult = await reader.IndexAsync(remotePath, workspace);
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseResult.Worksheets.Single(),
            localResult.Worksheets.Single(),
            remoteResult.Worksheets.Single(),
            new ThreeWayMergeOptions
            {
                SheetGroupId = "metadata-pair",
                Comparison = new WorksheetComparisonOptions { CompareRowMetadata = true },
            });
        var conflict = merge.Conflicts.ToArray().Single();
        Assert.AreEqual(ConflictKind.RowMetadata, conflict.Kind);
        var unresolvedPlan = MergePlanFactory.Create(
            baseResult.Metadata,
            localResult.Metadata,
            remoteResult.Metadata,
            [merge]);
        var plan = unresolvedPlan with
        {
            RowResolutions = new[] { RowResolution.UseRemote(conflict.Id) },
        };
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        var output = await reader.IndexAsync(destinationPath, workspace);
        var row = await output.Worksheets.Single().GetRowAsync(0);
        Assert.IsTrue(row.HasValue);
        Assert.AreEqual(22d, row.Value.Height);
        Assert.IsTrue(row.Value.IsHidden);
        Assert.AreEqual("value", row.Value.Cells.Span[0].Value.Scalar?.TextValue);
    }

    [TestMethod]
    public async Task Writer_applies_remote_worksheet_name_and_visibility_resolution()
    {
        using var temporaryDirectory = new TestDirectory();
        var worksheetXml =
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>value</v></c></row></sheetData>" +
            "</worksheet>";
        var basePath = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            worksheetXml,
            fileName: "base.xlsx",
            sheetName: "Data");
        var localPath = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            worksheetXml,
            fileName: "local.xlsx",
            sheetName: "LOCAL Name");
        var remotePath = OpenXmlTestWorkbook.Create(
            temporaryDirectory,
            worksheetXml,
            fileName: "remote.xlsx",
            sheetName: "REMOTE Name",
            sheetState: "hidden");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var reader = new OpenXmlWorkbookReader();
        var baseResult = await reader.IndexAsync(basePath, workspace);
        var localResult = await reader.IndexAsync(localPath, workspace);
        var remoteResult = await reader.IndexAsync(remotePath, workspace);
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseResult.Worksheets.Single(),
            localResult.Worksheets.Single(),
            remoteResult.Worksheets.Single(),
            new ThreeWayMergeOptions { SheetGroupId = "rename-pair" });
        var conflict = merge.Conflicts.ToArray().Single();
        Assert.AreEqual(ConflictKind.AmbiguousSheetRename, conflict.Kind);
        var unresolvedPlan = MergePlanFactory.Create(
            baseResult.Metadata,
            localResult.Metadata,
            remoteResult.Metadata,
            [merge]);
        var plan = unresolvedPlan with
        {
            RowResolutions = new[] { RowResolution.UseRemote(conflict.Id) },
        };
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        var metadata = await reader.ReadMetadataAsync(destinationPath);
        Assert.AreEqual("REMOTE Name", metadata.Sheets.Single().Name);
        Assert.AreEqual(SheetVisibility.Hidden, metadata.Sheets.Single().Visibility);
    }

    [TestMethod]
    public async Task Writer_row_override_both_covers_cell_conflicts_and_duplicates_complete_row()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateTextWorkbook(temporaryDirectory, "base.xlsx", "base");
        var localPath = CreateTextWorkbook(temporaryDirectory, "local.xlsx", "LOCAL");
        var remotePath = CreateTextWorkbook(temporaryDirectory, "remote.xlsx", "REMOTE");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var (unresolvedPlan, _, _, _) = await CreatePlanAsync(
            basePath,
            localPath,
            remotePath,
            workspace);
        var location = unresolvedPlan.Conflicts.Span[0].Location;
        var plan = unresolvedPlan with
        {
            RowOverrides = new[]
            {
                new RowMergeOverride(
                    location.SheetId,
                    location.BaseRowIndex,
                    location.LocalRowIndex,
                    location.RemoteRowIndex,
                    ResolutionKind.Both),
            },
        };
        Assert.IsFalse(plan.HasUnresolvedResolutions);
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);

        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });

        CollectionAssert.AreEqual(
            new[] { "LOCAL", "REMOTE" },
            await ReadFirstColumnAsync(destinationPath, workspace));
    }

    private static string CreateTextWorkbook(
        TestDirectory directory,
        string fileName,
        string value) =>
        OpenXmlTestWorkbook.Create(
            directory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            $"<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\"><v>{value}</v></c></row></sheetData>" +
            "</worksheet>",
            fileName: fileName);

    private static string CreateStyledTextWorkbook(
        TestDirectory directory,
        string fileName,
        string value,
        string styles) =>
        OpenXmlTestWorkbook.Create(
            directory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            $"<sheetData><row r=\"1\"><c r=\"A1\" t=\"str\" s=\"0\"><v>{value}</v></c></row></sheetData>" +
            "</worksheet>",
            stylesXml: styles,
            fileName: fileName);

    private static string CreateRowsWorkbook(
        TestDirectory directory,
        string fileName,
        params string[] values)
    {
        var rows = string.Concat(values.Select((value, index) =>
            $"<row r=\"{index + 1}\"><c r=\"A{index + 1}\" t=\"str\"><v>{value}</v></c></row>"));
        return OpenXmlTestWorkbook.Create(
            directory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            $"<sheetData>{rows}</sheetData>" +
            "</worksheet>",
            fileName: fileName);
    }

    private static string CreateFormulaWorkbook(
        TestDirectory directory,
        string fileName,
        bool insertRow)
    {
        var referencedRow = insertRow ? 3 : 2;
        var middleRow = insertRow
            ? "<row r=\"2\"><c r=\"A2\" t=\"str\"><v>inserted</v></c></row>"
            : string.Empty;
        return OpenXmlTestWorkbook.Create(
            directory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            "<sheetData>" +
            $"<row r=\"1\"><c r=\"A1\" t=\"str\"><v>header</v></c><c r=\"B1\">" +
            $"<f>IF(A{referencedRow}=&quot;A2&quot;,$A${referencedRow},A{referencedRow})</f><v>0</v></c></row>" +
            middleRow +
            $"<row r=\"{referencedRow}\"><c r=\"A{referencedRow}\" t=\"str\"><v>B</v></c></row>" +
            "</sheetData></worksheet>",
            fileName: fileName);
    }

    private static string CreateCoordinateWorkbook(
        TestDirectory directory,
        string fileName,
        bool insertRow)
    {
        var targetRow = insertRow ? 3 : 2;
        var middleRow = insertRow
            ? "<row r=\"2\"><c r=\"A2\" t=\"str\"><v>inserted</v></c></row>"
            : string.Empty;
        return OpenXmlTestWorkbook.Create(
            directory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            $"<dimension ref=\"A1:B{targetRow}\"/>" +
            "<sheetData>" +
            "<row r=\"1\"><c r=\"A1\" t=\"str\"><v>A</v></c></row>" +
            middleRow +
            $"<row r=\"{targetRow}\"><c r=\"A{targetRow}\" t=\"str\"><v>B</v></c></row>" +
            "</sheetData>" +
            $"<autoFilter ref=\"A1:A{targetRow}\"/>" +
            $"<mergeCells count=\"1\"><mergeCell ref=\"A{targetRow}:B{targetRow}\"/></mergeCells>" +
            $"<conditionalFormatting sqref=\"A{targetRow}\"><cfRule type=\"expression\" priority=\"1\">" +
            $"<formula>A{targetRow}</formula></cfRule></conditionalFormatting>" +
            $"<dataValidations count=\"1\"><dataValidation type=\"custom\" sqref=\"A{targetRow}\">" +
            $"<formula1>A{targetRow}</formula1></dataValidation></dataValidations>" +
            $"<hyperlinks><hyperlink ref=\"A{targetRow}\" location=\"A{targetRow}\"/></hyperlinks>" +
            $"<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"{targetRow}\" min=\"0\" max=\"16383\" man=\"1\"/></rowBreaks>" +
            "</worksheet>",
            fileName: fileName);
    }

    private static string CreateRowMetadataWorkbook(
        TestDirectory directory,
        string fileName,
        double height,
        bool hidden) =>
        OpenXmlTestWorkbook.Create(
            directory,
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
            $"<sheetData><row r=\"1\" ht=\"{height.ToString(System.Globalization.CultureInfo.InvariantCulture)}\" " +
            $"hidden=\"{(hidden ? 1 : 0)}\"><c r=\"A1\" t=\"str\"><v>value</v></c></row></sheetData>" +
            "</worksheet>",
            fileName: fileName);

    private static async Task<(
        MergePlan Plan,
        OpenXmlReaderResult Base,
        OpenXmlReaderResult Local,
        OpenXmlReaderResult Remote)> CreatePlanAsync(
        string basePath,
        string localPath,
        string remotePath,
        Workspace workspace)
    {
        var reader = new OpenXmlWorkbookReader();
        var baseResult = await reader.IndexAsync(basePath, workspace);
        var localResult = await reader.IndexAsync(localPath, workspace);
        var remoteResult = await reader.IndexAsync(remotePath, workspace);
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseResult.Worksheets.Single(),
            localResult.Worksheets.Single(),
            remoteResult.Worksheets.Single(),
            new ThreeWayMergeOptions { SheetGroupId = "sheet-1" });
        var plan = MergePlanFactory.Create(
            baseResult.Metadata,
            localResult.Metadata,
            remoteResult.Metadata,
            [merge]);
        return (plan, baseResult, localResult, remoteResult);
    }

    private static async Task<OpenXmlWriteRequest> CreateWriteRequestAsync(
        MergePlan plan,
        string basePath,
        string localPath,
        string remotePath,
        string destinationPath) =>
        new(
            plan,
            await OpenXmlWorkbookSource.CaptureAsync(basePath),
            await OpenXmlWorkbookSource.CaptureAsync(localPath),
            await OpenXmlWorkbookSource.CaptureAsync(remotePath),
            destinationPath);

    private static async Task<byte[]> HashFileAsync(string path)
    {
        await using var stream = File.OpenRead(path);
        return await SHA256.HashDataAsync(stream);
    }

    private static async Task<string[]> ReadFirstColumnAsync(string path, Workspace workspace)
    {
        var result = await new OpenXmlWorkbookReader().IndexAsync(path, workspace);
        var values = new List<string>();
        await foreach (var row in result.Worksheets.Single().CellStore.ReadRowsAsync())
        {
            values.Add(row.Cells.Span[0].Value.Scalar?.TextValue ?? string.Empty);
        }

        return values.ToArray();
    }

    private static void AssertNoTransactionFiles(string directory)
    {
        Assert.AreEqual(
            0,
            Directory.GetFiles(directory, ".*.tmp", SearchOption.TopDirectoryOnly).Length);
    }
}

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelMerge.Domain;
using ExcelMerge.Engine;
using ExcelMerge.OpenXml;
using ExcelMerge.Storage;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class OpenXmlWriterFidelityTests
{
    [TestMethod]
    public async Task Writer_updates_dimension_cross_sheet_formulas_and_defined_names()
    {
        using var temporaryDirectory = new TestDirectory();
        var definitions = new[]
        {
            new TestDefinedName("DataRange", "Data!$A$2:$A$2"),
            new TestDefinedName("_xlnm.Print_Area", "Data!$A$1:$A$2", 0),
            new TestDefinedName("_xlnm.Print_Titles", "Data!$1:$2", 0),
            new TestDefinedName("LocalCell", "$A$2", 0),
            new TestDefinedName("SummaryCell", "$A$2", 1),
        };
        var basePath = CreateWorkbook(
            temporaryDirectory,
            "base.xlsx",
            "Data",
            ["A", "B"],
            "Data!A2",
            definitions);
        var localPath = CreateWorkbook(
            temporaryDirectory,
            "local.xlsx",
            "Data",
            ["A", "B"],
            "Data!A2",
            definitions);
        var remotePath = CreateWorkbook(
            temporaryDirectory,
            "remote.xlsx",
            "Data",
            ["A", "inserted", "B"],
            "Data!A3",
            definitions);
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var plan = await CreateDataPlanAsync(basePath, localPath, remotePath, workspace);

        await WriteAsync(plan, basePath, localPath, remotePath, destinationPath);

        using var document = SpreadsheetDocument.Open(destinationPath, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var sheets = workbookPart.Workbook.Sheets!.Elements<Sheet>().ToArray();
        var dataPart = (WorksheetPart)workbookPart.GetPartById(sheets[0].Id!);
        var summaryPart = (WorksheetPart)workbookPart.GetPartById(sheets[1].Id!);
        Assert.AreEqual("A1:A3", dataPart.Worksheet.SheetDimension?.Reference?.Value);
        Assert.AreEqual("Data!A3", summaryPart.Worksheet.Descendants<CellFormula>().Single().Text);

        var names = workbookPart.Workbook.DefinedNames!.Elements<DefinedName>()
            .ToDictionary(static name => name.Name!.Value!, StringComparer.Ordinal);
        Assert.AreEqual("Data!$A$3:$A$3", names["DataRange"].Text);
        Assert.AreEqual("Data!$A$1:$A$3", names["_xlnm.Print_Area"].Text);
        Assert.AreEqual("Data!$1:$3", names["_xlnm.Print_Titles"].Text);
        Assert.AreEqual("$A$3", names["LocalCell"].Text);
        Assert.AreEqual("$A$2", names["SummaryCell"].Text);
    }

    [TestMethod]
    public async Task Writer_propagates_sheet_rename_into_formulas_and_defined_names()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateWorkbook(
            temporaryDirectory,
            "base.xlsx",
            "Data",
            ["value"],
            "Data!A1");
        var localPath = CreateWorkbook(
            temporaryDirectory,
            "local.xlsx",
            "LOCAL Name",
            ["value"],
            "'LOCAL Name'!A1",
            [new TestDefinedName("Selected", "'LOCAL Name'!$A$1")]);
        var remotePath = CreateWorkbook(
            temporaryDirectory,
            "remote.xlsx",
            "REMOTE Name",
            ["value"],
            "'REMOTE Name'!A1");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var reader = new OpenXmlWorkbookReader();
        var baseResult = await reader.IndexAsync(basePath, workspace);
        var localResult = await reader.IndexAsync(localPath, workspace);
        var remoteResult = await reader.IndexAsync(remotePath, workspace);
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseResult.Worksheets[0],
            localResult.Worksheets[0],
            remoteResult.Worksheets[0],
            new ThreeWayMergeOptions { SheetGroupId = "data" });
        var unresolved = MergePlanFactory.Create(
            baseResult.Metadata,
            localResult.Metadata,
            remoteResult.Metadata,
            [merge]);
        var conflict = unresolved.Conflicts.ToArray().Single();
        var plan = unresolved with
        {
            RowResolutions = new[] { RowResolution.UseRemote(conflict.Id) },
        };

        await WriteAsync(plan, basePath, localPath, remotePath, destinationPath);

        using var document = SpreadsheetDocument.Open(destinationPath, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        CollectionAssert.AreEqual(
            new[] { "REMOTE Name", "Summary" },
            workbookPart.Workbook.Sheets!.Elements<Sheet>()
                .Select(static sheet => sheet.Name!.Value!)
                .ToArray());
        var summary = workbookPart.WorksheetParts.Single(part =>
            part.Worksheet.Descendants<CellFormula>().Any());
        Assert.AreEqual("'REMOTE Name'!A1", summary.Worksheet.Descendants<CellFormula>().Single().Text);
        Assert.AreEqual(
            "'REMOTE Name'!$A$1",
            workbookPart.Workbook.DefinedNames!.Elements<DefinedName>().Single().Text);
    }

    [TestMethod]
    public async Task Writer_imports_table_graph_remaps_ids_and_sheet_scoped_names()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateWorkbook(temporaryDirectory, "base.xlsx", "Base", ["value"]);
        var localPath = CreateWorkbook(
            temporaryDirectory,
            "local.xlsx",
            "Other",
            ["Value", "local"],
            table: new TestTable(1, "LocalTable"));
        var remotePath = CreateWorkbook(
            temporaryDirectory,
            "remote.xlsx",
            "Data",
            ["Value", "remote"],
            definedNames: [new TestDefinedName("_xlnm.Print_Area", "Data!$A$1:$A$2", 0)],
            table: new TestTable(1, "RemoteTable"));
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var plan = await CreateRemoteOnlyPlanAsync(basePath, localPath, remotePath, workspace);

        await WriteAsync(plan, basePath, localPath, remotePath, destinationPath);

        using var document = SpreadsheetDocument.Open(destinationPath, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var tables = workbookPart.WorksheetParts
            .SelectMany(static part => part.TableDefinitionParts)
            .Select(static part => part.Table)
            .ToArray();
        Assert.AreEqual(2, tables.Length);
        Assert.AreEqual(2, tables.Select(static table => table.Id!.Value).Distinct().Count());
        CollectionAssert.AreEquivalent(
            new[] { "LocalTable", "RemoteTable" },
            tables.Select(static table => table.Name!.Value!).ToArray());
        var importedName = workbookPart.Workbook.DefinedNames!.Elements<DefinedName>().Single();
        Assert.AreEqual(1U, importedName.LocalSheetId?.Value);
        Assert.AreEqual("Data!$A$1:$A$2", importedName.Text);
    }

    [TestMethod]
    public async Task Writer_rejects_imported_table_name_collision_transactionally()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateWorkbook(temporaryDirectory, "base.xlsx", "Base", ["value"]);
        var localPath = CreateWorkbook(
            temporaryDirectory,
            "local.xlsx",
            "Other",
            ["Value", "local"],
            table: new TestTable(1, "Table1"));
        var remotePath = CreateWorkbook(
            temporaryDirectory,
            "remote.xlsx",
            "Data",
            ["Value", "remote"],
            table: new TestTable(1, "Table1"));
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await File.WriteAllTextAsync(destinationPath, "existing destination");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var plan = await CreateRemoteOnlyPlanAsync(basePath, localPath, remotePath, workspace);
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

        Assert.AreEqual(OpenXmlWriterError.UnsupportedOperation, exception.Error);
        Assert.AreEqual("existing destination", await File.ReadAllTextAsync(destinationPath));
        AssertNoTransactionFiles(temporaryDirectory.Path);
    }

    [TestMethod]
    public async Task Writer_rejects_duplicate_final_worksheet_names_transactionally()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateWorkbook(temporaryDirectory, "base.xlsx", "Base", ["value"]);
        var localPath = CreateWorkbook(temporaryDirectory, "local.xlsx", "Data", ["local"]);
        var remotePath = CreateWorkbook(temporaryDirectory, "remote.xlsx", "Data", ["remote"]);
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await File.WriteAllTextAsync(destinationPath, "existing destination");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var plan = await CreateRemoteOnlyPlanAsync(basePath, localPath, remotePath, workspace);
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

        Assert.AreEqual(OpenXmlWriterError.InvalidPlan, exception.Error);
        Assert.AreEqual("existing destination", await File.ReadAllTextAsync(destinationPath));
        AssertNoTransactionFiles(temporaryDirectory.Path);
    }

    [TestMethod]
    public async Task Writer_rejects_three_dimensional_formula_during_row_transform()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateWorkbook(
            temporaryDirectory,
            "base.xlsx",
            "Data",
            ["A", "B"],
            "SUM(Data:Summary!A2)");
        var localPath = CreateWorkbook(
            temporaryDirectory,
            "local.xlsx",
            "Data",
            ["A", "B"],
            "SUM(Data:Summary!A2)");
        var remotePath = CreateWorkbook(
            temporaryDirectory,
            "remote.xlsx",
            "Data",
            ["A", "inserted", "B"],
            "SUM(Data:Summary!A3)");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await File.WriteAllTextAsync(destinationPath, "existing destination");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var plan = await CreateDataPlanAsync(basePath, localPath, remotePath, workspace);
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

        Assert.AreEqual(OpenXmlWriterError.UnsupportedOperation, exception.Error);
        Assert.AreEqual("existing destination", await File.ReadAllTextAsync(destinationPath));
        AssertNoTransactionFiles(temporaryDirectory.Path);
    }

    [TestMethod]
    public async Task Writer_preflights_every_source_and_rejects_compound_encryption_container()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateWorkbook(temporaryDirectory, "base.xlsx", "Data", ["base"]);
        var localPath = CreateWorkbook(temporaryDirectory, "local.xlsx", "Data", ["base"]);
        var remotePath = CreateWorkbook(temporaryDirectory, "remote.xlsx", "Data", ["remote"]);
        var invalidBasePath = temporaryDirectory.GetPath("encrypted.xlsx");
        await File.WriteAllBytesAsync(
            invalidBasePath,
            [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await File.WriteAllTextAsync(destinationPath, "existing destination");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var plan = await CreateDataPlanAsync(basePath, localPath, remotePath, workspace);
        var request = new OpenXmlWriteRequest(
            plan,
            await OpenXmlWorkbookSource.CaptureAsync(invalidBasePath),
            await OpenXmlWorkbookSource.CaptureAsync(localPath),
            await OpenXmlWorkbookSource.CaptureAsync(remotePath),
            destinationPath);

        var exception = await Assert.ThrowsExceptionAsync<OpenXmlWriterException>(() =>
            new OpenXmlWorkbookWriter().WriteAsync(
                request,
                new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 }).AsTask());

        Assert.AreEqual(OpenXmlWriterError.UnsupportedOperation, exception.Error);
        Assert.AreEqual(invalidBasePath, exception.Path);
        Assert.AreEqual("existing destination", await File.ReadAllTextAsync(destinationPath));
        AssertNoTransactionFiles(temporaryDirectory.Path);
    }

    [TestMethod]
    public async Task Writer_rewrites_qualified_whole_row_formulas()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateWorkbook(
            temporaryDirectory,
            "base.xlsx",
            "Data",
            ["A", "B"],
            "SUM(Data!2:2)");
        var localPath = CreateWorkbook(
            temporaryDirectory,
            "local.xlsx",
            "Data",
            ["A", "B"],
            "SUM(Data!2:2)");
        var remotePath = CreateWorkbook(
            temporaryDirectory,
            "remote.xlsx",
            "Data",
            ["A", "inserted", "B"],
            "SUM(Data!3:3)");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var plan = await CreateDataPlanAsync(basePath, localPath, remotePath, workspace);

        await WriteAsync(plan, basePath, localPath, remotePath, destinationPath);

        using var document = SpreadsheetDocument.Open(destinationPath, isEditable: false);
        var formula = document.WorkbookPart!.WorksheetParts
            .SelectMany(static part => part.Worksheet.Descendants<CellFormula>())
            .Single();
        Assert.AreEqual("SUM(Data!3:3)", formula.Text);
    }

    [TestMethod]
    public async Task Writer_expands_dimension_for_nonstructural_cell_patch()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateWorkbook(temporaryDirectory, "base.xlsx", "Data", ["same"]);
        var localPath = CreateWorkbook(temporaryDirectory, "local.xlsx", "Data", ["same"]);
        var remotePath = CreateWorkbook(temporaryDirectory, "remote.xlsx", "Data", ["same"]);
        AddTextCell(remotePath, "D1", "remote");
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var plan = await CreateDataPlanAsync(basePath, localPath, remotePath, workspace);

        await WriteAsync(plan, basePath, localPath, remotePath, destinationPath);

        using var document = SpreadsheetDocument.Open(destinationPath, isEditable: false);
        var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet;
        Assert.AreEqual("A1:D1", worksheet.SheetDimension?.Reference?.Value);
        Assert.AreEqual(
            "remote",
            worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "D1")
                .Descendants<Text>()
                .Single()
                .Text);
    }

    [TestMethod]
    public async Task Writer_rejects_imported_scoped_name_with_foreign_sheet_dependency()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateWorkbook(temporaryDirectory, "base.xlsx", "Base", ["value"]);
        var localPath = CreateWorkbook(temporaryDirectory, "local.xlsx", "Other", ["local"]);
        var remotePath = CreateWorkbook(
            temporaryDirectory,
            "remote.xlsx",
            "Data",
            ["remote"],
            definedNames: [new TestDefinedName("Foreign", "Other!$A$1", 0)]);
        var destinationPath = temporaryDirectory.GetPath("result.xlsx");
        await File.WriteAllTextAsync(destinationPath, "existing destination");
        await using var workspace = CreateWorkspace(temporaryDirectory);
        var plan = await CreateRemoteOnlyPlanAsync(basePath, localPath, remotePath, workspace);
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

        Assert.AreEqual(OpenXmlWriterError.UnsupportedOperation, exception.Error);
        Assert.AreEqual("existing destination", await File.ReadAllTextAsync(destinationPath));
        AssertNoTransactionFiles(temporaryDirectory.Path);
    }

    private static Workspace CreateWorkspace(TestDirectory directory) =>
        new(new WorkspaceOptions { BaseDirectory = directory.Path });

    private static string CreateWorkbook(
        TestDirectory directory,
        string fileName,
        string dataSheetName,
        IReadOnlyList<string> dataValues,
        string? summaryFormula = null,
        IReadOnlyList<TestDefinedName>? definedNames = null,
        TestTable? table = null)
    {
        var path = directory.GetPath(fileName);
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        var dataPart = AddDataSheet(
            workbookPart,
            sheets,
            dataSheetName,
            sheetId: 1,
            dataValues,
            table);
        if (summaryFormula is not null)
        {
            AddSummarySheet(workbookPart, sheets, summaryFormula, sheetId: 2);
        }

        if (definedNames is { Count: > 0 })
        {
            var container = new DefinedNames();
            foreach (var definition in definedNames)
            {
                var name = new DefinedName
                {
                    Name = definition.Name,
                    LocalSheetId = definition.LocalSheetId,
                    Text = definition.Text,
                };
                container.Append(name);
            }

            workbookPart.Workbook.DefinedNames = container;
        }

        dataPart.Worksheet.Save();
        workbookPart.Workbook.Save();
        return path;
    }

    private static WorksheetPart AddDataSheet(
        WorkbookPart workbookPart,
        Sheets sheets,
        string name,
        uint sheetId,
        IReadOnlyList<string> values,
        TestTable? table)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        for (var index = 0; index < values.Count; index++)
        {
            var reference = $"A{index + 1}";
            sheetData.Append(new Row(
                new Cell(
                    new InlineString(new Text(values[index])))
                {
                    CellReference = reference,
                    DataType = CellValues.InlineString,
                })
            {
                RowIndex = checked((uint)index + 1),
            });
        }

        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension
            {
                Reference = values.Count == 0 ? "A1" : $"A1:A{values.Count}",
            },
            sheetData);
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = name,
        });
        if (table is not null)
        {
            AddTable(worksheetPart, table, values.Count);
        }

        return worksheetPart;
    }

    private static void AddSummarySheet(
        WorkbookPart workbookPart,
        Sheets sheets,
        string formula,
        uint sheetId)
    {
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetDimension { Reference = "A1" },
            new SheetData(
                new Row(
                    new Cell
                    {
                        CellReference = "A1",
                        CellFormula = new CellFormula(formula),
                        CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("0"),
                    })
                {
                    RowIndex = 1,
                }));
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = "Summary",
        });
        worksheetPart.Worksheet.Save();
    }

    private static void AddTextCell(string path, string reference, string value)
    {
        using var document = SpreadsheetDocument.Open(path, isEditable: true);
        var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet;
        var row = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Single();
        row.Append(new Cell(new InlineString(new Text(value)))
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
        });
        worksheet.SheetDimension!.Reference = $"A1:{reference}";
        worksheet.Save();
    }

    private static void AddTable(WorksheetPart worksheetPart, TestTable definition, int rowCount)
    {
        var tablePart = worksheetPart.AddNewPart<TableDefinitionPart>();
        var reference = $"A1:A{rowCount}";
        tablePart.Table = new Table(
            new AutoFilter { Reference = reference },
            new TableColumns(
                new TableColumn
                {
                    Id = 1,
                    Name = "Value",
                })
            {
                Count = 1,
            },
            new TableStyleInfo
            {
                Name = "TableStyleMedium2",
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false,
            })
        {
            Id = definition.Id,
            Name = definition.Name,
            DisplayName = definition.Name,
            Reference = reference,
            TotalsRowShown = false,
        };
        tablePart.Table.Save();
        worksheetPart.Worksheet.Append(new TableParts(
            new TablePart { Id = worksheetPart.GetIdOfPart(tablePart) })
        {
            Count = 1,
        });
    }

    private static async Task<MergePlan> CreateDataPlanAsync(
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
            baseResult.Worksheets[0],
            localResult.Worksheets[0],
            remoteResult.Worksheets[0],
            new ThreeWayMergeOptions { SheetGroupId = "data" });
        return MergePlanFactory.Create(
            baseResult.Metadata,
            localResult.Metadata,
            remoteResult.Metadata,
            [merge]);
    }

    private static async Task<MergePlan> CreateRemoteOnlyPlanAsync(
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
            baseSheet: null,
            localSheet: null,
            remoteResult.Worksheets[0],
            new ThreeWayMergeOptions { SheetGroupId = "remote-only" });
        return MergePlanFactory.Create(
            baseResult.Metadata,
            localResult.Metadata,
            remoteResult.Metadata,
            [merge]);
    }

    private static async Task WriteAsync(
        MergePlan plan,
        string basePath,
        string localPath,
        string remotePath,
        string destinationPath)
    {
        var request = await CreateWriteRequestAsync(
            plan,
            basePath,
            localPath,
            remotePath,
            destinationPath);
        await new OpenXmlWorkbookWriter().WriteAsync(
            request,
            new OpenXmlWriterOptions { MinimumFreeSpaceReserveBytes = 0 });
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

    private static void AssertNoTransactionFiles(string directory) =>
        Assert.AreEqual(
            0,
            Directory.GetFiles(directory, ".*.tmp", SearchOption.TopDirectoryOnly).Length);

    private sealed record TestDefinedName(string Name, string Text, uint? LocalSheetId = null);

    private sealed record TestTable(uint Id, string Name);
}

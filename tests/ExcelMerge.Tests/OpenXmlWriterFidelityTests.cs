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
        var workbookPart = document.WorkbookPart
            ?? throw new AssertFailedException("The result must contain a workbook part.");
        var workbook = workbookPart.Workbook
            ?? throw new AssertFailedException("The result must contain a workbook.");
        var sheetContainer = workbook.Sheets
            ?? throw new AssertFailedException("The result workbook must contain sheets.");
        var sheets = sheetContainer.Elements<Sheet>().ToArray();
        Assert.AreEqual(2, sheets.Length);
        var dataRelationshipId = sheets[0].Id?.Value
            ?? throw new AssertFailedException("The data sheet must have a relationship ID.");
        var summaryRelationshipId = sheets[1].Id?.Value
            ?? throw new AssertFailedException("The summary sheet must have a relationship ID.");
        var dataPart = workbookPart.GetPartById(dataRelationshipId) as WorksheetPart
            ?? throw new AssertFailedException("The data sheet relationship must target a worksheet part.");
        var summaryPart = workbookPart.GetPartById(summaryRelationshipId) as WorksheetPart
            ?? throw new AssertFailedException("The summary sheet relationship must target a worksheet part.");
        var dataWorksheet = dataPart.Worksheet
            ?? throw new AssertFailedException("The data worksheet part must contain a worksheet.");
        var summaryWorksheet = summaryPart.Worksheet
            ?? throw new AssertFailedException("The summary worksheet part must contain a worksheet.");
        var dimensionReference = dataWorksheet.SheetDimension?.Reference?.Value
            ?? throw new AssertFailedException("The data worksheet must have a dimension reference.");
        Assert.AreEqual("A1:A3", dimensionReference);
        Assert.AreEqual("Data!A3", summaryWorksheet.Descendants<CellFormula>().Single().Text);

        var definedNames = workbook.DefinedNames
            ?? throw new AssertFailedException("The result workbook must contain defined names.");
        var names = new Dictionary<string, DefinedName>(StringComparer.Ordinal);
        foreach (var definedName in definedNames.Elements<DefinedName>())
        {
            var name = definedName.Name?.Value
                ?? throw new AssertFailedException("Every defined name must have a name attribute.");
            Assert.IsTrue(names.TryAdd(name, definedName), $"Defined name '{name}' must be unique.");
        }

        Assert.AreEqual(5, names.Count);
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
        var workbookPart = document.WorkbookPart
            ?? throw new AssertFailedException("The result must contain a workbook part.");
        var workbook = workbookPart.Workbook
            ?? throw new AssertFailedException("The result must contain a workbook.");
        var sheetContainer = workbook.Sheets
            ?? throw new AssertFailedException("The result workbook must contain sheets.");
        var sheetNames = new List<string>();
        foreach (var sheet in sheetContainer.Elements<Sheet>())
        {
            var sheetName = sheet.Name?.Value
                ?? throw new AssertFailedException("Every result sheet must have a name attribute.");
            sheetNames.Add(sheetName);
        }

        CollectionAssert.AreEqual(
            new[] { "REMOTE Name", "Summary" },
            sheetNames);
        var formulaWorksheets = new List<Worksheet>();
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var worksheet = worksheetPart.Worksheet
                ?? throw new AssertFailedException("Every worksheet part must contain a worksheet.");
            if (worksheet.Descendants<CellFormula>().Any())
            {
                formulaWorksheets.Add(worksheet);
            }
        }

        var summaryWorksheet = formulaWorksheets.Single();
        Assert.AreEqual(
            "'REMOTE Name'!A1",
            summaryWorksheet.Descendants<CellFormula>().Single().Text);
        var definedNames = workbook.DefinedNames
            ?? throw new AssertFailedException("The result workbook must contain defined names.");
        Assert.AreEqual(
            "'REMOTE Name'!$A$1",
            definedNames.Elements<DefinedName>().Single().Text);
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
        var workbookPart = document.WorkbookPart
            ?? throw new AssertFailedException("The result must contain a workbook part.");
        var workbook = workbookPart.Workbook
            ?? throw new AssertFailedException("The result must contain a workbook.");
        var tables = new List<Table>();
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            foreach (var tablePart in worksheetPart.TableDefinitionParts)
            {
                var table = tablePart.Table
                    ?? throw new AssertFailedException("Every table definition part must contain a table.");
                tables.Add(table);
            }
        }

        Assert.AreEqual(2, tables.Count);
        var tableIds = new List<uint>(tables.Count);
        var tableNames = new List<string>(tables.Count);
        foreach (var table in tables)
        {
            var tableId = table.Id?.Value
                ?? throw new AssertFailedException("Every imported table must have an ID attribute.");
            var tableName = table.Name?.Value
                ?? throw new AssertFailedException("Every imported table must have a name attribute.");
            tableIds.Add(tableId);
            tableNames.Add(tableName);
        }

        Assert.AreEqual(2, tableIds.Distinct().Count());
        CollectionAssert.AreEquivalent(
            new[] { "LocalTable", "RemoteTable" },
            tableNames);
        var definedNames = workbook.DefinedNames
            ?? throw new AssertFailedException("The result workbook must contain defined names.");
        var importedName = definedNames.Elements<DefinedName>().Single();
        var localSheetId = importedName.LocalSheetId?.Value
            ?? throw new AssertFailedException("The imported defined name must have a local sheet ID.");
        Assert.AreEqual(1U, localSheetId);
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

        var exception = await Assert.ThrowsExactlyAsync<OpenXmlWriterException>(() =>
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

        var exception = await Assert.ThrowsExactlyAsync<OpenXmlWriterException>(() =>
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

        var exception = await Assert.ThrowsExactlyAsync<OpenXmlWriterException>(() =>
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

        var exception = await Assert.ThrowsExactlyAsync<OpenXmlWriterException>(() =>
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
        var workbookPart = document.WorkbookPart
            ?? throw new AssertFailedException("The result must contain a workbook part.");
        var formulas = new List<CellFormula>();
        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var worksheet = worksheetPart.Worksheet
                ?? throw new AssertFailedException("Every worksheet part must contain a worksheet.");
            formulas.AddRange(worksheet.Descendants<CellFormula>());
        }

        var formula = formulas.Single();
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
        var workbookPart = document.WorkbookPart
            ?? throw new AssertFailedException("The result must contain a workbook part.");
        var worksheet = workbookPart.WorksheetParts.Single().Worksheet
            ?? throw new AssertFailedException("The result worksheet part must contain a worksheet.");
        var dimensionReference = worksheet.SheetDimension?.Reference?.Value
            ?? throw new AssertFailedException("The result worksheet must have a dimension reference.");
        Assert.AreEqual("A1:D1", dimensionReference);
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

        var exception = await Assert.ThrowsExactlyAsync<OpenXmlWriterException>(() =>
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
        var workbook = new Workbook();
        workbookPart.Workbook = workbook;
        var sheets = workbook.AppendChild(new Sheets());
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

            workbook.DefinedNames = container;
        }

        var dataWorksheet = dataPart.Worksheet
            ?? throw new AssertFailedException("The data worksheet part must contain a worksheet.");
        dataWorksheet.Save();
        workbook.Save();
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
        var workbookPart = document.WorkbookPart
            ?? throw new AssertFailedException("The workbook must contain a workbook part.");
        var worksheet = workbookPart.WorksheetParts.Single().Worksheet
            ?? throw new AssertFailedException("The worksheet part must contain a worksheet.");
        var sheetData = worksheet.GetFirstChild<SheetData>()
            ?? throw new AssertFailedException("The worksheet must contain sheet data.");
        var row = sheetData.Elements<Row>().Single();
        row.Append(new Cell(new InlineString(new Text(value)))
        {
            CellReference = reference,
            DataType = CellValues.InlineString,
        });
        var sheetDimension = worksheet.SheetDimension
            ?? throw new AssertFailedException("The worksheet must contain a sheet dimension.");
        sheetDimension.Reference = $"A1:{reference}";
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
        var worksheet = worksheetPart.Worksheet
            ?? throw new AssertFailedException("The worksheet part must contain a worksheet.");
        worksheet.Append(new TableParts(
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

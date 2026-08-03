using ExcelMerge.Application;
using ExcelMerge.Domain;
using ExcelMerge.OpenXml;
using ExcelMerge.Storage;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class ApplicationTests
{
    [TestMethod]
    public async Task Compare_session_dispatches_csv_and_owns_workspace_lifetime()
    {
        using var temporaryDirectory = new TestDirectory();
        var localPath = temporaryDirectory.GetPath("local.csv");
        var remotePath = temporaryDirectory.GetPath("remote.csv");
        await File.WriteAllTextAsync(localPath, "id,value\r\n1,local");
        await File.WriteAllTextAsync(remotePath, "id,value\r\n1,remote");
        var reports = new List<ApplicationProgress>();
        await using var application = CreateApplication(temporaryDirectory);

        var session = await application.OpenCompareAsync(
            new CompareRequest(localPath, remotePath),
            new SynchronousProgress<ApplicationProgress>(reports.Add));

        Assert.AreEqual(WorkbookFormat.Csv, session.Local.Format);
        Assert.AreEqual(1, session.Sheets.Count);
        Assert.AreEqual(1L, session.Sheets[0].Difference.Statistics.ChangedCellCount);
        Assert.IsTrue(reports.Any(static report => report.Stage == ApplicationStage.Indexing));
        Assert.AreEqual(ApplicationStage.Completed, reports[^1].Stage);
        Assert.AreEqual(1, WorkspaceDirectories(temporaryDirectory).Length);

        await session.DisposeAsync();

        Assert.IsTrue(session.IsDisposed);
        Assert.AreEqual(0, WorkspaceDirectories(temporaryDirectory).Length);
    }

    [TestMethod]
    public async Task Application_rejects_mixed_and_legacy_formats_before_workspace_creation()
    {
        using var temporaryDirectory = new TestDirectory();
        var csvPath = temporaryDirectory.GetPath("data.csv");
        var xlsxPath = CreateTextWorkbook(temporaryDirectory, "data.xlsx", "value");
        var xlsPath = temporaryDirectory.GetPath("legacy.xls");
        await File.WriteAllTextAsync(csvPath, "value");
        await File.WriteAllTextAsync(xlsPath, "legacy");
        await using var application = CreateApplication(temporaryDirectory);

        var mixed = await Assert.ThrowsExactlyAsync<ExcelMergeApplicationException>(() =>
            application.OpenCompareAsync(new CompareRequest(csvPath, xlsxPath)).AsTask());
        var legacy = await Assert.ThrowsExactlyAsync<ExcelMergeApplicationException>(() =>
            application.OpenCompareAsync(new CompareRequest(xlsPath, xlsPath)).AsTask());

        Assert.AreEqual(ApplicationError.MixedFormats, mixed.Error);
        Assert.AreEqual(ApplicationError.UnsupportedLegacyWorkbook, legacy.Error);
        Assert.AreEqual(0, WorkspaceDirectories(temporaryDirectory).Length);
    }

    [TestMethod]
    public async Task Merge_session_saves_automatic_csv_result()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = temporaryDirectory.GetPath("base.csv");
        var localPath = temporaryDirectory.GetPath("local.csv");
        var remotePath = temporaryDirectory.GetPath("remote.csv");
        var resultPath = temporaryDirectory.GetPath("result.csv");
        await File.WriteAllTextAsync(basePath, "id,value\r\n1,base");
        await File.WriteAllTextAsync(localPath, "id,value\r\n1,base");
        await File.WriteAllTextAsync(remotePath, "id,value\r\n1,remote");
        await using var application = CreateApplication(temporaryDirectory);
        await using var session = await application.OpenMergeAsync(
            new MergeRequest(basePath, localPath, remotePath));

        var result = await session.SaveAsync(new SaveRequest(resultPath));

        Assert.AreEqual(WorkbookFormat.Csv, result.Format);
        Assert.IsFalse(session.HasUnresolvedConflicts);
        Assert.AreEqual("id,value\r\n1,remote\r\n", await File.ReadAllTextAsync(resultPath));
    }

    [TestMethod]
    public async Task Merge_session_requires_and_applies_cell_resolution()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = temporaryDirectory.GetPath("base.csv");
        var localPath = temporaryDirectory.GetPath("local.csv");
        var remotePath = temporaryDirectory.GetPath("remote.csv");
        var resultPath = temporaryDirectory.GetPath("result.csv");
        await File.WriteAllTextAsync(basePath, "base");
        await File.WriteAllTextAsync(localPath, "LOCAL");
        await File.WriteAllTextAsync(remotePath, "REMOTE");
        await using var application = CreateApplication(temporaryDirectory);
        await using var session = await application.OpenMergeAsync(
            new MergeRequest(basePath, localPath, remotePath));
        var conflict = session.Conflicts.Single();

        var unresolved = await Assert.ThrowsExactlyAsync<ExcelMergeApplicationException>(() =>
            session.SaveAsync(new SaveRequest(resultPath)).AsTask());
        session.ResolveCell(conflict.Id, ResolutionKind.Remote);
        await session.SaveAsync(new SaveRequest(resultPath));

        Assert.AreEqual(ApplicationError.UnresolvedConflicts, unresolved.Error);
        Assert.IsFalse(session.HasUnresolvedConflicts);
        Assert.AreEqual("REMOTE\r\n", await File.ReadAllTextAsync(resultPath));
        Assert.ThrowsExactly<ExcelMergeApplicationException>(() =>
            session.ResolveRow(conflict.Id, ResolutionKind.Local));
    }

    [TestMethod]
    public async Task Merge_session_emits_local_then_remote_for_row_both()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = temporaryDirectory.GetPath("base.tsv");
        var localPath = temporaryDirectory.GetPath("local.tsv");
        var remotePath = temporaryDirectory.GetPath("remote.tsv");
        var resultPath = temporaryDirectory.GetPath("result.tsv");
        await File.WriteAllTextAsync(basePath, "A\r\nB");
        await File.WriteAllTextAsync(localPath, "A\r\nLOCAL\r\nB");
        await File.WriteAllTextAsync(remotePath, "A\r\nREMOTE\r\nB");
        await using var application = CreateApplication(temporaryDirectory);
        await using var session = await application.OpenMergeAsync(
            new MergeRequest(basePath, localPath, remotePath));
        var conflict = session.Conflicts.Single();
        Assert.AreEqual(ConflictKind.RowAddAdd, conflict.Kind);

        session.ResolveRow(conflict.Id, ResolutionKind.Both);
        await session.SaveAsync(new SaveRequest(resultPath));

        Assert.AreEqual("A\r\nLOCAL\r\nREMOTE\r\nB\r\n", await File.ReadAllTextAsync(resultPath));
    }

    [TestMethod]
    public async Task Merge_session_both_override_duplicates_a_cell_conflict_row()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = temporaryDirectory.GetPath("base.csv");
        var localPath = temporaryDirectory.GetPath("local.csv");
        var remotePath = temporaryDirectory.GetPath("remote.csv");
        var resultPath = temporaryDirectory.GetPath("result.csv");
        await File.WriteAllTextAsync(basePath, "base-1,base-2");
        await File.WriteAllTextAsync(localPath, "LOCAL-1,LOCAL-2");
        await File.WriteAllTextAsync(remotePath, "REMOTE-1,REMOTE-2");
        await using var application = CreateApplication(temporaryDirectory);
        await using var session = await application.OpenMergeAsync(
            new MergeRequest(basePath, localPath, remotePath));
        Assert.AreEqual(2, session.Conflicts.Count);
        Assert.IsTrue(session.Conflicts.All(static item => item.Kind == ConflictKind.CellValue));
        var conflict = session.Conflicts[0];

        session.SetRowOverride(new RowMergeOverride(
            conflict.Location.SheetId,
            conflict.Location.BaseRowIndex,
            conflict.Location.LocalRowIndex,
            conflict.Location.RemoteRowIndex,
            ResolutionKind.Both));
        await session.SaveAsync(new SaveRequest(resultPath));

        Assert.IsFalse(session.HasUnresolvedConflicts);
        Assert.AreEqual(
            "LOCAL-1,LOCAL-2\r\nREMOTE-1,REMOTE-2\r\n",
            await File.ReadAllTextAsync(resultPath));
    }

    [TestMethod]
    public async Task Merge_session_rejects_changed_delimited_source_before_save()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = temporaryDirectory.GetPath("base.csv");
        var localPath = temporaryDirectory.GetPath("local.csv");
        var remotePath = temporaryDirectory.GetPath("remote.csv");
        var resultPath = temporaryDirectory.GetPath("result.csv");
        await File.WriteAllTextAsync(basePath, "base");
        await File.WriteAllTextAsync(localPath, "base");
        await File.WriteAllTextAsync(remotePath, "remote");
        await using var application = CreateApplication(temporaryDirectory);
        await using var session = await application.OpenMergeAsync(
            new MergeRequest(basePath, localPath, remotePath));
        await File.AppendAllTextAsync(remotePath, " changed");

        var exception = await Assert.ThrowsExactlyAsync<ExcelMergeApplicationException>(() =>
            session.SaveAsync(new SaveRequest(resultPath)).AsTask());

        Assert.AreEqual(ApplicationError.InvalidRequest, exception.Error);
        Assert.IsFalse(File.Exists(resultPath));
    }

    [TestMethod]
    public async Task Merge_session_routes_xlsx_save_through_transactional_writer()
    {
        using var temporaryDirectory = new TestDirectory();
        var basePath = CreateTextWorkbook(temporaryDirectory, "base.xlsx", "base");
        var localPath = CreateTextWorkbook(temporaryDirectory, "local.xlsx", "base");
        var remotePath = CreateTextWorkbook(temporaryDirectory, "remote.xlsx", "remote");
        var resultPath = temporaryDirectory.GetPath("result.xlsx");
        await using var application = CreateApplication(temporaryDirectory);
        await using var session = await application.OpenMergeAsync(
            new MergeRequest(basePath, localPath, remotePath));

        await session.SaveAsync(new SaveRequest(
            resultPath,
            MinimumFreeSpaceReserveBytes: 0));

        await using var workspace = new Workspace(new WorkspaceOptions
        {
            BaseDirectory = temporaryDirectory.Path,
        });
        var output = await new OpenXmlWorkbookReader().IndexAsync(resultPath, workspace);
        var row = await output.Worksheets.Single().GetRowAsync(0);
        Assert.AreEqual("remote", row.GetValueOrDefault().Cells.Span[0].Value.Scalar?.TextValue);
    }

    [TestMethod]
    public async Task Disposing_application_disposes_open_sessions()
    {
        using var temporaryDirectory = new TestDirectory();
        var localPath = temporaryDirectory.GetPath("local.csv");
        var remotePath = temporaryDirectory.GetPath("remote.csv");
        await File.WriteAllTextAsync(localPath, "local");
        await File.WriteAllTextAsync(remotePath, "remote");
        var application = CreateApplication(temporaryDirectory);
        var session = await application.OpenCompareAsync(new CompareRequest(localPath, remotePath));

        await application.DisposeAsync();

        Assert.IsTrue(session.IsDisposed);
        Assert.AreEqual(0, WorkspaceDirectories(temporaryDirectory).Length);
        await Assert.ThrowsExactlyAsync<ObjectDisposedException>(() =>
            application.OpenCompareAsync(new CompareRequest(localPath, remotePath)).AsTask());
    }

    [TestMethod]
    public async Task Json_state_store_round_trips_settings_and_bounded_recent_descriptors()
    {
        using var temporaryDirectory = new TestDirectory();
        var statePath = temporaryDirectory.GetPath("state/application.json");
        var localPath = temporaryDirectory.GetPath("local.csv");
        var remotePath = temporaryDirectory.GetPath("remote.csv");
        var basePath = temporaryDirectory.GetPath("base.csv");
        var firstId = Guid.NewGuid();
        var secondId = Guid.NewGuid();
        var thirdId = Guid.NewGuid();
        await using (var store = new JsonApplicationStateStore(statePath))
        {
            Assert.AreEqual(20, (await store.LoadAsync()).MaximumRecentSessions);
            await store.SaveAsync(new ApplicationSettings
            {
                MaximumRecentSessions = 2,
                KeyColumns = [0, 2],
                CompareCellStyles = true,
            });
            await store.RecordAsync(new RecentSessionDescriptor(
                firstId,
                RecentSessionMode.Compare,
                null,
                localPath,
                remotePath,
                DateTimeOffset.Parse("2026-07-31T10:00:00Z")));
            await store.RecordAsync(new RecentSessionDescriptor(
                secondId,
                RecentSessionMode.Merge,
                basePath,
                localPath,
                remotePath,
                DateTimeOffset.Parse("2026-07-31T11:00:00Z")));
            await store.RecordAsync(new RecentSessionDescriptor(
                thirdId,
                RecentSessionMode.Compare,
                null,
                localPath,
                remotePath,
                DateTimeOffset.Parse("2026-07-31T12:00:00Z")));
        }

        await using var reopened = new JsonApplicationStateStore(statePath);
        var settings = await reopened.LoadAsync();
        var recent = await reopened.ListAsync();

        CollectionAssert.AreEqual(new[] { 0, 2 }, settings.KeyColumns.ToArray());
        Assert.IsTrue(settings.CompareCellStyles);
        CollectionAssert.AreEqual(new[] { thirdId, secondId }, recent.Select(static item => item.Id).ToArray());
        Assert.AreEqual(Path.GetFullPath(basePath), recent[1].BasePath);
        Assert.IsTrue(await reopened.RemoveAsync(secondId));
        Assert.IsFalse(await reopened.RemoveAsync(firstId));
        Assert.AreEqual(1, (await reopened.ListAsync()).Count);
        var json = await File.ReadAllTextAsync(statePath);
        Assert.IsFalse(json.Contains("workspace", StringComparison.OrdinalIgnoreCase));
        Assert.IsFalse(json.Contains("mergePlan", StringComparison.OrdinalIgnoreCase));
    }

    private static ExcelMergeApplication CreateApplication(TestDirectory directory) =>
        new(new ExcelMergeApplicationOptions
        {
            Workspace = new WorkspaceOptions
            {
                BaseDirectory = directory.Path,
                DirectoryPrefix = "app-workspace-",
            },
            ProgressIntervalRows = 1,
        });

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

    private static string[] WorkspaceDirectories(TestDirectory directory) =>
        Directory.GetDirectories(directory.Path, "app-workspace-*");
}

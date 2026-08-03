using Avalonia;
using ExcelMerge.Application;
using ExcelMerge.Desktop;
using ExcelMerge.Domain;
using ExcelMerge.Engine;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class DesktopTests
{
    [TestMethod]
    public async Task Merge_grid_accepts_zero_conflict_rows_with_negative_start_sentinel()
    {
        var baseSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("unchanged"))),
            TestData.Row(1, (0, TestData.Text("base"))));
        var localSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("unchanged"))),
            TestData.Row(1, (0, TestData.Text("local"))));
        var remoteSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("unchanged"))),
            TestData.Row(1, (0, TestData.Text("remote"))));
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet,
            new ThreeWayMergeOptions { SheetGroupId = "sheet" });
        var result = new MergeSheetResult(
            "sheet",
            baseSheet,
            localSheet,
            remoteSheet,
            merge);

        Assert.AreEqual(-1, merge.ViewRows.Span[0].ConflictStartIndex);
        Assert.AreEqual(0, merge.ViewRows.Span[0].ConflictCount);

        var document = GridDocument.FromMerge(result, new Dictionary<long, ResolutionKind>());
        Assert.AreEqual(2, document.RowCount);
        Assert.AreEqual(GridRowVisualState.Unchanged, document.Rows[0].State);
        Assert.AreEqual(GridRowVisualState.Conflict, document.Rows[1].State);
        Assert.AreEqual(GridCellVisualState.Conflict, document.GetCellVisualState(1, 0));
        Assert.AreEqual(GridCellVisualState.None, document.GetCellVisualState(1, 1));
        Assert.AreEqual(GridCellVisualState.None, document.GetRowHeaderVisualState(1));
        Assert.AreEqual(new GridCellMarker(1, 0, GridCellVisualState.Conflict), document.CellMarkers.Single());
        Assert.AreEqual(0, document.RowMarkers.Count);
        var conflictItem = new ConflictItemViewModel(
            merge.Conflicts.Span[0],
            ExcelMerge.Domain.ResolutionKind.Unresolved);
        Assert.AreEqual("A2", conflictItem.Location);
        Assert.AreEqual("base", conflictItem.BaseValue);
        Assert.AreEqual("local", conflictItem.LocalValue);
        Assert.AreEqual("remote", conflictItem.RemoteValue);
        Assert.IsTrue(conflictItem.Matches(document.Rows[1], 0));
        Assert.IsFalse(conflictItem.Matches(document.Rows[1], 1));
        Assert.IsFalse(conflictItem.Matches(document.Rows[0], 0));

        var grid = new VirtualDiffGrid { Document = document };
        grid.Measure(new Size(1_000, 80));
        grid.Arrange(new Rect(0, 0, 1_000, 80));
        Assert.AreEqual(2, grid.VerticalViewportSize);
        Assert.AreEqual(0, grid.VerticalScrollMaximum);

        grid.ResizeRow(0, 80);
        Assert.AreEqual(1, grid.VerticalViewportSize);
        Assert.AreEqual(1, grid.VerticalScrollMaximum);
        grid.ScrollVerticalTo(double.MaxValue);
        Assert.AreEqual(1, grid.FirstVisibleRow);

        var resolved = GridDocument.FromMerge(
            result,
            new Dictionary<long, ResolutionKind>
            {
                [merge.Conflicts.Span[0].Id] = ResolutionKind.Local,
            },
            hideUnchanged: true);
        Assert.AreEqual(1, resolved.RowCount);
        Assert.AreEqual(GridRowVisualState.Resolved, resolved.Rows[0].State);
        Assert.AreEqual(GridCellVisualState.Resolved, resolved.GetCellVisualState(0, 0));
        Assert.AreEqual(GridCellVisualState.None, resolved.GetCellVisualState(0, 1));
        Assert.AreEqual(GridCellVisualState.None, resolved.GetRowHeaderVisualState(0));

        grid.ResizeRow(1, 70);
        grid.Document = resolved;
        Assert.AreEqual(70, grid.GetRowHeight(0));
    }

    [TestMethod]
    public async Task Structural_conflict_marks_row_headers_without_marking_cells()
    {
        var baseSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("A"))),
            TestData.Row(1, (0, TestData.Text("B"))));
        var localSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("A"))),
            TestData.Row(1, (0, TestData.Text("local insert"))),
            TestData.Row(2, (0, TestData.Text("B"))));
        var remoteSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("A"))),
            TestData.Row(1, (0, TestData.Text("remote insert"))),
            TestData.Row(2, (0, TestData.Text("B"))));
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet,
            new ThreeWayMergeOptions { SheetGroupId = "sheet" });
        var result = new MergeSheetResult("sheet", baseSheet, localSheet, remoteSheet, merge);
        var conflict = merge.Conflicts.ToArray().Single();
        var document = GridDocument.FromMerge(result, new Dictionary<long, ResolutionKind>());
        var conflictRow = Enumerable.Range(0, document.RowCount)
            .Single(index => document.Rows[index].State == GridRowVisualState.Conflict);

        Assert.IsNull(conflict.Location.ColumnIndex);
        Assert.AreEqual(GridCellVisualState.Conflict, document.GetRowHeaderVisualState(conflictRow));
        Assert.AreEqual(GridCellVisualState.None, document.GetCellVisualState(conflictRow, 0));
        Assert.AreEqual(
            new GridRowMarker(conflictRow, GridCellVisualState.Conflict),
            document.RowMarkers.Single());
        Assert.AreEqual(0, document.CellMarkers.Count);

        var resolved = GridDocument.FromMerge(
            result,
            new Dictionary<long, ResolutionKind> { [conflict.Id] = ResolutionKind.Local });
        Assert.AreEqual(GridCellVisualState.Resolved, resolved.GetRowHeaderVisualState(conflictRow));
        Assert.AreEqual(GridCellVisualState.None, resolved.GetCellVisualState(conflictRow, 0));
    }

    [TestMethod]
    public async Task Merge_grid_projects_both_as_local_then_remote_rows()
    {
        var baseSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("base"))));
        var localSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("LOCAL"))));
        var remoteSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("REMOTE"))));
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet,
            new ThreeWayMergeOptions { SheetGroupId = "sheet" });
        var result = new MergeSheetResult("sheet", baseSheet, localSheet, remoteSheet, merge);
        var conflict = merge.Conflicts.ToArray().Single();

        var document = GridDocument.FromMerge(
            result,
            new Dictionary<long, ResolutionKind> { [conflict.Id] = ResolutionKind.Both });

        Assert.AreEqual(2, document.RowCount);
        Assert.AreEqual(GridRowProjection.Local, document.Rows[0].Projection);
        Assert.AreEqual(0, document.Rows[0].LocalRowIndex);
        Assert.IsNull(document.Rows[0].RemoteRowIndex);
        Assert.AreEqual(GridRowProjection.Remote, document.Rows[1].Projection);
        Assert.IsNull(document.Rows[1].LocalRowIndex);
        Assert.AreEqual(0, document.Rows[1].RemoteRowIndex);
        Assert.AreEqual(GridCellVisualState.Resolved, document.GetCellVisualState(0, 0));
        Assert.AreEqual(GridCellVisualState.Resolved, document.GetCellVisualState(1, 0));
        Assert.AreEqual(0, document.FindViewRow(0, 0, 0));

        var conflictItem = new ConflictItemViewModel(conflict, ResolutionKind.Both);
        Assert.IsTrue(conflictItem.SupportsBoth);
        Assert.IsTrue(conflictItem.Matches(document.Rows[0], 0));
        Assert.IsTrue(conflictItem.Matches(document.Rows[1], 0));
    }

    [TestMethod]
    public async Task Map_position_centers_target_row_and_column_in_viewport()
    {
        var metadata = new ExcelMerge.Domain.SheetMetadata(
            "rId1",
            "Sheet1",
            0,
            LastRowIndex: 99,
            LastColumnIndex: 30);
        var rows = Enumerable.Range(0, 100)
            .Select(index => TestData.Row(
                index,
                (0, TestData.Text($"first-{index}")),
                (30, TestData.Text($"last-{index}"))))
            .ToArray();
        var sheet = TestData.Sheet(metadata, rows);
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            sheet,
            sheet,
            sheet,
            new ThreeWayMergeOptions { SheetGroupId = "sheet" });
        var document = GridDocument.FromMerge(
            new MergeSheetResult("sheet", sheet, sheet, sheet, merge),
            new Dictionary<long, ResolutionKind>());
        var grid = new VirtualDiffGrid { Document = document };
        grid.Measure(new Size(1_000, 500));
        grid.Arrange(new Rect(0, 0, 1_000, 500));

        grid.CenterOnMapPosition(0.5, 0.5);

        var visibleRowCenter = grid.FirstVisibleRow + ((grid.VerticalViewportSize - 1) / 2d);
        var visibleColumnCenter = grid.FirstVisibleColumn + ((grid.HorizontalViewportSize - 1) / 2d);
        Assert.AreEqual(50, visibleRowCenter, 0.5);
        Assert.AreEqual(15, visibleColumnCenter, 0.5);

        grid.CenterOnMapPosition(1, 1);
        Assert.AreEqual(grid.VerticalScrollMaximum, grid.FirstVisibleRow);
        Assert.AreEqual(grid.HorizontalScrollMaximum, grid.FirstVisibleColumn);
    }

    [TestMethod]
    public async Task Grid_horizontal_scroll_metrics_follow_pane_width_and_clamp_offset()
    {
        var metadata = new ExcelMerge.Domain.SheetMetadata(
            "rId1",
            "Sheet1",
            0,
            LastColumnIndex: 30);
        var baseSheet = TestData.Sheet(metadata, TestData.Row(
            0,
            (0, TestData.Text("first")),
            (30, TestData.Text("last"))));
        var localSheet = TestData.Sheet(metadata, TestData.Row(
            0,
            (0, TestData.Text("first")),
            (30, TestData.Text("last"))));
        var remoteSheet = TestData.Sheet(metadata, TestData.Row(
            0,
            (0, TestData.Text("first")),
            (30, TestData.Text("last"))));
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet,
            new ThreeWayMergeOptions { SheetGroupId = "sheet" });
        var document = GridDocument.FromMerge(
            new MergeSheetResult("sheet", baseSheet, localSheet, remoteSheet, merge),
            new Dictionary<long, ResolutionKind>());
        var grid = new VirtualDiffGrid { Document = document };

        grid.Measure(new Size(1_000, 500));
        grid.Arrange(new Rect(0, 0, 1_000, 500));

        Assert.AreEqual(4, grid.HorizontalViewportSize);
        Assert.AreEqual(27, grid.HorizontalScrollMaximum);
        var localPane = grid.GetPaneBounds(GridPane.Local);
        var remotePane = grid.GetPaneBounds(GridPane.Remote);
        Assert.AreEqual(500, localPane.Right);
        Assert.AreEqual(localPane.Right, remotePane.X);
        Assert.IsTrue(
            48 + (grid.HorizontalViewportSize * grid.GetColumnWidth(0)) > localPane.Right,
            "The partially visible LOCAL cell must be clipped at the pane boundary.");
        grid.ScrollHorizontalTo(double.MaxValue);
        Assert.AreEqual(27, grid.FirstVisibleColumn);

        grid.Measure(new Size(2_000, 500));
        grid.Arrange(new Rect(0, 0, 2_000, 500));

        Assert.AreEqual(9, grid.HorizontalViewportSize);
        Assert.AreEqual(22, grid.HorizontalScrollMaximum);
        Assert.AreEqual(22, grid.FirstVisibleColumn);

        grid.Measure(new Size(1_000, 500));
        grid.Arrange(new Rect(0, 0, 1_000, 500));
        grid.ScrollHorizontalTo(0);
        grid.ResizeColumn(0, 300);
        grid.ResizeRow(0, 80);

        Assert.AreEqual(300, grid.GetColumnWidth(0));
        Assert.AreEqual(80, grid.GetRowHeight(0));
        Assert.AreEqual(3, grid.HorizontalViewportSize);

        grid.ResizeColumn(30, 300);
        Assert.AreEqual(28, grid.HorizontalScrollMaximum);
        grid.ScrollHorizontalTo(double.MaxValue);
        Assert.AreEqual(28, grid.FirstVisibleColumn);
    }
}

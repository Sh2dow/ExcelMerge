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

    [TestMethod]
    public async Task Grid_document_surfaces_source_column_widths_and_row_heights()
    {
        var metadata = new SheetMetadata(
            "rId1",
            "Sheet1",
            0,
            ColumnWidths: [null, 20],
            DefaultColumnWidth: 10);
        var baseSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("a"))));
        var localSheet = TestData.Sheet(
            metadata,
            new RowRecord(
                0,
                new[] { new CellRecord(new CellAddress(0, 0), TestData.Text("a")) },
                height: 90));
        var remoteSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("a"))));
        var merge = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet,
            new ThreeWayMergeOptions { SheetGroupId = "sheet" });
        var result = new MergeSheetResult("sheet", baseSheet, localSheet, remoteSheet, merge);

        var document = GridDocument.FromMerge(result, new Dictionary<long, ResolutionKind>());

        Assert.IsNotNull(document.SourceColumnWidths);
        Assert.AreEqual(145d, document.SourceColumnWidths[1]);
        Assert.IsFalse(document.SourceColumnWidths.ContainsKey(0));
        Assert.AreEqual(75d, document.SourceDefaultColumnWidth);

        var grid = new VirtualDiffGrid { Document = document };
        Assert.AreEqual(145d, grid.GetColumnWidth(1));
        Assert.AreEqual(75d, grid.GetColumnWidth(0));

        var row = await document.LoadRowAsync(0);
        Assert.AreEqual(120d, GridDocument.GetSourceRowHeight(row));

        var plainDocument = GridDocument.FromMerge(
            new MergeSheetResult("sheet", baseSheet, baseSheet, remoteSheet, merge),
            new Dictionary<long, ResolutionKind>());
        Assert.IsNull(plainDocument.SourceColumnWidths);
        Assert.IsNull(GridDocument.GetSourceRowHeight(await plainDocument.LoadRowAsync(0)));

        var plainGrid = new VirtualDiffGrid { Document = plainDocument };
        Assert.AreEqual(GridSettings.ColumnWidth, plainGrid.GetColumnWidth(0));
    }

    [TestMethod]
    public void Content_width_fitter_grows_and_clamps_without_shrinking()
    {
        var fitter = new ContentWidthFitter();
        Assert.IsNull(fitter.GetWidth(0));

        Assert.IsTrue(fitter.TryGrow(0, 200));
        Assert.AreEqual(200 + ContentWidthFitter.CellPadding, fitter.GetWidth(0));

        Assert.IsFalse(fitter.TryGrow(0, 100));
        Assert.AreEqual(200 + ContentWidthFitter.CellPadding, fitter.GetWidth(0));

        Assert.IsFalse(fitter.TryGrow(0, double.NaN));
        Assert.IsTrue(fitter.TryGrow(1, 1));
        Assert.AreEqual(ContentWidthFitter.MinimumWidth, fitter.GetWidth(1));

        Assert.IsTrue(fitter.TryGrow(2, 10_000));
        Assert.AreEqual(ContentWidthFitter.MaximumWidth, fitter.GetWidth(2));

        fitter.Clear();
        Assert.IsNull(fitter.GetWidth(0));
    }

    [TestMethod]
    public void Text_differ_reports_only_changed_spans()
    {
        Assert.AreEqual(0, TextDiffer.GetChangedSpans("same", "same").Length);
        Assert.AreEqual(0, TextDiffer.GetChangedSpans(null, "").Length);
        Assert.AreEqual(0, TextDiffer.GetChangedSpans("anything", "").Length);

        CollectionAssert.AreEqual(
            new[] { new TextSpan(0, 5) },
            TextDiffer.GetChangedSpans("", "whole"));
        CollectionAssert.AreEqual(
            new[] { new TextSpan(0, 5) },
            TextDiffer.GetChangedSpans("xxxxx", "whole"));

        CollectionAssert.AreEqual(
            new[] { new TextSpan(6, 6) },
            TextDiffer.GetChangedSpans("hello world", "hello brave world"));
        CollectionAssert.AreEqual(
            new[] { new TextSpan(3, 3) },
            TextDiffer.GetChangedSpans("abc123xyz", "abc456xyz"));
        CollectionAssert.AreEqual(
            new[] { new TextSpan(3, 3) },
            TextDiffer.GetChangedSpans("abc456", "abcdef"));

        var spans = TextDiffer.GetChangedSpans("a1b2c3", "a1x2y3");
        Assert.AreEqual(2, spans.Length);
        Assert.AreEqual(new TextSpan(2, 1), spans[0]);
        Assert.AreEqual(new TextSpan(4, 1), spans[1]);
    }

    [TestMethod]
    public void Text_differ_falls_back_for_oversized_inputs()
    {
        var oldLong = new string('a', TextDiffer.MaxInputLength + 1);
        CollectionAssert.AreEqual(
            new[] { new TextSpan(0, 5) },
            TextDiffer.GetChangedSpans(oldLong, "short"));

        var prefix = new string('p', TextDiffer.MaxMiddleLength);
        var suffix = new string('s', TextDiffer.MaxMiddleLength);
        var oldMiddle = new string('o', TextDiffer.MaxMiddleLength + 1);
        var spans = TextDiffer.GetChangedSpans(prefix + oldMiddle + suffix, prefix + "new" + suffix);
        CollectionAssert.AreEqual(
            new[] { new TextSpan(TextDiffer.MaxMiddleLength, 3) },
            spans);

        var aligned = new string('z', TextDiffer.MaxMiddleLength);
        Assert.AreEqual(
            1,
            TextDiffer.GetChangedSpans(aligned + "1", aligned + "2").Length);
    }

    [TestMethod]
    public void Grid_palette_follows_the_effective_theme_variant()
    {
        Assert.AreSame(GridPalette.Light, GridPalette.ForVariant(Avalonia.Styling.ThemeVariant.Light));
        Assert.AreSame(GridPalette.Light, GridPalette.ForVariant(Avalonia.Styling.ThemeVariant.Default));
        Assert.AreSame(GridPalette.Light, GridPalette.ForVariant(null));
        Assert.AreSame(GridPalette.Dark, GridPalette.ForVariant(Avalonia.Styling.ThemeVariant.Dark));

        var lightText = (Avalonia.Media.SolidColorBrush)GridPalette.Light.Text;
        var darkText = (Avalonia.Media.SolidColorBrush)GridPalette.Dark.Text;
        Assert.AreEqual(Avalonia.Media.Color.Parse("#202428"), lightText.Color);
        Assert.AreNotEqual(lightText.Color, darkText.Color);
        var lightBackground = (Avalonia.Media.SolidColorBrush)GridPalette.Light.Background;
        Assert.AreEqual(Avalonia.Media.Colors.White, lightBackground.Color);
    }
}

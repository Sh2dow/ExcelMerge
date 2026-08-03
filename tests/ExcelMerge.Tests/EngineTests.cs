using ExcelMerge.Domain;
using ExcelMerge.Engine;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelMerge.Tests;

[TestClass]
public sealed class EngineTests
{
    [TestMethod]
    public async Task Two_way_diff_reports_cell_changes_and_compact_column_runs()
    {
        var baseSheet = TestData.Sheet(TestData.Row(
            0,
            (0, TestData.Text("same")),
            (1, TestData.Text("old-1")),
            (2, TestData.Text("old-2")),
            (4, TestData.Text("old-4"))));
        var sourceSheet = TestData.Sheet(TestData.Row(
            0,
            (0, TestData.Text("same")),
            (1, TestData.Text("new-1")),
            (2, TestData.Text("new-2")),
            (3, TestData.Text("added")),
            (4, TestData.Text("new-4"))));

        var result = await new TwoWayDiffEngine().CompareAsync(baseSheet, sourceSheet);
        var rowChange = result.SheetChange.RowChanges.ToArray().Single();
        var cellChanges = rowChange.CellChanges.ToArray();
        var runs = result.ViewRows.ToArray().Single().ChangedColumnRuns.ToArray();

        Assert.AreEqual(ChangeKind.Modified, result.SheetChange.Kind);
        Assert.AreEqual(SourceSide.Local, result.SheetChange.Source);
        Assert.AreEqual(ChangeKind.Modified, rowChange.Kind);
        Assert.AreEqual(4, cellChanges.Length);
        Assert.AreEqual(ChangeKind.Modified, cellChanges[0].Kind);
        Assert.AreEqual(ChangeKind.Modified, cellChanges[1].Kind);
        Assert.AreEqual(ChangeKind.Added, cellChanges[2].Kind);
        Assert.AreEqual(ChangeKind.Modified, cellChanges[3].Kind);
        CollectionAssert.AreEqual(
            new[]
            {
                new ColumnChangeRun(1, 2, ChangeKind.Modified),
                new ColumnChangeRun(3, 1, ChangeKind.Added),
                new ColumnChangeRun(4, 1, ChangeKind.Modified),
            },
            runs);
        Assert.AreEqual(1, result.Statistics.ModifiedRowCount);
        Assert.AreEqual(4L, result.Statistics.ChangedCellCount);
    }

    [TestMethod]
    public async Task Two_way_diff_aligns_equal_rows_after_physical_insertions()
    {
        var baseSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("A"))),
            TestData.Row(1, (0, TestData.Text("B"))));
        var sourceSheet = TestData.Sheet(
            TestData.Row(0, (0, TestData.Text("inserted"))),
            TestData.Row(1, (0, TestData.Text("A"))),
            TestData.Row(2, (0, TestData.Text("B"))));

        var result = await new TwoWayDiffEngine().CompareAsync(baseSheet, sourceSheet);
        var viewRows = result.ViewRows.ToArray();

        Assert.AreEqual(1, result.Statistics.AddedRowCount);
        Assert.AreEqual(0, result.Statistics.ModifiedRowCount);
        Assert.AreEqual(ChangeKind.Added, viewRows[0].Kind);
        Assert.IsNull(viewRows[0].BaseRowIndex);
        Assert.AreEqual(0, viewRows[0].SourceRowIndex);
        Assert.AreEqual(ChangeKind.Unchanged, viewRows[1].Kind);
        Assert.AreEqual(0, viewRows[1].BaseRowIndex);
        Assert.AreEqual(1, viewRows[1].SourceRowIndex);
    }

    [TestMethod]
    public async Task Two_way_diff_compares_typed_formula_data_not_display_text_by_default()
    {
        var baseSheet = TestData.Sheet(TestData.Row(
            0,
            (0, CellValue.FromNumber(1, "1")),
            (1, CellValue.FromFormula("A1+1", CellScalar.FromNumber(2), "2"))));
        var sourceSheet = TestData.Sheet(TestData.Row(
            0,
            (0, CellValue.FromNumber(1, "1.00")),
            (1, CellValue.FromFormula("A1+1", CellScalar.FromNumber(3), "3"))));

        var defaultResult = await new TwoWayDiffEngine().CompareAsync(baseSheet, sourceSheet);
        var relaxedResult = await new TwoWayDiffEngine().CompareAsync(
            baseSheet,
            sourceSheet,
            new WorksheetComparisonOptions
            {
                CellValues = new CellComparisonOptions
                {
                    CompareFormulaCachedValues = false,
                },
            });
        var displayResult = await new TwoWayDiffEngine().CompareAsync(
            baseSheet,
            sourceSheet,
            new WorksheetComparisonOptions
            {
                CellValues = new CellComparisonOptions
                {
                    CompareFormulaCachedValues = false,
                    CompareDisplayText = true,
                },
            });

        Assert.AreEqual(1L, defaultResult.Statistics.ChangedCellCount);
        Assert.AreEqual(ChangeKind.Unchanged, relaxedResult.SheetChange.Kind);
        Assert.AreEqual(2L, displayResult.Statistics.ChangedCellCount);
    }

    [TestMethod]
    public async Task Three_way_merge_automatically_uses_the_only_changed_side()
    {
        var baseSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("base"))));
        var localSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("local"))));
        var remoteSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("base"))));

        var result = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet);
        var decision = result.AutomaticDecisions.ToArray().Single();

        Assert.IsFalse(result.HasConflicts);
        Assert.AreEqual(MergeDecisionScope.Cell, decision.Scope);
        Assert.AreEqual(AutomaticMergeKind.UseLocal, decision.Kind);
        Assert.AreEqual(TestData.Text("local"), decision.ResultValue);
        Assert.AreEqual(MergeViewRowState.Automatic, result.ViewRows.ToArray().Single().State);
    }

    [TestMethod]
    public async Task Three_way_merge_creates_cell_conflict_and_resolution()
    {
        var baseSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("base"))));
        var localSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("local"))));
        var remoteSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("remote"))));
        var options = new ThreeWayMergeOptions { FirstConflictId = 100 };

        var result = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet,
            options);
        var conflict = result.Conflicts.ToArray().Single();
        var resolution = result.CellResolutions.ToArray().Single();

        Assert.AreEqual(100L, conflict.Id);
        Assert.AreEqual(ConflictKind.CellValue, conflict.Kind);
        Assert.AreEqual(new ConflictLocation("rId1", 0, 0, 0, 0), conflict.Location);
        Assert.AreEqual(TestData.Text("base"), conflict.BaseValue);
        Assert.AreEqual(TestData.Text("local"), conflict.LocalValue);
        Assert.AreEqual(TestData.Text("remote"), conflict.RemoteValue);
        Assert.AreEqual(ResolutionKind.Unresolved, resolution.Kind);
        Assert.AreEqual(101L, result.NextConflictId);
        Assert.AreEqual(MergeViewRowState.Conflict, result.ViewRows.ToArray().Single().State);
    }

    [TestMethod]
    public async Task Three_way_merge_classifies_row_delete_edit_as_structural_conflict()
    {
        var baseSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("base"))));
        var localSheet = TestData.Sheet();
        var remoteSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("remote"))));

        var result = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet);
        var conflict = result.Conflicts.ToArray().Single();

        Assert.AreEqual(ConflictKind.RowDeleteEdit, conflict.Kind);
        Assert.AreEqual(1, result.RowResolutions.Length);
        Assert.AreEqual(0, result.CellResolutions.Length);
        Assert.IsFalse(result.RowResolutions.Span[0].IsResolved);
    }

    [TestMethod]
    public async Task Three_way_merge_detects_incompatible_same_slot_insertions()
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

        var result = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet);

        Assert.IsTrue(result.HasConflicts);
        Assert.AreEqual(ConflictKind.RowAddAdd, result.Conflicts.Span[0].Kind);
        Assert.IsTrue(result.ViewRows.ToArray().Any(row => row.State == MergeViewRowState.Conflict));
    }

    [TestMethod]
    public async Task Merge_plan_factory_collects_results_and_rejects_duplicate_conflict_ids()
    {
        var baseSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("base"))));
        var localSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("local"))));
        var remoteSheet = TestData.Sheet(TestData.Row(0, (0, TestData.Text("remote"))));
        var result = await new ThreeWayMergeEngine().MergeAsync(
            baseSheet,
            localSheet,
            remoteSheet);
        var workbook = TestData.Workbook();

        var plan = MergePlanFactory.Create(workbook, workbook, workbook, [result]);

        Assert.AreEqual(2, plan.SheetChanges.Length);
        Assert.AreEqual(1, plan.Conflicts.Length);
        Assert.IsTrue(plan.HasUnresolvedResolutions);
        Assert.ThrowsExactly<ArgumentException>(() =>
            MergePlanFactory.Create(workbook, workbook, workbook, [result, result]));
    }

    [TestMethod]
    public async Task Merge_keeps_complete_row_topology_when_unchanged_view_rows_are_hidden()
    {
        var rows = new[]
        {
            TestData.Row(0, (0, TestData.Text("A"))),
            TestData.Row(3, (0, TestData.Text("B"))),
        };
        var result = await new ThreeWayMergeEngine().MergeAsync(
            TestData.Sheet(rows),
            TestData.Sheet(rows),
            TestData.Sheet(rows),
            new ThreeWayMergeOptions
            {
                SheetGroupId = "pair-1",
                Comparison = new WorksheetComparisonOptions
                {
                    IncludeUnchangedViewRows = false,
                },
            });

        Assert.AreEqual("pair-1", result.SheetGroupId);
        Assert.AreEqual(0, result.ViewRows.Length);
        CollectionAssert.AreEqual(
            new[]
            {
                new MergeRowMapping(0, 0, 0),
                new MergeRowMapping(3, 3, 3),
            },
            result.RowMappings.ToArray());
    }
}

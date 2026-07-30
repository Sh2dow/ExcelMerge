using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using NetDiff;

namespace ExcelMerge
{
    public class ExcelSheet
    {
        public SortedDictionary<int, ExcelRow> Rows { get; private set; }
        public List<ExcelMergedRegion> MergedRegions { get; private set; }

        public ExcelSheet()
        {
            Rows = new SortedDictionary<int, ExcelRow>();
            MergedRegions = new List<ExcelMergedRegion>();
        }

        public static ExcelSheet Create(ISheet srcSheet, ExcelSheetReadConfig config)
        {
            var rows = ExcelReader.Read(srcSheet);
            var sheet = CreateSheet(rows, config);
            for (var index = 0; index < srcSheet.NumMergedRegions; index++)
            {
                var region = srcSheet.GetMergedRegion(index);
                sheet.MergedRegions.Add(new ExcelMergedRegion(
                    region.FirstRow,
                    region.LastRow,
                    region.FirstColumn,
                    region.LastColumn));
            }

            return sheet;
        }

        public static ExcelSheet CreateFromCsv(string path, ExcelSheetReadConfig config)
        {
            var rows = CsvReader.Read(path);

            return CreateSheet(rows, config);
        }

        public static ExcelSheet CreateFromTsv(string path, ExcelSheetReadConfig config)
        {
            var rows = TsvReader.Read(path);

            return CreateSheet(rows, config);
        }

        private static ExcelSheet CreateSheet(IEnumerable<ExcelRow> rows, ExcelSheetReadConfig config)
        {
            var sheet = CreateSheet(rows);

            if (config.TrimFirstBlankRows)
                sheet.TrimFirstBlankRows();

            if (config.TrimFirstBlankColumns)
                sheet.TrimFirstBlankColumns();

            if (config.TrimLastBlankRows)
                sheet.TrimLastBlankRows();

            if (config.TrimLastBlankColumns)
                sheet.TrimLastBlankColumns();

            return sheet;
        }

        public void TrimFirstBlankRows()
        {
            var rows = new SortedDictionary<int, ExcelRow>();
            var index = 0;
            foreach (var row in Rows.SkipWhile(r => r.Value.IsBlank()))
            {
                rows.Add(index, new ExcelRow(index, row.Value.Cells));
                index++;
            }

            Rows = rows;
        }

        public void TrimFirstBlankColumns()
        {
            var columns = CreateColumns();
            var indices = columns.Select((v, i) => new { v, i }).TakeWhile(c => c.v.IsBlank()).Select(c => c.i);

            foreach (var i in indices)
                RemoveColumn(i);
        }

        public void TrimLastBlankRows()
        {
            var rows = new SortedDictionary<int, ExcelRow>();
            var index = 0;
            foreach (var row in Rows.Reverse().SkipWhile(r => r.Value.IsBlank()).Reverse())
            {
                rows.Add(index, new ExcelRow(index, row.Value.Cells));
                index++;
            }

            Rows = rows;
        }

        public void TrimLastBlankColumns()
        {
            var columns = CreateColumns();
            var indices = columns.Select((v, i) => new { v, i }).Reverse().TakeWhile(c => c.v.IsBlank()).Select(c => c.i);

            foreach (var i in indices)
                RemoveColumn(i);
        }

        public void RemoveColumn(int column)
        {
            foreach (var row in Rows)
            {
                if (row.Value.Cells.Count > column)
                    row.Value.Cells.RemoveAt(column);
            }
        }

        private static ExcelSheet CreateSheet(IEnumerable<ExcelRow> rows)
        {
            var sheet = new ExcelSheet();
            foreach (var row in rows)
            {
                sheet.Rows.Add(row.Index, row);
            }

            return sheet;
        }

        public static ExcelSheetDiff Diff(ExcelSheet src, ExcelSheet dst, ExcelSheetDiffConfig config)
        {
            var srcColumns = src.CreateColumns();
            var dstColumns = dst.CreateColumns();
            var columnStatusMap = CreateColumnStatusMap(srcColumns, dstColumns, config);
            var srcRows = AlignRows(src.Rows.Values, columnStatusMap, ExcelColumnStatus.Inserted);
            var dstRows = AlignRows(dst.Rows.Values, columnStatusMap, ExcelColumnStatus.Deleted);

            var option = new DiffOption<ExcelRow>();
            option.EqualityComparer =
                new RowComparer(columnStatusMap
                    .Select((status, index) => new { status, index })
                    .Where(item => item.status != ExcelColumnStatus.None)
                    .Select(item => item.index)
                    .ToHashSet());

            var r = DiffUtil.Diff(srcRows, dstRows, option);
            r = DiffUtil.Order(r, DiffOrderType.LazyDeleteFirst);
            var resultArray = DiffUtil.OptimizeCaseDeletedFirst(r).ToArray();
            if (resultArray.Length > 10000)
            {
                var included = new bool[resultArray.Length];
                for (var index = 0; index < 100; index++)
                    included[index] = true;
                for (var index = 0; index < resultArray.Length; index++)
                {
                    if (resultArray[index].Status == DiffStatus.Equal)
                        continue;
                    var start = Math.Max(0, index - 100);
                    var end = Math.Min(resultArray.Length, start + 200);
                    for (var retainedIndex = start; retainedIndex < end; retainedIndex++)
                        included[retainedIndex] = true;
                }
                resultArray = resultArray.Where((_, index) => included[index]).ToArray();
            }

            var sheetDiff = new ExcelSheetDiff();
            DiffCells(resultArray, sheetDiff, columnStatusMap);

            return sheetDiff;
        }

        private static List<ExcelRow> AlignRows(
            IEnumerable<ExcelRow> rows,
            IReadOnlyList<ExcelColumnStatus> columnStatuses,
            ExcelColumnStatus placeholderStatus)
        {
            if (!columnStatuses.Contains(placeholderStatus))
                return rows.ToList();

            var alignedRows = new List<ExcelRow>();
            foreach (var row in rows)
            {
                var cells = new List<ExcelCell>(columnStatuses.Count);
                var cellIndex = 0;
                var columnIndex = 0;
                while (cellIndex < row.Cells.Count)
                {
                    if (columnStatuses[columnIndex] == placeholderStatus)
                        cells.Add(new ExcelCell(string.Empty, 0, 0));
                    else
                        cells.Add(row.Cells[cellIndex++]);
                    columnIndex++;
                }
                alignedRows.Add(ExcelRow.FromCells(row.Index, cells));
            }
            return alignedRows;
        }

        private static IReadOnlyList<ExcelColumnStatus> CreateColumnStatusMap(
            IEnumerable<ExcelColumn> srcColumns, IEnumerable<ExcelColumn> dstColumns, ExcelSheetDiffConfig config)
        {
            var option = new DiffOption<ExcelColumn>();

            if (config.SrcHeaderIndex >= 0)
            {
                option.EqualityComparer = new HeaderComparer();
                foreach (var sc in srcColumns)
                    sc.HeaderIndex = config.SrcHeaderIndex;
            }

            if (config.DstHeaderIndex >= 0)
            {
                foreach (var dc in dstColumns)
                    dc.HeaderIndex = config.DstHeaderIndex;
            }

            var results = DiffUtil.Diff(srcColumns, dstColumns, option);
            results = DiffUtil.Order(results, DiffOrderType.LazyDeleteFirst);
            results = DiffUtil.OptimizeCaseDeletedFirst(results);
            var ret = new List<ExcelColumnStatus>();
            foreach (var result in results)
            {
                var status = ExcelColumnStatus.None;
                if (result.Status == DiffStatus.Deleted)
                    status = ExcelColumnStatus.Deleted;
                else if (result.Status == DiffStatus.Inserted)
                    status = ExcelColumnStatus.Inserted;

                ret.Add(status);
            }

            return ret;
        }

        private IReadOnlyList<ExcelColumn> CreateColumns()
        {
            var columns = new List<ExcelColumn>();
            foreach (var row in Rows)
            {
                var columnIndex = 0;
                foreach (var cell in row.Value.Cells)
                {
                    if (columnIndex == columns.Count)
                        columns.Add(new ExcelColumn());

                    columns[columnIndex].Cells.Add(cell);
                    columnIndex++;
                }
            }

            return columns;
        }

        private static void DiffCells(
            IEnumerable<DiffResult<ExcelRow>> results, ExcelSheetDiff sheetDiff, IReadOnlyList<ExcelColumnStatus> columnStatusMap)
        {
            foreach (var result in results)
            {
                switch (result.Status)
                {
                    case DiffStatus.Equal:
                        DiffCellsCaseEqual(result, sheetDiff, columnStatusMap);
                        break;
                    case DiffStatus.Modified:
                        DiffCellsCaseEqual(result, sheetDiff, columnStatusMap);
                        break;
                    case DiffStatus.Deleted:
                        DiffCellsCaseDeleted(result, sheetDiff);
                        break;
                    case DiffStatus.Inserted:
                        DiffCellsCaseInserted(result, sheetDiff);
                        break;
                }
            }
        }

        private static void DiffCellsCaseEqual(
            DiffResult<ExcelRow> result, ExcelSheetDiff sheetDiff, IReadOnlyList<ExcelColumnStatus> columnStatusMap)
        {
            var row = sheetDiff.CreateRow();

            for (var columnIndex = 0; columnIndex < columnStatusMap.Count; columnIndex++)
            {
                var srcCell = columnIndex < result.Obj1.Cells.Count ? result.Obj1.Cells[columnIndex] : null;
                var dstCell = columnIndex < result.Obj2.Cells.Count ? result.Obj2.Cells[columnIndex] : null;

                if (srcCell != null && dstCell != null)
                {
                    var status = srcCell.Value.Equals(dstCell.Value) ? ExcelCellStatus.None : ExcelCellStatus.Modified;
                    if (columnStatusMap[columnIndex] == ExcelColumnStatus.Deleted)
                        status = ExcelCellStatus.Removed;
                    else if (columnStatusMap[columnIndex] == ExcelColumnStatus.Inserted)
                        status = ExcelCellStatus.Added;

                    row.CreateCell(srcCell, dstCell, columnIndex, status);
                }
                else if (srcCell != null && dstCell == null)
                {
                    dstCell = new ExcelCell(string.Empty, srcCell.OriginalColumnIndex, srcCell.OriginalRowIndex);
                    row.CreateCell(srcCell, dstCell, columnIndex, ExcelCellStatus.Removed);
                }
                else if (srcCell == null && dstCell != null)
                {
                    srcCell = new ExcelCell(string.Empty, dstCell.OriginalColumnIndex, dstCell.OriginalRowIndex);
                    row.CreateCell(srcCell, dstCell, columnIndex, ExcelCellStatus.Added);
                }
                else
                {
                    srcCell = new ExcelCell(string.Empty, 0, 0);
                    dstCell = new ExcelCell(string.Empty, 0, 0);
                    row.CreateCell(srcCell, dstCell, columnIndex, ExcelCellStatus.None);
                }

            }
        }

        private static void DiffCellsCaseDeleted(
            DiffResult<ExcelRow> result, ExcelSheetDiff sheetDiff)
        {
            var row = sheetDiff.CreateRow();

            var columnIndex = 0;
            foreach (var cell1 in result.Obj1.Cells)
            {
                var cell2 = new ExcelCell(string.Empty, cell1.OriginalColumnIndex, cell1.OriginalRowIndex);
                row.CreateCell(cell1, cell2, columnIndex, ExcelCellStatus.Removed);

                columnIndex++;
            }
        }

        private static void DiffCellsCaseInserted(
            DiffResult<ExcelRow> result, ExcelSheetDiff sheetDiff)
        {
            var row = sheetDiff.CreateRow();

            var columnIndex = 0;
            foreach (var cell2 in result.Obj2.Cells)
            {
                var cell1 = new ExcelCell(string.Empty, cell2.OriginalColumnIndex, cell2.OriginalRowIndex);
                row.CreateCell(cell1, cell2, columnIndex, ExcelCellStatus.Added);

                columnIndex++;
            }
        }
    }
}

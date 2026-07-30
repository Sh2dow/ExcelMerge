using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelSheetDiff
    {
        public SortedDictionary<int, ExcelRowDiff> Rows { get; private set; }

        public ExcelSheetDiff()
        {
            Rows = new SortedDictionary<int, ExcelRowDiff>();
        }

        public ExcelRowDiff CreateRow()
        {
            var row = new ExcelRowDiff(Rows.Count);
            Rows.Add(row.Index, row);

            return row;
        }

        public ExcelSheetDiffSummary CreateSummary()
        {
            var addedRowCount = 0;
            var removedRowCount = 0;
            var modifiedRowCount = 0;
            var modifiedCellCount = 0;
            foreach (var row in Rows)
            {
                var isAdded = true;
                var isRemoved = true;
                var isModified = false;
                foreach (var cell in row.Value.Cells.Values)
                {
                    var status = cell.Status;
                    isAdded &= status == ExcelCellStatus.Added;
                    isRemoved &= status == ExcelCellStatus.Removed;
                    if (status != ExcelCellStatus.None)
                    {
                        isModified = true;
                        modifiedCellCount++;
                    }
                }

                if (isAdded)
                    addedRowCount++;
                else if (isRemoved)
                    removedRowCount++;

                if (isModified)
                    modifiedRowCount++;
            }

            return new ExcelSheetDiffSummary
            {
                AddedRowCount = addedRowCount,
                RemovedRowCount = removedRowCount,
                ModifiedRowCount = modifiedRowCount,
                ModifiedCellCount = modifiedCellCount,
            };
        }
    }
}

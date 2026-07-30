using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    internal class ExcelReader
    {
        internal static IEnumerable<ExcelRow> Read(ISheet sheet)
        {
            var actualRowIndex = 0;
            foreach (IRow row in sheet)
            {
                var cells = new List<ExcelCell>(System.Math.Max(0, (int)row.LastCellNum));
                for (int columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex);
                    var stringValue = ExcelUtility.GetCellStringValue(cell);

                    cells.Add(new ExcelCell(stringValue, columnIndex, row.RowNum));
                }

                yield return ExcelRow.FromCells(actualRowIndex++, cells);
            }
        }
    }
}

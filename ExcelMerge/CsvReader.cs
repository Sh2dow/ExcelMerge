using System.IO;
using System.Collections.Generic;
using System.Text;

namespace ExcelMerge
{
    internal class CsvReader
    {
        internal static IEnumerable<ExcelRow> Read(string path)
        {
            return Read(path, ',');
        }

        internal static IEnumerable<ExcelRow> Read(string path, char separator)
        {
            using (var sr = new StreamReader(path, Encoding.UTF8))
            {
                var rowIndex = 0;
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    var columnIndex = 0;
                    var cells = new List<ExcelCell>();
                    var start = 0;
                    while (true)
                    {
                        var end = line.IndexOf(separator, start);
                        var value = end < 0 ? line.Substring(start) : line.Substring(start, end - start);
                        cells.Add(new ExcelCell(value, columnIndex++, rowIndex));
                        if (end < 0)
                            break;
                        start = end + 1;
                    }

                    yield return ExcelRow.FromCells(rowIndex++, cells);
                }
            }
        }
    }
}

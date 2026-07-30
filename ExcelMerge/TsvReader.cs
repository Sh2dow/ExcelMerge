using System.Collections.Generic;

namespace ExcelMerge
{
    public class TsvReader
    {
        internal static IEnumerable<ExcelRow> Read(string path)
        {
            return CsvReader.Read(path, '\t');
        }
    }
}

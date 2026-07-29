using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;

namespace ExcelMerge
{
    public class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; }

        public ExcelWorkbook()
        {
            Sheets = new Dictionary<string, ExcelSheet>();
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config)
        {
            var extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension == ".csv")
                return CreateFromCsv(path, config);

            if (extension == ".tsv")
                return CreateFromTsv(path, config);

            using (var srcWb = WorkbookFactory.Create(path))
            {
                var wb = new ExcelWorkbook();
                for (int i = 0; i < srcWb.NumberOfSheets; i++)
                {
                    var srcSheet = srcWb.GetSheetAt(i);
                    wb.Sheets.Add(srcSheet.SheetName, ExcelSheet.Create(srcSheet, config));
                }

                return wb;
            }
        }

        public static IEnumerable<string> GetSheetNames(string path)
        {
            var extension = Path.GetExtension(path).ToLowerInvariant();
            if (extension == ".csv" || extension == ".tsv")
                return new[] { Path.GetFileName(path) };

            using (var workbook = WorkbookFactory.Create(path))
            {
                var names = new List<string>();
                for (int i = 0; i < workbook.NumberOfSheets; i++)
                    names.Add(workbook.GetSheetAt(i).SheetName);

                return names;
            }
        }

        private static ExcelWorkbook CreateFromCsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromCsv(path, config));

            return wb;
        }

        private static ExcelWorkbook CreateFromTsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromTsv(path, config));

            return wb;
        }
    }
}

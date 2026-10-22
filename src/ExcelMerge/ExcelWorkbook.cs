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
            if (Path.GetExtension(path).ToLowerInvariant() == ".csv")
                return CreateFromCsv(path, config);

            if (Path.GetExtension(path).ToLowerInvariant() == ".tsv")
                return CreateFromTsv(path, config);

            var wb = new ExcelWorkbook();
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var srcWb = WorkbookFactory.Create(stream);
                try
                {
                    for (int i = 0; i < srcWb.NumberOfSheets; i++)
                    {
                        var srcSheet = srcWb.GetSheetAt(i);
                        wb.Sheets.Add(srcSheet.SheetName, ExcelSheet.Create(srcSheet, config));
                    }
                }
                finally { srcWb.Close(); }
            }

            return wb;
        }

        public static IEnumerable<string> GetSheetNames(string path)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".csv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else if (Path.GetExtension(path).ToLowerInvariant() == ".tsv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var wb = WorkbookFactory.Create(stream);
                    try
                    {
                        for (int i = 0; i < wb.NumberOfSheets; i++)
                            yield return wb.GetSheetAt(i).SheetName;
                    }
                    finally { wb.Close(); }
                }
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

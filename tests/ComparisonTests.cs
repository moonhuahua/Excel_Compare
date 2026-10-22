using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using ExcelMerge;

class ComparisonTests
{
    static int checks;
    static void Check(bool value, string name)
    {
        if (!value) throw new Exception("FAIL: " + name);
        Console.WriteLine("PASS: " + name);
        checks++;
    }
    static string Hash(string p)
    {
        using (var stream = File.OpenRead(p))
        using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(stream));
    }
    static void Create(string path, bool updated, bool blank)
    {
        IWorkbook wb = Path.GetExtension(path) == ".xls" ? (IWorkbook)new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = wb.CreateSheet("材料清单");
        if (!blank)
        {
            string[][] rows = {
                new[] {"编号", "名称", "数量"}, new[] {"A", "水泵", updated ? "10" : "1"},
                new[] {"B", "阀门", "2"}, new[] {"C", "膜组件", "3"}
            };
            for (int i = 0; i < rows.Length; i++)
            {
                var row = sheet.CreateRow(i);
                for (int j = 0; j < rows[i].Length; j++) row.CreateCell(j).SetCellValue(rows[i][j]);
            }
            if (updated)
            {
                var row = sheet.CreateRow(4);
                row.CreateCell(0).SetCellValue("D"); row.CreateCell(1).SetCellValue("新增设备"); row.CreateCell(2).SetCellValue("4");
            }
        }
        wb.CreateSheet("备注").CreateRow(0).CreateCell(0).SetCellValue("第二张工作表");
        using (var file = File.Create(path)) wb.Write(file);
        wb.Close();
    }
    static ExcelSheetDiffSummary Diff(string left, string right, string name)
    {
        var l = ExcelWorkbook.Create(left, new ExcelSheetReadConfig());
        var r = ExcelWorkbook.Create(right, new ExcelSheetReadConfig());
        return ExcelSheet.Diff(l.Sheets[name], r.Sheets[name], new ExcelSheetDiffConfig()).CreateSummary();
    }
    static void Main(string[] args)
    {
        try { Run(args); }
        catch (Exception ex)
        {
            Console.WriteLine(ex.GetType().FullName);
            Console.WriteLine(ex.Message);
            Console.WriteLine(ex.StackTrace);
            Environment.ExitCode = 1;
        }
    }
    static void Run(string[] args)
    {
        var dir = args[0]; Directory.CreateDirectory(dir);
        foreach (string ext in new[] { ".xlsx", ".xls" })
        {
            string l = Path.Combine(dir, "左侧 原文件" + ext), r = Path.Combine(dir, "右侧 新文件" + ext);
            Create(l, false, false); Create(r, true, false);
            string lh = Hash(l), rh = Hash(r);
            Check(ExcelWorkbook.GetSheetNames(l).SequenceEqual(new[] {"材料清单", "备注"}), ext + " sheet names and order");
            var summary = Diff(l, r, "材料清单");
            Check(summary.ModifiedCellCount == 4 && summary.AddedRowCount == 1 && summary.RemovedRowCount == 0,
                ext + " expected four changed cells (one edit plus three added cells) and one added row");
            var reverse = Diff(r, l, "材料清单");
            Check(reverse.ModifiedCellCount == 4 && reverse.RemovedRowCount == 1 && reverse.AddedRowCount == 0, ext + " swapped directions");
            Check(!Diff(l, l, "材料清单").HasDiff, ext + " identical files");
            Check(!Diff(l, r, "备注").HasDiff, ext + " second sheet independently compared");
            Check(lh == Hash(l) && rh == Hash(r), ext + " sources unchanged");
            using (var file = new FileStream(l, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                Check(file.Length > 0, ext + " read handles released");
        }
        string empty = Path.Combine(dir, "空表.xlsx"); Create(empty, false, true);
        Check(!Diff(empty, empty, "材料清单").HasDiff, "empty worksheets");
        string csv = Path.Combine(dir, "英文扩展名.CSV"); File.WriteAllText(csv, "id,value\r\nA,1\r\n");
        Check(!Diff(csv, csv, Path.GetFileName(csv)).HasDiff, "uppercase CSV");
        string bad = Path.Combine(dir, "损坏文件.xlsx"); File.WriteAllText(bad, "not an xlsx");
        bool rejected = false;
        try { ExcelWorkbook.GetSheetNames(bad).ToList(); } catch { rejected = true; }
        Check(rejected, "corrupt file rejected");
        Console.WriteLine("All " + checks + " checks passed.");
    }
}

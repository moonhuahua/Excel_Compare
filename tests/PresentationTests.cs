using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Data;
using System.Windows.Documents;
using ExcelMerge;
using ExcelMerge.GUI;
using ExcelMerge.GUI.Models;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.ViewModels;
using ExcelMerge.GUI.Views;
using FastWpfGrid;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using ExcelMerge.GUI.Commands;
using HorizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment;
using VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment;

class PresentationTests
{
    class TestApp : App { protected override void OnStartup(StartupEventArgs e) { } }
    class LongTextModel : FastGridModelBase
    {
        public override int RowCount { get { return 2; } }
        public override int ColumnCount { get { return 1; } }
        public override string GetCellText(int row, int column) { return "长文本换行启动回归测试，未设置最大行高时也应能正常测量并显示。"; }
    }
    static int checks;
    static readonly BindingFlags Private = BindingFlags.Instance | BindingFlags.NonPublic;
    static void Check(bool ok, string name) { if (!ok) throw new Exception("FAIL: " + name); Console.WriteLine("PASS: " + name); checks++; }
    static object Call(object o, string name, params object[] args)
    {
        var method = o.GetType().GetMethods(Private).Single(m => m.Name == name && m.GetParameters().Length == args.Length);
        return method.Invoke(o, args);
    }
    static T Control<T>(DiffView view, string name) where T : class { return view.FindName(name) as T; }
    static string Hash(string path) { using (var f = File.OpenRead(path)) using (var h = SHA256.Create()) return BitConverter.ToString(h.ComputeHash(f)); }
    static void Pump() { Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.ApplicationIdle, new Action(delegate { })); }
    static int ColumnWidthTotal(FastGridControl grid)
    {
        return Enumerable.Range(0, grid.Model.ColumnCount).Sum(c => grid.GetPresentationColumnWidth(c));
    }
    static bool ColumnsFill(FastGridControl grid)
    {
        return Math.Abs(ColumnWidthTotal(grid) - grid.GridScrollAreaWidth) <= 1;
    }
    static double Luminance(Color color)
    {
        Func<byte, double> linear = c => c / 255.0 <= 0.04045 ? c / 255.0 / 12.92 : Math.Pow((c / 255.0 + 0.055) / 1.055, 2.4);
        return 0.2126 * linear(color.R) + 0.7152 * linear(color.G) + 0.0722 * linear(color.B);
    }
    static double ButtonContrast(Button button)
    {
        button.ApplyTemplate();
        var content = (ContentPresenter)button.Template.FindName("ButtonContent", button);
        var border = (Border)button.Template.FindName("ButtonBorder", button);
        double a = Luminance(((SolidColorBrush)content.GetValue(TextElement.ForegroundProperty)).Color);
        double b = Luminance(((SolidColorBrush)border.Background).Color);
        return (Math.Max(a, b) + 0.05) / (Math.Min(a, b) + 0.05);
    }
    static void ResetViewListeners()
    {
        // Each production process owns one comparison view; isolate sequential test windows accordingly.
        foreach (string name in new[] { "DataGridEventDispatcher", "LocationGridEventDispatcher", "ViewportEventDispatcher", "ValueTextBoxEventDispatcher" })
        {
            var type = typeof(DiffView).Assembly.GetType("ExcelMerge.GUI.Views." + name);
            var instance = type.GetProperty("Instance", BindingFlags.Public | BindingFlags.Static).GetValue(null, null);
            ((System.Collections.IList)type.GetField("Listeners").GetValue(instance)).Clear();
        }
    }
    static void Create(string path, bool updated, bool styleOnly)
    {
        IWorkbook wb = path.EndsWith(".xls") ? (IWorkbook)new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = wb.CreateSheet("格式与换行");
        var font = wb.CreateFont(); font.FontName = "微软雅黑"; font.FontHeightInPoints = 15;
        font.IsBold = true; font.IsItalic = true; font.Underline = FontUnderlineType.Single; font.Color = IndexedColors.DarkRed.Index;
        var style = wb.CreateCellStyle(); style.SetFont(font); style.FillPattern = FillPattern.SolidForeground;
        style.FillForegroundColor = styleOnly ? IndexedColors.Yellow.Index : IndexedColors.LightYellow.Index;
        style.Alignment = HorizontalAlignment.Center; style.VerticalAlignment = VerticalAlignment.Top;
        var money = wb.CreateCellStyle(); money.DataFormat = wb.CreateDataFormat().GetFormat("#,##0.00"); money.Alignment = HorizontalAlignment.Right;
        var percent = wb.CreateCellStyle(); percent.DataFormat = wb.CreateDataFormat().GetFormat("0.0%");
        var date = wb.CreateCellStyle(); date.DataFormat = wb.CreateDataFormat().GetFormat("yyyy-mm-dd");
        string[][] rows = {
            new[] { "编号", "设备说明", "单价 / 参数" },
            new[] { "A01", "不锈钢循环泵，含变频控制、压力监测与自动保护功能；这是用于验证缩窄列宽后自动换行的长文本。", "" },
            new[] { "A02", updated ? "FL30S-50 就地显示 T≤60℃ P≤0.3MPA 银好" : "FL30S-50 就地显示 T≤60℃ P≤0.3MPA 银锌", "" },
            new[] { "A03", "交付日期", "" }
        };
        for (int r = 0; r < rows.Length; r++) for (int c = 0; c < 3; c++)
        {
            var row = sheet.GetRow(r) ?? sheet.CreateRow(r); row.CreateCell(c).SetCellValue(rows[r][c]);
        }
        sheet.GetRow(1).GetCell(1).CellStyle = style;
        sheet.GetRow(1).GetCell(2).SetCellValue(updated ? 1434.5 : 1234.5); sheet.GetRow(1).GetCell(2).CellStyle = money;
        sheet.GetRow(2).GetCell(2).SetCellValue(0.125); sheet.GetRow(2).GetCell(2).CellStyle = percent;
        sheet.GetRow(3).GetCell(2).SetCellValue(new DateTime(2026, 10, 20)); sheet.GetRow(3).GetCell(2).CellStyle = date;
        if (updated) { var r = sheet.CreateRow(4); r.CreateCell(0).SetCellValue("A04"); r.CreateCell(1).SetCellValue("新增过滤器"); r.CreateCell(2).SetCellValue(80); }
        wb.CreateSheet("备注").CreateRow(0).CreateCell(0).SetCellValue("第二工作表仍可选择");
        using (var f = File.Create(path)) wb.Write(f); wb.Close();
    }
    [STAThread]
    static void Main(string[] args)
    {
        try { Run(args[0]); }
        catch (Exception e) { Console.WriteLine(e.ToString()); Environment.ExitCode = 1; }
    }
    static void Run(string root)
    {
        string configBefore = File.Exists(ApplicationSetting.Location) ? Hash(ApplicationSetting.Location) : null;
        Directory.CreateDirectory(root);
        foreach (string extension in new[] { ".xlsx", ".xls" })
        {
            var left = Path.Combine(root, "格式左侧" + extension); var right = Path.Combine(root, "格式右侧" + extension); var styleOnly = Path.Combine(root, "仅样式变化" + extension);
            Create(left, false, false); Create(right, true, false); Create(styleOnly, false, true);
            string lh = Hash(left), rh = Hash(right);
            var l = ExcelWorkbook.Create(left, new ExcelSheetReadConfig()).Sheets["格式与换行"];
            var r = ExcelWorkbook.Create(right, new ExcelSheetReadConfig()).Sheets["格式与换行"];
            var s = ExcelWorkbook.Create(styleOnly, new ExcelSheetReadConfig()).Sheets["格式与换行"];
            var cell = l.Rows[1].Cells[1];
            Check(cell.Appearance.FontName == "微软雅黑" && Math.Abs(cell.Appearance.FontSize - 20) < 0.1 && cell.Appearance.Bold && cell.Appearance.Italic && cell.Appearance.Underline, extension + " font name/size/bold/italic/underline");
            Check(cell.Appearance.Foreground == "#800000" && cell.Appearance.Background == "#FFFF99" && cell.Appearance.HorizontalAlignment == "Center", extension + " font/fill colors and alignment");
            Check(l.Rows[1].Cells[2].DisplayValue == "1,234.50" && l.Rows[2].Cells[2].DisplayValue == "12.5%" && l.Rows[3].Cells[2].DisplayValue == "2026-10-20", extension + " numeric/percentage/date display formats");
            Check(!ExcelSheet.Diff(l, s, new ExcelSheetDiffConfig()).CreateSummary().HasDiff, extension + " format-only changes keep content comparison unchanged");
            Check(ExcelSheet.Diff(l, r, new ExcelSheetDiffConfig()).CreateSummary().ModifiedCellCount == 5, extension + " two edits plus three added cells");
            Check(lh == Hash(left) && rh == Hash(right), extension + " sources unchanged");
        }
        var app = new TestApp();
        app.ShutdownMode = ShutdownMode.OnExplicitShutdown;
        app.Resources["InverseBooleanConverter"] = new ExcelMerge.GUI.ValueConverters.InverseBooleanConverter();
        var setting = new ApplicationSetting(); setting.NotifyEqual = false; setting.Ensure(); setting.EnsureCulture();
        typeof(App).GetProperty("Setting").SetValue(app, setting, null);
        typeof(App).GetProperty("CommandLineOption").SetValue(app, new CommandLineOption { KeepFileHistory = true }, null);
        app.UpdateResourceCulture();
        var emptyGrid = new FastGridControl();
        var gridWindow = new Window { Content = emptyGrid, Width = 600, Height = 400, Left = -20000, Top = -20000, ShowInTaskbar = false };
        gridWindow.Show(); Pump();
        emptyGrid.AllowFlexibleRows = true;
        emptyGrid.RefreshPresentation(); Pump();
        Check(emptyGrid.Model == null && emptyGrid.RealRowCount == 0, "empty grid permits flexible rows before a height limit is initialized");
        emptyGrid.WrapText = true;
        emptyGrid.Model = new LongTextModel(); emptyGrid.SetPresentationColumnWidth(0, 80); emptyGrid.RefreshPresentation(); Pump();
        Check(emptyGrid.GetModelRowHeight(0) > 40, "populated grid wraps without requiring an optional maximum row height");
        emptyGrid.Model = null; emptyGrid.AllowFlexibleRows = false; emptyGrid.AllowFlexibleRows = true; emptyGrid.RefreshPresentation(); Pump();
        Check(emptyGrid.RealRowCount == 0, "clearing a comparison while flexible rows are enabled remains safe");
        gridWindow.Close();
        for (int option = 0; option < 4; option++)
        {
            ResetViewListeners();
            setting.WrapText = (option & 1) != 0; setting.PreserveFormatting = (option & 2) != 0;
            var blankView = new DiffView { DataContext = new DiffViewModel() };
            var blankWindow = new Window { Content = blankView, Width = 1200, Height = 700, Left = -20000, Top = -20000, ShowInTaskbar = false };
            blankWindow.Show(); Pump();
            var blankLeft = Control<FastGridControl>(blankView, "SrcDataGrid");
            var blankRight = Control<FastGridControl>(blankView, "DstDataGrid");
            Check(blankLeft.Model == null && blankRight.Model == null, "no-file startup with saved wrap/format options " + option);
            setting.WrapText = !setting.WrapText; setting.PreserveFormatting = !setting.PreserveFormatting;
            Call(blankView, "ApplyGridPresentation"); blankWindow.Height = 450; Pump();
            Check(blankLeft.RealRowCount == 0 && blankRight.RealRowCount == 0, "blank view toggles display options and resizes safely " + option);
            blankWindow.Close();
        }
        setting.WrapText = false; setting.PreserveFormatting = false;
        ResetViewListeners();
        var vm = new DiffViewModel(); vm.SrcPath = Path.Combine(root, "格式左侧.xlsx"); vm.DstPath = Path.Combine(root, "格式右侧.xlsx");
        var view = new DiffView(); view.DataContext = vm;
        var window = new MainWindow { DataContext = new MainWindowViewModel(view), WindowState = WindowState.Normal, Width = 1600, Height = 850, Left = -20000, Top = -20000, ShowInTaskbar = false };
        window.Show(); Pump();
        var menu = (Menu)window.FindName("ApplicationMenu"); var menuIcon = (MenuItem)window.FindName("ApplicationMenuIcon");
        Check(menu.Items.Count == 1 && menuIcon.Items.Cast<MenuItem>().Select(i => i.Header.ToString()).SequenceEqual(new[] { "文件", "设置", "工具", "视图" }), "one menu icon retains all four original menu groups in order");
        Check(((MenuItem)((MenuItem)menuIcon.Items[0]).Items[0]).Command != null && ((MenuItem)((MenuItem)menuIcon.Items[1]).Items[0]).Command != null, "nested file and settings menu commands remain bound");
        var comparePosition = Control<Button>(view, "DiffButton").TranslatePoint(new Point(0, 0), window);
        var menuPosition = menu.TranslatePoint(new Point(0, 0), window);
        Check(Math.Abs(comparePosition.Y - menuPosition.Y) <= 6 && menuPosition.X + menu.ActualWidth <= comparePosition.X && ((Grid)window.FindName("MainGrid")).RowDefinitions[0].ActualHeight == 0, "menu icon is immediately left of compare on one toolbar with no separate menu row");
        Check(Control<Border>(view, "SrcFileBoxBorder").ActualHeight <= 38 && Control<Border>(view, "DstFileBoxBorder").ActualHeight <= 38 && Control<TextBox>(view, "SrcPathTextBox").FontSize == 13, "file selection row is compact without reducing the text size");
        var leftGrid = Control<FastGridControl>(view, "SrcDataGrid"); var rightGrid = Control<FastGridControl>(view, "DstDataGrid");
        Check(leftGrid.Model != null && rightGrid.Model != null && vm.ModifiedCellCount == 5, "WPF view loads two workbooks and compares automatically");
        Check(ColumnsFill(leftGrid) && ColumnsFill(rightGrid), "both tables fill their available width after initial comparison");
        Check(((ScrollBar)leftGrid.FindName("hscroll")).Maximum == 0, "fitted columns do not allow horizontal scrolling into empty space");
        Check(leftGrid.HeaderWidth >= 20 && leftGrid.HeaderWidth <= 40, "row-number header occupies about three digits");
        leftGrid.SetPresentationColumnWidth(0, 4); leftGrid.RefreshPresentation();
        Check(leftGrid.GetPresentationColumnWidth(0) == 4, "column may shrink to four pixels");
        leftGrid.SetPresentationColumnWidth(0, 65); leftGrid.SetPresentationColumnWidth(1, 135); rightGrid.SetPresentationColumnWidth(1, 135);
        setting.WrapText = false; Call(view, "ApplyGridPresentation"); var shortHeight = leftGrid.GetModelRowHeight(1);
        setting.WrapText = true; Call(view, "ApplyGridPresentation"); Pump();
        Check(leftGrid.GetModelRowHeight(1) > shortHeight * 2 && leftGrid.GetModelRowHeight(1) == rightGrid.GetModelRowHeight(1), "long Chinese text wraps and both row heights remain aligned");
        setting.PreserveFormatting = true; Call(view, "ApplyGridPresentation");
        var model = (DiffGridModel)leftGrid.Model;
        var formatted = model.GetCell(leftGrid, 1, 1) as IFormattedGridCell;
        Check(formatted != null && formatted.Presentation.Bold && formatted.Presentation.FontSize == 20 && model.GetCellText(1, 2) == "1,234.50", "preserve-format renders source font and formatted numeric text");
        setting.PreserveFormatting = false; Call(view, "ApplyGridPresentation");
        Check(model.GetCellText(1, 2) != "1,234.50" && vm.ModifiedCellCount == 5, "format toggle restores raw display without changing results");
        var search = Control<ToggleButton>(view, "SearchOptionsButton"); var expander = Control<Expander>(view, "ToolExpander");
        search.IsChecked = true; Pump(); Check(expander.IsExpanded, "top-right search button opens original search controls");
        search.IsChecked = false; Pump(); Check(!expander.IsExpanded, "search button collapses search controls");
        var list = Control<ListBox>(view, "ChangeList");
        Check(list.Items.Count == 5 && list.Items.Cast<ChangeEntry>().Any(e => e.Address == "C2"), "change panel has five cards and original Excel addresses");
        Control<ToggleButton>(view, "FilterAdded").IsChecked = false; Call(view, "RefreshChangeList");
        Check(list.Items.Count == 2 && list.Items.Cast<ChangeEntry>().All(e => e.Status == ExcelCellStatus.Modified), "panel filters added cells");
        Call(view, "MoveChange", 1); Pump(); var first = (ChangeEntry)list.SelectedItem;
        var selected = model.GetRealCellAddress(leftGrid.CurrentCell);
        Check(selected.Row == first.Row && selected.Column == first.Column && leftGrid.CurrentCell.Equals(rightGrid.CurrentCell), "card navigation selects corresponding cell on both sides");
        Call(view, "MoveChange", 1); Call(view, "MoveChange", 1);
        Check(list.SelectedIndex == 0, "next-change navigation cycles through visible cards");
        Control<ToggleButton>(view, "FilterAdded").IsChecked = true; Call(view, "RefreshChangeList");
        int expandedWidth = ColumnWidthTotal(leftGrid), expandedHeight = leftGrid.GetModelRowHeight(1);
        Control<ToggleButton>(view, "ChangePanelButton").IsChecked = false; Call(view, "SetChangePanelVisibility"); Pump();
        Check(Control<Border>(view, "ChangePanel").Visibility == Visibility.Collapsed && ((ColumnDefinition)view.FindName("ChangePanelWidth")).Width.Value == 0, "collapse releases panel width");
        Check(ColumnWidthTotal(leftGrid) > expandedWidth + 100 && ColumnsFill(leftGrid) && ColumnsFill(rightGrid), "closing the change panel expands columns to fill both tables");
        Check(leftGrid.GetModelRowHeight(1) <= expandedHeight && leftGrid.GetModelRowHeight(1) == rightGrid.GetModelRowHeight(1), "panel resize recalculates wrapping with matching row heights");
        Control<ToggleButton>(view, "ChangePanelButton").IsChecked = true; Call(view, "SetChangePanelVisibility"); Pump();
        Check(Math.Abs(ColumnWidthTotal(leftGrid) - expandedWidth) <= 1 && ColumnsFill(rightGrid), "reopening the panel shrinks columns to the available area");
        window.Width += 220; Pump();
        Check(ColumnsFill(leftGrid) && ColumnsFill(rightGrid), "columns fit after widening the main window");
        window.Width -= 350; Pump();
        Check(ColumnsFill(leftGrid) && ColumnsFill(rightGrid), "columns fit after narrowing the main window");
        leftGrid.ResizeColumn(4, 0); Pump();
        Check(leftGrid.GetPresentationColumnWidth(0) <= 12 && ColumnsFill(leftGrid) && ColumnsFill(rightGrid), "manual narrow-column resizing remains supported while other columns fill the space");
        setting.DarkMode = true; Call(view, "ApplyTheme"); Call(view, "ApplyGridPresentation"); Pump();
        Check(leftGrid.CanvasColor == Color.FromRgb(17, 22, 29) && ((SolidColorBrush)app.Resources["TextBrush"]).Color == Color.FromRgb(227, 233, 242), "dark mode applies to grid and window text");
        Check((int)Call(window, "ApplyTitleBarTheme", true) == 0, "Windows accepts the dark caption background and light title text colors");
        Check(window.WindowStyle == WindowStyle.SingleBorderWindow && window.ResizeMode == ResizeMode.CanResize, "standard title bar and native window controls are retained");
        var compareButton = Control<Button>(view, "DiffButton");
        Check(compareButton.IsEnabled && ButtonContrast(compareButton) >= 4.5, "enabled compare button has sufficient text contrast in dark mode");
        var compareBinding = BindingOperations.GetBinding(compareButton, UIElement.IsEnabledProperty);
        compareButton.IsEnabled = false; Pump();
        Check(ButtonContrast(compareButton) >= 4.5, "disabled compare button remains readable in dark mode");
        setting.DarkMode = false; Call(view, "ApplyTheme"); Call(view, "ApplyGridPresentation");
        Pump(); Check(ButtonContrast(compareButton) >= 4.5, "disabled compare button remains readable after switching to light mode");
        BindingOperations.SetBinding(compareButton, UIElement.IsEnabledProperty, compareBinding);
        Check((int)Call(window, "ApplyTitleBarTheme", false) == 0, "Windows accepts restoring the system title bar color in light mode");
        Check(leftGrid.CanvasColor == Colors.White && vm.ModifiedCellCount == 5, "light mode restored and comparison stable");
        window.Close();
        Check(configBefore == (File.Exists(ApplicationSetting.Location) ? Hash(ApplicationSetting.Location) : null), "automated tests leave the user's configuration unchanged");
        Console.WriteLine("All " + checks + " presentation checks passed.");
    }
}

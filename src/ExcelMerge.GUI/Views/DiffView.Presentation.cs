using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using FastWpfGrid;
using ExcelMerge.GUI.Models;
using NetDiff;

namespace ExcelMerge.GUI.Views
{
    public sealed class ChangeEntry
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int Number { get; set; }
        public ExcelCellStatus Status { get; set; }
        public string Address { get; set; }
        public string Left { get; set; }
        public string Right { get; set; }
        public string Kind { get { return Status == ExcelCellStatus.Added ? "新增" : Status == ExcelCellStatus.Removed ? "删除" : "已修改"; } }
        public string Tooltip { get { return "左侧：" + Left + "\n右侧：" + Right; } }
    }

    public partial class DiffView
    {
        private bool presentationReady;
        private bool updatingChangeSelection;
        private bool columnFitQueued;
        private readonly List<ChangeEntry> allChanges = new List<ChangeEntry>();

        private void InitializePresentation()
        {
            WrapTextButton.IsChecked = App.Instance.Setting.WrapText;
            PreserveFormatButton.IsChecked = App.Instance.Setting.PreserveFormatting;
            DarkModeButton.IsChecked = App.Instance.Setting.DarkMode;
            ChangePanelButton.IsChecked = App.Instance.Setting.ShowChangePanel;
            presentationReady = true;
            SrcDataGrid.ColumnWidthChanged += (sender, args) => QueueFitColumns();
            DstDataGrid.ColumnWidthChanged += (sender, args) => QueueFitColumns();
            ApplyTheme();
            SetChangePanelVisibility();
        }

        private void PresentationOption_Click(object sender, RoutedEventArgs e)
        {
            if (!presentationReady) return;
            App.Instance.Setting.WrapText = WrapTextButton.IsChecked == true;
            App.Instance.Setting.PreserveFormatting = PreserveFormatButton.IsChecked == true;
            App.Instance.Setting.DarkMode = DarkModeButton.IsChecked == true;
            App.Instance.Setting.ShowChangePanel = ChangePanelButton.IsChecked == true;
            App.Instance.Setting.Save();
            ApplyTheme();
            SetChangePanelVisibility();
            ApplyGridPresentation();
            RefreshChangeList();
        }

        private void HideChangePanel_Click(object sender, RoutedEventArgs e)
        {
            ChangePanelButton.IsChecked = false;
            PresentationOption_Click(sender, e);
        }

        private void SetChangePanelVisibility()
        {
            bool visible = ChangePanelButton.IsChecked == true;
            ChangePanelWidth.Width = new GridLength(visible ? 280 : 0);
            ChangePanel.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
            QueueFitColumns();
        }

        private static SolidColorBrush Brush(string value) { return new SolidColorBrush((Color)ColorConverter.ConvertFromString(value)); }

        private void ApplyTheme()
        {
            bool dark = App.Instance.Setting.DarkMode;
            var colors = new Dictionary<string, string>
            {
                ["SurfaceBrush"] = dark ? "#191E26" : "#FFFFFF",
                ["CanvasBrush"] = dark ? "#11161D" : "#F5F7FA",
                ["TextBrush"] = dark ? "#E3E9F2" : "#263445",
                ["MutedBrush"] = dark ? "#A6B2C3" : "#64748B",
                ["BorderBrush"] = dark ? "#3B4656" : "#CBD5E1",
                ["HoverBrush"] = dark ? "#314151" : "#E6F2F0",
                ["AccentBrush"] = dark ? "#176F65" : "#087F72",
                ["LeftBrush"] = dark ? "#F1A7B4" : "#99394A",
                ["RightBrush"] = dark ? "#87DABB" : "#087F48",
                ["ModifiedBrush"] = dark ? "#F5D07B" : "#946200",
                ["DisabledButtonBrush"] = dark ? "#34414F" : "#DCE3EA",
                ["DisabledTextBrush"] = dark ? "#D3DBE6" : "#526071"
            };
            foreach (var item in colors) Application.Current.Resources[item.Key] = Brush(item.Value);
            Application.Current.Resources[SystemColors.WindowBrushKey] = Brush(colors["SurfaceBrush"]);
            Application.Current.Resources[SystemColors.WindowTextBrushKey] = Brush(colors["TextBrush"]);
            Application.Current.Resources[SystemColors.ControlBrushKey] = Brush(colors["SurfaceBrush"]);
            Application.Current.Resources[SystemColors.ControlTextBrushKey] = Brush(colors["TextBrush"]);
            Application.Current.Resources[SystemColors.HighlightBrushKey] = Brush(colors["AccentBrush"]);
            var window = Window.GetWindow(this);
            if (window != null)
            {
                window.Background = Brush(colors["CanvasBrush"]);
                window.Foreground = Brush(colors["TextBrush"]);
                (window as MainWindow)?.ApplyTitleBarTheme(dark);
            }
        }

        private void ApplyGridPresentation()
        {
            bool dark = App.Instance.Setting.DarkMode;
            foreach (var grid in new[] { SrcDataGrid, DstDataGrid })
            {
                grid.WrapText = App.Instance.Setting.WrapText;
                grid.SetMaxRowSize(2000);
                grid.AllowFlexibleRows = App.Instance.Setting.WrapText || App.Instance.Setting.PreserveFormatting || App.Instance.Setting.FitRowHeight;
                grid.SetMinColumnSize(1);
                grid.CanvasColor = dark ? Color.FromRgb(17, 22, 29) : Colors.White;
                grid.CellFontColor = dark ? Color.FromRgb(227, 233, 242) : Colors.Black;
                grid.GridLineColor = dark ? Color.FromRgb(60, 70, 83) : Color.FromRgb(210, 215, 222);
                grid.HeaderBackground = dark ? Color.FromRgb(39, 47, 59) : Color.FromRgb(232, 238, 244);
                grid.HeaderCurrentBackground = dark ? Color.FromRgb(57, 77, 91) : Color.FromRgb(190, 207, 220);
                grid.AlternatingColors = dark ? new[] { Color.FromRgb(24, 30, 39) } : App.Instance.Setting.AlternatingColors;
                grid.SelectedColor = dark ? Color.FromRgb(39, 89, 116) : Color.FromRgb(51, 153, 255);
                grid.MouseOverRowColor = dark ? Color.FromRgb(39, 49, 64) : Color.FromRgb(240, 244, 250);
                grid.RefreshPresentation();
            }
            FitColumnsToView();
            SynchronizePresentationRowHeights();
            QueueFitColumns();
        }

        private void QueueFitColumns()
        {
            if (!presentationReady || columnFitQueued) return;
            columnFitQueued = true;
            Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
            {
                columnFitQueued = false;
                if (FitColumnsToView()) SynchronizePresentationRowHeights();
            }));
        }

        private bool FitColumnsToView()
        {
            int count = SrcDataGrid.Model?.ColumnCount ?? 0;
            if (count == 0 || count != (DstDataGrid.Model?.ColumnCount ?? 0)) return false;
            var widths = new int[count];
            for (int column = 0; column < count; column++)
                widths[column] = Math.Max(SrcDataGrid.GetPresentationColumnWidth(column), DstDataGrid.GetPresentationColumnWidth(column));
            bool changed = SrcDataGrid.FitPresentationColumns(widths);
            return DstDataGrid.FitPresentationColumns(widths) || changed;
        }

        private void SynchronizePresentationRowHeights()
        {
            // Use the same heights on both sides after wrapping or a font-size change.
            int count = Math.Max(SrcDataGrid.ModelRowCount, DstDataGrid.ModelRowCount);
            SrcDataGrid.ClearRowSizeOverrides();
            DstDataGrid.ClearRowSizeOverrides();
            for (int row = 0; row < count; row++)
            {
                int height = Math.Max(SrcDataGrid.CalculateModelRowHeight(row), DstDataGrid.CalculateModelRowHeight(row));
                SrcDataGrid.PutSizeOverride(row, height);
                DstDataGrid.PutSizeOverride(row, height);
            }
            foreach (var grid in new[] { SrcDataGrid, DstDataGrid }) { grid.BuildRowIndex(); grid.InvalidateAll(); }
        }

        private static string ExcelAddress(ExcelCell cell)
        {
            int column = cell.OriginalColumnIndex + 1;
            string letters = "";
            while (column > 0) { column--; letters = (char)('A' + column % 26) + letters; column /= 26; }
            return letters + (cell.OriginalRowIndex + 1);
        }

        private void BuildChangePanel(ExcelSheetDiff diff)
        {
            allChanges.Clear();
            foreach (var row in diff.Rows)
                foreach (var cell in row.Value.Cells.Values.Where(c => c.Status != ExcelCellStatus.None))
                {
                    var original = cell.Status == ExcelCellStatus.Removed ? cell.SrcCell : cell.DstCell;
                    allChanges.Add(new ChangeEntry
                    {
                        Row = cell.RowIndex, Column = cell.ColumnIndex, Number = allChanges.Count + 1,
                        Status = cell.Status, Address = ExcelAddress(original), Left = cell.SrcCell.Value, Right = cell.DstCell.Value
                    });
                }
            RefreshChangeList();
        }

        private void ChangeFilter_Click(object sender, RoutedEventArgs e) { RefreshChangeList(); }

        private void RefreshChangeList()
        {
            if (!presentationReady) return;
            FilterRemoved.Content = "− " + allChanges.Count(c => c.Status == ExcelCellStatus.Removed);
            FilterAdded.Content = "+ " + allChanges.Count(c => c.Status == ExcelCellStatus.Added);
            FilterModified.Content = "≈ " + allChanges.Count(c => c.Status == ExcelCellStatus.Modified);
            var selected = ChangeList.SelectedItem as ChangeEntry;
            updatingChangeSelection = true;
            var visible = allChanges.Where(c => c.Status == ExcelCellStatus.Added ? FilterAdded.IsChecked == true :
                c.Status == ExcelCellStatus.Removed ? FilterRemoved.IsChecked == true : FilterModified.IsChecked == true).ToList();
            ChangeList.ItemsSource = visible;
            if (selected != null) ChangeList.SelectedItem = visible.FirstOrDefault(c => c.Number == selected.Number);
            updatingChangeSelection = false;
            UpdateChangePosition();
            ChangeEmpty.Text = allChanges.Count == 0 ? "没有差异" : "没有符合筛选条件的差异";
            ChangeEmpty.Visibility = visible.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private void UpdateChangePosition()
        {
            ChangePosition.Text = (ChangeList.SelectedIndex + 1) + " / " + ChangeList.Items.Count;
        }

        private void ChangeList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!presentationReady || updatingChangeSelection) return;
            var selected = ChangeList.SelectedItem as ChangeEntry;
            UpdateChangePosition();
            var model = SrcDataGrid.Model as DiffGridModel;
            if (selected == null || model == null) return;
            var address = model.GetVisualCellAddress(new FastGridCellAddress(selected.Row, selected.Column));
            if (address.IsEmpty)
            {
                ShowAllRadioButton.IsChecked = true;
                address = model.GetVisualCellAddress(new FastGridCellAddress(selected.Row, selected.Column));
            }
            if (address.IsEmpty) return;
            updatingChangeSelection = true;
            SrcDataGrid.CurrentCell = address;
            DstDataGrid.CurrentCell = address;
            updatingChangeSelection = false;
            UpdateValueDiff(selected.Left, selected.Right);
            ChangeList.ScrollIntoView(selected);
        }

        private void SyncChangeSelection(FastGridControl grid)
        {
            if (!presentationReady || updatingChangeSelection) return;
            var model = grid.Model as DiffGridModel;
            if (model == null || grid.CurrentCell.IsEmpty) return;
            var address = model.GetRealCellAddress(grid.CurrentCell);
            updatingChangeSelection = true;
            ChangeList.SelectedItem = ChangeList.Items.Cast<ChangeEntry>().FirstOrDefault(c => c.Row == address.Row && c.Column == address.Column);
            if (ChangeList.SelectedItem != null) ChangeList.ScrollIntoView(ChangeList.SelectedItem);
            updatingChangeSelection = false;
            UpdateChangePosition();
        }

        private void ChangePrevious_Click(object sender, RoutedEventArgs e) { MoveChange(-1); }
        private void ChangeNext_Click(object sender, RoutedEventArgs e) { MoveChange(1); }
        private void MoveChange(int direction)
        {
            int count = ChangeList.Items.Count;
            if (count == 0) return;
            ChangeList.SelectedIndex = (ChangeList.SelectedIndex + direction + count) % count;
        }

        private void ChangePreview_Loaded(object sender, RoutedEventArgs e)
        {
            RenderChangePreview((TextBlock)sender);
        }

        private void ChangePreview_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            RenderChangePreview((TextBlock)sender);
        }

        private void RenderChangePreview(TextBlock text)
        {
            var entry = text.DataContext as ChangeEntry;
            text.Inlines.Clear();
            if (entry == null) return;
            bool right = (string)text.Tag == "Right";
            string value = right ? entry.Right : entry.Left;
            var prefix = new Run(right ? "新：" : "原：");
            prefix.SetResourceReference(TextElement.ForegroundProperty, "MutedBrush");
            text.Inlines.Add(prefix);
            if (string.IsNullOrEmpty(value)) { text.Inlines.Add("（空）"); return; }
            if (value.Length + (right ? entry.Left.Length : entry.Right.Length) > 4000)
            {
                text.Inlines.Add(value.Length > 1000 ? value.Substring(0, 1000) + "…" : value);
                return;
            }
            foreach (var part in DiffUtil.Diff(entry.Left.ToCharArray(), entry.Right.ToCharArray()))
            {
                if ((right && part.Status == DiffStatus.Deleted) || (!right && part.Status == DiffStatus.Inserted)) continue;
                char character = right ? part.Obj2 : part.Obj1;
                var run = new Run(character.ToString());
                if (part.Status != DiffStatus.Equal)
                {
                    run.SetResourceReference(TextElement.ForegroundProperty, right ? "RightBrush" : "LeftBrush");
                    run.FontWeight = FontWeights.Bold;
                }
                text.Inlines.Add(run);
            }
        }
    }
}

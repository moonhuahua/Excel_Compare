using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.IO;
using System.Windows;
using Prism.Mvvm;
using FastWpfGrid;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Behaviors;

namespace ExcelMerge.GUI.ViewModels
{
    public class DiffViewModel : BindableBase
    {
        private bool showLocationGridLine;
        public bool ShowLocationGridLine
        {
            get { return showLocationGridLine; }
            set { SetProperty(ref showLocationGridLine, value); }
        }

        private string srcPath;
        public string SrcPath
        {
            get { return srcPath; }
            set
            {
                if (!SetProperty(ref srcPath, value)) return;
                Settings.EMEnvironmentValue.Set("SRC", value);
                UpdateExecutableFlag(true);
            }
        }

        private string dstPath;
        public string DstPath
        {
            get { return dstPath; }
            set
            {
                if (!SetProperty(ref dstPath, value)) return;
                Settings.EMEnvironmentValue.Set("DST", value);
                UpdateExecutableFlag(false);
            }
        }

        private List<string> srcSheetNames;
        public List<string> SrcSheetNames
        {
            get { return srcSheetNames; }
            private set { SetProperty(ref srcSheetNames, value); }
        }

        private List<string> dstSheetNames;
        public List<string> DstSheetNames
        {
            get { return dstSheetNames; }
            private set { SetProperty(ref dstSheetNames, value); }
        }

        private int selectedSrcSheetIndex;
        public int SelectedSrcSheetIndex
        {
            get { return selectedSrcSheetIndex; }
            set { SetProperty(ref selectedSrcSheetIndex, value); }
        }

        private int selectedDstSheetIndex;
        public int SelectedDstSheetIndex
        {
            get { return selectedDstSheetIndex; }
            set { SetProperty(ref selectedDstSheetIndex, value); }
        }

        private bool executable;
        public bool Executable
        {
            get { return executable; }
            private set { SetProperty(ref executable, value); }
        }

        private int modifiedCellCount;
        public int ModifiedCellCount
        {
            get { return modifiedCellCount; }
            private set { SetProperty(ref modifiedCellCount, value); }
        }

        private int modifiedRowCount;
        public int ModifiedRowCount
        {
            get { return modifiedRowCount; }
            private set { SetProperty(ref modifiedRowCount, value); }
        }

        private int addedRowCount;
        public int AddedRowCount
        {
            get { return addedRowCount; }
            private set { SetProperty(ref addedRowCount, value); }
        }

        private int removedRowCount;
        public int RemovedRowCount
        {
            get { return removedRowCount; }
            private set { SetProperty(ref removedRowCount, value); }
        }

        private DragAcceptDescription description;
        public DragAcceptDescription Description
        {
            get { return description; }
            private set { SetProperty(ref description, value); }
        }

        public DiffViewModel()
        {
            Description = new DragAcceptDescription();
            Description.DragDrop += DragDrop;
            Description.DragOver += DragOver;

            SrcPath = string.Empty;
            DstPath = string.Empty;
        }

        public DiffViewModel(string src, string dst, MainWindowViewModel mwv) : this()
        {
            SrcPath = src;
            DstPath = dst;

            mwv.PropertyChanged += Mwv_PropertyChanged;
        }

        public void UpdateDiffSummary(ExcelSheetDiffSummary summary)
        {
            ModifiedCellCount = summary.ModifiedCellCount;
            ModifiedRowCount = summary.ModifiedRowCount;
            AddedRowCount = summary.AddedRowCount;
            RemovedRowCount = summary.RemovedRowCount;
        }

        private void Mwv_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SrcPath))
            {
                var vm = sender as MainWindowViewModel;
                if (vm != null)
                {
                    var prop = typeof(MainWindowViewModel).GetProperties().FirstOrDefault(p => p.Name == e.PropertyName);
                    if (prop != null)
                    {
                        SrcPath = prop.GetValue(vm) as string;
                    }
                }
            }
            else if (e.PropertyName == nameof(DstPath))
            {
                var vm = sender as MainWindowViewModel;
                if (vm != null)
                {
                    var prop = typeof(MainWindowViewModel).GetProperties().FirstOrDefault(p => p.Name == e.PropertyName);
                    if (prop != null)
                    {
                        DstPath = prop.GetValue(vm) as string;
                    }
                }
            }
        }

        private void DragDrop(DragEventArgs e)
        {
            var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (paths == null || !paths.Any())
                return;

            var target = e.Source as FrameworkElement;
            if (target == null)
                return;

            OnDragDrop(paths, target);
        }

        protected virtual void OnDragDrop(string[] filePath, FrameworkElement target)
        {
            if (filePath.Length > 1)
            {
                SrcPath = filePath[0];
                DstPath = filePath[1];

                return;
            }

            var tag = Convert.ToInt32(target.Tag);
            if (tag == 0)
                SrcPath = filePath[0];
            else if (tag == 1)
                DstPath = filePath[0];
        }

        private void DragOver(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;

            e.Handled = true;
        }

        private string srcError;
        public string SrcError
        {
            get { return srcError; }
            private set { SetProperty(ref srcError, value); }
        }

        private string dstError;
        public string DstError
        {
            get { return dstError; }
            private set { SetProperty(ref dstError, value); }
        }

        private void UpdateExecutableFlag(bool source)
        {
            string path = source ? SrcPath : DstPath;
            var sheets = new List<string>();
            string error = string.Empty;
            if (!string.IsNullOrWhiteSpace(path))
            {
                try
                {
                    string extension = Path.GetExtension(path).ToLowerInvariant();
                    if (!new[] { ".xls", ".xlsx", ".xlsm", ".csv", ".tsv" }.Contains(extension))
                        error = "请选择 XLSX、XLS、XLSM、CSV 或 TSV 文件";
                    else if (!File.Exists(path))
                        error = "文件不存在，请检查完整路径";
                    else
                    {
                        sheets = ExcelWorkbook.GetSheetNames(path).ToList();
                        if (sheets.Count == 0) error = "文件中没有可读取的工作表";
                    }
                }
                catch (Exception ex) { error = "无法读取文件：" + ex.Message; }
            }
            if (source)
            {
                SrcError = error;
                SrcSheetNames = sheets;
                SelectedSrcSheetIndex = sheets.Count > 0 ? 0 : -1;
            }
            else
            {
                DstError = error;
                DstSheetNames = sheets;
                SelectedDstSheetIndex = sheets.Count > 0 ? 0 : -1;
            }
            Executable = SrcSheetNames != null && SrcSheetNames.Count > 0
                && DstSheetNames != null && DstSheetNames.Count > 0
                && string.IsNullOrEmpty(SrcError) && string.IsNullOrEmpty(DstError);
        }
    }
}

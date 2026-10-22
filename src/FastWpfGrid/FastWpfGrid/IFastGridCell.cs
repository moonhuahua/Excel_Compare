using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace FastWpfGrid
{
    public sealed class CellPresentation
    {
        public string FontFamily { get; set; }
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikeout { get; set; }
        public string HorizontalAlignment { get; set; }
        public string VerticalAlignment { get; set; }
    }
    public interface IFormattedGridCell
    {
        CellPresentation Presentation { get; }
    }
    public enum CellDecoration
    {
        None,
        StrikeOutHorizontal,
    }

    public enum TooltipVisibilityMode
    {
        Always,
        OnlyWhenTrimmed,
    }

    public interface IFastGridCell
    {
        Color? BackgroundColor { get; }

        int BlockCount { get; }
        int RightAlignBlockCount { get; }
        IFastGridCellBlock GetBlock(int blockIndex);
        CellDecoration Decoration { get; }
        Color? DecorationColor { get; }

        /// <summary>
        /// return NULL disables inline editor
        /// </summary>
        /// <returns></returns>
        string GetEditText();
        void SetEditText(string value);

        string ToolTipText { get; }
        TooltipVisibilityMode ToolTipVisibility { get; }
    }
}

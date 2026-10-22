using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace FastWpfGrid
{
    public partial class FastGridControl
    {
        public bool WrapText { get; set; }
        public Color CanvasColor { get; set; } = Colors.White;
        private readonly Dictionary<string, WriteableBitmap> presentationCache = new Dictionary<string, WriteableBitmap>();

        public void RefreshPresentation()
        {
            presentationCache.Clear();
            RecountRowHeights();
            RecalculateHeaderSize();
            AdjustScrollbars();
            InvalidateAll();
        }

        public void SetPresentationColumnWidth(int column, int width)
        {
            _columnSizes.Resize(_columnSizes.ModelToReal(column), Math.Max(1, width));
        }

        public int GetPresentationColumnWidth(int column) { return _columnSizes.GetSizeByModelIndex(column); }

        public bool FitPresentationColumns(int[] proportions)
        {
            if (_drawBuffer == null || IsTransposed || Model == null || _modelColumnCount == 0 ||
                proportions.Length != _modelColumnCount || _realColumnCount != _modelColumnCount) return false;
            int available = _drawBuffer.PixelWidth - HeaderWidth;
            if (available < proportions.Length) return false;
            double total = 0;
            foreach (int width in proportions) total += Math.Max(1, width);
            int remaining = available - proportions.Length;
            int used = 0;
            double cumulative = 0;
            bool changed = false;
            for (int column = 0; column < proportions.Length; column++)
            {
                cumulative += Math.Max(1, proportions[column]);
                int end = column == proportions.Length - 1 ? available :
                    column + 1 + (int)Math.Round(remaining * cumulative / total);
                int width = end - used;
                if (width != GetPresentationColumnWidth(column))
                {
                    SetPresentationColumnWidth(column, width);
                    changed = true;
                }
                used = end;
            }
            FirstVisibleColumnScrollIndex = 0;
            AdjustScrollbars();
            hscroll.Maximum = 0;
            AdjustScrollBarPositions();
            InvalidateAll();
            return changed;
        }

        private FormattedText CreatePresentationText(IFastGridCell cell, int width, Color color, bool wrap)
        {
            var style = (cell as IFormattedGridCell)?.Presentation;
            var block = cell.GetBlock(0);
            string family = style?.FontFamily ?? CellFontName;
            double size = style != null && style.FontSize > 0 ? style.FontSize : CellFontSize;
            var face = new Typeface(new FontFamily(family),
                (style?.Italic ?? block.IsItalic) ? FontStyles.Italic : FontStyles.Normal,
                (style?.Bold ?? block.IsBold) ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal);
            var text = new FormattedText(block.TextData ?? string.Empty, CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight, face, Math.Max(1, size), new SolidColorBrush(color));
            text.MaxTextWidth = Math.Max(1, width);
            text.MaxLineCount = wrap ? int.MaxValue : 1;
            text.Trimming = wrap ? TextTrimming.None : TextTrimming.CharacterEllipsis;
            if (style?.HorizontalAlignment == "Right") text.TextAlignment = TextAlignment.Right;
            else if (style?.HorizontalAlignment == "Center") text.TextAlignment = TextAlignment.Center;
            if (style?.Underline == true) text.SetTextDecorations(TextDecorations.Underline);
            if (style?.Strikeout == true) text.SetTextDecorations(TextDecorations.Strikethrough);
            return text;
        }

        private int GetPresentationHeight(IFastGridCell cell, int width)
        {
            if (cell == null || cell.BlockCount != 1 || cell.GetBlock(0).BlockType != FastGridBlockType.Text) return 0;
            if (!WrapText && (cell as IFormattedGridCell)?.Presentation == null) return 0;
            return (int)Math.Ceiling(CreatePresentationText(cell, Math.Max(1, width - 2 * CellPaddingHorizontal), CellFontColor, WrapText).Height);
        }

        private bool RenderPresentationCell(IFastGridCell cell, IntRect rect, Color? selectedTextColor, Color background)
        {
            if (cell.BlockCount != 1 || cell.GetBlock(0).BlockType != FastGridBlockType.Text) return false;
            var style = (cell as IFormattedGridCell)?.Presentation;
            if (!WrapText && style == null) return false;
            int width = rect.Width, height = rect.Height;
            if (width < 1 || height < 1) return true;
            var block = cell.GetBlock(0);
            Color color = selectedTextColor ?? block.FontColor ?? CellFontColor;
            string key = string.Join("|", block.TextData, width, height, color, background, WrapText,
                style?.FontFamily ?? CellFontName, style?.FontSize ?? CellFontSize, style?.Bold ?? block.IsBold,
                style?.Italic ?? block.IsItalic, style?.Underline, style?.Strikeout, style?.HorizontalAlignment, style?.VerticalAlignment);
            WriteableBitmap bitmap;
            if (!presentationCache.TryGetValue(key, out bitmap))
            {
                var text = CreatePresentationText(cell, width, color, WrapText);
                double y = style?.VerticalAlignment == "Top" ? 0 :
                    style?.VerticalAlignment == "Bottom" ? Math.Max(0, height - text.Height) : Math.Max(0, (height - text.Height) / 2);
                var visual = new DrawingVisual();
                using (var drawing = visual.RenderOpen())
                {
                    drawing.PushClip(new RectangleGeometry(new Rect(0, 0, width, height)));
                    drawing.DrawRectangle(new SolidColorBrush(background), null, new Rect(0, 0, width, height));
                    drawing.DrawText(text, new Point(0, y));
                }
                var target = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
                target.Render(visual);
                bitmap = new WriteableBitmap(target);
                if (presentationCache.Count >= 1500) presentationCache.Clear();
                presentationCache[key] = bitmap;
            }
            _drawBuffer.Blit(new Rect(rect.Left, rect.Top, width, height), bitmap, new Rect(0, 0, width, height));
            return true;
        }
    }
}

using System.Collections.Generic;
using NPOI.SS.UserModel;
using System;
using System.Globalization;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace ExcelMerge
{
    internal class ExcelReader
    {
        internal static IEnumerable<ExcelRow> Read(ISheet sheet)
        {
            var styles = new Dictionary<short, ExcelCellAppearance>();
            var formatter = new DataFormatter(CultureInfo.CurrentCulture);
            var actualRowIndex = 0;
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null)
                    continue;

                var cells = new List<ExcelCell>();
                for (int columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++)
                {
                    var cell = row.GetCell(columnIndex);
                    var stringValue = ExcelUtility.GetCellStringValue(cell);

                    var result = new ExcelCell(stringValue, columnIndex, rowIndex);
                    if (cell != null)
                    {
                        var style = cell.CellStyle;
                        if (!styles.ContainsKey(style.Index)) styles[style.Index] = ReadAppearance(sheet.Workbook, style);
                        result.Appearance = styles[style.Index];
                        result.IsNumeric = cell.CellType == CellType.Numeric ||
                            (cell.CellType == CellType.Formula && cell.CachedFormulaResultType == CellType.Numeric);
                        if (result.IsNumeric)
                            result.DisplayValue = formatter.FormatRawCellContents(cell.NumericCellValue, style.DataFormat, style.GetDataFormatString());
                    }
                    cells.Add(result);
                }

                yield return new ExcelRow(actualRowIndex++, cells);
            }
        }
        private static string Rgb(byte[] rgb)
        {
            if (rgb == null || rgb.Length < 3) return null;
            int offset = rgb.Length - 3;
            return string.Format("#{0:X2}{1:X2}{2:X2}", rgb[offset], rgb[offset + 1], rgb[offset + 2]);
        }

        private static string IndexedColor(IWorkbook book, short index)
        {
            if (index < 0 || index >= 64) return null;
            var hssf = book as HSSFWorkbook;
            if (hssf != null) return Rgb(hssf.GetCustomPalette().GetColor(index)?.RGB);
            var colors = NPOI.HSSF.Util.HSSFColor.GetIndexHash();
            return colors.ContainsKey(index) ? Rgb((colors[index] as NPOI.HSSF.Util.HSSFColor)?.RGB) : null;
        }

        private static ExcelCellAppearance ReadAppearance(IWorkbook book, ICellStyle style)
        {
            var font = book.GetFontAt(style.FontIndex);
            var xf = font as XSSFFont;
            var xs = style as XSSFCellStyle;
            string foreground = xf != null ? XssfColor(book, xf.GetXSSFColor()) : IndexedColor(book, font.Color);
            string background = null;
            if (style.FillPattern == FillPattern.SolidForeground)
                background = xs != null ? XssfColor(book, xs.FillForegroundXSSFColor) : IndexedColor(book, style.FillForegroundColor);
            return new ExcelCellAppearance
            {
                FontName = font.FontName, FontSize = font.FontHeightInPoints * 96.0 / 72.0,
                Bold = font.IsBold, Italic = font.IsItalic, Underline = font.Underline != FontUnderlineType.None,
                Strikeout = font.IsStrikeout, Foreground = foreground, Background = background,
                HorizontalAlignment = style.Alignment.ToString(), VerticalAlignment = style.VerticalAlignment.ToString(),
                NumberFormat = style.GetDataFormatString()
            };
        }

        private static string XssfColor(IWorkbook book, XSSFColor color)
        {
            if (color == null || color.IsAuto) return null;
            (book as XSSFWorkbook)?.GetTheme()?.InheritFromThemeAsRequired(color);
            return Rgb(color.GetRgbWithTint() ?? color.RGB);
        }
    }
}

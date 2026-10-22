namespace ExcelMerge
{
    public class ExcelCell
    {
        public ExcelCellAppearance Appearance { get; set; }
        public string DisplayValue { get; set; }
        public bool IsNumeric { get; set; }
        public string Value { get; private set; }
        public int OriginalColumnIndex { get; private set; }
        public int OriginalRowIndex { get; private set; }

        public ExcelCell(string value, int originalColumnIndex, int originalRowIndex)
        {
            Value = value;
            DisplayValue = value;
            OriginalColumnIndex = originalColumnIndex;
            OriginalRowIndex = originalRowIndex;
        }
    }

    public sealed class ExcelCellAppearance
    {
        public string FontName { get; set; }
        public double FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikeout { get; set; }
        public string Foreground { get; set; }
        public string Background { get; set; }
        public string HorizontalAlignment { get; set; }
        public string VerticalAlignment { get; set; }
        public string NumberFormat { get; set; }
    }
}

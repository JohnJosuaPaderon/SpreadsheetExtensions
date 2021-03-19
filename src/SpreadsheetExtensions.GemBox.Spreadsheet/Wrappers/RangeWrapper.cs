using GemBox.Spreadsheet;
using SpreadsheetExtensions.Styles;

namespace SpreadsheetExtensions.Wrappers
{
    internal sealed class RangeWrapper : IRangeWrapper
    {
        private readonly CellRange _range;
        public object Range => _range;

        public RangeWrapper(CellRange range)
        {
            _range = range;
        }

        public IRangeWrapper Merge(bool isMerged)
        {
            _range.Merged = isMerged;
            return this;
        }

        public IRangeWrapper SetBorderBottom(SpreadsheetBorder border)
        {
            _range.Style.Borders.Apply(IndividualBorder.Bottom, border);
            return this;
        }

        public IRangeWrapper SetBorderLeft(SpreadsheetBorder border)
        {
            _range.Style.Borders.Apply(IndividualBorder.Left, border);
            return this;
        }

        public IRangeWrapper SetBorderRight(SpreadsheetBorder border)
        {
            _range.Style.Borders.Apply(IndividualBorder.Right, border);
            return this;
        }

        public IRangeWrapper SetBorderTop(SpreadsheetBorder border)
        {
            _range.Style.Borders.Apply(IndividualBorder.Top, border);
            return this;
        }

        public IRangeWrapper SetFontBold(bool isBold)
        {
            _range.Style.Font.Weight = isBold ? ExcelFont.BoldWeight : ExcelFont.NormalWeight;
            return this;
        }

        public IRangeWrapper SetFontItalic(bool isItalic)
        {
            _range.Style.Font.Italic = isItalic;
            return this;
        }

        public IRangeWrapper SetFontUnderline(bool isUnderline)
        {
            _range.Style.Font.UnderlineStyle = isUnderline ? UnderlineStyle.Single : UnderlineStyle.None;
            return this;
        }

        public IRangeWrapper SetHorizontalAlignment(SpreadsheetHorizontalAlignment alignment)
        {
            _range.Style.HorizontalAlignment = SpreadsheetHorizontalAlignmentMapper.MapStyle(alignment);
            return this;
        }

        public IRangeWrapper SetValue(object value)
        {
            _range.Value = value;
            return this;
        }

        public IRangeWrapper SetVerticalAlignment(SpreadsheetVerticalAlignment alignment)
        {
            _range.Style.VerticalAlignment = SpreadsheetVerticalAlignmentMapper.MapStyle(alignment);
            return this;
        }
    }
}

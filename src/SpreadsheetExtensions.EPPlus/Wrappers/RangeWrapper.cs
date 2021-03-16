using OfficeOpenXml;
using SpreadsheetExtensions.Styles;

namespace SpreadsheetExtensions.Wrappers
{
    internal sealed class RangeWrapper : IRangeWrapper
    {
        private readonly ExcelRange _range;
        public object Range => _range;

        public RangeWrapper(ExcelRange range)
        {
            _range = range;
        }

        public IRangeWrapper Merge(bool isMerged)
        {
            _range.Merge = isMerged;
            return this;
        }

        public IRangeWrapper SetHorizontalAlignment(SpreadsheetHorizontalAlignment alignment)
        {
            _range.Style.HorizontalAlignment = SpreadsheetHorizontalAlignmentMapper.Map(alignment);
            return this;
        }

        public IRangeWrapper SetValue(object value)
        {
            _range.Value = value;
            return this;
        }

        public IRangeWrapper SetVerticalAlignment(SpreadsheetVerticalAlignment alignment)
        {
            _range.Style.VerticalAlignment = SpreadsheetVerticalAlignmentMapper.Map(alignment);
            return this;
        }

        public IRangeWrapper SetFontBold(bool isBold)
        {
            _range.Style.Font.Bold = isBold;
            return this;
        }

        public IRangeWrapper SetBorderTop(SpreadsheetBorder border)
        {
            _range.Style.Border.Top.Apply(border);
            return this;
        }

        public IRangeWrapper SetBorderBottom(SpreadsheetBorder border)
        {
            _range.Style.Border.Bottom.Apply(border);
            return this;
        }

        public IRangeWrapper SetBorderLeft(SpreadsheetBorder border)
        {
            _range.Style.Border.Left.Apply(border);
            return this;
        }

        public IRangeWrapper SetBorderRight(SpreadsheetBorder border)
        {
            _range.Style.Border.Right.Apply(border);
            return this;
        }

        public IRangeWrapper SetFontItalic(bool isItalic)
        {
            _range.Style.Font.Italic = isItalic;
            return this;
        }

        public IRangeWrapper SetFontUnderline(bool isUnderline)
        {
            _range.Style.Font.UnderLine = isUnderline;
            return this;
        }
    }
}

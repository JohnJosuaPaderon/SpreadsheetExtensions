using SpreadsheetExtensions.Styles;

namespace SpreadsheetExtensions.Wrappers
{
    public interface IRangeWrapper
    {
        object Range { get; }
        IRangeWrapper SetValue(object value);
        IRangeWrapper Merge(bool isMerged);
        IRangeWrapper SetHorizontalAlignment(SpreadsheetHorizontalAlignment alignment);
        IRangeWrapper SetFontBold(bool isBold);
        IRangeWrapper SetVerticalAlignment(SpreadsheetVerticalAlignment alignment);
        IRangeWrapper SetBorderTop(SpreadsheetBorder border);
        IRangeWrapper SetBorderBottom(SpreadsheetBorder border);
        IRangeWrapper SetBorderLeft(SpreadsheetBorder border);
        IRangeWrapper SetBorderRight(SpreadsheetBorder border);
        IRangeWrapper SetFontItalic(bool isItalic);
        IRangeWrapper SetFontUnderline(bool isUnderline);
    }
}

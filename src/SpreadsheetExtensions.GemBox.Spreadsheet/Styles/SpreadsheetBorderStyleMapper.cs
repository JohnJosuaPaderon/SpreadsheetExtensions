using GemBox.Spreadsheet;

namespace SpreadsheetExtensions.Styles
{
    internal static class SpreadsheetBorderStyleMapper
    {
        public static LineStyle Map(SpreadsheetBorderStyle style)
        {
            return style switch
            {
                SpreadsheetBorderStyle.None => LineStyle.None,
                SpreadsheetBorderStyle.Thin => LineStyle.Thin,
                SpreadsheetBorderStyle.Thick => LineStyle.Thick,
                SpreadsheetBorderStyle.Dotted => LineStyle.Dotted,
                SpreadsheetBorderStyle.Dashed => LineStyle.Dashed,
                SpreadsheetBorderStyle.Double => LineStyle.Double,
                _ => SpreadsheetLibraryDefaults.LineStyle
            };
        }
    }
}

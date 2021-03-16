using OfficeOpenXml.Style;

namespace SpreadsheetExtensions.Styles
{
    internal static class SpreadsheetVerticalAlignmentMapper
    {
        public static ExcelVerticalAlignment Map(SpreadsheetVerticalAlignment alignment)
        {
            return alignment switch
            {
                SpreadsheetVerticalAlignment.Top => ExcelVerticalAlignment.Top,
                SpreadsheetVerticalAlignment.Center => ExcelVerticalAlignment.Center,
                SpreadsheetVerticalAlignment.Bottom => ExcelVerticalAlignment.Bottom,
                _ => SpreadsheetLibraryDefaults.VerticalAlignment
            };
        }
    }
}

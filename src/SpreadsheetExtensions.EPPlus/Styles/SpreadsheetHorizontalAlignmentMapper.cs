using OfficeOpenXml.Style;

namespace SpreadsheetExtensions.Styles
{
    internal static class SpreadsheetHorizontalAlignmentMapper
    {
        public static ExcelHorizontalAlignment Map(SpreadsheetHorizontalAlignment alignment)
        {
            return alignment switch
            {
                SpreadsheetHorizontalAlignment.Left => ExcelHorizontalAlignment.Left,
                SpreadsheetHorizontalAlignment.Center => ExcelHorizontalAlignment.Center,
                SpreadsheetHorizontalAlignment.Right => ExcelHorizontalAlignment.Right,
                _ => SpreadsheetLibraryDefaults.HorizontalAlignment
            };
        }
    }
}

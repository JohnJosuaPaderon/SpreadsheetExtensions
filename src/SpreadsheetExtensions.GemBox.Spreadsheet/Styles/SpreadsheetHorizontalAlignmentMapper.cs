using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Drawing;

namespace SpreadsheetExtensions.Styles
{
    internal static class SpreadsheetHorizontalAlignmentMapper
    {
        public static HorizontalAlignment Map(SpreadsheetHorizontalAlignment alignment)
        {
            return alignment switch
            {
                SpreadsheetHorizontalAlignment.Left => HorizontalAlignment.Left,
                SpreadsheetHorizontalAlignment.Center => HorizontalAlignment.Center,
                SpreadsheetHorizontalAlignment.Right => HorizontalAlignment.Right,
                _ => SpreadsheetLibraryDefaults.HorizontalAlignment
            };
        }

        public static HorizontalAlignmentStyle MapStyle(SpreadsheetHorizontalAlignment alignment)
        {
            return alignment switch
            {
                SpreadsheetHorizontalAlignment.Left => HorizontalAlignmentStyle.Left,
                SpreadsheetHorizontalAlignment.Center => HorizontalAlignmentStyle.Center,
                SpreadsheetHorizontalAlignment.Right => HorizontalAlignmentStyle.Right,
                _ => SpreadsheetLibraryDefaults.HorizontalAlignmentStyle
            };
        }
    }
}

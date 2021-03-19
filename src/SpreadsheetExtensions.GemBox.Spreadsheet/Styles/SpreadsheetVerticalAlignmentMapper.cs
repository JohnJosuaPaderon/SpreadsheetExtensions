using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Drawing;

namespace SpreadsheetExtensions.Styles
{
    internal static class SpreadsheetVerticalAlignmentMapper
    {
        public static VerticalAlignment Map(SpreadsheetVerticalAlignment alignment)
        {
            return alignment switch
            {
                SpreadsheetVerticalAlignment.Top => VerticalAlignment.Top,
                SpreadsheetVerticalAlignment.Center => VerticalAlignment.Middle,
                SpreadsheetVerticalAlignment.Bottom => VerticalAlignment.Bottom,
                _ => SpreadsheetLibraryDefaults.VerticalAlignment
            };
        }

        public static VerticalAlignmentStyle MapStyle(SpreadsheetVerticalAlignment alignment)
        {
            return alignment switch
            {
                SpreadsheetVerticalAlignment.Top => VerticalAlignmentStyle.Top,
                SpreadsheetVerticalAlignment.Center => VerticalAlignmentStyle.Center,
                SpreadsheetVerticalAlignment.Bottom => VerticalAlignmentStyle.Bottom,
                _ => SpreadsheetLibraryDefaults.VerticalAlignmentStyle
            };
        }
    }
}

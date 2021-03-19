using GemBox.Spreadsheet;
using GemBox.Spreadsheet.Drawing;

namespace SpreadsheetExtensions
{
    public static class SpreadsheetLibraryDefaults
    {
        public static HorizontalAlignment HorizontalAlignment { get; set; } = HorizontalAlignment.Left;
        public static HorizontalAlignmentStyle HorizontalAlignmentStyle { get; set; } = HorizontalAlignmentStyle.Left;
        public static VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Middle;
        public static VerticalAlignmentStyle VerticalAlignmentStyle { get; set; } = VerticalAlignmentStyle.Center;
        public static LineStyle LineStyle { get; set; } = LineStyle.Thin;
    }
}

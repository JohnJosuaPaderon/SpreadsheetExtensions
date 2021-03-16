using OfficeOpenXml.Style;

namespace SpreadsheetExtensions
{
    public static class SpreadsheetLibraryDefaults
    {
        public static ExcelHorizontalAlignment HorizontalAlignment { get; set; } = ExcelHorizontalAlignment.Left;
        public static ExcelVerticalAlignment VerticalAlignment { get; set; } = ExcelVerticalAlignment.Center;
        public static ExcelBorderStyle BorderStyle { get; set; } = ExcelBorderStyle.Thin;
    }
}

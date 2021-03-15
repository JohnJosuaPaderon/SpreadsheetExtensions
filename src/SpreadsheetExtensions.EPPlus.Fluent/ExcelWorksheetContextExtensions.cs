using OfficeOpenXml.Style;

namespace SpreadsheetExtensions
{
    public static class ExcelWorksheetContextExtensions
    {
        public static ExcelWorksheetContext Write(this ExcelWorksheetContext instance, int row, int column, object value)
        {
            var range = instance.GetRange(row, column);
            range.Value = value;
            return instance;
        }

        public static ExcelWorksheetContext SetHorizontalAlignment(this ExcelWorksheetContext instance, ExcelHorizontalAlignment alignment)
        {
            instance.CurrentRange.Style.HorizontalAlignment = alignment;
            return instance;
        }

        public static ExcelWorksheetContext SetVerticalAlignment(this ExcelWorksheetContext instance, ExcelVerticalAlignment alignment)
        {
            instance.CurrentRange.Style.VerticalAlignment = alignment;
            return instance;
        }

        public static ExcelWorksheetContext SetFontFold(this ExcelWorksheetContext instance, bool isBold)
        {
            instance.CurrentRange.Style.Font.Bold = isBold;
            return instance;
        }
    }
}

using OfficeOpenXml;

namespace SpreadsheetExtensions
{
    public static class ExcelWorksheetExtensions
    {
        public static ExcelWorksheetContext CreateContext(this ExcelWorksheet instance) => new(instance);

        public static ExcelWorksheetContext Write(this ExcelWorksheet instance, int row, int column, object value) => instance.CreateContext().Write(row, column, value);
    }
}

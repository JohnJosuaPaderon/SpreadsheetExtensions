using GemBox.Spreadsheet;
using SpreadsheetExtensions.Wrappers;

namespace SpreadsheetExtensions
{
    public static class ExcelWorksheetExtensions
    {
        public static WorksheetContext CreateContext(this ExcelWorksheet instance) => new(new WorksheetWrapper(instance));

        public static WorksheetContext Write(this ExcelWorksheet instance, string address, object value) => instance
            .CreateContext()
            .Write(address, value);

        public static WorksheetContext Write(this ExcelWorksheet instance, int row, int column, object value) => instance
            .CreateContext()
            .Write(row, column, value);

        public static WorksheetContext Write(this ExcelWorksheet instance, int fromRow, int fromColumn, int toRow, int toColumn, object value) => instance
            .CreateContext()
            .Write(fromRow, fromColumn, toRow, toColumn, value);
    }
}

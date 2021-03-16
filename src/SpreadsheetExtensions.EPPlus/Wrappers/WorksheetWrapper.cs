using OfficeOpenXml;

namespace SpreadsheetExtensions.Wrappers
{
    internal sealed class WorksheetWrapper : WorksheetWrapperBase<ExcelWorksheet>, IWorksheetWrapper
    {
        public WorksheetWrapper(ExcelWorksheet worksheet) : base(worksheet, SpreadsheetLibrarySettings.Instance)
        {
        }

        public IRangeWrapper GetRange(string address)
        {
            var range = GetUnderlyingWorksheet().Cells[address];
            return new RangeWrapper(range);
        }

        public IRangeWrapper GetRange(int row, int column)
        {
            var range = GetUnderlyingWorksheet().Cells[row, column];
            return new RangeWrapper(range);
        }

        public IRangeWrapper GetRange(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            var range = GetUnderlyingWorksheet().Cells[fromRow, fromColumn, toRow, toColumn];
            return new RangeWrapper(range);
        }
    }
}

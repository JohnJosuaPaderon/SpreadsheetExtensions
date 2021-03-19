using GemBox.Spreadsheet;

namespace SpreadsheetExtensions.Wrappers
{
    internal sealed class WorksheetWrapper : WorksheetWrapperBase<ExcelWorksheet>, IWorksheetWrapper
    {
        public WorksheetWrapper(ExcelWorksheet worksheet) : base(worksheet, SpreadsheetLibrarySettings.Instance)
        {
        }

        public IRangeWrapper GetRange(string address)
        {
            var range = GetUnderlyingWorksheet().Cells.GetSubrange(address);
            return new RangeWrapper(range);
        }

        public IRangeWrapper GetRange(int row, int column)
        {
            var range = GetUnderlyingWorksheet().Cells.GetSubrangeRelative(row, column, 1, 1);
            return new RangeWrapper(range);
        }

        public IRangeWrapper GetRange(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            var range = GetUnderlyingWorksheet().Cells.GetSubrangeAbsolute(fromRow, fromColumn, toRow, toColumn);
            return new RangeWrapper(range);
        }
    }
}

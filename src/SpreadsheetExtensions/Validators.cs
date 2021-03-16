using SpreadsheetExtensions.Exceptions;
using SpreadsheetExtensions.Wrappers;

namespace SpreadsheetExtensions
{
    public static class Validators
    {
        public static void MinRow(int row, int minRow, string propertyName)
        {
            if (row < minRow)
                throw new SpreadsheetMinRowException($"Value of '{propertyName}' is less than the minimum row: {minRow}");
        }

        public static void MaxRow(int row, int maxRow, string propertyName)
        {
            if (row > maxRow)
                throw new SpreadsheetException($"Value of '{propertyName}' is less than the maximum row: {maxRow}");
        }

        public static void MinColumn(int column, int minColumn, string propertyName)
        {
            if (column < minColumn)
                throw new SpreadsheetException($"Value of '{propertyName}' is less than the minimum column: {minColumn}");
        }

        public static void MaxColumn(int column, int maxColumn, string propertyName)
        {
            if (column > maxColumn)
                throw new SpreadsheetException($"Value of '{propertyName}' is greater than the maximum column: {maxColumn}");
        }

        public static void WorksheetWrapper(IWorksheetWrapper wrapper)
        {
            if (wrapper is null)
                throw new SpreadsheetNullWorksheetWrapperException();
        }

        public static void RangeWrapper(IRangeWrapper wrapper)
        {
            if (wrapper is null)
                throw new SpreadsheetNullRangeWrapperException();
        }

        public static void RangeWrapper(IRangeWrapper wrapper, string message)
        {
            if (wrapper is null)
                throw new SpreadsheetNullRangeWrapperException(message);
        }

        public static void Address(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                throw new SpreadsheetAddressException();
        }

        public static void IndexRange(int from, int to)
        {
            if (from > to)
                throw new SpreadsheetIndexRangeException();
        }
    }
}

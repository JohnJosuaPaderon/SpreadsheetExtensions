using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetNullRangeWrapperException : SpreadsheetException
    {
        public SpreadsheetNullRangeWrapperException()
        {
        }

        public SpreadsheetNullRangeWrapperException(string message) : base(message)
        {
        }

        public SpreadsheetNullRangeWrapperException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

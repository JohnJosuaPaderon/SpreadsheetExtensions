using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetIndexRangeException : SpreadsheetException
    {
        public SpreadsheetIndexRangeException()
        {
        }

        public SpreadsheetIndexRangeException(string message) : base(message)
        {
        }

        public SpreadsheetIndexRangeException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

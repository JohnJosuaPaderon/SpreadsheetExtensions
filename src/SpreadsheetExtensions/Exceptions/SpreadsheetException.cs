using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetException : Exception
    {
        public SpreadsheetException()
        {
        }

        public SpreadsheetException(string message) : base(message)
        {
        }

        public SpreadsheetException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

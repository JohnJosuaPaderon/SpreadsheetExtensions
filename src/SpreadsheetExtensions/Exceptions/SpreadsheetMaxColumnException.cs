using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetMaxColumnException : SpreadsheetException
    {
        public SpreadsheetMaxColumnException()
        {
        }

        public SpreadsheetMaxColumnException(string message) : base(message)
        {
        }

        public SpreadsheetMaxColumnException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

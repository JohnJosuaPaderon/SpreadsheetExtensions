using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetMinColumnException : SpreadsheetException
    {
        public SpreadsheetMinColumnException()
        {
        }

        public SpreadsheetMinColumnException(string message) : base(message)
        {
        }

        public SpreadsheetMinColumnException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

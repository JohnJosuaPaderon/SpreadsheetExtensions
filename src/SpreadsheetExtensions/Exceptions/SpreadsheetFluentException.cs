using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetFluentException : SpreadsheetException
    {
        public SpreadsheetFluentException()
        {
        }

        public SpreadsheetFluentException(string message) : base(message)
        {
        }

        public SpreadsheetFluentException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

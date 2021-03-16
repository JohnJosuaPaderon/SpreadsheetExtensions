using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetMinRowException : SpreadsheetException
    {
        public SpreadsheetMinRowException()
        {
        }

        public SpreadsheetMinRowException(string message) : base(message)
        {
        }

        public SpreadsheetMinRowException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

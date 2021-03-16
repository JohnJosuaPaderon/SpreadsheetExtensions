using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetMaxRowException : SpreadsheetException
    {
        public SpreadsheetMaxRowException()
        {
        }

        public SpreadsheetMaxRowException(string message) : base(message)
        {
        }

        public SpreadsheetMaxRowException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

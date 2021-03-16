using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetNullWorksheetWrapperException : SpreadsheetException
    {
        public SpreadsheetNullWorksheetWrapperException()
        {
        }

        public SpreadsheetNullWorksheetWrapperException(string message) : base(message)
        {
        }

        public SpreadsheetNullWorksheetWrapperException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

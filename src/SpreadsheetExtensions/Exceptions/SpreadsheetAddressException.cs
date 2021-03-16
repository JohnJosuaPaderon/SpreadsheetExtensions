using System;

namespace SpreadsheetExtensions.Exceptions
{
    public class SpreadsheetAddressException : SpreadsheetException
    {
        public SpreadsheetAddressException()
        {
        }

        public SpreadsheetAddressException(string message) : base(message)
        {
        }

        public SpreadsheetAddressException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}

using System;

namespace SpreadsheetExtensions
{
    public class SpreadhsheetExtensionsException : Exception
    {
        public SpreadhsheetExtensionsException()
        {
        }

        public SpreadhsheetExtensionsException(string message) : base(message)
        {
        }

        public SpreadhsheetExtensionsException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }

    public class SpreadsheetExtensionsMinRowException : SpreadhsheetExtensionsException
    {
        public string PropertyName { get; }

        public SpreadsheetExtensionsMinRowException(string propertyName) : base($"Value of '{propertyName}' cannot be less than the minimum row which is 1")
        {
            PropertyName = propertyName;
        }

        public SpreadsheetExtensionsMinRowException(string propertyName, string message) : base(message)
        {
            PropertyName = propertyName;
        }

        public SpreadsheetExtensionsMinRowException(string propertyName, string message, Exception innerException) : base(message, innerException)
        {
            PropertyName = propertyName;
        }
    }
}

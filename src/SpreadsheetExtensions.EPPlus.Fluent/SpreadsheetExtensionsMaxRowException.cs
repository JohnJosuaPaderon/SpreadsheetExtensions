using System;

namespace SpreadsheetExtensions
{
    public class SpreadsheetExtensionsMaxRowException : SpreadhsheetExtensionsException
    {
        public string PropertyName { get; }
        public SpreadsheetExtensionsMaxRowException(string propertyName) : base($"Value of '{propertyName}' cannot be greater than the maximum row which is ${ExcelPackage.MaxRows}")
        {
            PropertyName = propertyName;
        }

        public SpreadsheetExtensionsMaxRowException(string propertyName, string message) : base(message)
        {
            PropertyName = propertyName;
        }

        public SpreadsheetExtensionsMaxRowException(string propertyName, string message, Exception innerException) : base(message, innerException)
        {
            PropertyName = propertyName;
        }
    }
}

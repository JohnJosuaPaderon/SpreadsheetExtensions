namespace SpreadsheetExtensions
{
    internal static class Validators
    {
        public static void MinRow(int row, string propertyName)
        {
            if (row < 1)
                throw new SpreadsheetExtensionsMinRowException(propertyName);
        }

        public static void MaxRow(int row, string propertyName)
        {
            if (row > ExcelPackage.MaxRows)
                throw new SpreadsheetExtensionsMaxRowException(propertyName);
        }
    }
}

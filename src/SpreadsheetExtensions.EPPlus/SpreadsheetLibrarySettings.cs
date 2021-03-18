using OfficeOpenXml;

namespace SpreadsheetExtensions
{
    internal sealed class SpreadsheetLibrarySettings : ISpreadsheetLibrarySettings
    {
        public int MinRow { get; }
        public int MaxRow { get; }
        public int MinColumn { get; }
        public int MaxColumn { get; }

        private SpreadsheetLibrarySettings()
        {
            MinRow = 1;
            MinColumn = 1;
            MaxRow = ExcelPackage.MaxRows;
            MaxColumn = ExcelPackage.MaxColumns;
        }

        public static SpreadsheetLibrarySettings Instance { get; } = new SpreadsheetLibrarySettings();
    }
}

namespace SpreadsheetExtensions
{
    internal sealed class SpreadsheetLibrarySettings : ISpreadsheetLibrarySettings
    {
        public int MinRow { get; set; }
        public int MaxRow { get; set; }
        public int MinColumn { get; set; }
        public int MaxColumn { get; set; }

        private SpreadsheetLibrarySettings()
        {
            MinRow = 0;
            MinColumn = 0;
            MaxRow = 1048576;
            MaxColumn = 16384;
        }

        public static SpreadsheetLibrarySettings Instance { get; } = new SpreadsheetLibrarySettings();
    }
}

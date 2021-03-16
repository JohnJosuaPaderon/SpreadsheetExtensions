namespace SpreadsheetExtensions.Wrappers
{
    public abstract class WorksheetWrapperBase<TWorksheet>
    {
        private readonly TWorksheet _worksheet;
        public object Worksheet => _worksheet;
        public ISpreadsheetLibrarySettings LibrarySettings { get; }

        public WorksheetWrapperBase(TWorksheet worksheet, ISpreadsheetLibrarySettings librarySettings)
        {
            _worksheet = worksheet;
            LibrarySettings = librarySettings;
        }

        public TWorksheet GetUnderlyingWorksheet() => _worksheet;
    }
}

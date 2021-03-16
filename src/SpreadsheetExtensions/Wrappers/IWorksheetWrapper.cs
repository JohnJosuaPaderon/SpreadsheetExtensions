namespace SpreadsheetExtensions.Wrappers
{
    public interface IWorksheetWrapper
    {

        object Worksheet { get; }
        ISpreadsheetLibrarySettings LibrarySettings { get; }
        IRangeWrapper GetRange(string address);
        IRangeWrapper GetRange(int row, int column);
        IRangeWrapper GetRange(int fromRow, int fromColumn, int toRow, int toColumn);
    }
}

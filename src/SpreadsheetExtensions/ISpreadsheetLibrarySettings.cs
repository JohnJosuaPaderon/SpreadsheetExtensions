namespace SpreadsheetExtensions
{
    public interface ISpreadsheetLibrarySettings
    {
        int MinRow { get; }
        int MaxRow { get; }
        int MinColumn { get; }
        int MaxColumn { get; }
    }
}

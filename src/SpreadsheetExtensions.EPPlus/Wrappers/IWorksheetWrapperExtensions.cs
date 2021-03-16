using SpreadsheetExtensions.Exceptions;

namespace SpreadsheetExtensions.Wrappers
{
    internal static class IWorksheetWrapperExtensions
    {
        public static WorksheetWrapper AsWorksheetWrapper(this IWorksheetWrapper instance)
        {
            if (instance is WorksheetWrapper result)
                return result;

            throw new SpreadsheetException($"'{nameof(instance)}' is not an instance of '{typeof(WorksheetWrapper).FullName}'");
        }
    }
}

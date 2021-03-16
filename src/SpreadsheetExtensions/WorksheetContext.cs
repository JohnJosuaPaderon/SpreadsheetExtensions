using SpreadsheetExtensions.Wrappers;

namespace SpreadsheetExtensions
{
    public class WorksheetContext
    {
        private readonly object _lockObj = new();
        private readonly IWorksheetWrapper _wrapper;

        private IRangeWrapper _currentRange;
        public IRangeWrapper CurrentRange
        {
            get
            {
                lock (_lockObj)
                {
                    Validators.RangeWrapper(_currentRange);
                    return _currentRange;
                }
            }
            private set
            {
                lock (_lockObj)
                {
                    _currentRange = value;
                }
            }
        }

        public WorksheetContext(IWorksheetWrapper wrapper)
        {
            Validators.WorksheetWrapper(wrapper);
            _wrapper = wrapper;
        }

        private IRangeWrapper TrySetCurrentRange(IRangeWrapper range, bool setAsCurrent)
        {
            if (setAsCurrent)
            {
                Validators.RangeWrapper(range, "Cannot set current range to null");
                CurrentRange = range;
            }

            return range;
        }

        public void RemoveCurrentRange()
        {
            CurrentRange = null;
        }

        public IRangeWrapper GetRange(string address, bool setAsCurrent = true)
        {
            Validators.Address(address);

            var range = _wrapper.GetRange(address);
            return TrySetCurrentRange(range, setAsCurrent);
        }

        public IRangeWrapper GetRange(int row, int column, bool setAsCurrent = true)
        {
            Validators.MinRow(row, _wrapper.LibrarySettings.MinRow, nameof(row));
            Validators.MaxRow(row, _wrapper.LibrarySettings.MaxRow, nameof(row));
            Validators.MinColumn(column, _wrapper.LibrarySettings.MinColumn, nameof(column));
            Validators.MaxColumn(column, _wrapper.LibrarySettings.MaxColumn, nameof(column));

            var range = _wrapper.GetRange(row, column);
            return TrySetCurrentRange(range, setAsCurrent);
        }
        public IRangeWrapper GetRange(int fromRow, int fromColumn, int toRow, int toColumn, bool setAsCurrent = true)
        {
            Validators.MinRow(fromRow, _wrapper.LibrarySettings.MinRow, nameof(fromRow));
            Validators.MaxRow(fromRow, _wrapper.LibrarySettings.MaxRow, nameof(fromRow));
            Validators.MinColumn(fromColumn, _wrapper.LibrarySettings.MinColumn, nameof(fromColumn));
            Validators.MaxColumn(fromColumn, _wrapper.LibrarySettings.MaxColumn, nameof(fromColumn));
            Validators.MinRow(toRow, _wrapper.LibrarySettings.MinRow, nameof(toRow));
            Validators.MaxRow(toRow, _wrapper.LibrarySettings.MaxRow, nameof(toRow));
            Validators.MinColumn(toColumn, _wrapper.LibrarySettings.MinColumn, nameof(toColumn));
            Validators.MaxColumn(toColumn, _wrapper.LibrarySettings.MaxColumn, nameof(toColumn));
            Validators.IndexRange(fromRow, toRow);
            Validators.IndexRange(fromColumn, toColumn);

            var range = _wrapper.GetRange(fromRow, fromColumn, toRow, toColumn);
            return TrySetCurrentRange(range, setAsCurrent);
        }
    }
}

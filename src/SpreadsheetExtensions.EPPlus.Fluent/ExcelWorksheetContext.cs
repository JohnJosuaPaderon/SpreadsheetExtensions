using OfficeOpenXml;
using System;

namespace SpreadsheetExtensions
{
    public class ExcelWorksheetContext
    {
        private readonly object _lockObj;
        private readonly ExcelWorksheet _worksheet;

        private ExcelRange _currentRange;
        public ExcelRange CurrentRange
        {
            get
            {
                lock (_lockObj)
                {
                    if (_currentRange is null)
                        throw new SpreadhsheetExtensionsException($"{nameof(CurrentRange)} is null");

                    return _currentRange;
                }
            }
        }

        public ExcelWorksheetContext(ExcelWorksheet worksheet)
        {
            if (worksheet is null)
                throw new SpreadhsheetExtensionsException($"Argument {nameof(worksheet)} is null");

            _worksheet = worksheet;
        }

        internal void SetCurrentRange(ExcelRange currentRange)
        {
            if (currentRange is null)
                throw new ArgumentNullException(nameof(currentRange));

            lock (_lockObj)
            {
                _currentRange = currentRange;
            }
        }

        internal void RemoveCurrentRange()
        {
            lock (_lockObj)
            {
                _currentRange = null;
            }
        }

        internal ExcelRange GetRange(int row, int column, bool setAsCurrent = true)
        {
            if (row < 0)
                throw new SpreadhsheetExtensionsException($"Argument {nameof(row)} cannot be less than zero");

            if (column < 0)
                throw new SpreadhsheetExtensionsException($"Argument {nameof(column)} cannot be less than zero");

            var range = _worksheet.Cells[row, column];

            if (setAsCurrent)
                SetCurrentRange(range);

            return range;
        }

        internal ExcelRange GetRange(string address, bool setAsCurrent = true)
        {
            if (string.IsNullOrWhiteSpace(address))
                throw new SpreadhsheetExtensionsException($"Argument {nameof(address)} cannot be null or empty");

            var range = _worksheet.Cells[address];

            if (setAsCurrent)
                SetCurrentRange(range);

            return range;
        }

        internal ExcelRange GetRange(int fromRow, int fromColumn, int toRow, int toColumn, bool setAsCurrent = true)
        {
            if (fromRow < 0)
                throw new SpreadhsheetExtensionsException($"Argument '{nameof(fromRow)}' cannot be less than zero");

            if (fromColumn < 0)
                throw new SpreadhsheetExtensionsException($"Argument '{nameof(fromColumn)}' cannot be less than zero");

            if (toRow < 0)
                throw new SpreadhsheetExtensionsException($"Argument '{nameof(toRow)}' cannot be less than zero");

            if (toColumn < 0)
                throw new SpreadhsheetExtensionsException($"Argument '{nameof(toColumn)}' cannot be less than zero");
        }
    }
}

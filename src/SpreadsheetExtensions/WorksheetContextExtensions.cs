using SpreadsheetExtensions.Exceptions;
using SpreadsheetExtensions.Styles;
using SpreadsheetExtensions.Wrappers;
using System;

namespace SpreadsheetExtensions
{
    public static class WorksheetContextExtensions
    {
        public static IRangeWrapper GetVerticalRange(this WorksheetContext instance, int column, int fromRow, int toRow, bool setAsCurrent = true)
        {
            return instance.GetRange(fromRow, column, toRow, column, setAsCurrent);
        }

        public static IRangeWrapper GetHorizontalRange(this WorksheetContext instance, int row, int fromColumn, int toColumn, bool setAsCurrent = true)
        {
            return instance.GetRange(row, fromColumn, row, toColumn, setAsCurrent);
        }

        public static WorksheetContext Write(this WorksheetContext instance, string address, object value)
        {
            instance
                .GetRange(address)
                .SetValue(value);

            return instance;
        }

        public static WorksheetContext Write(this WorksheetContext instance, int row, int column, object value)
        {
            instance
                .GetRange(row, column)
                .SetValue(value);

            return instance;
        }

        public static WorksheetContext Write(this WorksheetContext instance, int fromRow, int fromColumn, int toRow, int toColumn, object value)
        {
            instance
                .GetRange(fromRow, fromColumn, toRow, toColumn)
                .SetValue(value);

            return instance;
        }

        public static WorksheetContext Merge(this WorksheetContext instance, bool isMerged) => instance.WrapOperation(() =>
            {
                instance.CurrentRange.Merge(isMerged);
            });

        public static WorksheetContext SetHorizontalAlignment(this WorksheetContext instance, SpreadsheetHorizontalAlignment alignment) => instance.WrapOperation(() =>
        {
            instance.CurrentRange.SetHorizontalAlignment(alignment);
        });

        public static WorksheetContext SetVerticalAlignment(this WorksheetContext instance, SpreadsheetVerticalAlignment alignment) => instance.WrapOperation(() =>
        {
            instance.CurrentRange.SetVerticalAlignment(alignment);
        });

        public static WorksheetContext SetFontBold(this WorksheetContext instance, bool isBold) => instance.WrapOperation(() =>
        {
            instance.CurrentRange.SetFontBold(isBold);
        });

        public static WorksheetContext SetBorderTop(this WorksheetContext instance, SpreadsheetBorder border) => instance.WrapOperation(() =>
        {
            instance.CurrentRange.SetBorderTop(border);
        });

        public static WorksheetContext SetBorderBottom(this WorksheetContext instance, SpreadsheetBorder border) => instance.WrapOperation(() =>
        {
            instance.CurrentRange.SetBorderBottom(border);
        });

        public static WorksheetContext SetBorderLeft(this WorksheetContext instance, SpreadsheetBorder border) => instance.WrapOperation(() =>
        {
            instance.CurrentRange.SetBorderLeft(border);
        });

        public static WorksheetContext SetBorderRight(this WorksheetContext instance, SpreadsheetBorder border) => instance.WrapOperation(() =>
        {
            instance.CurrentRange.SetBorderRight(border);
        });

        public static WorksheetContext SetBorder(this WorksheetContext instance, SpreadsheetBorder top, SpreadsheetBorder left, SpreadsheetBorder bottom, SpreadsheetBorder right) => instance.WrapOperation(() =>
        {
            instance.CurrentRange
                .SetBorderTop(top)
                .SetBorderBottom(bottom)
                .SetBorderLeft(left)
                .SetBorderRight(right);
        });

        public static WorksheetContext SetBorder(this WorksheetContext instance, SpreadsheetBorder all) => instance.SetBorder(all, all, all, all);

        public static WorksheetContext SetBorder(this WorksheetContext instance, SpreadsheetBorder topBottom, SpreadsheetBorder leftRight) => instance.SetBorder(topBottom, leftRight, topBottom, leftRight);

        public static WorksheetContext SetBorder(this WorksheetContext instance, SpreadsheetBorder top, SpreadsheetBorder leftRight, SpreadsheetBorder bottom) => instance.SetBorder(top, leftRight, bottom, leftRight);

        public static WorksheetContext SetFontItalic(this WorksheetContext instance, bool isItalic) => instance.WrapOperation(() =>
        {
            instance.CurrentRange.SetFontItalic(isItalic);
        });

        public static WorksheetContext SetFontUnderline(this WorksheetContext instance, bool isUnderline) => instance.WrapOperation(() =>
        {
            instance.CurrentRange.SetFontUnderline(isUnderline);
        });

        private static WorksheetContext WrapOperation(this WorksheetContext instance, Action operation)
        {
            try
            {
                operation?.Invoke();
                return instance;
            }
            catch (SpreadsheetNullRangeWrapperException ex)
            {
                throw new SpreadsheetFluentException("Cannot perform operation on current range which is not set", ex);
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

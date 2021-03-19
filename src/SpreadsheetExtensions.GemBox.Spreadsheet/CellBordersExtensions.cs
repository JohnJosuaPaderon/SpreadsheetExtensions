using GemBox.Spreadsheet;
using SpreadsheetExtensions.Styles;

namespace SpreadsheetExtensions
{
    public static class CellBordersExtensions
    {
        public static CellBorders Apply(this CellBorders instance, IndividualBorder side, SpreadsheetBorder border)
        {
            if (border.Style is not null)
                instance[side].LineStyle = SpreadsheetBorderStyleMapper.Map(border.Style.Value);

            if (border.Color is not null)
                instance[side].LineColor = border.Color.Value.ToDotNetColor();

            return instance;
        }
    }
}

using OfficeOpenXml.Style;
using SpreadsheetExtensions.Styles;

namespace SpreadsheetExtensions
{
    public static class ExcelBorderItemExtensions
    {
        public static ExcelBorderItem Apply(this ExcelBorderItem instance, SpreadsheetBorder border)
        {
            if (border.Style is not null)
                instance.Style = SpreadsheetBorderStyleMapper.Map(border.Style.Value);

            if (border.Color is not null)
                instance.Color.SetColor(border.Color.Value.ToDotNetColor());

            return instance;
        }
    }
}

using OfficeOpenXml.Style;

namespace SpreadsheetExtensions.Styles
{
    internal static class SpreadsheetBorderStyleMapper
    {
        public static ExcelBorderStyle Map(SpreadsheetBorderStyle style)
        {
            return style switch
            {
                SpreadsheetBorderStyle.None => ExcelBorderStyle.None,
                SpreadsheetBorderStyle.Thin => ExcelBorderStyle.Thin,
                SpreadsheetBorderStyle.Thick => ExcelBorderStyle.Thick,
                SpreadsheetBorderStyle.Dotted => ExcelBorderStyle.Dotted,
                SpreadsheetBorderStyle.Dashed => ExcelBorderStyle.Dashed,
                SpreadsheetBorderStyle.Double => ExcelBorderStyle.Double,
                _ => SpreadsheetLibraryDefaults.BorderStyle
            };
        }
    }
}

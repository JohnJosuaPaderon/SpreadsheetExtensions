using GemBox.Spreadsheet;
using SpreadsheetExtensions.Styles;
using System;
using System.IO;

namespace SpreadsheetExtensions.GemBox.Spreadsheet.SampleConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            Console.WriteLine("SpreadsheetExtensions.GemBox.Spreadsheet.SampleConsoleApp");

            var file = Path.Combine(Directory.GetCurrentDirectory(), "sample_gembox_spreadsheet.xlsx");

            if (!File.Exists(file))
            {
                Console.WriteLine($"The file '{file}' does not exists");
                return;
            }

            var excelFile = ExcelFile.Load(file);
            var worksheet = excelFile.Worksheets[0];

            worksheet.Write(0, 0, 4, 4, "SAMPLE TEXT")
                .Merge(true)
                .SetFontBold(true)
                .SetFontItalic(true)
                .SetFontUnderline(true)
                .SetBorderBottom(new SpreadsheetBorder
                {
                    Style = SpreadsheetBorderStyle.Thick,
                    Color = new Styles.SpreadsheetColor(255, 0, 0, 255)
                })
                .SetBorderRight(new SpreadsheetBorder
                {
                    Style = SpreadsheetBorderStyle.Dotted,
                    Color = Styles.SpreadsheetColor.FromDotNetColor(System.Drawing.Color.Green)
                })
                .SetVerticalAlignment(SpreadsheetVerticalAlignment.Center)
                .SetHorizontalAlignment(SpreadsheetHorizontalAlignment.Center);
            excelFile.Save(file);
        }
    }
}

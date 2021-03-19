using OfficeOpenXml;
using SpreadsheetExtensions.Styles;
using System;
using System.IO;

namespace SpreadsheetExtensions.EPPlus.SampleConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Console.WriteLine("SpreadsheetExtensions.EPPlus.SampleConsoleApp");

            var file = Path.Combine(Directory.GetCurrentDirectory(), "sample_epplus.xlsx");

            if (!File.Exists(file))
            {
                Console.WriteLine($"The file '{file}' does not exists");
                return;
            }

            using var package = new ExcelPackage(new FileInfo(file));
            using var workbook = package.Workbook;
            using var worksheet = workbook.Worksheets[0];

            worksheet.Write(1, 1, 5, 5, "SAMPLE TEXT")
                .Merge(true)
                .SetFontBold(true)
                .SetFontItalic(true)
                .SetFontUnderline(true)
                .SetBorderBottom(new SpreadsheetBorder
                {
                    Style = SpreadsheetBorderStyle.Thick,
                    Color = new SpreadsheetColor(255, 0, 0, 255)
                })
                .SetBorderRight(new SpreadsheetBorder
                {
                    Style = SpreadsheetBorderStyle.Dotted,
                    Color = SpreadsheetColor.FromDotNetColor(System.Drawing.Color.Green)
                })
                .SetVerticalAlignment(SpreadsheetVerticalAlignment.Center)
                .SetHorizontalAlignment(SpreadsheetHorizontalAlignment.Center);
            package.Save();
        }
    }
}

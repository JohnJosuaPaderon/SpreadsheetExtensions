using System.Drawing;

namespace SpreadsheetExtensions.Styles
{
    public struct SpreadsheetColor
    {
        public byte Alpha { get; }
        public byte Red { get; }
        public byte Green { get; }
        public byte Blue { get; }

        public SpreadsheetColor(byte red, byte green, byte blue, byte alpha)
        {
            Red = red;
            Green = green;
            Blue = blue;
            Alpha = alpha;
        }

        public Color ToDotNetColor()
        {
            return Color.FromArgb(Alpha, Red, Green, Blue);
        }

        public static SpreadsheetColor FromRgba(byte red, byte green, byte blue, byte alpha)
        {
            return new SpreadsheetColor(red, green, blue, alpha);
        }

        public static SpreadsheetColor FromDotNetColor(Color color)
        {
            return new SpreadsheetColor(color.R, color.G, color.B, color.A);
        }
    }
}

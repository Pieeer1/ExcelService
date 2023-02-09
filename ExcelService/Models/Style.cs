using ExcelService.Enums;
using System.Drawing;

namespace ExcelService.Models
{
    public class Style
    {
        public Style() { }
        public Style(Font? font, Color? color, double? fontSize)
        {
            Font = font;
            Color = color;
            FontSize = fontSize;
        }

        public Font? Font { get; set; }
        public Color? Color { get; set; }
        public double? FontSize { get; set; }

        public static Style Empty()
        { 
            return new Style();
        }
    }
}

using ExcelService.Enums;
using ExcelService.Models.Styles;
using System.Drawing;

namespace ExcelService.Models
{
    public class Style
    {
        public Style() { }
        public Style(Font? font = null, Color? color = null, double? fontSize = null, FontStyle? fontStyle = null, Color? textColor = null, Border? border = null)
        {
            Font = font;
            Color = color;
            FontSize = fontSize;
            FontStyle = fontStyle;
            TextColor = textColor;
            Border = border;
        }

        public Font? Font { get; set; }
        public Color? Color { get; set; }
        public double? FontSize { get; set; }
        public FontStyle? FontStyle { get; set; }
        public Color? TextColor { get; set; }
        public Border? Border { get; set; }

        public static Style Empty => new Style();
    }
}

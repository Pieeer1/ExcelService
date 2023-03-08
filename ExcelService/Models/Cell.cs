using ExcelService.Enums;
using ExcelService.Models.Styles;
using System.Drawing;
using System.Linq.Expressions;

namespace ExcelService.Models
{
    public class Cell
    {
        public Cell(string data, Style style)
        {
            Data = data;
            Style = style;
        }
        public Cell(string data, Font? font = null, Color? color = null, double? fontSize = null, FontStyle? fontStyle = null, Color? textColor = null, Border? border = null)
        {
            Data = data;
            Style = new Style(font, color, fontSize, fontStyle, textColor, border);
        }

        public string Data { get; private set; }
        public Style Style { get; private set; }

        public void SetCell(Cell cell)
        {
            Data = cell.Data;
            Style = cell.Style;
        }
        public void SetStyle(Style style)
        {
            Style = style;
        }
        public void SetStyle(Font? font = null, Color? color = null, double? fontSize = null, FontStyle? fontStyle = null, Color? textColor = null, Border? border = null)
        {
            Style = new Style(font, color, fontSize, fontStyle, textColor, border);
        }
    }
}

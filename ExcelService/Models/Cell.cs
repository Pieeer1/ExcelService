using ExcelService.Enums;
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
        public Cell(string data, Font? font = null, Color? color = null, double? fontSize = null)
        {
            Data = data;
            Style = new Style(font, color, fontSize);
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
        public void SetStyle(Font? font = null, Color? color = null, double? fontSize = null)
        {
            Style = new Style(font, color, fontSize);
        }
    }
}

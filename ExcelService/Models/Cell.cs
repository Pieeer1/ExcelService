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
            Font = style.Font;
            Color = style.Color;
            FontSize = style.FontSize;
        }
        public Cell(string data, Font? font = null, Color? color = null, double? fontSize = null)
        {
            Data = data;
            Font = font;
            Color = color;
            FontSize = fontSize;
        }

        public string Data { get; private set; }
        public Font? Font { get; private set; }
        public Color? Color { get; private set; }
        public double? FontSize { get; private set; }

        public void SetCell(Cell cell)
        { 
            Data = cell.Data;
            Font = cell.Font;
            Color = cell.Color;
            FontSize= cell.FontSize;
        }
        public void SetStyle(Style style)
        {
            Font = style.Font;
            Color = style.Color;
            FontSize = style.FontSize;
        }
        public void SetStyle(Font? font = null, Color? color = null, double? fontSize = null)
        {
            Font = font;
            Color = color;
            FontSize = fontSize;
        }
    }
}

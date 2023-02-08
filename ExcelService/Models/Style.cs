using ExcelService.Enums;
using System.Drawing;

namespace ExcelService.Models
{
    public class Style
    {
        public Font? Font { get; set; }
        public Color? Color { get; set; }
        public double? FontSize { get; set; }
    }
}

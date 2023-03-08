using System.Drawing;

namespace ExcelService.Models.Styles
{
    public class Border
    {
        public Border(uint thickness, Color? color)
        {
            Thickness = thickness;
            Color = color;
        }

        public uint Thickness { get; set; }
        public Color? Color { get; set; }
    }
}

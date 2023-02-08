using System.Data;
using System.Reflection;
namespace ExcelService.Models
{
    public class Row
    {
        public Row(IEnumerable<Cell> cells, IEnumerable<Style>? styles = null)
        {
            Cells = cells;
            Styles = styles;
        }

        public IEnumerable<Cell> Cells { get; private set; }
        public IEnumerable<Style>? Styles { get; private set; }
        public static Row GenerateRowFromObject<T>(T obj, IEnumerable<Style>? styles = null)
        {
            List<Cell> cells = new List<Cell>();

            for (int i = 0; i < typeof(T).GetProperties().Where(p => p.CanRead).Count(); i++)
            {
                cells.Add(new Cell(typeof(T).GetProperties().Where(p => p.CanRead).ElementAt(i).GetValue(obj, null)?.ToString() ?? string.Empty, styles.ElementAt(i)));
            }

            return new Row(cells);
        }
    }
}

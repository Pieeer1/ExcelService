using System.Data;

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
                int count = typeof(T).GetProperties().Where(p => p.CanRead).Count();
                List<string> names = typeof(T).GetProperties().Where(p => p.CanRead).Select(x => x.Name).ToList();
                try
                {
                    cells.Add(new Cell(typeof(T).GetProperties().Where(p => p.CanRead).ElementAt(i).GetValue(obj, null)?.ToString() ?? string.Empty, styles?.ElementAt(i) ?? new Style()));
                }
                catch (Exception ex)
                {
                    return new Row(new List<Cell>() { new Cell(Convert.ToString(obj) ?? string.Empty) });
                }
            }

            return new Row(cells);
        }
    }
}

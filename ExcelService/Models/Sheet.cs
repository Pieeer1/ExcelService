using System.Reflection;

namespace ExcelService.Models
{
    public class Sheet
    {
        private Sheet(Row headerRow, IEnumerable<Row> rows, string name)
        {
            HeaderRow = headerRow;
            Rows = rows;
            Name = name;
        }
        public Row HeaderRow { get; private set; }
        public IEnumerable<Row> Rows { get; private set; }
        public string Name { get; private set; }


        public static Sheet GetSheetFromDataSet<T>(IEnumerable<T> objects, IEnumerable<IEnumerable<Style>>? styles, string sheetName)
        {
            List<Row> rows = new List<Row>();
            string firstProperty = string.Empty;
            List<Cell> cells = new List<Cell>();
            foreach (PropertyInfo property in typeof(T).GetProperties().Where(p => p.CanRead))
            {
                if (property.Name == firstProperty)
                {
                    break;
                }
                firstProperty = firstProperty == string.Empty ? property.Name : firstProperty;
                cells.Add(new Cell(property.Name));
            }
            for (int i =0; i < objects.Count(); i++)
            {
                rows.Add(Row.GenerateRowFromObject(objects.ElementAt(i), styles?.ElementAt(i+1)));
            }

            return new Sheet(new Row(cells, styles?.ElementAt(0)), rows, sheetName);
        }
    }
}

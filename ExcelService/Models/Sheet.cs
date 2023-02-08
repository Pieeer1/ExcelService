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
        public Cell this[uint x, uint y]
        {
            get => GetCell(x, y);
            set => SetCell(x, y, value);
        }
        public Cell this[char x, uint y]
        {
            get => GetCell(x, y);
            set => SetCell(x, y, value);
        }

        public Cell GetCell(uint x, uint y)
        {
            return Rows.ElementAt((int)x).Cells.ElementAt((int)y);
        }
        public Cell GetCell(char x, uint y)
        {
            if ((int)x < 65 || (int)x > 90)
            {
                throw new InvalidOperationException("First Argument must be argument of type A-Z");
            }
            return Rows.ElementAt((int)y).Cells.ElementAt((int)(x - 65));
        }

        public void SetCell(uint x, uint y, Cell value)
        {
            Rows.ElementAt((int)x).Cells.ElementAt((int)y).SetCell(value);
        }
        public void SetCell(char x, uint y, Cell value)
        {
            if ((int)x < 65 || (int)x > 90)
            {
                throw new InvalidOperationException("First Argument must be argument of type A-Z");
            }
            Rows.ElementAt((int)y).Cells.ElementAt((int)(x - 65)).SetCell(value);
        }
    }
}

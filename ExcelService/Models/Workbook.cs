using ExcelService.Enums;
using ExcelService.Extensions;
using System.Drawing;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelService.Models
{
    public class Workbook
    {
        private Workbook(IEnumerable<Sheet> sheets, string? name)
        {
            Sheets = sheets;
            Name = name ?? "Workbook";
        }
        public IEnumerable<Sheet> Sheets { get; set; }
        public string Name { get; set; }

        /// <summary>
        /// Creating a WorkBook with a single Sheet
        /// </summary>
        /// <typeparam name="T">Type of Object to Convert</typeparam>
        /// <param name="objects">Enumerable to Convert into an excel table</param>
        /// <param name="styles">Styles to add: optional, must be at least as long as the actual excel sheets if you want to declare it at startup</param>
        /// <param name="sheetName">optional name of sheet</param>
        /// <returns>Returns a new Workbook object</returns>
        public static Workbook GetWorkbookFromDataSet<T>(IEnumerable<T> objects, IEnumerable<IEnumerable<Style>>? styles = null, string? workbookName = null, string? sheetName = null) 
        {
            if (!typeof(T).IsClass) { throw new TargetException("Must be a Reference Type to Create an Excel Sheet From a Dynamic Type"); }//currently only supports objects will change to take structs later
            return new Workbook(new List<Sheet>() { Sheet.GetSheetFromDataSet(objects, styles, sheetName ?? "Sheet") }, workbookName);
        }
        public void AddSheetToWorkBook(Sheet sheet)
        {
            List<Sheet> newSheetList = Sheets.ToList();
            newSheetList.Add(sheet);
            Sheets = newSheetList;
        }
        public void AddWorkbookSheetsToWorkBook(Workbook workbook)
        {
            List<Sheet> newSheetList = Sheets.ToList();
            foreach (Sheet sheet in workbook.Sheets)
            {
                newSheetList.Add(sheet);
            }
            Sheets = newSheetList;
        }
        public void StyleWhere(string header, Func<string, bool> operation, Style style)
        {
            foreach (Sheet sheet in Sheets)
            {
                int headerPosition = sheet.HeaderRow.Cells.IndexOf(sheet.HeaderRow.Cells.FirstOrDefault(x => x.Data == header));
                if (headerPosition == -1) { continue; }
                foreach (Row row in sheet.Rows)
                {
                    Cell cell = row.Cells.ElementAt(headerPosition);
                    if (operation.Invoke(cell.Data))
                    {
                        cell.SetStyle(style);
                    }
                }
            }
        }
        public void StyleWhere(string header, Func<string, bool> operation, Font? font, Color? color, double? fontSize)
        {
            foreach (Sheet sheet in Sheets)
            {
                int headerPosition = sheet.HeaderRow.Cells.IndexOf(sheet.HeaderRow.Cells.FirstOrDefault(x => x.Data == header));
                if (headerPosition == -1) { continue; }
                foreach (Row row in sheet.Rows)
                {
                    Cell cell = row.Cells.ElementAt(headerPosition);
                    if (operation.Invoke(cell.Data))
                    {
                        cell.SetStyle(font, color, fontSize);
                    }
                }
            }
        }
        public void StyleRowWhere<T>(Expression<Func<T, bool>> expression, Style style)
        {
            foreach (Sheet sheet in Sheets)
            {
                sheet.StyleRowWhere(expression, style);
            }
        }
        public Cell this[uint sheet, uint x, uint y]
        {
            get => GetCell(sheet, x, y);
            set => SetCell(sheet, x, y, value);
        }
        public Cell this[uint sheet, char x, uint y]
        {
            get => GetCell(sheet, x, y);
            set => SetCell(sheet, x, y, value);
        }

        public Cell GetCell(uint sheet, uint x, uint y)
        {
            if (sheet >= Sheets.Count())
            {
                throw new IndexOutOfRangeException($"Cannot find sheet {sheet}");
            }

            return Sheets.ElementAt((int)sheet)[x, y];
        }
        public Cell GetCell(uint sheet, char x, uint y)
        {
            if ((int)x < 65 || (int)x > 90)
            {
                throw new InvalidOperationException("First Argument must be argument of type A-Z");
            }
            if (sheet >= Sheets.Count())
            {
                throw new IndexOutOfRangeException($"Cannot find sheet {sheet}");
            }

            return Sheets.ElementAt((int)sheet)[x, y];
        }
        public void SetCell(uint sheet, uint x, uint y, Cell value)
        {
            if (sheet >= Sheets.Count())
            {
                throw new IndexOutOfRangeException($"Cannot find sheet {sheet}");
            }

            Sheets.ElementAt((int)sheet)[x, y] = value;
        }
        public void SetCell(uint sheet, char x, uint y, Cell value)
        {
            if ((int)x < 65 || (int)x > 90)
            {
                throw new InvalidOperationException("First Argument must be argument of type A-Z");
            }
            if (sheet >= Sheets.Count())
            {
                throw new IndexOutOfRangeException($"Cannot find sheet {sheet}");
            }

            Sheets.ElementAt((int)sheet)[x, y] = value; 
        }
        public IEnumerable<Style> GetDistinctStyles()
        {
            HashSet<Style> styles = new HashSet<Style>
            {
                Style.Empty(), // default
                new Style(Font.Arial ,Color.Gray, 409) // need to add a second default to the registry to load in
            };
            foreach (Sheet sheets in Sheets)
            {
                styles.UnionWith(sheets.HeaderRow.Cells.Select(x => x.Style).Distinct()); //add header columns
                foreach (Row row in sheets.Rows)
                {
                    styles.UnionWith(row.Cells.Select(x => x.Style).Distinct());
                }
            }

            return styles.Where(x => x != null && !(x.Color == null && x.Font == null && x.FontSize == null));
        }
    }
}

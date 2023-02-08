using ExcelService.Enums;
using ExcelService.Extensions;
using System.Drawing;

namespace ExcelService.Models
{
    public class Workbook
    {
        private Workbook(IEnumerable<Sheet> sheets, string name = "Workbook")
        {
            Sheets = sheets;
            Name = name;
        }
        public IEnumerable<Sheet> Sheets { get; set; }
        public string Name { get; set; }


        public static Workbook GetWorkbookFromDataSet<T>(IEnumerable<IEnumerable<T>> objects, IEnumerable<IEnumerable<IEnumerable<Style>>>? styles = null, string[]? sheetNames = null)
        {
            List<Sheet> sheets = new List<Sheet>();
            for (int i = 0; i < objects.Count();  i++)
            {
                sheets.Add(Sheet.GetSheetFromDataSet(objects.ElementAt(i), styles?.ElementAt(i), sheetNames?[i] ?? $"Sheet {i+1}"));
            }

            return new Workbook(new List<Sheet>(sheets));
        }
        public static Workbook GetWorkbookFromDataSet<T>(IEnumerable<T> objects, IEnumerable<IEnumerable<Style>>? styles = null, string[]? sheetNames = null)
        {
            return new Workbook(new List<Sheet>() { Sheet.GetSheetFromDataSet(objects, styles, sheetNames?[0] ?? "Sheet") });
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

    }
}

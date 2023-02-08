using ExcelService.Models;

namespace ExcelService.Interfaces
{
    public interface IExcel
    {
        public void GenerateNewWorkBook(Workbook workbook);
        public void GenerateNewWorkBook<T>(IEnumerable<IEnumerable<T>> objects, IEnumerable<IEnumerable<IEnumerable<Style>>>? styles = null, string[]? sheetNames = null);
        public void GenerateNewWorkBook<T>(IEnumerable<T> objects, IEnumerable<IEnumerable<Style>>? styles = null, string? sheetName = null);
        public Workbook? this[Workbook workbook]
        {
            get => GetWorkbook(workbook);
        }
        public Workbook? this[string workbookName]
        {
            get => GetWorkbook(workbookName);
        }
        public Workbook? GetWorkbook(Workbook workbook);
        public Workbook? GetWorkbook(string workbookName);
        public void GetExcelFromWorkBook(Stream stream, Workbook workbook);
        public void SaveExcelFileFromWorkbook(string fileName, Workbook workbook);
        public void RemoveWorkbook(Workbook workbook);
        public void RemoveWorkbook(string workbookName);
    }
}

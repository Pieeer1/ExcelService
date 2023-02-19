using ExcelService.Models;

namespace ExcelService.Interfaces
{
    public interface IExcel
    {
        public void GenerateNewWorkBook(Workbook workbook);
        public void GenerateNewWorkBook<T>(IEnumerable<T> objects, IEnumerable<IEnumerable<Style>>? styles = null, string? workbookName = null, string? sheetName = null);
        public Workbook this[uint index]
        {
            get => GetWorkbook(index);
        }
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
        public Workbook GetWorkbook(uint index);
        public Stream GetExcelFromWorkBook(Workbook workbook);
        public void SaveExcelFileFromWorkbook(string fileName, Workbook workbook);
        public Workbook GetWorkbookFromExcelFile(string filePath);
        public Workbook GetWorkbookFromExcelFile(Stream stream);
        public void RemoveWorkbook(Workbook workbook);
        public void RemoveWorkbook(string workbookName);
        public void CombineWorkbooks(Workbook baseWorkbook, Workbook additonalWorkbook);
        public int WorkbookCount();
        public void ClearWorkbooks();
    }
}

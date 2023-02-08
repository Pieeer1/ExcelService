using ExcelService.Interfaces;
using ExcelService.Models;

namespace ExcelService
{
    public class Excel : IExcel
    {
        public readonly HashSet<Workbook> Workbooks;
        private Excel() 
        {
            Workbooks = new HashSet<Workbook>();   
        }


        public void GenerateNewWorkBook(Workbook workbook) => Workbooks.Add(workbook);
        public void GenerateNewWorkBook<T>(IEnumerable<IEnumerable<T>> objects, IEnumerable<IEnumerable<IEnumerable<Style>>>? styles = null, string[]? sheetNames = null) => Workbooks.Add(Workbook.GetWorkbookFromDataSet(objects, styles, sheetNames));
        public void GenerateNewWorkBook<T>(IEnumerable<T> objects, IEnumerable<IEnumerable<Style>>? styles = null, string? sheetName = null) => Workbooks.Add(Workbook.GetWorkbookFromDataSet(objects, styles, sheetName));

        public Workbook? this[Workbook workbook]
        {
            get => GetWorkbook(workbook);
        }
        public Workbook? this[string workbookName]
        {
            get => GetWorkbook(workbookName);
        }
        public Workbook? GetWorkbook(Workbook workbook)
        {
            return Workbooks.FirstOrDefault(workbook);
        }
        public Workbook? GetWorkbook(string workbookName)
        { 
            return Workbooks.FirstOrDefault(x => x.Name == workbookName);
        }
        public void GetExcelFromWorkBook(Stream stream, Workbook workbook)
        {

            stream.Position = 0; // keep before eof
            throw new NotImplementedException();
        }
        public void SaveExcelFileFromWorkbook(string fileName, Workbook workbook)
        {


            throw new NotImplementedException();
        }


        public void RemoveWorkbook(Workbook workbook)
        {
            Workbooks.Remove(workbook);
        }
        public void RemoveWorkbook(string workbookName)
        {
            RemoveWorkbook(Workbooks.First(x => x.Name == workbookName));
        }
    }
}
using ExcelService.Interfaces;
using ExcelService.Models;

namespace ExcelService
{
    public class Excel : IExcel
    {
        private readonly HashSet<Workbook> Workbooks;
        public Excel() 
        {
            Workbooks = new HashSet<Workbook>();   
        }

        public void GenerateNewWorkBook(Workbook workbook) => Workbooks.Add(workbook);
        public void GenerateNewWorkBook<T>(IEnumerable<T> objects, IEnumerable<IEnumerable<Style>>? styles = null, string? sheetName = null) => Workbooks.Add(Workbook.GetWorkbookFromDataSet(objects, styles, sheetName));

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
        public Workbook GetWorkbook(uint index)
        {
            return Workbooks.ElementAt((int)index);
        }
        public Workbook? GetWorkbook(Workbook workbook)
        {
            return Workbooks.FirstOrDefault(workbook);
        }
        public Workbook? GetWorkbook(string workbookName)
        { 
            return Workbooks.FirstOrDefault(x => x.Name == workbookName);
        }
        public Stream GetExcelFromWorkBook(Workbook workbook)
        {
            return OpenXMLService.OpenXMLService.GetXLSXStreamFromWorkbook(workbook);
        }
        public void SaveExcelFileFromWorkbook(string fileName, Workbook workbook)
        {
            MemoryStream ms = new MemoryStream();
            OpenXMLService.OpenXMLService.GetXLSXStreamFromWorkbook(workbook).CopyTo(ms);
            File.WriteAllBytes(fileName, ms.ToArray());
        }
        public void CombineWorkbooks(Workbook baseWorkbook, Workbook additonalWorkbook)
        { 
            baseWorkbook.AddWorkbookSheetsToWorkBook(additonalWorkbook);
            RemoveWorkbook(additonalWorkbook);
        }

        public void RemoveWorkbook(Workbook workbook)
        {
            Workbooks.Remove(workbook);
        }
        public void RemoveWorkbook(string workbookName)
        {
            RemoveWorkbook(Workbooks.First(x => x.Name == workbookName));
        }
        public int WorkbookCount()
        {
            return Workbooks.Count;
        }
    }
}
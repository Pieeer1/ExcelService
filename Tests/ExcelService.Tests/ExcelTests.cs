using ExcelService.Models;
using ExcelService.Tests.Objects;
using Xunit.Sdk;

namespace ExcelService.Tests
{
    public class ExcelTests
    {
        private readonly Excel excel;
        private readonly Workbook noStyleWorkbook;
        private readonly Workbook styleWorkbook;
        public ExcelTests()
        {
            excel = new Excel();
            noStyleWorkbook = TestObjects.noStyleWorkbook;
            styleWorkbook = TestObjects.styleWorkbook;
        }

        [Fact]
        public void TestGenerateExcel()
        {
            excel.GenerateNewWorkBook(noStyleWorkbook);
            excel.GenerateNewWorkBook(styleWorkbook);
            excel.GenerateNewWorkBook(Workbook.GetWorkbookFromDataSet(new List<TestClass>() { new TestClass() { TestString = "a"} }, null, null)); //currently only supports objects will change to take structs later
        }
        private class TestClass
        {
            public string? TestString { get; set; }
        }
    }
}

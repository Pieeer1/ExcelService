using ExcelService.Models;
using System.Drawing;
using Xunit.Sdk;

namespace ExcelService.Tests
{
    public class ExcelTests
    {
        private readonly Excel excel;
        private Workbook noStyleWorkbook = Workbook.GetWorkbookFromDataSet(new List<TestObject>()
        {
            new TestObject()
            {
                A = "A String",
                B = "B String",
                C = 1,
                D = 2,
            },
            new TestObject()
            {
                A = "A String1",
                B = "B String2",
                C = 3,
                D = 4,
            }
        });
        private Workbook styleWorkbook = Workbook.GetWorkbookFromDataSet(new List<TestObject>()
        {
            new TestObject()
            {
                A = "A String",
                B = "B String",
                C = 1,
                D = 2,
            },
            new TestObject()
            {
                A = "A String1",
                B = "B String2",
                C = 3,
                D = 4,
            }
        },
        new List<List<Style>>()
        {
            new List<Style>() // first one is always header row
            {
                new Style()
                {
                    Color = Color.Green,
                    Font = Enums.Font.Calibri,
                    FontSize= 12,
                },
                new Style()
                {
                    Color = Color.Red,
                    Font = Enums.Font.Arial,
                    FontSize= 8,
                },
                Style.Empty(),
                Style.Empty()
            },
            new List<Style>()
            {
                new Style()
                {
                    Color = Color.Green,
                    Font = Enums.Font.Calibri,
                    FontSize= 12,
                },
                new Style()
                {
                    Color = Color.Red,
                    Font = Enums.Font.Arial,
                    FontSize= 8,
                },
                Style.Empty(),
                Style.Empty()
            },
            new List<Style>()
            {
                new Style()
                {
                    Color = Color.Green,
                    Font = Enums.Font.Calibri,
                    FontSize= 12,
                },
                new Style()
                {
                    Color = Color.Red,
                    Font = Enums.Font.Arial,
                    FontSize= 8,
                },
                Style.Empty(),
                Style.Empty()
            },
        }
        );

        public ExcelTests()
        {
            excel = new Excel();

            excel.GenerateNewWorkBook(noStyleWorkbook);
            excel.GenerateNewWorkBook(styleWorkbook);
            excel.GenerateNewWorkBook(Workbook.GetWorkbookFromDataSet(new List<TestClass>() { new TestClass() { TestString = "a" } }, null, null));
            excel.GenerateNewWorkBook(Workbook.GetWorkbookFromDataSet(new List<string>() { "a", "b", "c", "d" }, null, "stringBook", "stringSheet"));

        }

        [Fact]
        public void TestGenerateExcelGetFromIndex()
        {
            Workbook? noStyleBook = excel[0];
            Workbook? styleBook = excel[1];
            Workbook testBook = excel[2];
            Workbook? stringBook = excel["stringBook"];

            Assert.NotNull(noStyleBook);
            Assert.NotNull(styleBook);
            Assert.NotNull(testBook);
            Assert.NotNull(stringBook);
        }
        [Fact]
        public void TestAddExcelSheets()
        {
            excel.CombineWorkbooks(excel[0], excel[1]);
            excel.CombineWorkbooks(excel[1], excel[2]);

            Assert.Equal(2, excel.WorkbookCount());
            Assert.Equal(2, excel[0].Sheets.Count());
            Assert.Equal(2, excel[1].Sheets.Count());
            Assert.Throws<ArgumentOutOfRangeException>(() => excel[2]);
        }
        [Fact]

        public void TestExcelSheetDownload()
        {
            Workbook workbook = excel.GetWorkbookFromExcelFile("../../../Examples/test.xlsx");

            Assert.Equal("a", workbook[0, 'A', 1].Data);
            Assert.Equal("Column2", workbook[0, 0, 1].Data);
        }
        [Fact]
        public void TestExcelSheetDownloadFromStream()
        {
            Workbook workbook = excel.GetWorkbookFromExcelFile(new MemoryStream(File.ReadAllBytes("../../../Examples/test.xlsx")));

            Assert.Equal("a", workbook[0, 'A', 1].Data);
            Assert.Equal("Column2", workbook[0, 0, 1].Data);
        }
        private class TestClass
        {
            public string? TestString { get; set; }
        }
        private class TestObject
        {
            public string? A { get; set; }
            public string? B { get; set; }
            public int C { get; set; }
            public int D { get; set; }
        }
    }
}

using ExcelService;
using ExcelService.Models;
using System.Drawing;

namespace ExcelService.IntegrationTests
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Excel excel = new Excel();

            excel.GenerateNewWorkBook(new List<TestClass>()
            {
                new TestClass("a", "b", "c", "d", "e", "f", "g", 1, new DateTime(1999, 12, 08)),
                new TestClass("h", "i", "j", "k", "l", "m", "m", 2, new DateTime(1998, 11, 07)),
                new TestClass("n", "l", "o", "p", "q", "r", "s", 2, new DateTime(1998, 11, 07)),
            },
            new List<List<Style>>()
            { 
                new List<Style>()
                { 
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                },
                new List<Style>()
                {
                    new Style(Enums.Font.Calibri, Color.Red, 8),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    new Style(Enums.Font.Calibri, Color.Red, 8),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                },
                new List<Style>()
                {
                    new Style(Enums.Font.Calibri, Color.Pink, 8),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    new Style(Enums.Font.Arial, Color.Green, 20),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                },
                new List<Style>()
                {
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                    new Style(Enums.Font.Arial, Color.Red, 25),
                    new Style(Enums.Font.Arial, Color.Yellow, 48),
                    Style.Empty(),
                    Style.Empty(),
                    Style.Empty(),
                },
            },
            "TestWorkbook",
            "TestSheet");

            excel["TestWorkbook"]?.StyleRowWhere<TestClass>(x => x.Column1 == "a", new Style(Enums.Font.Arial, Color.Red, 65));


            excel["TestWorkbook"]?.StyleRowWhere<TestClass>(x => x.Column1 == "n", new Style(Enums.Font.Arial, Color.Red, 25));

            excel.SaveExcelFileFromWorkbook("../../../test.xlsx", excel["TestWorkbook"] ?? throw new NullReferenceException("Invalid Container"));
            
        }

        public class TestClass
        {
            public TestClass(string? column1, string? column2, string? column3, string? column4, string? column5, string? column6, string? column7, int column8, DateTime? column9)
            {
                Column1 = column1;
                Column2 = column2;
                Column3 = column3;
                Column4 = column4;
                Column5 = column5;
                Column6 = column6;
                Column7 = column7;
                Column8 = column8;
                Column9 = column9;
            }

            public string? Column1 { get; set; }
            public string? Column2 { get; set; }
            public string? Column3 { get; set; }
            public string? Column4 { get; set; }
            public string? Column5 { get; set; }
            public string? Column6 { get; set; }
            public string? Column7 { get; set; }
            public int Column8 { get; set; }
            public DateTime? Column9 { get; set; }
        }



    }
}
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

            Console.WriteLine("Welcome to The ExcelService Integrated Test Environment");
            Console.WriteLine("=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=");
            Console.WriteLine("Select Between the Following 2 options: ");
            Console.WriteLine("1 - Create a workbook example");
            Console.WriteLine("2 - Read a workbook example");
            string? response = Console.ReadLine();

            int responseInt = ParseReader(response).Invoke();
            if (responseInt == 1)
            {

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
                        new Style(textColor: Color.Red, border: new Models.Styles.Border(1, Color.Black)),
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

                excel["TestWorkbook"]?.StyleRowWhere<TestClass>(x => x.Column1 == "a", new Style(Enums.Font.Arial, Color.Green, 65, Enums.FontStyle.Underline, Color.AliceBlue, new Models.Styles.Border(6, Color.Black)));

                excel["TestWorkbook"]?.StyleRowWhere<TestClass>(x => x.Column1 == "n", new Style(Enums.Font.Arial, Color.Red, 25));

                excel["TestWorkbook"]?.StyleCellWhere<TestClass>(x => x.Column2 == "b", new Style(Enums.Font.Calibri, Color.Aqua, 30));

                excel.SaveExcelFileFromWorkbook("../../../test.xlsx", excel["TestWorkbook"] ?? throw new NullReferenceException("Invalid Container"));

            }
            else if (responseInt == 2)
            {
                Workbook workbook = excel.GetWorkbookFromExcelFile("../../../test.xlsx");

                Console.WriteLine(workbook[0, 'A', 1].Data);
                Console.WriteLine(workbook[0, 'A', 2].Data);
                Console.WriteLine(workbook[0, 'A', 3].Data);
            }
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

        private static Func<int> ParseReader(string? rawInput)
        {
            return rawInput switch
            {
                "1" => () => { return 1;},
                "2" => () => { return 2;},
                _ => () => 
                { 
                    Console.WriteLine("Invalid Input, Please Input 1 or 2:");
                    string? newInput = Console.ReadLine(); 
                    return ParseReader(newInput).Invoke(); 
                }
            };
        }



    }
}
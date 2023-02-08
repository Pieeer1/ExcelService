using ExcelService.Models;

namespace ExcelService.Tests
{
    public class WorkbookTests
    {


        [Fact]
        public void TestCreationFromObjectDefaultNames()
        {
            Workbook workbook = Workbook.GetWorkbookFromDataSet(new List<List<TestObject>>()
            {
                new List<TestObject>()
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
                }
            });
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
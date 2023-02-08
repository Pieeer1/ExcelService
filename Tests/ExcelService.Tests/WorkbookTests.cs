using ExcelService.Models;

namespace ExcelService.Tests
{
    public class WorkbookTests
    {


        [Fact]
        public void TestCreationFromObjectDefaultNames()
        {
            Workbook workbook = Workbook.GetWorkbookFromDataSet(new List<TestObject>()
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

            Assert.Equal("A String", workbook[0, 0, 0].Data);
            Assert.Equal("B String", workbook[0, 0, 1].Data);
            Assert.Equal("1", workbook[0, 0, 2].Data);
            Assert.Equal("2", workbook[0, 0, 3].Data);
            Assert.Equal("A String", workbook[0, 'A', 0].Data);
            Assert.Equal("B String", workbook[0, 'B', 0].Data);
            Assert.Equal("1", workbook[0, 'C', 0].Data);
            Assert.Equal("2", workbook[0, 'D', 0].Data);
            Assert.Equal("A String1", workbook[0, 1, 0].Data);
            Assert.Equal("B String2", workbook[0, 1, 1].Data);
            Assert.Equal("3", workbook[0, 1, 2].Data);
            Assert.Equal("4", workbook[0, 1, 3].Data);
            Assert.Equal("A String1", workbook[0, 'A', 1].Data);
            Assert.Equal("B String2", workbook[0, 'B', 1].Data);
            Assert.Equal("3", workbook[0, 'C', 1].Data);
            Assert.Equal("4", workbook[0, 'D', 1].Data);

            Assert.Throws<IndexOutOfRangeException>(() => workbook[1, 1, 1]);
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook[0, 'A', 2]);
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook[0, 2, 0]);
            Assert.Throws<InvalidOperationException>(() => workbook[0, 'a', 0]);
            Assert.Throws<InvalidOperationException>(() => workbook[0, '~', 0]);
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
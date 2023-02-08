using ExcelService.Models;
using System.Drawing;

namespace ExcelService.Tests
{
    public class WorkbookTests
    {
        private readonly Workbook noStyleWorkbook;
        private readonly Workbook styleWorkbook;

        public WorkbookTests() 
        {
            noStyleWorkbook = Workbook.GetWorkbookFromDataSet(new List<TestObject>()
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
            styleWorkbook = Workbook.GetWorkbookFromDataSet(new List<TestObject>()
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
        }


        [Fact]
        public void TestGetCells()
        {
            Assert.Equal("A String", noStyleWorkbook[0, 0, 0].Data);
            Assert.Equal("B String", noStyleWorkbook[0, 0, 1].Data);
            Assert.Equal("1", noStyleWorkbook[0, 0, 2].Data);
            Assert.Equal("2", noStyleWorkbook[0, 0, 3].Data);
            Assert.Equal("A String", noStyleWorkbook[0, 'A', 0].Data);
            Assert.Equal("B String", noStyleWorkbook[0, 'B', 0].Data);
            Assert.Equal("1", noStyleWorkbook[0, 'C', 0].Data);
            Assert.Equal("2", noStyleWorkbook[0, 'D', 0].Data);
            Assert.Equal("A String1", noStyleWorkbook[0, 1, 0].Data);
            Assert.Equal("B String2", noStyleWorkbook[0, 1, 1].Data);
            Assert.Equal("3", noStyleWorkbook[0, 1, 2].Data);
            Assert.Equal("4", noStyleWorkbook[0, 1, 3].Data);
            Assert.Equal("A String1", noStyleWorkbook[0, 'A', 1].Data);
            Assert.Equal("B String2", noStyleWorkbook[0, 'B', 1].Data);
            Assert.Equal("3", noStyleWorkbook[0, 'C', 1].Data);
            Assert.Equal("4", noStyleWorkbook[0, 'D', 1].Data);
        }
        [Fact]
        public void TestSetCells()
        {
            noStyleWorkbook[0, 0, 0] = new Cell("some data", new Style());
            noStyleWorkbook[0, 'C', 0] = new Cell("99", new Style());
            noStyleWorkbook[0, 1, 0] = new Cell("some different data", new Style());
            noStyleWorkbook[0,'C', 1] = new Cell("53", new Style());

            Assert.Equal("some data", noStyleWorkbook[0, 0, 0].Data);
            Assert.Equal("B String", noStyleWorkbook[0, 0, 1].Data);
            Assert.Equal("99", noStyleWorkbook[0, 0, 2].Data);
            Assert.Equal("2", noStyleWorkbook[0, 0, 3].Data);
            Assert.Equal("some data", noStyleWorkbook[0, 'A', 0].Data);
            Assert.Equal("B String", noStyleWorkbook[0, 'B', 0].Data);
            Assert.Equal("99", noStyleWorkbook[0, 'C', 0].Data);
            Assert.Equal("2", noStyleWorkbook[0, 'D', 0].Data);
            Assert.Equal("some different data", noStyleWorkbook[0, 1, 0].Data);
            Assert.Equal("B String2", noStyleWorkbook[0, 1, 1].Data);
            Assert.Equal("53", noStyleWorkbook[0, 1, 2].Data);
            Assert.Equal("4", noStyleWorkbook[0, 1, 3].Data);
            Assert.Equal("some different data", noStyleWorkbook[0, 'A', 1].Data);
            Assert.Equal("B String2", noStyleWorkbook[0, 'B', 1].Data);
            Assert.Equal("53", noStyleWorkbook[0, 'C', 1].Data);
            Assert.Equal("4", noStyleWorkbook[0, 'D', 1].Data);
        }
        [Fact]

        public void TestRanges()
        {
            Assert.Throws<IndexOutOfRangeException>(() => noStyleWorkbook[1, 1, 1]);
            Assert.Throws<ArgumentOutOfRangeException>(() => noStyleWorkbook[0, 'A', 2]);
            Assert.Throws<ArgumentOutOfRangeException>(() => noStyleWorkbook[0, 2, 0]);
            Assert.Throws<InvalidOperationException>(() => noStyleWorkbook[0, 'a', 0]);
            Assert.Throws<InvalidOperationException>(() => noStyleWorkbook[0, '~', 0]);
        }
        [Fact]
        public void TestGetDistinctColors()
        {
            IEnumerable<Color?> colors = styleWorkbook.GetDistinctColors().Where(x => x is not null);

            Assert.NotEmpty(colors);
            Assert.Equal(Color.Green, colors.ElementAt(0));
            Assert.Equal(Color.Red, colors.ElementAt(1));
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
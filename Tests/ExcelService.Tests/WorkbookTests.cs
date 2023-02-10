using ExcelService.Models;
using System.Drawing;
using ExcelService.Enums;
using Microsoft.CSharp.RuntimeBinder;

namespace ExcelService.Tests
{
    public class WorkbookTests
    {
        public Workbook noStyleWorkbook = Workbook.GetWorkbookFromDataSet(new List<TestObject>()
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
        public Workbook styleWorkbook = Workbook.GetWorkbookFromDataSet(new List<TestObject>()
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
                    Font = Font.Calibri,
                    FontSize= 12,
                },
                new Style()
                {
                    Color = Color.Red,
                    Font = Font.Arial,
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
                    Font = Font.Calibri,
                    FontSize= 12,
                },
                new Style()
                {
                    Color = Color.Red,
                    Font = Font.Arial,
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
                    Font = Font.Calibri,
                    FontSize= 12,
                },
                new Style()
                {
                    Color = Color.Red,
                    Font = Font.Arial,
                    FontSize= 8,
                },
                Style.Empty(),
                Style.Empty()
            },
        }
        );

        public WorkbookTests() 
        {

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
        public void TestGetDistinctStyles()
        {
            IEnumerable<Style?> styles = styleWorkbook.GetDistinctStyles();

            IEnumerable<Style?> noStyles = noStyleWorkbook.GetDistinctStyles();

            Assert.Single(noStyles);

            Assert.NotEmpty(styles);
            Assert.Equal(Color.Gray, styles.Select(x => x?.Color).ElementAt(0));
            Assert.Equal(Color.Green, styles.Select(x => x?.Color).ElementAt(1));
            Assert.Equal(Color.Red, styles.Select(x => x?.Color).ElementAt(2));

            Assert.Equal(Font.Arial, styles.Select(x => x?.Font).ElementAt(0));
            Assert.Equal(Font.Calibri, styles.Select(x => x?.Font).ElementAt(1));
            Assert.Equal(Font.Arial, styles.Select(x => x?.Font).ElementAt(2));

            Assert.Equal(409, styles.Select(x => x?.FontSize).ElementAt(0));
            Assert.Equal(12, styles.Select(x => x?.FontSize).ElementAt(1));
            Assert.Equal(8, styles.Select(x => x?.FontSize).ElementAt(2));
        }
        [Fact]
        public void TestStyleWhere()
        {
            Style aquaLargeText = new Style(Font.Calibri, Color.Aquamarine, 32);

            noStyleWorkbook.StyleRowWhere((TestObject x) => x.A == "A String", aquaLargeText);

            Assert.Equal(aquaLargeText, noStyleWorkbook[0, 0, 0].Style);
            Assert.Equal(aquaLargeText, noStyleWorkbook[0, 0, 1].Style);
            Assert.Equal(aquaLargeText, noStyleWorkbook[0, 0, 2].Style);
            Assert.Equal(aquaLargeText, noStyleWorkbook[0, 0, 3].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 1, 0].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 1, 1].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 1, 2].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 1, 3].Style);
        }
        [Fact]
        public void TestStyleCellWhere()
        {
            Style aquaLargeText = new Style(Font.Calibri, Color.Aquamarine, 32);

            noStyleWorkbook.StyleCellWhere((TestObject x) => x.A == "A String", aquaLargeText);

            Assert.Equal(aquaLargeText, noStyleWorkbook[0, 0, 0].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 0, 1].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 0, 2].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 0, 3].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 1, 0].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 1, 1].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 1, 2].Style);
            Assert.NotEqual(aquaLargeText, noStyleWorkbook[0, 1, 3].Style);
        }
        [Fact]
        public void TestStylingWithWrongType()
        {
            Assert.Throws<RuntimeBinderException>(() => noStyleWorkbook.StyleRowWhere<DifferentTestObject>(x => x.E == "A", new Style(Font.Arial, Color.Red, 40)));
        }

        public class TestObject
        {
            public string? A { get; set; }
            public string? B { get; set; }
            public int C { get; set; }
            public int D { get; set; }
        }
        public class DifferentTestObject
        { 
            public string? E { get; set; }
            public string? F { get; set; }
            public string? G { get; set; }
            public string? H { get; set; }
        }
    }
}
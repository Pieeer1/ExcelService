using ExcelService.Models;
using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using System.Drawing;

namespace ExcelService.Tests.Objects
{
    public static class TestObjects
    {
        public static Workbook noStyleWorkbook = Workbook.GetWorkbookFromDataSet(new List<TestObject>()
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
        public static Workbook styleWorkbook = Workbook.GetWorkbookFromDataSet(new List<TestObject>()
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

        private class TestObject
        {
            public string? A { get; set; }
            public string? B { get; set; }
            public int C { get; set; }
            public int D { get; set; }
        }
    }
}

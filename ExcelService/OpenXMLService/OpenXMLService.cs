﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelService.Models;
using System.Drawing;
using System.Text;

namespace ExcelService.OpenXMLService
{
    public static class OpenXMLService
    {
        public static Models.Workbook GetWorkbookFromFile(string filePath)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            { 
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart ?? throw new NullReferenceException("Invalid Workbook Part");
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                List<Models.Row> rowList = new List<Models.Row>();
                foreach (DocumentFormat.OpenXml.Spreadsheet.Row row in sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>())
                {
                    List<Models.Cell> cellList = new List<Models.Cell>();
                    foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                    {
                        cellList.Add(new Models.Cell(cell.CellValue?.Text ?? string.Empty));
                    }
                    rowList.Add(new Models.Row(cellList));
                }
                return new Models.Workbook(new List<Models.Sheet>() { new Models.Sheet(new Models.Row(new List<Models.Cell>()), rowList, "A Temporary Sheet Name") }, "A Temporary Workbook Name");
            }
        }
        public static Stream GetXLSXStreamFromWorkbook(Models.Workbook excelServiceWorkbook)
        {
            if (excelServiceWorkbook.Sheets.Select(x => x.Name) is not null)
            {
                foreach (var sheetName in excelServiceWorkbook.Sheets.Select(x => x.Name))
                {
                    if (sheetName.Length >= 32)
                    {
                        throw new InvalidDataException("Length of Sheet Names Cannot be Longer than 31 Characters");
                    }
                }
            }
            MemoryStream stream = new MemoryStream();
            using (var workbook = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart!.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();

                StyleSheetMapperObject? mapper = null;
                if (excelServiceWorkbook.GetDistinctStyles().Any())
                {
                    mapper = CreateStyleSheet(excelServiceWorkbook.GetDistinctStyles()); // use dictionary from here

                    //idk why this works but it does not work without it so
                    WorkbookStylesPart workBookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                    workBookStylesPart.Stylesheet = mapper.StyleSheet;
                    workBookStylesPart.Stylesheet.Save();

                }

                uint sheetId = 1;
                foreach (Models.Sheet excelServiceSheet in excelServiceWorkbook.Sheets)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>() ?? throw new NullReferenceException("Invalid Sheets");
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets!.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                    {
                        sheetId = sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId!.Value).Max() + 1;
                    }

                    if (excelServiceSheet.Name.Length >= 32)
                    {
                        throw new InvalidDataException("Length of Sheet Names Cannot be Longer than 31 Characters");
                    }

                    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = excelServiceSheet.Name};
                    sheets.Append(sheet);

                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    List<string> columns = new List<string>();
                    foreach (Models.Cell excelServiceCell in excelServiceSheet.HeaderRow.Cells)
                    {
                        columns.Add(excelServiceCell.Data);

                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(excelServiceCell.Data);
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (Models.Row excelServiceRow in excelServiceSheet.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        for (int i = 0; i < excelServiceRow.Cells.Count(); i++)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();

                            //style magic here
                            if (mapper is not null)
                            {
                                cell.StyleIndex = mapper.StyleMapperDictionary.TryGetValue(excelServiceRow.Cells.ElementAt(i).Style, out uint value) ? value : 0U;
                            }

                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(excelServiceRow.Cells.ElementAt(i).Data);
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                }
            }
            stream.Position = 0;
            return stream;
        }
        private static StyleSheetMapperObject CreateStyleSheet(IEnumerable<Models.Style> distinctStyles)
        {
            StyleSheetMapperObject mapper = new StyleSheetMapperObject();


            //do not remove namespaces, breaks sheet
            Stylesheet stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts = new Fonts() { Count = (UInt32Value)(uint)distinctStyles.Count() + 1, KnownFonts = true };
            Fills fills = new Fills() { Count = (UInt32Value)(uint)distinctStyles.Count() + 1 };
            Borders borders = new Borders() { Count = (UInt32Value)(uint)distinctStyles.Count() + 1 };
            CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = (UInt32Value)(uint)distinctStyles.Count() + 1 };
            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)(uint)distinctStyles.Count() + 1 };
            CellStyles cellStyles = new CellStyles() { Count = (UInt32Value)(uint)distinctStyles.Count() + 1 };
            DifferentialFormats differentialFormats = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleMedium9" };
            SetDefaults(fonts, fills, borders, cellStyleFormats, cellFormats, cellStyles, mapper);
            //start of defaults
            uint iterator = 1;
            foreach (Models.Style style in distinctStyles)
            {
                //might need to change all 1U to iterator

                Font font = new Font();
                FontSize fontSize = new FontSize() { Val = style.FontSize ?? 11D };
                DocumentFormat.OpenXml.Spreadsheet.Color color = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U }; // 1U IS BLACK 0U IS WHITE
                FontName fontName = new FontName() { Val = style.Font.ToString() ?? "Calibri" };
                FontFamilyNumbering fontFamilyNumbering = new FontFamilyNumbering() { Val = 2 };
                FontScheme fontScheme = new FontScheme() { Val = FontSchemeValues.Minor };

                font.Append(fontSize);
                font.Append(color);
                font.Append(fontName);
                font.Append(fontFamilyNumbering);
                font.Append(fontScheme);
                fonts.Append(font);


                Fill fill = new Fill();
                if (style.Color is not null)
                {
                    PatternFill patternFill = new PatternFill() { PatternType = PatternValues.Solid };
                    ForegroundColor foregroundColor = new ForegroundColor() { Rgb = HexConverter(style.Color ?? throw new NullReferenceException("Invalid Color")) };
                    BackgroundColor backgroundColor = new BackgroundColor() { Indexed = (UInt32Value)64U };
                    patternFill.Append(foregroundColor);
                    patternFill.Append(backgroundColor);
                    fill.Append(patternFill);
                }
                else
                {
                    PatternFill patternFill = new PatternFill() { PatternType = PatternValues.None };
                    fill.Append(patternFill);
                }

                fills.Append(fill);

                Border border = new Border();
                LeftBorder leftBorder = new LeftBorder();
                RightBorder rightBorder = new RightBorder();
                TopBorder topBorder = new TopBorder();
                BottomBorder bottomBorder = new BottomBorder();
                DiagonalBorder diagonalBorder = new DiagonalBorder();

                border.Append(leftBorder);
                border.Append(rightBorder);
                border.Append(topBorder);
                border.Append(bottomBorder);
                border.Append(diagonalBorder);

                borders.Append(border);

                CellFormat cellStyleFormat = new CellFormat() { NumberFormatId = (UInt32Value)iterator, FontId = (UInt32Value)iterator, FillId = (UInt32Value)iterator, BorderId = (UInt32Value)iterator, FormatId = (UInt32Value)iterator, ApplyFill = true };

                cellStyleFormats.Append(cellStyleFormat);

                CellFormat cellFormat = new CellFormat() { NumberFormatId = (UInt32Value)iterator, FontId = (UInt32Value)iterator, FillId = (UInt32Value)iterator, BorderId = (UInt32Value)iterator, FormatId = (UInt32Value)iterator, ApplyFill = true };

                cellFormats.Append(cellFormat);

                CellStyle cellStyle = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)iterator, BuiltinId = (UInt32Value)iterator };

                cellStyles.Append(cellStyle);

                mapper.StyleMapperDictionary.Add(style, iterator++);
            }

            StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" }; // we love random guids
            stylesheetExtensionList.Append(stylesheetExtension);

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles);
            stylesheet.Append(differentialFormats);
            stylesheet.Append(tableStyles);
            stylesheet.Append(stylesheetExtensionList);

            mapper.StyleSheet = stylesheet;

            return mapper;
        }
        private static void SetDefaults(Fonts fonts, Fills fills, Borders borders, CellStyleFormats cellStyleFormats, CellFormats cellFormats, CellStyles cellStyles, StyleSheetMapperObject mapper)
        {
            Font font = new Font();
            FontSize fontSize = new FontSize() { Val = 12D };
            DocumentFormat.OpenXml.Spreadsheet.Color color = new DocumentFormat.OpenXml.Spreadsheet.Color() { Theme = (UInt32Value)1U };
            FontName fontName = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme = new FontScheme() { Val = FontSchemeValues.Minor };

            font.Append(fontSize);
            font.Append(color);
            font.Append(fontName);
            font.Append(fontFamilyNumbering);
            font.Append(fontScheme);
            fonts.Append(font);


            Fill fill = new Fill();

            PatternFill patternFill = new PatternFill() { PatternType = PatternValues.None };
            fill.Append(patternFill);
            
            fills.Append(fill);

            Border border = new Border();
            LeftBorder leftBorder = new LeftBorder();
            RightBorder rightBorder = new RightBorder();
            TopBorder topBorder = new TopBorder();
            BottomBorder bottomBorder = new BottomBorder();
            DiagonalBorder diagonalBorder = new DiagonalBorder();

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);
            borders.Append(border);
            CellFormat cellStyleFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
            cellStyleFormats.Append(cellStyleFormat);
            CellFormat cellFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true };
            cellFormats.Append(cellFormat);
            CellStyle cellStyle = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };
            cellStyles.Append(cellStyle);
            mapper.StyleMapperDictionary.Add(Models.Style.Empty(), 0);
        }
        private class StyleSheetMapperObject
        {
            public StyleSheetMapperObject()
            {
                StyleMapperDictionary = new Dictionary<Models.Style, uint>();
            }

            public Stylesheet StyleSheet { get; set; } = null!;
            public Dictionary<Models.Style, uint> StyleMapperDictionary { get; private set; }
        }
        private static string HexConverter(System.Drawing.Color c)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(c.R.ToString("X2"));
            sb.Append(c.G.ToString("X2"));
            sb.Append(c.B.ToString("X2"));
            return sb.ToString();
        }

    }
}

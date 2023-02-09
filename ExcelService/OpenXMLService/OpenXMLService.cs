using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace ExcelService.OpenXMLService
{
    public static class OpenXMLService
    {
        // no styles
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
                workbook.WorkbookPart!.Workbook = new Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();

                uint sheetId = 1;
                foreach (Models.Sheet excelServiceSheet in excelServiceWorkbook.Sheets)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>() ?? throw new NullReferenceException("Invalid Sheets");
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets!.Elements<Sheet>().Count() > 0)
                    {
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId!.Value).Max() + 1;
                    }

                    if (excelServiceSheet.Name.Length >= 32)
                    {
                        throw new InvalidDataException("Length of Sheet Names Cannot be Longer than 31 Characters");
                    }

                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = excelServiceSheet.Name};
                    sheets.Append(sheet);

                    Row headerRow = new Row();

                    List<string> columns = new List<string>();
                    foreach (Models.Cell excelServiceCell in excelServiceSheet.HeaderRow.Cells)
                    {
                        columns.Add(excelServiceCell.Data);

                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(excelServiceCell.Data);
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (Models.Row excelServiceRow in excelServiceSheet.Rows)
                    {
                        Row newRow = new Row();
                        for (int i = 0; i < excelServiceRow.Cells.Count(); i++)
                        {
                            Cell cell = new Cell();
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

    }
}

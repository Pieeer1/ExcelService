using System.Net;
using ExcelService.Interfaces;
using ExcelService.Models;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;

namespace ExcelService.APITests
{
    public class ExcelAPITests
    {
        private readonly ILogger _logger;
        private readonly IExcel _excel;

        public ExcelAPITests(ILoggerFactory loggerFactory, IExcel excel)
        {
            _logger = loggerFactory.CreateLogger<ExcelAPITests>();
            _excel = excel;
        }

        [Function("ExcelAPITests")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
        {

            //application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
            _logger.LogInformation("Excel Tests HTTP trigger function processed a request.");
            var response = req.CreateResponse();
            response.StatusCode = HttpStatusCode.OK;

            Workbook workbook = _excel.GetWorkbookFromExcelFile(req.Body);

            List<Cell> cells = new List<Cell>();
            workbook.Sheets.ToList().ForEach(sheet =>
            {
                sheet.Rows.ToList().ForEach(row =>
                {
                    row.Cells.ToList().ForEach(cell =>
                    {
                        cells.Add(cell);
                    });
                });
            });
            await response.WriteStringAsync(string.Join(',', cells.Select(x => x.Data)));

            return response;
        }
    }
}

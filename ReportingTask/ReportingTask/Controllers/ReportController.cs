using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using ReportingTask.Models;
using ReportingTask.Services;
using ReportingTask.Services.Contracts;
using System.Diagnostics;

namespace ReportingTask.Controllers
{
    public class ReportController : Controller
    {
        private readonly ILogger<ReportController> _logger;
        private readonly IJsonConverterService _jsonConverterService;
        private readonly IExcelReportService _excelReportService;

        public ReportController(ILogger<ReportController> logger)
        {
            _logger = logger;
            _jsonConverterService = new JsonConverterService();
            _excelReportService = new ExcelReportService();
        }

        [HttpPost]
        [Route("api/excel/download")]
        public async Task<IActionResult> DownloadExcelReport(IFormFile file)
        {
            if (file == null && file?.Length < 0)
                return BadRequest("No file has been loaded");

            var hotelData = await _jsonConverterService.ConvertJsonFileToObjectAsync(file);

            var excelFile = _excelReportService.GenerateExcelFile(hotelData);

            return File(excelFile, "application/octet-stream", "AwesomeHotel.xlsx") ;
        }
    }
}
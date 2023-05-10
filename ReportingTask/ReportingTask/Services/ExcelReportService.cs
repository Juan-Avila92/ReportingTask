using OfficeOpenXml;
using ReportingTask.ExcelExport;
using ReportingTask.ExcelExport.Contract;
using ReportingTask.Models;
using ReportingTask.Services.Contracts;

namespace ReportingTask.Services
{
    public class ExcelReportService : IExcelReportService
    {
        private readonly IHotelDataSheetGenerator _dataSheetGenerator;
        public ExcelReportService()
        {
            _dataSheetGenerator = new HotelDataSheetGenerator();
        }

        public MemoryStream GenerateExcelFile(HotelDataModel hotelData)
        {
            MemoryStream memoryStream = new MemoryStream();

            ExcelPackage excel = new ExcelPackage(memoryStream);

            excel = _dataSheetGenerator.GenerateHotelDataSheet(excel, hotelData);

            excel.Save();

            memoryStream.Position = 0;

            return memoryStream;
        }
    }
}

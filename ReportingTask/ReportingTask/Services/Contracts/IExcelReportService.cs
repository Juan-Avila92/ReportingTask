using ReportingTask.Models;

namespace ReportingTask.Services.Contracts
{
    public interface IExcelReportService
    {
        public MemoryStream GenerateExcelFile(HotelDataModel hotelData);
    }
}

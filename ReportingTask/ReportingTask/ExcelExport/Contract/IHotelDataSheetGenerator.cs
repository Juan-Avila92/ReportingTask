using OfficeOpenXml;
using ReportingTask.Models;

namespace ReportingTask.ExcelExport.Contract
{
    public interface IHotelDataSheetGenerator
    {
        public ExcelPackage GenerateHotelDataSheet(ExcelPackage excel, HotelDataModel hotelData);
    }
}

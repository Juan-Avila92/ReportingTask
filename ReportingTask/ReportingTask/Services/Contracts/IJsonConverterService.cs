using ReportingTask.Models;

namespace ReportingTask.Services.Contracts
{
    public interface IJsonConverterService
    {
        public Task<HotelDataModel> ConvertJsonFileToObjectAsync(IFormFile file);
    }
}

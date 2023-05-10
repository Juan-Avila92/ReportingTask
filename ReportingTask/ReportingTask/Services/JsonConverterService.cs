using Newtonsoft.Json;
using ReportingTask.Models;
using ReportingTask.Services.Contracts;

namespace ReportingTask.Services
{
    public class JsonConverterService : IJsonConverterService
    {
        public async Task<HotelDataModel> ConvertJsonFileToObjectAsync(IFormFile file)
        {
            var hotelData = new HotelDataModel();

            using (var streamReader = new StreamReader(file.OpenReadStream()))
            {
                var fileContent = await streamReader.ReadToEndAsync();

                hotelData = JsonConvert.DeserializeObject<HotelDataModel>(fileContent);
            }

            return hotelData;
        }
    }
}

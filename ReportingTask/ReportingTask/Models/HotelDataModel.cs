namespace ReportingTask.Models
{
    public class HotelDataModel
    {
        public HotelDataModel() 
        { 
            Hotel = new HotelModel();
            HotelRates = new List<HotelRatesModel>();
        }

        public HotelModel Hotel { get; set; }

        public List<HotelRatesModel> HotelRates { get; set; } 
    }
}
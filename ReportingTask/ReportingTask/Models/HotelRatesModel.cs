namespace ReportingTask.Models
{
    public class HotelRatesModel
    {
        public HotelRatesModel()
        {
            Adults = default;
            Los = default;
            Price = new PriceModel();
            RateDescription = string.Empty;
            RateId = string.Empty;
            RateName = string.Empty;
            RateTags = new List<RateTagModel>();
        }

        public int Adults { get; set; }
        public int Los { get; set; }
        public PriceModel Price { get; set; }
        public string RateDescription { get; set; }
        public string RateId { get; set; }
        public string RateName { get; set; }
        public List<RateTagModel> RateTags { get; set; }
        public DateTime TargetDay { get; set; }
    }
}
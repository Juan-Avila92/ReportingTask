namespace ReportingTask.Models
{
    [Serializable]
    public class HotelModel
    {        
        public int HotelID { get; set; }
        public int Classification { get; set; }
        public string Name { get; set; }
        public double Reviewscore { get; set; }
    }
}
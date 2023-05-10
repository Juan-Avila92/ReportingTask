namespace ReportingTask.Models
{
    public class RateTagModel
    {       
        public RateTagModel() {
            Name = string.Empty;
            Shape = false;
        }

        public string Name { get; set; }
        public bool Shape { get; set; }
    }
}
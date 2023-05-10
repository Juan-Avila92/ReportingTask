namespace ReportingTask.Models
{
    public class PriceModel
    {        
        public PriceModel() 
        {
            Currency = string.Empty;
            NumericInteger = default(int);
            NumericFloat= default(float);
        }

        public string Currency { get; set; }
        public float NumericFloat { get; set; }
        public int NumericInteger{ get; set; }
    }
}
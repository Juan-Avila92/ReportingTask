using OfficeOpenXml;
using ReportingTask.ExcelExport;
using ReportingTask.ExcelExport.Contract;
using ReportingTask.Models;

namespace ExcelExport.Test
{
    public class HotelDataSheetGeneratorTest
    {
        private IHotelDataSheetGenerator _service;

        [SetUp]
        public void Setup()
        {
            _service = new HotelDataSheetGenerator();
        }

        [Test]
        public void GenerateHotelDataSheet_HotelDataModel_ReturnsExcelSheets()
        {
            var dummyHotelDataModel = GetDummyHotelDataModel();

            MemoryStream memoryStream = new MemoryStream();

            ExcelPackage excel = new ExcelPackage(memoryStream);

            var results = _service.GenerateHotelDataSheet(excel, dummyHotelDataModel);

            var expectedSheetNames = new List<string> { "Hotel", "HotelRates" };
            var expectedHotelSheetColumValues = new List< List<string>> { 
                new List<string>()
                {
                    "HotelID", "Classification", "Name", "Reviewscore"
                },
                new List<string>()
                {
                    "123456", "10", "The Prancing Pony", "10"
                }
            };

            Assert.That(results.Workbook.Worksheets[0].Name, Is.EqualTo(expectedSheetNames[0]));
            Assert.That(results.Workbook.Worksheets[1].Name, Is.EqualTo(expectedSheetNames[1]));

            int rowCounter = 1;
            foreach(var expectedRowValue in expectedHotelSheetColumValues[0])
            {
                var rowValue = results.Workbook.Worksheets[0].GetValue(rowCounter++, 1);
                Assert.AreEqual(expectedRowValue, rowValue.ToString());
            }

            rowCounter = 1;
            foreach (var expectedRowValue in expectedHotelSheetColumValues[1])
            {
                var rowValue = results.Workbook.Worksheets[0].GetValue(rowCounter++, 2);
                Assert.AreEqual(expectedRowValue, rowValue.ToString());
            }
        }

        [Test]
        public void GenerateHotelDataSheet_HotelRates_ReturnsExcelSheets()
        {
            var dummyHotelDataModel = GetDummyHotelDataModel();

            MemoryStream memoryStream = new MemoryStream();

            ExcelPackage excel = new ExcelPackage(memoryStream);

            var results = _service.GenerateHotelDataSheet(excel, dummyHotelDataModel);

            var expectedHotelSheetColumValues = new List<List<string>> {
                new List<string>()
                {
                    "Adults", "Los", "Currency", "Price", "RateDescription", "RateId", "RateName", "Breakfast_Included", "TargetDay"
                },
                new List<string>()
                {
                    "2", "2", "EUR", "100,00", "Unsere Classic Zimmer", "123456", "Classic Zimmer - Frühbucher Rate", "0", DateTime.Now.ToString()
                }
            };

            int columnCounter = 1;
            foreach (var expectedColumnValue in expectedHotelSheetColumValues[0])
            {
                var columnValue = results.Workbook.Worksheets[1].GetValue(1, columnCounter++);
                Assert.AreEqual(expectedColumnValue, columnValue.ToString());
            }

            columnCounter = 1;
            foreach (var expectedColumnValue in expectedHotelSheetColumValues[1])
            {
                var columnValue = results.Workbook.Worksheets[1].GetValue(2, columnCounter++);
                Assert.AreEqual(expectedColumnValue, columnValue.ToString());
            }
        }



        private HotelDataModel GetDummyHotelDataModel()
        {
            return new HotelDataModel
            {
                Hotel = new HotelModel()
                {
                    HotelID = 123456,
                    Classification = 10,
                    Name = "The Prancing Pony",
                    Reviewscore = 10
                },
                HotelRates = new List<HotelRatesModel>()
                {
                    new HotelRatesModel()
                    {
                        Adults = 2,
                        Los = 2,
                        Price = new PriceModel()
                        {
                            Currency = "EUR",
                            NumericFloat = 100f,
                            NumericInteger = 100
                        },
                        RateDescription = "Unsere Classic Zimmer",
                        RateId = "123456",
                        RateName = "Classic Zimmer - Frühbucher Rate",
                        RateTags = new List<RateTagModel>{ new RateTagModel() { Name = "breakfast", Shape = false } },
                        TargetDay = DateTime.Now,
                    }
                }
            };
        }
    }
}
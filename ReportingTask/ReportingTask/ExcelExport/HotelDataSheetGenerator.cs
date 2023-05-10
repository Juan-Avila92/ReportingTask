using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using ReportingTask.ExcelExport.Contract;
using ReportingTask.Models;
using System.Drawing;

namespace ReportingTask.ExcelExport
{
    public class HotelDataSheetGenerator : IHotelDataSheetGenerator
    {
        public ExcelPackage GenerateHotelDataSheet(ExcelPackage excel, HotelDataModel hotelData)
        {
            CreateDataSheets(excel, hotelData);

            CreateHotelData(excel, hotelData);

            CreateHotelRatesData(excel, hotelData);

            CreateTableContent(excel, hotelData);

            SetCellColour(excel, hotelData);

            return excel;
        }

        private ExcelPackage CreateDataSheets(ExcelPackage excel, HotelDataModel hotelData)
        {
            foreach (var hotelDataProperty in hotelData.GetType().GetProperties())
            {
                excel.Workbook.Worksheets.Add(hotelDataProperty.Name);
            }

            return excel;
        }

        private ExcelPackage CreateHotelData(ExcelPackage excel, HotelDataModel hotelData)
        {
            var hotelWorksheet = excel.Workbook.Worksheets.Where(w => w.Name.Equals(nameof(hotelData.Hotel))).First();

            hotelWorksheet.Cells[$"A1"].Value = nameof(hotelData.Hotel.HotelID);
            hotelWorksheet.Cells[$"B1"].Value = hotelData.Hotel.HotelID;
            hotelWorksheet.Cells["A1:B1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            hotelWorksheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSteelBlue);

            hotelWorksheet.Cells[$"A2"].Value = nameof(hotelData.Hotel.Classification); ;
            hotelWorksheet.Cells[$"B2"].Value = hotelData.Hotel.Classification;

            hotelWorksheet.Cells[$"A3"].Value = nameof(hotelData.Hotel.Name); ;
            hotelWorksheet.Cells[$"B3"].Value = hotelData.Hotel.Name;
            hotelWorksheet.Cells["A3:B3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            hotelWorksheet.Cells["A3:B3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSteelBlue);

            hotelWorksheet.Cells[$"A4"].Value = nameof(hotelData.Hotel.Reviewscore); ;
            hotelWorksheet.Cells[$"B4"].Value = hotelData.Hotel.Reviewscore;

            var totalRows = hotelWorksheet.Dimension.End.Row;
            var totalColumns = hotelWorksheet.Dimension.End.Column;

            hotelWorksheet.Cells[1, 1, totalRows, totalColumns].Style.Font.Color.SetColor(Color.Navy);

            hotelWorksheet.Cells[hotelWorksheet.Dimension.Address].AutoFitColumns();

            return excel;
        }

        private ExcelPackage CreateHotelRatesData(ExcelPackage excel, HotelDataModel hotelData)
        {
            CreateHeaders(excel, hotelData);

            return excel;
        }

        private ExcelPackage CreateHeaders(ExcelPackage excel, HotelDataModel hotelData)
        {
            var hotelWorksheet = excel.Workbook.Worksheets.Where(w => w.Name.Equals(nameof(hotelData.HotelRates))).First();

            int columnCounter = 1;

            foreach (var hotelRateProperty in hotelData.HotelRates.First().GetType().GetProperties())
            {
                if (hotelRateProperty.PropertyType.Name.Equals(nameof(PriceModel)))
                {
                    hotelWorksheet.Cells[1, columnCounter].Value = nameof(PriceModel.Currency);
                    columnCounter++;
                    hotelWorksheet.Cells[1, columnCounter].Value = "Price";
                    columnCounter++;
                    continue;
                }

                if (hotelRateProperty.Name.Equals(nameof(HotelRatesModel.RateTags)))
                {
                    hotelWorksheet.Cells[1, columnCounter].Value = "Breakfast_Included";
                    columnCounter++;
                    continue;
                }

                hotelWorksheet.Cells[1, columnCounter].Value = hotelRateProperty.Name;

                columnCounter++;
            }

            return excel;
        }
        private ExcelPackage CreateTableContent(ExcelPackage excel, HotelDataModel hotelData)
        {
            var hotelWorksheet = excel.Workbook.Worksheets.Where(w => w.Name.Equals(nameof(hotelData.HotelRates))).First();

            int columnCounter = 1;
            int rowCounter = 2;
            var cellColour = Color.PowderBlue;

            foreach (var hotelRate in hotelData.HotelRates)
            {

                hotelWorksheet.Cells[rowCounter, columnCounter++].Value = hotelRate.Adults;
                hotelWorksheet.Cells[rowCounter, columnCounter++].Value = hotelRate.Los;
                hotelWorksheet.Cells[rowCounter, columnCounter++].Value = hotelRate.Price.Currency;
                hotelWorksheet.Cells[rowCounter, columnCounter++].Value = String.Format("{0:0.00}", hotelRate.Price.NumericFloat);
                hotelWorksheet.Cells[rowCounter, columnCounter++].Value = hotelRate.RateDescription;
                hotelWorksheet.Cells[rowCounter, columnCounter++].Value = hotelRate.RateId;
                hotelWorksheet.Cells[rowCounter, columnCounter++].Value = hotelRate.RateName;
                hotelWorksheet.Cells[rowCounter, columnCounter++].Value = hotelRate.RateTags.First().Shape ? 1 : 0;
                hotelWorksheet.Cells[rowCounter, columnCounter++].Value = hotelRate.TargetDay.ToString();

                rowCounter++;

                hotelWorksheet.Cells[hotelWorksheet.Dimension.Address].AutoFitColumns();

                columnCounter = 1;
            }

            return excel;
        }

        private ExcelPackage SetCellColour(ExcelPackage excel, HotelDataModel hotelData)
        {
            var hotelWorksheet = excel.Workbook.Worksheets.Where(w => w.Name.Equals(nameof(hotelData.HotelRates))).First();

            var totalRows = hotelWorksheet.Dimension.End.Row;
            var totalColumns = hotelWorksheet.Dimension.End.Column;

            hotelWorksheet.Cells[1, 1, totalRows, totalColumns].Style.Font.Color.SetColor(Color.Navy);

            for (int rowNumber = 2; rowNumber < totalRows; rowNumber++)
            {
                hotelWorksheet.Cells[rowNumber, 1, rowNumber, totalColumns].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                hotelWorksheet.Cells[rowNumber, 1, rowNumber, totalColumns].Style.Fill.BackgroundColor.SetColor(Color.LightSteelBlue);
                rowNumber++;
            }

            return excel;
        }
    }
}

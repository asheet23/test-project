using GrapeCity.Documents.Excel;
using Org.BouncyCastle.Crypto.Engines;

namespace ShowScenario
{
    internal class Program
    {
        static void Main(string[] args)
        {
            
            // Create a new workbook.
            var workbook = new Workbook();

            // Open an Excel file.
            workbook.Open("What-If-Analysis-Scenarios.xlsx");

            // Get the active sheet.
            var worksheet = workbook.ActiveSheet;

            // Create and add different scenarios which represent the different discount rates. 
            // Create a scenario with less discount rates.
            // The changing cells are D2:D6 and the comment of the scenario is "Created by Document Solutions for Excel".
            var lessDiscountRatesValues = new List<object> { 0.05, 0.02, 0.03, 0.02, 0.05 };
            var lessDiscountRates = worksheet.Scenarios.Add("Less Discount Rates", worksheet.Range["D2:D6"], lessDiscountRatesValues, "Created by Document Solutions for Excel");

            // Create a scenario with normal discount rates.
            // The changing cells are D2:D6.
            var normalDiscountRatesValues = new List<object> { 0.1, 0.05, 0.05, 0.05, 0.1 };
            var normalDiscountRates = worksheet.Scenarios.Add("Normal Discount Rates", worksheet.Range["D2:D6"], normalDiscountRatesValues);

            // Create a scenario with selling without discount.
            // The changing cells are D2:D6.
            var sellingWithoutDiscountValues = new List<object> { 0, 0, 0, 0, 0 };
            var sellingWithoutDiscount = worksheet.Scenarios.Add("Selling Without Discount", worksheet.Range["D2:D6"], sellingWithoutDiscountValues);

            // Create a scenario with bulk quantity sold.
            // The changing cells are E2:E6.
            var bulkQuantitySoldValues = new List<object> { 1000, 1000, 1000, 1000, 1000 };
            var bulkQuantitySold = worksheet.Scenarios.Add("Bulk Quantity Sold", worksheet.Range["E2:E6"], bulkQuantitySoldValues);

            #region SS
            // Show "Less Discount Rates" scenario.
            //worksheet.Scenarios["Less Discount Rates"].Show();

            worksheet.Scenarios["Normal Discount Rates"].Show();
            #endregion

            // Save the workbook.
            workbook.Save("ShowScenario.xlsx");
            
        }
    }
}

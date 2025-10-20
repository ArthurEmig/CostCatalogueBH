using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using CostsViewer.Services;

namespace CostsViewer
{
    public class TestCorrectionFactors
    {
        public static void CreateSampleExcelFile()
        {
            try
            {
                // Create a sample Excel file for testing
                var filePath = Path.Combine(Directory.GetCurrentDirectory(), "SampleCorrectionFactors.xlsx");
                
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Correction Factors");

                // Add headers
                worksheet.Cell(1, 1).Value = "Year";
                worksheet.Cell(1, 2).Value = "Correction Factor";

                // Add sample data with inflation-like progression
                var sampleData = new[]
                {
                    (1999, 1.0),
                    (2000, 1.02),
                    (2001, 1.04),
                    (2002, 1.06),
                    (2003, 1.08),
                    (2004, 1.10),
                    (2005, 1.12),
                    (2010, 1.22),
                    (2015, 1.32),
                    (2020, 1.42),
                    (2024, 1.50)
                };

                int row = 2;
                foreach (var (year, factor) in sampleData)
                {
                    worksheet.Cell(row, 1).Value = year;
                    worksheet.Cell(row, 2).Value = factor;
                    row++;
                }

                // Format the worksheet
                worksheet.Columns().AdjustToContents();
                worksheet.Column(2).Style.NumberFormat.Format = "0.00";

                workbook.SaveAs(filePath);
                Console.WriteLine($"Sample Excel file created: {filePath}");

                // Test the import functionality
                var importedSettings = CorrectionFactorService.ImportFromExcel(filePath);
                Console.WriteLine($"Successfully imported {importedSettings.YearFactors.Count} correction factors:");
                
                foreach (var kvp in importedSettings.YearFactors.OrderBy(x => x.Key))
                {
                    Console.WriteLine($"  {kvp.Key}: {kvp.Value:F2}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void TestTemplateCreation()
        {
            try
            {
                var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "CorrectionFactors_Template.xlsx");
                CorrectionFactorService.CreateExcelTemplate(templatePath);
                Console.WriteLine($"Template created successfully: {templatePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating template: {ex.Message}");
            }
        }
    }
}
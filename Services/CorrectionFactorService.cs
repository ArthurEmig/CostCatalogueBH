using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using CostsViewer.Models;

namespace CostsViewer.Services
{
    public class CorrectionFactorService
    {
        private static CorrectionFactorSettings? _cachedSettings;
        
        public static CorrectionFactorSettings LoadSettings()
        {
            if (_cachedSettings == null)
            {
                _cachedSettings = CorrectionFactorSettings.CreateDefault();
            }
            return _cachedSettings;
        }
        
        public static void SaveSettings(CorrectionFactorSettings settings)
        {
            _cachedSettings = settings;
            // Settings are now stored in memory only and will be included in Excel exports
            // No external files needed for single .exe deployment
        }
        
        public static void UpdateSettings(CorrectionFactorSettings settings)
        {
            _cachedSettings = settings;
        }

        /// <summary>
        /// Imports correction factors from an Excel file
        /// Expected format: Column A = Year, Column B = Correction Factor
        /// </summary>
        public static CorrectionFactorSettings ImportFromExcel(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Excel file not found: {filePath}");

            var settings = new CorrectionFactorSettings();

            try
            {
                using var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                    throw new InvalidOperationException("No worksheets found in the Excel file.");

                // Find the data range (skip header row if it exists)
                var rows = worksheet.RowsUsed().Skip(1); // Skip first row (assumed to be header)
                
                // Check if first row looks like a header
                var firstRow = worksheet.Row(1);
                var firstCellValue = firstRow.Cell(1).GetString().Trim();
                bool hasHeader = firstCellValue.Equals("Year", StringComparison.OrdinalIgnoreCase) ||
                               firstCellValue.Equals("Jahre", StringComparison.OrdinalIgnoreCase);

                if (!hasHeader)
                {
                    // If no header, include the first row in data processing
                    rows = worksheet.RowsUsed();
                }

                foreach (var row in rows)
                {
                    try
                    {
                        // Skip empty rows
                        if (row.CellsUsed().Count() < 2)
                            continue;

                        // Parse Year (Column A)
                        var yearCell = row.Cell(1).GetString().Trim();
                        if (!int.TryParse(yearCell, out int year))
                            continue;

                        // Parse Correction Factor (Column B)
                        var factorCell = row.Cell(2).GetString().Trim();
                        if (!double.TryParse(factorCell, out double factor))
                            continue;

                        // Validate reasonable ranges
                        if (year < 1900 || year > 2100)
                            continue;

                        if (factor <= 0 || factor > 10) // Reasonable factor range
                            continue;

                        settings.SetFactorForYear(year, factor);
                    }
                    catch (Exception ex)
                    {
                        // Log the error for debugging but continue processing
                        Console.WriteLine($"CorrectionFactorService: Error parsing row {row.RowNumber()} - {ex.Message}");
                    }
                }

                // Ensure we have at least some data
                if (settings.YearFactors.Count == 0)
                {
                    throw new InvalidOperationException("No valid correction factors found in the Excel file. Please check the format: Column A = Year, Column B = Correction Factor");
                }

                Console.WriteLine($"Successfully imported {settings.YearFactors.Count} correction factors from Excel file.");
                return settings;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error importing correction factors from Excel: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Creates an Excel template file for correction factors
        /// </summary>
        public static void CreateExcelTemplate(string filePath)
        {
            try
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Correction Factors");

                // Add headers
                worksheet.Cell(1, 1).Value = "Year";
                worksheet.Cell(1, 2).Value = "Correction Factor";

                // Style headers
                var headerRange = worksheet.Range(1, 1, 1, 2);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                // Add sample data from 1999 to current year
                var currentYear = DateTime.Now.Year;
                int row = 2;

                for (int year = 1999; year <= currentYear; year++)
                {
                    worksheet.Cell(row, 1).Value = year;
                    worksheet.Cell(row, 2).Value = 1.0; // Default factor
                    row++;
                }

                // Auto-fit columns
                worksheet.Columns().AdjustToContents();

                // Add some formatting
                var dataRange = worksheet.Range(2, 1, row - 1, 2);
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // Format correction factor column as number with 2 decimal places
                worksheet.Column(2).Style.NumberFormat.Format = "0.00";

                workbook.SaveAs(filePath);
                Console.WriteLine($"Excel template created successfully: {filePath}");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error creating Excel template: {ex.Message}", ex);
            }
        }
    }
}

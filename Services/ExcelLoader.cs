using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using CostsViewer.Models;

namespace CostsViewer.Services
{
    public static class ExcelLoader
    {
        public static List<ProjectRecord> Load(string filePath)
        {
            if (!File.Exists(filePath)) throw new FileNotFoundException(filePath);

            var records = new List<ProjectRecord>();

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            
            if (worksheet == null)
                throw new InvalidOperationException("No worksheets found in the Excel file.");

            // Find the data range (skip empty rows)
            var rows = worksheet.RowsUsed().Skip(1); // Skip header row

            foreach (var row in rows)
            {
                try
                {
                    // Check if the row has enough columns and is not empty
                    if (row.CellsUsed().Count() < 18) // Minimum expected columns
                        continue;

                    var rec = new ProjectRecord
                    {
                        // Column mapping to match Excel export format:
                        // Include, Project ID, Title, Types, Area, KG220, KG230, KG410, KG420, KG434, KG430, KG440, KG450, KG460, KG474, KG475, KG480, KG550
                        Include = ParseBooleanField(row.Cell(1).GetString()),     // Column A: Include
                        ProjectId = row.Cell(2).GetString(),                      // Column B: Project ID
                        ProjectTitle = row.Cell(3).GetString(),                   // Column C: Title
                        ProjectTypes = ParseTypes(row.Cell(4).GetString()),       // Column D: Types
                        TotalArea = ParseIntField(row.Cell(5).GetString()),       // Column E: Area
                        CostPerSqmKG220 = ParseIntField(row.Cell(6).GetString()), // Column F: KG220
                        CostPerSqmKG230 = ParseIntField(row.Cell(7).GetString()), // Column G: KG230
                        CostPerSqmKG410 = ParseIntField(row.Cell(8).GetString()), // Column H: KG410
                        CostPerSqmKG420 = ParseIntField(row.Cell(9).GetString()), // Column I: KG420
                        CostPerSqmKG434 = ParseIntField(row.Cell(10).GetString()),// Column J: KG434
                        CostPerSqmKG430 = ParseIntField(row.Cell(11).GetString()),// Column K: KG430
                        CostPerSqmKG440 = ParseIntField(row.Cell(12).GetString()),// Column L: KG440
                        CostPerSqmKG450 = ParseIntField(row.Cell(13).GetString()),// Column M: KG450
                        CostPerSqmKG460 = ParseIntField(row.Cell(14).GetString()),// Column N: KG460
                        CostPerSqmKG474 = ParseIntField(row.Cell(15).GetString()),// Column O: KG474
                        CostPerSqmKG475 = ParseIntField(row.Cell(16).GetString()),// Column P: KG475
                        CostPerSqmKG480 = ParseIntField(row.Cell(17).GetString()),// Column Q: KG480
                        CostPerSqmKG550 = ParseIntField(row.Cell(18).GetString()),// Column R: KG550
                    };
                    records.Add(rec);
                }
                catch (Exception ex)
                {
                    // Log the error for debugging but continue processing
                    Console.WriteLine($"ExcelLoader: Error parsing row {row.RowNumber()} - {ex.Message}");
                }
            }

            return records;
        }

        private static bool ParseBooleanField(string? field)
        {
            if (string.IsNullOrWhiteSpace(field)) return false;
            
            // Handle TRUE/FALSE, true/false, 1/0, yes/no
            return field.Trim().ToUpperInvariant() switch
            {
                "TRUE" => true,
                "FALSE" => false,
                "1" => true,
                "0" => false,
                "YES" => true,
                "NO" => false,
                _ => false
            };
        }

        private static int ParseIntField(string? field)
        {
            if (string.IsNullOrWhiteSpace(field)) return 0;
            
            // Try to parse as integer, return 0 if failed
            return int.TryParse(field.Trim(), out var result) ? result : 0;
        }

        private static List<string> ParseTypes(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new List<string>();
            
            // Remove surrounding quotes if present
            raw = raw.Trim().Trim('"');
            
            // Split by comma and clean up each type
            return raw.Split(',', StringSplitOptions.RemoveEmptyEntries)
                      .Select(s => s.Trim())
                      .Where(s => !string.IsNullOrEmpty(s))
                      .ToList();
        }
    }
}

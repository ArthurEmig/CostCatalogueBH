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
            var headerRow = worksheet.Row(1);

            foreach (var row in rows)
            {
                try
                {
                    // Check if the row has enough columns and is not empty
                    if (row.CellsUsed().Count() < 18) // Allow legacy files too; we'll map by headers below
                        continue;

                    var rec = new ProjectRecord
                    {
                        // Column mapping to match Excel export format:
                        // Include, Project ID, Title, Types, Area, Correction Factor, KG220, KG230, KG410, KG420, KG434, KG430, KG440, KG450, KG460, KG474, KG475, KG480, KG490, KG550, Year, Corrected...
                        Include = ParseBooleanField(row.Cell(1).GetString()),     // Column 1: Include
                        ProjectId = row.Cell(2).GetString(),                      // Column 2: Project ID
                        ProjectTitle = row.Cell(3).GetString(),                   // Column 3: Title
                        ProjectTypes = ParseTypes(row.Cell(4).GetString()),       // Column 4: Types
                        TotalArea = ParseIntField(row.Cell(5).GetString()),       // Column 5: Area
                        // Column 6 is Correction Factor - skip for import
                        CostPerSqmKG220 = ParseIntField(row.Cell(7).GetString()), // Column 7: KG220
                        CostPerSqmKG230 = ParseIntField(row.Cell(8).GetString()), // Column 8: KG230
                        CostPerSqmKG410 = ParseIntField(row.Cell(9).GetString()), // Column 9: KG410
                        CostPerSqmKG420 = ParseIntField(row.Cell(10).GetString()),// Column 10: KG420
                        CostPerSqmKG434 = ParseIntField(row.Cell(11).GetString()),// Column 11: KG434
                        CostPerSqmKG430 = ParseIntField(row.Cell(12).GetString()),// Column 12: KG430
                        CostPerSqmKG440 = ParseIntField(row.Cell(13).GetString()),// Column 13: KG440
                        CostPerSqmKG450 = ParseIntField(row.Cell(14).GetString()),// Column 14: KG450
                        CostPerSqmKG460 = ParseIntField(row.Cell(15).GetString()),// Column 15: KG460
                        CostPerSqmKG474 = ParseIntField(row.Cell(16).GetString()),// Column 16: KG474
                        CostPerSqmKG475 = ParseIntField(row.Cell(17).GetString()),// Column 17: KG475
                        CostPerSqmKG480 = ParseIntField(row.Cell(18).GetString()),// Column 18: KG480
                        CostPerSqmKG490 = ParseIntField(row.Cell(19).GetString()),// Column 19: KG490
                        CostPerSqmKG550 = ParseIntField(row.Cell(20).GetString()),// Column 20: KG550
                    };
                    // Year column is at position 21 in the new export format
                    try
                    {
                        int year = 0;
                        // Try by position first - Year is now in column 21
                        if (row.Cell(21) != null && !row.Cell(21).IsEmpty())
                        {
                            year = ParseIntField(row.Cell(21).GetString());
                        }

                        // If still zero, try header lookup for backward compatibility
                        if (year == 0)
                        {
                            int lastCell = headerRow.LastCellUsed().Address.ColumnNumber;
                            for (int c = 1; c <= lastCell; c++)
                            {
                                var header = headerRow.Cell(c).GetString()?.Trim();
                                if (string.Equals(header, "Year", StringComparison.OrdinalIgnoreCase) || string.Equals(header, "Year of cost calculation", StringComparison.OrdinalIgnoreCase))
                                {
                                    year = ParseIntField(row.Cell(c).GetString());
                                    break;
                                }
                            }
                        }
                        rec.Year = year > 0 ? year : DateTime.Now.Year;
                    }
                    catch
                    {
                        rec.Year = DateTime.Now.Year;
                    }
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

        private static int GetByHeaderOrIndex(IXLRow row, IXLRow headerRow, string headerName, params int[] indices)
        {
            // Try header first
            int lastCell = headerRow.LastCellUsed().Address.ColumnNumber;
            for (int c = 1; c <= lastCell; c++)
            {
                var header = headerRow.Cell(c).GetString()?.Trim();
                if (string.Equals(header, headerName, StringComparison.OrdinalIgnoreCase) || string.Equals(header, headerName.Replace(" €/sqm", string.Empty), StringComparison.OrdinalIgnoreCase))
                {
                    return ParseIntField(row.Cell(c).GetString());
                }
            }
            // Fallback by indices
            foreach (var idx in indices)
            {
                if (idx > 0)
                {
                    var cell = row.Cell(idx);
                    if (cell != null && !cell.IsEmpty())
                    {
                        return ParseIntField(cell.GetString());
                    }
                }
            }
            return 0;
        }
    }
}

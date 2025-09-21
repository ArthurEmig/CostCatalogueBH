using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using CostsViewer.Models;
using System.Linq;
using System.Text.Json;

namespace CostsViewer.Services
{
    public static class CsvLoader
    {
        public static List<ProjectRecord> Load(string filePath)
        {
            if (!File.Exists(filePath)) throw new FileNotFoundException(filePath);

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",  // Changed from semicolon to comma to match Excel export format
                HasHeaderRecord = true,
                IgnoreBlankLines = true,
                BadDataFound = null,
            };

            using var reader = new StreamReader(filePath);
            using var csv = new CsvReader(reader, config);
            csv.Read();
            csv.ReadHeader();

            var records = new List<ProjectRecord>();
            while (csv.Read())
            {
                try
                {
                    var rec = new ProjectRecord
                    {
                        // Column mapping to match Excel export format:
                        // Include, Project ID, Title, Types, Area, KG220, KG230, KG410, KG420, KG434, KG430, KG440, KG450, KG460, KG474, KG475, KG480, KG550
                        Include = ParseBooleanField(csv.GetField(0)),  // Column 0: Include
                        ProjectId = csv.GetField(1) ?? string.Empty,   // Column 1: Project ID
                        ProjectTitle = csv.GetField(2) ?? string.Empty, // Column 2: Title
                        ProjectTypes = ParseTypes(csv.GetField(3)),    // Column 3: Types
                        TotalArea = csv.GetField<int>(4),              // Column 4: Area
                        CostPerSqmKG220 = csv.GetField<int>(5),        // Column 5: KG220
                        CostPerSqmKG230 = csv.GetField<int>(6),        // Column 6: KG230
                        CostPerSqmKG410 = csv.GetField<int>(7),        // Column 7: KG410
                        CostPerSqmKG420 = csv.GetField<int>(8),        // Column 8: KG420
                        CostPerSqmKG434 = csv.GetField<int>(9),        // Column 9: KG434
                        CostPerSqmKG430 = csv.GetField<int>(10),       // Column 10: KG430
                        CostPerSqmKG440 = csv.GetField<int>(11),       // Column 11: KG440
                        CostPerSqmKG450 = csv.GetField<int>(12),       // Column 12: KG450
                        CostPerSqmKG460 = csv.GetField<int>(13),       // Column 13: KG460
                        CostPerSqmKG474 = csv.GetField<int>(14),       // Column 14: KG474
                        CostPerSqmKG475 = csv.GetField<int>(15),       // Column 15: KG475
                        CostPerSqmKG480 = csv.GetField<int>(16),       // Column 16: KG480
                        CostPerSqmKG550 = csv.GetField<int>(17),       // Column 17: KG550
                    };
                    records.Add(rec);
                }
                catch (Exception ex)
                {
                    // Log the error for debugging but continue processing
                    Console.WriteLine($"CsvLoader: Error parsing row - {ex.Message}");
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



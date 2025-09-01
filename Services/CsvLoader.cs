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
                Delimiter = ";",
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
                        ProjectId = csv.GetField(0),
                        ProjectTitle = csv.GetField(1),
                        ProjectTypes = ParseTypes(csv.GetField(2)),
                        TotalArea = csv.GetField<int>(3),
                        // Totals exist in cols 4..13 but we recompute on demand, not needed here
                        CostPerSqmKG220 = csv.GetField<int>(14),
                        CostPerSqmKG410 = csv.GetField<int>(15),
                        CostPerSqmKG420 = csv.GetField<int>(16),
                        CostPerSqmKG434 = csv.GetField<int>(17),
                        CostPerSqmKG430 = csv.GetField<int>(18),
                        CostPerSqmKG440 = csv.GetField<int>(19),
                        CostPerSqmKG450 = csv.GetField<int>(20),
                        CostPerSqmKG460 = csv.GetField<int>(21),
                        CostPerSqmKG480 = csv.GetField<int>(22),
                        CostPerSqmKG550 = csv.GetField<int>(23),
                    };
                    records.Add(rec);
                }
                catch
                {
                    // Skip malformed rows
                }
            }
            return records;
        }

        private static List<string> ParseTypes(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return new List<string>();
            try
            {
                var list = JsonSerializer.Deserialize<List<string>>(raw);
                return list ?? new List<string>();
            }
            catch
            {
                // Attempt to parse simple comma-separated
                return raw.Split(',', StringSplitOptions.RemoveEmptyEntries)
                          .Select(s => s.Trim().Trim('"'))
                          .ToList();
            }
        }
    }
}



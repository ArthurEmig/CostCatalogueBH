using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using CostsViewer.Models;

namespace CostsViewer.ExportServices
{
    public static class ExcelExporter
    {
        public static void Export(List<ProjectRecord> records, List<CostGroupSummary> costGroupSummary)
        {
            if (records.Count == 0) return;

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var file = Path.Combine(desktop, $"Costs_Export_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");

            using var wb = new XLWorkbook();

            // Create Projects worksheet
            var projectsWs = wb.AddWorksheet("Projects");

            string[] headers = {
                "Include","Project ID","Title","Types","Area",
                "KG220 €/sqm","KG230 €/sqm","KG410 €/sqm","KG420 €/sqm","KG434 €/sqm","KG430 €/sqm","KG440 €/sqm","KG450 €/sqm","KG460 €/sqm","KG490 €/sqm","KG474 €/sqm","KG475 €/sqm","KG480 €/sqm","KG550 €/sqm","Year"
            };
            for (int i = 0; i < headers.Length; i++) projectsWs.Cell(1, i + 1).Value = headers[i];

            int r = 2;
            foreach (var p in records)
            {
                projectsWs.Cell(r, 1).Value = p.Include;
                projectsWs.Cell(r, 2).Value = p.ProjectId;
                projectsWs.Cell(r, 3).Value = p.ProjectTitle;
                projectsWs.Cell(r, 4).Value = string.Join(", ", p.ProjectTypes);
                projectsWs.Cell(r, 5).Value = p.TotalArea;
                projectsWs.Cell(r, 6).Value = p.CostPerSqmKG220;
                projectsWs.Cell(r, 7).Value = p.CostPerSqmKG230;
                projectsWs.Cell(r, 8).Value = p.CostPerSqmKG410;
                projectsWs.Cell(r, 9).Value = p.CostPerSqmKG420;
                projectsWs.Cell(r,10).Value = p.CostPerSqmKG434;
                projectsWs.Cell(r,11).Value = p.CostPerSqmKG430;
                projectsWs.Cell(r,12).Value = p.CostPerSqmKG440;
                projectsWs.Cell(r,13).Value = p.CostPerSqmKG450;
                projectsWs.Cell(r,14).Value = p.CostPerSqmKG460;
                projectsWs.Cell(r,15).Value = p.CostPerSqmKG490;
                projectsWs.Cell(r,16).Value = p.CostPerSqmKG474;
                projectsWs.Cell(r,17).Value = p.CostPerSqmKG475;
                projectsWs.Cell(r,18).Value = p.CostPerSqmKG480;
                projectsWs.Cell(r,19).Value = p.CostPerSqmKG550;
                projectsWs.Cell(r,20).Value = p.Year;
                r++;
            }

            projectsWs.Columns().AdjustToContents();

            // Create Summary worksheet
            var summaryWs = wb.AddWorksheet("Cost Group Summary (DIN 276)");

            string[] summaryHeaders = {
                "Cost Group", "Description", "Average €/sqm", "Min €/sqm", "Max €/sqm", "Standard Deviation"
            };
            for (int i = 0; i < summaryHeaders.Length; i++) summaryWs.Cell(1, i + 1).Value = summaryHeaders[i];

            // Style the header row
            var headerRange = summaryWs.Range(1, 1, 1, summaryHeaders.Length);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

            r = 2;
            foreach (var summary in costGroupSummary)
            {
                summaryWs.Cell(r, 1).Value = summary.CostGroup;
                summaryWs.Cell(r, 2).Value = summary.Description;
                summaryWs.Cell(r, 3).Value = Math.Round(summary.AverageCost, 2);
                summaryWs.Cell(r, 4).Value = Math.Round(summary.MinCost, 2);
                summaryWs.Cell(r, 5).Value = Math.Round(summary.MaxCost, 2);
                summaryWs.Cell(r, 6).Value = Math.Round(summary.StandardDeviation, 2);
                r++;
            }

            // Format the summary worksheet
            summaryWs.Columns().AdjustToContents();

            // Add borders to the summary table
            if (costGroupSummary.Count > 0)
            {
                var dataRange = summaryWs.Range(1, 1, costGroupSummary.Count + 1, summaryHeaders.Length);
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }

            wb.SaveAs(file);

            // Also export as CSV for consistency with import format
            ExportCsv(records);
        }

        public static void ExportCsv(List<ProjectRecord> records)
        {
            if (records.Count == 0) return;

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var file = Path.Combine(desktop, $"CSV_Costs_Export_{DateTime.Now:yyyyMMdd_HHmm}.csv");

            var csv = new StringBuilder();

            // Add header row matching the import format
            csv.AppendLine("Include,Project ID,Title,Types,Area,KG220 €/sqm,KG230 €/sqm,KG410 €/sqm,KG420 €/sqm,KG434 €/sqm,KG430 €/sqm,KG440 €/sqm,KG450 €/sqm,KG460 €/sqm,KG490 €/sqm,KG474 €/sqm,KG475 €/sqm,KG480 €/sqm,KG550 €/sqm,Year");

            // Add data rows
            foreach (var record in records)
            {
                var typesString = string.Join(", ", record.ProjectTypes);
                // Wrap types in quotes if it contains commas
                if (typesString.Contains(','))
                {
                    typesString = $"\"{typesString}\"";
                }

                csv.AppendLine($"{(record.Include ? "TRUE" : "FALSE")},{record.ProjectId},{record.ProjectTitle},{typesString},{record.TotalArea},{record.CostPerSqmKG220},{record.CostPerSqmKG230},{record.CostPerSqmKG410},{record.CostPerSqmKG420},{record.CostPerSqmKG434},{record.CostPerSqmKG430},{record.CostPerSqmKG440},{record.CostPerSqmKG450},{record.CostPerSqmKG460},{record.CostPerSqmKG490},{record.CostPerSqmKG474},{record.CostPerSqmKG475},{record.CostPerSqmKG480},{record.CostPerSqmKG550},{record.Year}");
            }

            File.WriteAllText(file, csv.ToString(), Encoding.UTF8);
        }
    }
}

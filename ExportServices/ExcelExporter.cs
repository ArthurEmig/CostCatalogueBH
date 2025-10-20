using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ClosedXML.Excel;
using CostsViewer.Models;
using CostsViewer.Services;

namespace CostsViewer.ExportServices
{
    public static class ExcelExporter
    {
        public static void Export(List<ProjectRecord> records, List<CostGroupSummary> costGroupSummary)
        {
            if (records.Count == 0) return;

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var file = Path.Combine(desktop, $"Costs_Export_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");

            // Load correction factor settings
            var correctionFactorSettings = CorrectionFactorService.LoadSettings();

            using var wb = new XLWorkbook();

            // Create Projects worksheet
            var projectsWs = wb.AddWorksheet("Projects");

            string[] headers = {
                "Include","Project ID","Title","Types","Area","Correction Factor",
                "KG220 €/sqm","KG230 €/sqm","KG410 €/sqm","KG420 €/sqm","KG434 €/sqm","KG430 €/sqm","KG440 €/sqm","KG450 €/sqm","KG460 €/sqm","KG474 €/sqm","KG475 €/sqm","KG480 €/sqm","KG490 €/sqm","KG550 €/sqm","Year",
                "Corrected KG220","Corrected KG230","Corrected KG410","Corrected KG420","Corrected KG434","Corrected KG430","Corrected KG440","Corrected KG450","Corrected KG460","Corrected KG474","Corrected KG475","Corrected KG480","Corrected KG490","Corrected KG550"
            };
            for (int i = 0; i < headers.Length; i++) projectsWs.Cell(1, i + 1).Value = headers[i];

            int r = 2;
            foreach (var p in records)
            {
                var correctionFactor = correctionFactorSettings.GetFactorForYear(p.Year);
                
                projectsWs.Cell(r, 1).Value = p.Include;
                projectsWs.Cell(r, 2).Value = p.ProjectId;
                projectsWs.Cell(r, 3).Value = p.ProjectTitle;
                projectsWs.Cell(r, 4).Value = string.Join(", ", p.ProjectTypes);
                projectsWs.Cell(r, 5).Value = p.TotalArea;
                projectsWs.Cell(r, 6).Value = correctionFactor;
                
                // Original values
                projectsWs.Cell(r, 7).Value = p.CostPerSqmKG220;
                projectsWs.Cell(r, 8).Value = p.CostPerSqmKG230;
                projectsWs.Cell(r, 9).Value = p.CostPerSqmKG410;
                projectsWs.Cell(r,10).Value = p.CostPerSqmKG420;
                projectsWs.Cell(r,11).Value = p.CostPerSqmKG434;
                projectsWs.Cell(r,12).Value = p.CostPerSqmKG430;
                projectsWs.Cell(r,13).Value = p.CostPerSqmKG440;
                projectsWs.Cell(r,14).Value = p.CostPerSqmKG450;
                projectsWs.Cell(r,15).Value = p.CostPerSqmKG460;
                projectsWs.Cell(r,16).Value = p.CostPerSqmKG474;
                projectsWs.Cell(r,17).Value = p.CostPerSqmKG475;
                projectsWs.Cell(r,18).Value = p.CostPerSqmKG480;
                projectsWs.Cell(r,19).Value = p.CostPerSqmKG490;
                projectsWs.Cell(r,20).Value = p.CostPerSqmKG550;
                
                // Year before corrected values
                projectsWs.Cell(r,21).Value = p.Year;
                
                // Corrected values
                projectsWs.Cell(r,22).Value = Math.Round(p.CostPerSqmKG220 * correctionFactor, 0);
                projectsWs.Cell(r,23).Value = Math.Round(p.CostPerSqmKG230 * correctionFactor, 0);
                projectsWs.Cell(r,24).Value = Math.Round(p.CostPerSqmKG410 * correctionFactor, 0);
                projectsWs.Cell(r,25).Value = Math.Round(p.CostPerSqmKG420 * correctionFactor, 0);
                projectsWs.Cell(r,26).Value = Math.Round(p.CostPerSqmKG434 * correctionFactor, 0);
                projectsWs.Cell(r,27).Value = Math.Round(p.CostPerSqmKG430 * correctionFactor, 0);
                projectsWs.Cell(r,28).Value = Math.Round(p.CostPerSqmKG440 * correctionFactor, 0);
                projectsWs.Cell(r,29).Value = Math.Round(p.CostPerSqmKG450 * correctionFactor, 0);
                projectsWs.Cell(r,30).Value = Math.Round(p.CostPerSqmKG460 * correctionFactor, 0);
                projectsWs.Cell(r,31).Value = Math.Round(p.CostPerSqmKG474 * correctionFactor, 0);
                projectsWs.Cell(r,32).Value = Math.Round(p.CostPerSqmKG475 * correctionFactor, 0);
                projectsWs.Cell(r,33).Value = Math.Round(p.CostPerSqmKG480 * correctionFactor, 0);
                projectsWs.Cell(r,34).Value = Math.Round(p.CostPerSqmKG490 * correctionFactor, 0);
                projectsWs.Cell(r,35).Value = Math.Round(p.CostPerSqmKG550 * correctionFactor, 0);
                
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
                summaryWs.Cell(r, 3).Value = Math.Round(summary.AverageCost, 0);
                summaryWs.Cell(r, 4).Value = Math.Round(summary.MinCost, 0);
                summaryWs.Cell(r, 5).Value = Math.Round(summary.MaxCost, 0);
                summaryWs.Cell(r, 6).Value = Math.Round(summary.StandardDeviation, 0);
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

            // Create Correction Factors worksheet
            var correctionWs = wb.AddWorksheet("Correction Factors");
            
            string[] correctionHeaders = { "Year", "Correction Factor", "Percentage" };
            for (int i = 0; i < correctionHeaders.Length; i++) correctionWs.Cell(1, i + 1).Value = correctionHeaders[i];
            
            // Style the header row
            var correctionHeaderRange = correctionWs.Range(1, 1, 1, correctionHeaders.Length);
            correctionHeaderRange.Style.Font.Bold = true;
            correctionHeaderRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            
            r = 2;
            var currentYear = DateTime.Now.Year;
            for (int year = 1999; year <= currentYear; year++)
            {
                var factor = correctionFactorSettings.GetFactorForYear(year);
                correctionWs.Cell(r, 1).Value = year;
                correctionWs.Cell(r, 2).Value = Math.Round(factor, 4);
                correctionWs.Cell(r, 3).Value = $"{(factor * 100):F2}%";
                r++;
            }
            
            correctionWs.Columns().AdjustToContents();
            
            // Add borders to the correction factors table
            var correctionDataRange = correctionWs.Range(1, 1, currentYear - 1999 + 2, correctionHeaders.Length);
            correctionDataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            correctionDataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            wb.SaveAs(file);

            // Also export as CSV for consistency with import format
            ExportCsv(records);
        }

        public static void ExportCsv(List<ProjectRecord> records)
        {
            if (records.Count == 0) return;

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var file = Path.Combine(desktop, $"CSV_Costs_Export_{DateTime.Now:yyyyMMdd_HHmm}.csv");

            // Load correction factor settings
            var correctionFactorSettings = CorrectionFactorService.LoadSettings();

            var csv = new StringBuilder();

            // Add header row with correction factors
            csv.AppendLine("Include,Project ID,Title,Types,Area,Correction Factor,KG220 €/sqm,KG230 €/sqm,KG410 €/sqm,KG420 €/sqm,KG434 €/sqm,KG430 €/sqm,KG440 €/sqm,KG450 €/sqm,KG460 €/sqm,KG474 €/sqm,KG475 €/sqm,KG480 €/sqm,KG490 €/sqm,KG550 €/sqm,Year,Corrected KG220,Corrected KG230,Corrected KG410,Corrected KG420,Corrected KG434,Corrected KG430,Corrected KG440,Corrected KG450,Corrected KG460,Corrected KG474,Corrected KG475,Corrected KG480,Corrected KG490,Corrected KG550");

            // Add data rows
            foreach (var record in records)
            {
                var correctionFactor = correctionFactorSettings.GetFactorForYear(record.Year);
                var typesString = string.Join(", ", record.ProjectTypes);
                // Wrap types in quotes if it contains commas
                if (typesString.Contains(','))
                {
                    typesString = $"\"{typesString}\"";
                }

                csv.AppendLine($"{(record.Include ? "TRUE" : "FALSE")},{record.ProjectId},{record.ProjectTitle},{typesString},{record.TotalArea},{correctionFactor:F4},{record.CostPerSqmKG220},{record.CostPerSqmKG230},{record.CostPerSqmKG410},{record.CostPerSqmKG420},{record.CostPerSqmKG434},{record.CostPerSqmKG430},{record.CostPerSqmKG440},{record.CostPerSqmKG450},{record.CostPerSqmKG460},{record.CostPerSqmKG474},{record.CostPerSqmKG475},{record.CostPerSqmKG480},{record.CostPerSqmKG490},{record.CostPerSqmKG550},{record.Year},{Math.Round(record.CostPerSqmKG220 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG230 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG410 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG420 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG434 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG430 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG440 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG450 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG460 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG474 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG475 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG480 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG490 * correctionFactor, 0)},{Math.Round(record.CostPerSqmKG550 * correctionFactor, 0)}");
            }

            File.WriteAllText(file, csv.ToString(), Encoding.UTF8);
        }
    }
}

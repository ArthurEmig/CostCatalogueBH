using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using CostsViewer.Models;

namespace CostsViewer.ExportServices
{
    public static class ExcelExporter
    {
        public static void Export(List<ProjectRecord> records)
        {
            if (records.Count == 0) return;

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var file = Path.Combine(desktop, $"Costs_Export_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Projects");

            string[] headers = {
                "Include","Project ID","Title","Types","Area",
                "KG220 €/sqm","KG410 €/sqm","KG420 €/sqm","KG434 €/sqm","KG430 €/sqm","KG440 €/sqm","KG450 €/sqm","KG460 €/sqm","KG480 €/sqm","KG550 €/sqm"
            };
            for (int i = 0; i < headers.Length; i++) ws.Cell(1, i + 1).Value = headers[i];

            int r = 2;
            foreach (var p in records)
            {
                ws.Cell(r, 1).Value = p.Include;
                ws.Cell(r, 2).Value = p.ProjectId;
                ws.Cell(r, 3).Value = p.ProjectTitle;
                ws.Cell(r, 4).Value = string.Join(", ", p.ProjectTypes);
                ws.Cell(r, 5).Value = p.TotalArea;
                ws.Cell(r, 6).Value = p.CostPerSqmKG220;
                ws.Cell(r, 7).Value = p.CostPerSqmKG410;
                ws.Cell(r, 8).Value = p.CostPerSqmKG420;
                ws.Cell(r, 9).Value = p.CostPerSqmKG434;
                ws.Cell(r,10).Value = p.CostPerSqmKG430;
                ws.Cell(r,11).Value = p.CostPerSqmKG440;
                ws.Cell(r,12).Value = p.CostPerSqmKG450;
                ws.Cell(r,13).Value = p.CostPerSqmKG460;
                ws.Cell(r,14).Value = p.CostPerSqmKG480;
                ws.Cell(r,15).Value = p.CostPerSqmKG550;
                r++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(file);
        }
    }
}



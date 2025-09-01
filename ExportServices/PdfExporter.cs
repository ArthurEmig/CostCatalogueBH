using System;
using System.Collections.Generic;
using System.IO;
using CostsViewer.Models;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace CostsViewer.ExportServices
{
    public static class PdfExporter
    {
        public static void Export(List<ProjectRecord> records, double avgArea, params double[] avgKgs)
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var file = Path.Combine(desktop, $"Costs_Export_{DateTime.Now:yyyyMMdd_HHmm}.pdf");

            using var doc = new PdfDocument();
            var page = doc.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            var font = new XFont("Arial", 12);
            double y = 40;
            gfx.DrawString("Costs Export", new XFont("Arial", 18, XFontStyle.Bold), XBrushes.Black, new XPoint(40, y));
            y += 30;
            gfx.DrawString($"Included records: {records.Count}", font, XBrushes.Black, new XPoint(40, y));
            y += 20;
            gfx.DrawString($"Average Area: {avgArea:F1} sqm", font, XBrushes.Black, new XPoint(40, y));
            y += 20;

            string[] labels = {"KG220","KG410","KG420","KG434","KG430","KG440","KG450","KG460","KG480","KG550"};
            for (int i = 0; i < Math.Min(labels.Length, avgKgs.Length); i++)
            {
                gfx.DrawString($"Average {labels[i]}: {avgKgs[i]:F1} €/sqm", font, XBrushes.Black, new XPoint(40, y));
                y += 18;
            }

            y += 10;
            gfx.DrawString("Projects:", new XFont("Arial", 14, XFontStyle.Bold), XBrushes.Black, new XPoint(40, y));
            y += 22;
            foreach (var p in records)
            {
                var line = $"{(p.Include ? "[x]" : "[ ]")} {p.ProjectId} | {p.ProjectTitle} | {string.Join(", ", p.ProjectTypes)} | {p.TotalArea} sqm";
                gfx.DrawString(line, font, XBrushes.Black, new XPoint(40, y));
                y += 16;
                if (y > page.Height - 40)
                {
                    page = doc.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    y = 40;
                }
            }

            doc.Save(file);
        }
    }
}



using System;
using System.Collections.Generic;
using System.IO;
using CostsViewer.Models;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Text;

namespace CostsViewer.ExportServices
{
    public static class PdfExporter
    {
        public static void Export(List<ProjectRecord> records, List<CostGroupSummary> costGroupSummary, double avgArea, params double[] avgKgs)
        {
            try
            {
                // Set encoding to ensure compatibility
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                
                var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                var file = Path.Combine(desktop, $"Costs_Export_{DateTime.Now:yyyyMMdd_HHmm}.pdf");

                using var doc = new PdfDocument();
                var page = doc.AddPage();
                var gfx = XGraphics.FromPdfPage(page);
                
                // Use more compatible font creation
                XFont font;
                XFont boldFont;
                try
                {
                    font = new XFont("Arial", 12, XFontStyle.Regular);
                    boldFont = new XFont("Arial", 12, XFontStyle.Bold);
                }
                catch
                {
                    // Fallback to default fonts if Arial is not available
                    font = new XFont("Times New Roman", 12, XFontStyle.Regular);
                    boldFont = new XFont("Times New Roman", 12, XFontStyle.Bold);
                }
                
                double y = 40;
            
                // Create title font with fallback
                XFont titleFont;
                try
                {
                    titleFont = new XFont("Arial", 18, XFontStyle.Bold);
                }
                catch
                {
                    titleFont = new XFont("Times New Roman", 18, XFontStyle.Bold);
                }
                
                gfx.DrawString("Costs Export", titleFont, XBrushes.Black, new XPoint(40, y));
            y += 30;
            gfx.DrawString($"Included records: {records.Count}", font, XBrushes.Black, new XPoint(40, y));
            y += 20;
            gfx.DrawString($"Average Area: {avgArea:F1} sqm", font, XBrushes.Black, new XPoint(40, y));
            y += 30;

                // Cost Group Summary Section
                XFont sectionFont;
                try
                {
                    sectionFont = new XFont("Arial", 14, XFontStyle.Bold);
                }
                catch
                {
                    sectionFont = new XFont("Times New Roman", 14, XFontStyle.Bold);
                }
                
                gfx.DrawString("Cost Group Summary (DIN 276):", sectionFont, XBrushes.Black, new XPoint(40, y));
            y += 25;

            // Summary table headers
            gfx.DrawString("Cost Group", boldFont, XBrushes.Black, new XPoint(40, y));
            gfx.DrawString("Description", boldFont, XBrushes.Black, new XPoint(120, y));
            gfx.DrawString("Avg €/sqm", boldFont, XBrushes.Black, new XPoint(250, y));
            gfx.DrawString("Min €/sqm", boldFont, XBrushes.Black, new XPoint(320, y));
            gfx.DrawString("Max €/sqm", boldFont, XBrushes.Black, new XPoint(390, y));
            gfx.DrawString("Std Dev", boldFont, XBrushes.Black, new XPoint(460, y));
            y += 20;

            // Summary table data
            foreach (var summary in costGroupSummary)
            {
                gfx.DrawString(summary.CostGroup, font, XBrushes.Black, new XPoint(40, y));
                gfx.DrawString(summary.Description, font, XBrushes.Black, new XPoint(120, y));
                gfx.DrawString($"{summary.AverageCost:F2}", font, XBrushes.Black, new XPoint(250, y));
                gfx.DrawString($"{summary.MinCost:F2}", font, XBrushes.Black, new XPoint(320, y));
                gfx.DrawString($"{summary.MaxCost:F2}", font, XBrushes.Black, new XPoint(390, y));
                gfx.DrawString($"{summary.StandardDeviation:F2}", font, XBrushes.Black, new XPoint(460, y));
                y += 16;
                
                if (y > page.Height - 80)
                {
                    page = doc.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    y = 40;
                }
            }

            y += 20;

                // Legacy averages section for backward compatibility
                string[] labels = {
                    "KG220 (Site Prep)", 
                    "KG230 (Earthworks)", 
                    "KG410 (Water/Gas)", 
                    "KG420 (Heating)", 
                    "KG434 (Process)", 
                    "KG430 (HVAC)", 
                    "KG440 (Electrical)", 
                    "KG450 (Comm/Safety)", 
                    "KG460 (Conveying)", 
                    "KG474 (Fire Protection)", 
                    "KG475 (Security/Access)", 
                    "KG480 (Automation)", 
                    "KG550 (Outdoor Tech)"
                };
                gfx.DrawString("Overall Averages (DIN 276):", sectionFont, XBrushes.Black, new XPoint(40, y));
            y += 22;
            for (int i = 0; i < Math.Min(labels.Length, avgKgs.Length); i++)
            {
                gfx.DrawString($"Average {labels[i]}: {avgKgs[i]:F2} €/sqm", font, XBrushes.Black, new XPoint(40, y));
                y += 18;
                
                if (y > page.Height - 80)
                {
                    page = doc.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    y = 40;
                }
            }

                y += 20;
                gfx.DrawString("Projects:", sectionFont, XBrushes.Black, new XPoint(40, y));
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
                Console.WriteLine($"PDF exported successfully to: {file}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"PDF Export Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
                
                // Try to create a simple text-based PDF as fallback
                try
                {
                    CreateFallbackPdf(records, costGroupSummary, avgArea, avgKgs);
                }
                catch (Exception fallbackEx)
                {
                    Console.WriteLine($"Fallback PDF creation also failed: {fallbackEx.Message}");
                    throw new InvalidOperationException($"PDF export failed: {ex.Message}. Fallback also failed: {fallbackEx.Message}", ex);
                }
            }
        }

        private static void CreateFallbackPdf(List<ProjectRecord> records, List<CostGroupSummary> costGroupSummary, double avgArea, double[] avgKgs)
        {
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var file = Path.Combine(desktop, $"Costs_Export_Fallback_{DateTime.Now:yyyyMMdd_HHmm}.pdf");

            using var doc = new PdfDocument();
            var page = doc.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            
            // Use the most basic font possible
            var font = new XFont("Courier", 10);
            double y = 40;
            
            gfx.DrawString("Costs Export (Fallback Mode)", font, XBrushes.Black, new XPoint(40, y));
            y += 20;
            gfx.DrawString($"Records: {records.Count}, Average Area: {avgArea:F1} sqm", font, XBrushes.Black, new XPoint(40, y));
            y += 20;
            
            gfx.DrawString("Cost Group Summary:", font, XBrushes.Black, new XPoint(40, y));
            y += 15;
            
            foreach (var summary in costGroupSummary)
            {
                var line = $"{summary.CostGroup}: Avg={summary.AverageCost:F0}, Min={summary.MinCost:F0}, Max={summary.MaxCost:F0}";
                gfx.DrawString(line, font, XBrushes.Black, new XPoint(40, y));
                y += 12;
                
                if (y > page.Height - 40)
                {
                    page = doc.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    y = 40;
                }
            }
            
            doc.Save(file);
            Console.WriteLine($"Fallback PDF exported successfully to: {file}");
        }
    }
}



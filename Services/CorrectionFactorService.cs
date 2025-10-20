using System;
using System.IO;
using System.Text.Json;
using CostsViewer.Models;

namespace CostsViewer.Services
{
    public class CorrectionFactorService
    {
        private const string SettingsFileName = "appsettings.json";
        
        public static CorrectionFactorSettings LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsFileName))
                {
                    var defaultSettings = CorrectionFactorSettings.CreateDefault();
                    SaveSettings(defaultSettings);
                    return defaultSettings;
                }

                var json = File.ReadAllText(SettingsFileName);
                var document = JsonDocument.Parse(json);
                
                if (document.RootElement.TryGetProperty("CorrectionFactors", out var correctionFactorsElement))
                {
                    var settings = new CorrectionFactorSettings();
                    
                    if (correctionFactorsElement.TryGetProperty("YearFactors", out var yearFactorsElement))
                    {
                        foreach (var property in yearFactorsElement.EnumerateObject())
                        {
                            if (int.TryParse(property.Name, out var year) && property.Value.TryGetDouble(out var factor))
                            {
                                settings.YearFactors[year] = factor;
                            }
                        }
                    }
                    
                    // Ensure all years from 1999 to current year are present
                    var currentYear = DateTime.Now.Year;
                    for (int year = 1999; year <= currentYear; year++)
                    {
                        if (!settings.YearFactors.ContainsKey(year))
                        {
                            settings.YearFactors[year] = 1.0;
                        }
                    }
                    
                    return settings;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading correction factor settings: {ex.Message}");
            }
            
            return CorrectionFactorSettings.CreateDefault();
        }
        
        public static void SaveSettings(CorrectionFactorSettings settings)
        {
            try
            {
                // Read existing settings or create new structure
                JsonDocument? existingDocument = null;
                if (File.Exists(SettingsFileName))
                {
                    var existingJson = File.ReadAllText(SettingsFileName);
                    if (!string.IsNullOrWhiteSpace(existingJson))
                    {
                        existingDocument = JsonDocument.Parse(existingJson);
                    }
                }

                using var stream = new MemoryStream();
                using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true });
                
                writer.WriteStartObject();
                
                // Copy existing properties except CorrectionFactors
                if (existingDocument != null)
                {
                    foreach (var property in existingDocument.RootElement.EnumerateObject())
                    {
                        if (property.Name != "CorrectionFactors")
                        {
                            property.WriteTo(writer);
                        }
                    }
                }
                
                // Write CorrectionFactors
                writer.WritePropertyName("CorrectionFactors");
                writer.WriteStartObject();
                
                writer.WritePropertyName("YearFactors");
                writer.WriteStartObject();
                foreach (var kvp in settings.YearFactors)
                {
                    writer.WriteNumber(kvp.Key.ToString(), kvp.Value);
                }
                writer.WriteEndObject();
                
                writer.WriteEndObject();
                writer.WriteEndObject();
                
                var json = System.Text.Encoding.UTF8.GetString(stream.ToArray());
                File.WriteAllText(SettingsFileName, json);
                
                existingDocument?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving correction factor settings: {ex.Message}");
            }
        }
    }
}

using System;
using CostsViewer.Models;
using CostsViewer.Services;

namespace CostsViewer
{
    public class TestCorrectionFactors
    {
        public static void TestExport()
        {
            Console.WriteLine("=== Testing Correction Factors Export ===");
            
            // Load correction factor settings
            var correctionFactorSettings = CorrectionFactorService.LoadSettings();
            
            // Test a few years
            for (int year = 2020; year <= 2024; year++)
            {
                var factor = correctionFactorSettings.GetFactorForYear(year);
                Console.WriteLine($"Year {year}: Factor = {factor:F4}");
            }
            
            // Test with a sample project
            var testProject = new ProjectRecord
            {
                ProjectId = "TEST001",
                ProjectTitle = "Test Project",
                Year = 2020,
                CostPerSqmKG220 = 100,
                CostPerSqmKG230 = 200,
                CostPerSqmKG410 = 150
            };
            
            var testFactor = correctionFactorSettings.GetFactorForYear(testProject.Year);
            Console.WriteLine($"\nTest Project (Year {testProject.Year}):");
            Console.WriteLine($"Correction Factor: {testFactor:F4}");
            Console.WriteLine($"Original KG220: {testProject.CostPerSqmKG220}");
            Console.WriteLine($"Corrected KG220: {Math.Round(testProject.CostPerSqmKG220 * testFactor, 2)}");
            Console.WriteLine($"Original KG230: {testProject.CostPerSqmKG230}");
            Console.WriteLine($"Corrected KG230: {Math.Round(testProject.CostPerSqmKG230 * testFactor, 2)}");
            Console.WriteLine($"Original KG410: {testProject.CostPerSqmKG410}");
            Console.WriteLine($"Corrected KG410: {Math.Round(testProject.CostPerSqmKG410 * testFactor, 2)}");
        }
    }
}

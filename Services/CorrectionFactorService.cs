using System;
using CostsViewer.Models;

namespace CostsViewer.Services
{
    public class CorrectionFactorService
    {
        private static CorrectionFactorSettings? _cachedSettings;
        
        public static CorrectionFactorSettings LoadSettings()
        {
            if (_cachedSettings == null)
            {
                _cachedSettings = CorrectionFactorSettings.CreateDefault();
            }
            return _cachedSettings;
        }
        
        public static void SaveSettings(CorrectionFactorSettings settings)
        {
            _cachedSettings = settings;
            // Settings are now stored in memory only and will be included in Excel exports
            // No external files needed for single .exe deployment
        }
        
        public static void UpdateSettings(CorrectionFactorSettings settings)
        {
            _cachedSettings = settings;
        }
    }
}

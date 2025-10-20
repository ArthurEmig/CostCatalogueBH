using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CostsViewer.Models
{
    public class CorrectionFactor : INotifyPropertyChanged
    {
        private int _year;
        private double _factor = 1.0;

        public int Year
        {
            get => _year;
            set
            {
                if (_year != value)
                {
                    _year = value;
                    OnPropertyChanged();
                }
            }
        }

        public double Factor
        {
            get => _factor;
            set
            {
                if (Math.Abs(_factor - value) > 0.0001)
                {
                    _factor = value;
                    OnPropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public class CorrectionFactorSettings
    {
        public Dictionary<int, double> YearFactors { get; set; } = new();

        public double GetFactorForYear(int year)
        {
            return YearFactors.TryGetValue(year, out var factor) ? factor : 1.0;
        }

        public void SetFactorForYear(int year, double factor)
        {
            YearFactors[year] = factor;
        }

        public static CorrectionFactorSettings CreateDefault()
        {
            var settings = new CorrectionFactorSettings();
            var currentYear = DateTime.Now.Year;
            
            // Initialize all years from 1999 to current year with factor 1.0
            for (int year = 1999; year <= currentYear; year++)
            {
                settings.YearFactors[year] = 1.0;
            }
            
            return settings;
        }
    }
}

using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using CostsViewer.Models;

namespace CostsViewer.ViewModels
{
    public class CorrectionFactorSettingsViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<CorrectionFactor> CorrectionFactors { get; } = new();
        
        public ICommand ResetToDefaultCommand { get; }
        public ICommand ApplyInflationCommand { get; }
        public ICommand OkCommand { get; }
        public ICommand CancelCommand { get; }

        public bool DialogResult { get; private set; }

        public CorrectionFactorSettingsViewModel(CorrectionFactorSettings settings)
        {
            LoadFromSettings(settings);
            
            ResetToDefaultCommand = new RelayCommand(_ => ResetToDefault());
            ApplyInflationCommand = new RelayCommand(_ => ApplyInflation());
            OkCommand = new RelayCommand(_ => { DialogResult = true; CloseRequested?.Invoke(); });
            CancelCommand = new RelayCommand(_ => { DialogResult = false; CloseRequested?.Invoke(); });
        }

        public event Action? CloseRequested;

        private void LoadFromSettings(CorrectionFactorSettings settings)
        {
            CorrectionFactors.Clear();
            var currentYear = DateTime.Now.Year;
            
            for (int year = 1999; year <= currentYear; year++)
            {
                var factor = settings.GetFactorForYear(year);
                CorrectionFactors.Add(new CorrectionFactor { Year = year, Factor = factor });
            }
        }

        public CorrectionFactorSettings GetSettings()
        {
            var settings = new CorrectionFactorSettings();
            foreach (var factor in CorrectionFactors)
            {
                settings.SetFactorForYear(factor.Year, factor.Factor);
            }
            return settings;
        }

        private void ResetToDefault()
        {
            foreach (var factor in CorrectionFactors)
            {
                factor.Factor = 1.0;
            }
        }

        private void ApplyInflation()
        {
            // Apply a simple inflation model - 2% per year from 1999
            var baseYear = 1999;
            var inflationRate = 0.02; // 2% per year
            
            foreach (var factor in CorrectionFactors)
            {
                var yearsFromBase = factor.Year - baseYear;
                factor.Factor = Math.Pow(1 + inflationRate, yearsFromBase);
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}

using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Microsoft.Win32;
using System.Windows;
using CostsViewer.Models;
using CostsViewer.Services;

namespace CostsViewer.ViewModels
{
    public class CorrectionFactorSettingsViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<CorrectionFactor> CorrectionFactors { get; } = new();
        
        public ICommand ResetToDefaultCommand { get; }
        public ICommand ApplyInflationCommand { get; }
        public ICommand ImportFromExcelCommand { get; }
        public ICommand ExportTemplateCommand { get; }
        public ICommand OkCommand { get; }
        public ICommand CancelCommand { get; }

        public bool DialogResult { get; private set; }

        public CorrectionFactorSettingsViewModel(CorrectionFactorSettings settings)
        {
            LoadFromSettings(settings);
            
            ResetToDefaultCommand = new RelayCommand(_ => ResetToDefault());
            ApplyInflationCommand = new RelayCommand(_ => ApplyInflation());
            ImportFromExcelCommand = new RelayCommand(_ => ImportFromExcel());
            ExportTemplateCommand = new RelayCommand(_ => ExportTemplate());
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

        private void ImportFromExcel()
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Title = "Import Correction Factors from Excel",
                    Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    FilterIndex = 1
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    var importedSettings = CorrectionFactorService.ImportFromExcel(openFileDialog.FileName);
                    LoadFromSettings(importedSettings);
                    
                    MessageBox.Show(
                        $"Successfully imported {importedSettings.YearFactors.Count} correction factors from Excel file.",
                        "Import Successful",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error importing correction factors:\n\n{ex.Message}",
                    "Import Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void ExportTemplate()
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Title = "Export Correction Factors Template",
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FilterIndex = 1,
                    FileName = "CorrectionFactors_Template.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    CorrectionFactorService.CreateExcelTemplate(saveFileDialog.FileName);
                    
                    MessageBox.Show(
                        $"Excel template created successfully:\n{saveFileDialog.FileName}",
                        "Template Created",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error creating Excel template:\n\n{ex.Message}",
                    "Export Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}

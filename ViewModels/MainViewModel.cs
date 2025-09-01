using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Data;
using System.Windows.Input;
using CostsViewer.Models;
using CostsViewer.Services;

namespace CostsViewer.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly ObservableCollection<ProjectRecord> _projects = new();
        public ICollectionView ProjectsView { get; }

        public ObservableCollection<string> AllProjectTypes { get; } = new();
        public IList<object> SelectedProjectTypes { get; } = new ObservableCollection<object>();

        private int? _minArea;
        public int? MinArea { get => _minArea; set { _minArea = value; OnPropertyChanged(); RefreshView(); } }

        private int? _maxArea;
        public int? MaxArea { get => _maxArea; set { _maxArea = value; OnPropertyChanged(); RefreshView(); } }

        public ICommand LoadCsvCommand { get; }
        public ICommand ApplyFilterCommand { get; }
        public ICommand ResetFilterCommand { get; }
        public ICommand IncludeMatchesCommand { get; }
        public ICommand ExcludeMatchesCommand { get; }
        public ICommand ExportExcelCommand { get; }
        public ICommand ExportPdfCommand { get; }

        public MainViewModel()
        {
            ProjectsView = CollectionViewSource.GetDefaultView(_projects);
            ProjectsView.Filter = FilterProject;

            LoadCsvCommand = new RelayCommand(_ => LoadCsv());
            ApplyFilterCommand = new RelayCommand(_ => RefreshView());
            ResetFilterCommand = new RelayCommand(_ => ResetFilters());
            IncludeMatchesCommand = new RelayCommand(_ => SetIncludeForMatches(true));
            ExcludeMatchesCommand = new RelayCommand(_ => SetIncludeForMatches(false));
            ExportExcelCommand = new RelayCommand(_ => ExportExcel());
            ExportPdfCommand = new RelayCommand(_ => ExportPdf());
        }

        private void LoadCsv()
        {
            try
            {
                var path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "2025-08-25_Sample_Table_Costs.csv");
                var fullPath = System.IO.Path.GetFullPath(path);
                var items = CsvLoader.Load(fullPath);
                _projects.Clear();
                foreach (var it in items) _projects.Add(it);
                RebuildProjectTypes();
                RefreshView();
                OnPropertyChanged(nameof(IncludedCount));
                UpdateAverages();
            }
            catch { }
        }

        private void RebuildProjectTypes()
        {
            AllProjectTypes.Clear();
            foreach (var t in _projects.SelectMany(p => p.ProjectTypes).Distinct().OrderBy(s => s))
                AllProjectTypes.Add(t);
        }

        private bool FilterProject(object obj)
        {
            if (obj is not ProjectRecord p) return false;
            if (MinArea.HasValue && p.TotalArea < MinArea.Value) return false;
            if (MaxArea.HasValue && p.TotalArea > MaxArea.Value) return false;
            if (SelectedProjectTypes.Count > 0)
            {
                var set = SelectedProjectTypes.Cast<string>().ToHashSet();
                if (!p.ProjectTypes.Any(set.Contains)) return false;
            }
            return true;
        }

        private void ResetFilters()
        {
            MinArea = null;
            MaxArea = null;
            SelectedProjectTypes.Clear();
            RefreshView();
        }

        private void SetIncludeForMatches(bool include)
        {
            foreach (var obj in _projects.Where(p => FilterProject(p)))
                obj.Include = include;
            UpdateAverages();
            OnPropertyChanged(nameof(IncludedCount));
        }

        private IEnumerable<ProjectRecord> Included => _projects.Where(p => p.Include);

        public int IncludedCount => Included.Count();

        private double _averageArea;
        public double AverageArea { get => _averageArea; private set { _averageArea = value; OnPropertyChanged(); } }
        public double AverageKG220 { get; private set; }
        public double AverageKG410 { get; private set; }
        public double AverageKG420 { get; private set; }
        public double AverageKG434 { get; private set; }
        public double AverageKG430 { get; private set; }
        public double AverageKG440 { get; private set; }
        public double AverageKG450 { get; private set; }
        public double AverageKG460 { get; private set; }
        public double AverageKG480 { get; private set; }
        public double AverageKG550 { get; private set; }

        private void UpdateAverages()
        {
            var list = Included.ToList();
            if (list.Count == 0)
            {
                AverageArea = 0;
                AverageKG220 = AverageKG410 = AverageKG420 = AverageKG434 = AverageKG430 = AverageKG440 = AverageKG450 = AverageKG460 = AverageKG480 = AverageKG550 = 0;
            }
            else
            {
                AverageArea = list.Average(p => p.TotalArea);
                AverageKG220 = list.Average(p => p.CostPerSqmKG220);
                AverageKG410 = list.Average(p => p.CostPerSqmKG410);
                AverageKG420 = list.Average(p => p.CostPerSqmKG420);
                AverageKG434 = list.Average(p => p.CostPerSqmKG434);
                AverageKG430 = list.Average(p => p.CostPerSqmKG430);
                AverageKG440 = list.Average(p => p.CostPerSqmKG440);
                AverageKG450 = list.Average(p => p.CostPerSqmKG450);
                AverageKG460 = list.Average(p => p.CostPerSqmKG460);
                AverageKG480 = list.Average(p => p.CostPerSqmKG480);
                AverageKG550 = list.Average(p => p.CostPerSqmKG550);
            }

            OnPropertyChanged(nameof(AverageKG220));
            OnPropertyChanged(nameof(AverageKG410));
            OnPropertyChanged(nameof(AverageKG420));
            OnPropertyChanged(nameof(AverageKG434));
            OnPropertyChanged(nameof(AverageKG430));
            OnPropertyChanged(nameof(AverageKG440));
            OnPropertyChanged(nameof(AverageKG450));
            OnPropertyChanged(nameof(AverageKG460));
            OnPropertyChanged(nameof(AverageKG480));
            OnPropertyChanged(nameof(AverageKG550));
        }

        private void RefreshView()
        {
            ProjectsView.Refresh();
            UpdateAverages();
            OnPropertyChanged(nameof(IncludedCount));
        }

        private void ExportExcel()
        {
            try { ExportServices.ExcelExporter.Export(Included.ToList()); } catch { }
        }

        private void ExportPdf()
        {
            try { ExportServices.PdfExporter.Export(Included.ToList(), AverageArea, AverageKG220, AverageKG410, AverageKG420, AverageKG434, AverageKG430, AverageKG440, AverageKG450, AverageKG460, AverageKG480, AverageKG550); } catch { }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action<object?> _execute;
        private readonly Func<object?, bool>? _canExecute;
        public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
        { _execute = execute; _canExecute = canExecute; }
        public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;
        public void Execute(object? parameter) => _execute(parameter);
        public event EventHandler? CanExecuteChanged { add { CommandManager.RequerySuggested += value; } remove { CommandManager.RequerySuggested -= value; } }
    }
}



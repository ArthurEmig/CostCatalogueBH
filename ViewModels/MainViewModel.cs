using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Data;
using System.Windows.Input;
using CostsViewer.Models;
using CostsViewer.Services;
using CostsViewer.ExportServices;

namespace CostsViewer.ViewModels
{
    public enum ProjectTypeMatchMode
    {
        Any,
        All
    }

    public class MainViewModel : INotifyPropertyChanged
    {
        private readonly ObservableCollection<ProjectRecord> _projects = new();
        public ICollectionView ProjectsView { get; }

        public ObservableCollection<string> AllProjectTypes { get; } = new();
        public IList<object> SelectedProjectTypes { get; } = new ObservableCollection<object>();
        public Array ProjectTypeMatchModes => Enum.GetValues(typeof(ProjectTypeMatchMode));
        public ObservableCollection<CostGroupSummary> CostGroupSummary { get; } = new();

        private ProjectTypeMatchMode _projectTypeMatchMode = ProjectTypeMatchMode.Any;
        public ProjectTypeMatchMode ProjectTypeMatchMode
        {
            get => _projectTypeMatchMode;
            set 
            { 
                Console.WriteLine($"ProjectTypeMatchMode: Changed from {_projectTypeMatchMode} to {value}");
                _projectTypeMatchMode = value; 
                OnPropertyChanged(); 
                RefreshView(); 
            }
        }

        private int? _minArea;
        public int? MinArea 
        { 
            get => _minArea; 
            set 
            { 
                Console.WriteLine($"MinArea: Changed from {_minArea} to {value}");
                _minArea = value; 
                OnPropertyChanged(); 
                RefreshView(); 
            } 
        }

        private int? _maxArea;
        public int? MaxArea 
        { 
            get => _maxArea; 
            set 
            { 
                Console.WriteLine($"MaxArea: Changed from {_maxArea} to {value}");
                _maxArea = value; 
                OnPropertyChanged(); 
                RefreshView(); 
            } 
        }

        public ICommand LoadFileCommand { get; }
        public ICommand ApplyFilterCommand { get; }
        public ICommand ResetFilterCommand { get; }
        public ICommand IncludeMatchesCommand { get; }
        public ICommand ExcludeMatchesCommand { get; }
        public ICommand ExportExcelCommand { get; }
        public ICommand ExportPdfCommand { get; }

        public MainViewModel()
        {
            Console.WriteLine("=== MainViewModel: Constructor starting ===");
            ProjectsView = CollectionViewSource.GetDefaultView(_projects);
            ProjectsView.Filter = FilterProject;
            Console.WriteLine("MainViewModel: ProjectsView created and filter set");

            _projects.CollectionChanged += OnProjectsCollectionChanged;
            Console.WriteLine("MainViewModel: Projects collection change handler attached");

            if (SelectedProjectTypes is INotifyCollectionChanged selectedTypesChanges)
            {
                selectedTypesChanges.CollectionChanged += (_, __) => RefreshView();
                Console.WriteLine("MainViewModel: SelectedProjectTypes change handler attached");
            }

            LoadFileCommand = new RelayCommand(_ => LoadFile());
            ApplyFilterCommand = new RelayCommand(_ => RefreshView());
            ResetFilterCommand = new RelayCommand(_ => ResetFilters());
            IncludeMatchesCommand = new RelayCommand(_ => SetIncludeForMatches(true));
            ExcludeMatchesCommand = new RelayCommand(_ => SetIncludeForMatches(false));
            ExportExcelCommand = new RelayCommand(_ => ExportExcel());
            ExportPdfCommand = new RelayCommand(_ => ExportPdf());
            Console.WriteLine("MainViewModel: All commands initialized");
            Console.WriteLine("=== MainViewModel: Constructor completed ===");
        }

        private void OnProjectsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Reset)
            {
                foreach (var rec in _projects)
                {
                    rec.PropertyChanged -= OnProjectPropertyChanged;
                    rec.PropertyChanged += OnProjectPropertyChanged;
                }
                UpdateAverages();
                OnPropertyChanged(nameof(IncludedCount));
                return;
            }

            if (e.OldItems != null)
            {
                foreach (var item in e.OldItems)
                {
                    if (item is ProjectRecord oldRec)
                    {
                        oldRec.PropertyChanged -= OnProjectPropertyChanged;
                    }
                }
            }

            if (e.NewItems != null)
            {
                foreach (var item in e.NewItems)
                {
                    if (item is ProjectRecord newRec)
                    {
                        newRec.PropertyChanged += OnProjectPropertyChanged;
                    }
                }
            }
            UpdateAverages();
            OnPropertyChanged(nameof(IncludedCount));
        }

        private void OnProjectPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ProjectRecord.Include))
            {
                var project = sender as ProjectRecord;
                Console.WriteLine($"OnProjectPropertyChanged: Project {project?.ProjectId} Include changed to {project?.Include}");
                UpdateAverages();
                OnPropertyChanged(nameof(IncludedCount));
            }
        }

        private void LoadFile()
        {
            try
            {
                Console.WriteLine("=== LoadFile: Starting file load process ===");
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Select costs file",
                    Filter = "Supported files (*.csv;*.xlsx)|*.csv;*.xlsx|CSV files (*.csv)|*.csv|Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    CheckFileExists = true,
                    Multiselect = false
                };

                var result = dlg.ShowDialog();
                if (result != true) 
                {
                    Console.WriteLine("LoadFile: User cancelled file dialog");
                    return;
                }

                Console.WriteLine($"LoadFile: Loading file: {dlg.FileName}");
                
                List<ProjectRecord> items;
                var extension = Path.GetExtension(dlg.FileName).ToLowerInvariant();
                
                switch (extension)
                {
                    case ".csv":
                        Console.WriteLine("LoadFile: Loading as CSV file");
                        items = CsvLoader.Load(dlg.FileName);
                        break;
                    case ".xlsx":
                        Console.WriteLine("LoadFile: Loading as Excel file");
                        items = ExcelLoader.Load(dlg.FileName);
                        break;
                    default:
                        Console.WriteLine($"LoadFile: Unsupported file extension: {extension}");
                        throw new NotSupportedException($"Unsupported file format: {extension}. Please select a CSV (.csv) or Excel (.xlsx) file.");
                }
                
                Console.WriteLine($"LoadFile: Loaded {items.Count()} items from {extension.ToUpper()} file");
                
                _projects.Clear();
                foreach (var it in items) _projects.Add(it);
                Console.WriteLine($"LoadFile: Added {_projects.Count} projects to collection");
                
                RebuildProjectTypes();
                RefreshView();
                OnPropertyChanged(nameof(IncludedCount));
                UpdateAverages();
                Console.WriteLine("LoadFile: Completed file load process");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"LoadFile: ERROR - {ex.Message}");
                Console.WriteLine($"LoadFile: Stack trace - {ex.StackTrace}");
            }
        }

        private void RebuildProjectTypes()
        {
            Console.WriteLine("RebuildProjectTypes: Starting rebuild");
            AllProjectTypes.Clear();
            AllProjectTypes.Add("All types");
            var projectTypes = _projects.SelectMany(p => p.ProjectTypes).Distinct().OrderBy(s => s).ToList();
            Console.WriteLine($"RebuildProjectTypes: Found {projectTypes.Count} unique project types");
            foreach (var t in projectTypes)
            {
                AllProjectTypes.Add(t);
                Console.WriteLine($"RebuildProjectTypes: Added type '{t}'");
            }
            Console.WriteLine($"RebuildProjectTypes: Total types in collection: {AllProjectTypes.Count}");
        }

        private bool FilterProject(object obj)
        {
            if (obj is not ProjectRecord p) 
            {
                Console.WriteLine("FilterProject: Object is not ProjectRecord");
                return false;
            }
            
            if (MinArea.HasValue && p.TotalArea < MinArea.Value) 
            {
                Console.WriteLine($"FilterProject: Project {p.ProjectId} filtered out by MinArea ({p.TotalArea} < {MinArea.Value})");
                return false;
            }
            
            if (MaxArea.HasValue && p.TotalArea > MaxArea.Value) 
            {
                Console.WriteLine($"FilterProject: Project {p.ProjectId} filtered out by MaxArea ({p.TotalArea} > {MaxArea.Value})");
                return false;
            }
            
            if (SelectedProjectTypes.Count > 0)
            {
                var selected = SelectedProjectTypes.Cast<string>().ToList();
                Console.WriteLine($"FilterProject: Project {p.ProjectId} - Selected types: [{string.Join(", ", selected)}]");
                Console.WriteLine($"FilterProject: Project {p.ProjectId} - Project types: [{string.Join(", ", p.ProjectTypes)}]");
                
                // If "All types" is selected, skip project type filtering
                if (!selected.Contains("All types"))
                {
                    if (ProjectTypeMatchMode == ProjectTypeMatchMode.All)
                    {
                        var projectTypeSet = p.ProjectTypes.ToHashSet();
                        if (!selected.All(projectTypeSet.Contains)) 
                        {
                            Console.WriteLine($"FilterProject: Project {p.ProjectId} filtered out by ProjectType (All mode - missing types)");
                            return false;
                        }
                    }
                    else
                    {
                        if (!p.ProjectTypes.Any(t => selected.Contains(t))) 
                        {
                            Console.WriteLine($"FilterProject: Project {p.ProjectId} filtered out by ProjectType (Any mode - no matching types)");
                            return false;
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"FilterProject: Project {p.ProjectId} - 'All types' selected, skipping type filtering");
                }
            }
            
            Console.WriteLine($"FilterProject: Project {p.ProjectId} passed all filters");
            return true;
        }

        private void ResetFilters()
        {
            Console.WriteLine("=== ResetFilters: Resetting all filters ===");
            Console.WriteLine($"ResetFilters: Before - MinArea: {MinArea}, MaxArea: {MaxArea}, SelectedTypes: {SelectedProjectTypes.Count}");
            
            MinArea = null;
            MaxArea = null;
            SelectedProjectTypes.Clear();
            ProjectTypeMatchMode = ProjectTypeMatchMode.Any;
            
            Console.WriteLine("ResetFilters: All filters cleared, refreshing view");
            RefreshView();
            Console.WriteLine("ResetFilters: Completed filter reset");
        }

        private void SetIncludeForMatches(bool include)
        {
            Console.WriteLine($"=== SetIncludeForMatches: Setting Include={include} for all filtered items ===");
            var items = ProjectsView.Cast<ProjectRecord>().ToList();
            Console.WriteLine($"SetIncludeForMatches: Found {items.Count} filtered items");
            
            foreach (var obj in items)
            {
                Console.WriteLine($"SetIncludeForMatches: Setting Project {obj.ProjectId} Include to {include}");
                obj.Include = include;
            }
            
            UpdateAverages();
            OnPropertyChanged(nameof(IncludedCount));
            Console.WriteLine("SetIncludeForMatches: Completed batch include/exclude operation");
        }

        private IEnumerable<ProjectRecord> FilteredItems => ProjectsView.Cast<ProjectRecord>();
        private IEnumerable<ProjectRecord> FilteredIncluded => FilteredItems.Where(p => p.Include);

        public int IncludedCount => FilteredIncluded.Count();

        private double _averageArea;
        public double AverageArea { get => _averageArea; private set { _averageArea = value; OnPropertyChanged(); } }
        public double AverageKG220 { get; private set; }
        public double AverageKG230 { get; private set; }
        public double AverageKG410 { get; private set; }
        public double AverageKG420 { get; private set; }
        public double AverageKG434 { get; private set; }
        public double AverageKG430 { get; private set; }
        public double AverageKG440 { get; private set; }
        public double AverageKG450 { get; private set; }
        public double AverageKG460 { get; private set; }
        public double AverageKG474 { get; private set; }
        public double AverageKG475 { get; private set; }
        public double AverageKG480 { get; private set; }
        public double AverageKG550 { get; private set; }

        // Min values for each cost group
        public double MinKG220 { get; private set; }
        public double MinKG230 { get; private set; }
        public double MinKG410 { get; private set; }
        public double MinKG420 { get; private set; }
        public double MinKG434 { get; private set; }
        public double MinKG430 { get; private set; }
        public double MinKG440 { get; private set; }
        public double MinKG450 { get; private set; }
        public double MinKG460 { get; private set; }
        public double MinKG474 { get; private set; }
        public double MinKG475 { get; private set; }
        public double MinKG480 { get; private set; }
        public double MinKG550 { get; private set; }

        // Max values for each cost group
        public double MaxKG220 { get; private set; }
        public double MaxKG230 { get; private set; }
        public double MaxKG410 { get; private set; }
        public double MaxKG420 { get; private set; }
        public double MaxKG434 { get; private set; }
        public double MaxKG430 { get; private set; }
        public double MaxKG440 { get; private set; }
        public double MaxKG450 { get; private set; }
        public double MaxKG460 { get; private set; }
        public double MaxKG474 { get; private set; }
        public double MaxKG475 { get; private set; }
        public double MaxKG480 { get; private set; }
        public double MaxKG550 { get; private set; }

        // Standard Deviation values for each cost group
        public double StdDevKG220 { get; private set; }
        public double StdDevKG230 { get; private set; }
        public double StdDevKG410 { get; private set; }
        public double StdDevKG420 { get; private set; }
        public double StdDevKG434 { get; private set; }
        public double StdDevKG430 { get; private set; }
        public double StdDevKG440 { get; private set; }
        public double StdDevKG450 { get; private set; }
        public double StdDevKG460 { get; private set; }
        public double StdDevKG474 { get; private set; }
        public double StdDevKG475 { get; private set; }
        public double StdDevKG480 { get; private set; }
        public double StdDevKG550 { get; private set; }

        private void UpdateAverages()
        {
            Console.WriteLine("=== UpdateAverages: Starting calculation ===");
            var list = FilteredIncluded.ToList();
            Console.WriteLine($"UpdateAverages: Found {list.Count} included items for calculation");
            
            if (list.Count == 0)
            {
                Console.WriteLine("UpdateAverages: No items included, setting all values to 0");
                AverageArea = 0;
                AverageKG220 = AverageKG230 = AverageKG410 = AverageKG420 = AverageKG434 = AverageKG430 = AverageKG440 = AverageKG450 = AverageKG460 = AverageKG474 = AverageKG475 = AverageKG480 = AverageKG550 = 0;
                MinKG220 = MinKG230 = MinKG410 = MinKG420 = MinKG434 = MinKG430 = MinKG440 = MinKG450 = MinKG460 = MinKG474 = MinKG475 = MinKG480 = MinKG550 = 0;
                MaxKG220 = MaxKG230 = MaxKG410 = MaxKG420 = MaxKG434 = MaxKG430 = MaxKG440 = MaxKG450 = MaxKG460 = MaxKG474 = MaxKG475 = MaxKG480 = MaxKG550 = 0;
                StdDevKG220 = StdDevKG230 = StdDevKG410 = StdDevKG420 = StdDevKG434 = StdDevKG430 = StdDevKG440 = StdDevKG450 = StdDevKG460 = StdDevKG474 = StdDevKG475 = StdDevKG480 = StdDevKG550 = 0;
            }
            else
            {
                Console.WriteLine("UpdateAverages: Calculating averages, mins, and maxs...");
                foreach (var item in list)
                {
                    Console.WriteLine($"UpdateAverages: Item {item.ProjectId} - Include: {item.Include}, Area: {item.TotalArea}, KG220: {item.CostPerSqmKG220}");
                }
                
                AverageArea = list.Average(p => p.TotalArea);
                AverageKG220 = list.Average(p => p.CostPerSqmKG220);
                AverageKG230 = list.Average(p => p.CostPerSqmKG230);
                AverageKG410 = list.Average(p => p.CostPerSqmKG410);
                AverageKG420 = list.Average(p => p.CostPerSqmKG420);
                AverageKG434 = list.Average(p => p.CostPerSqmKG434);
                AverageKG430 = list.Average(p => p.CostPerSqmKG430);
                AverageKG440 = list.Average(p => p.CostPerSqmKG440);
                AverageKG450 = list.Average(p => p.CostPerSqmKG450);
                AverageKG460 = list.Average(p => p.CostPerSqmKG460);
                AverageKG474 = list.Average(p => p.CostPerSqmKG474);
                AverageKG475 = list.Average(p => p.CostPerSqmKG475);
                AverageKG480 = list.Average(p => p.CostPerSqmKG480);
                AverageKG550 = list.Average(p => p.CostPerSqmKG550);

                // Calculate Min values (excluding zeros)
                MinKG220 = list.Where(p => p.CostPerSqmKG220 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG220 ?? 0);
                MinKG230 = list.Where(p => p.CostPerSqmKG230 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG230 ?? 0);
                MinKG410 = list.Where(p => p.CostPerSqmKG410 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG410 ?? 0);
                MinKG420 = list.Where(p => p.CostPerSqmKG420 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG420 ?? 0);
                MinKG434 = list.Where(p => p.CostPerSqmKG434 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG434 ?? 0);
                MinKG430 = list.Where(p => p.CostPerSqmKG430 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG430 ?? 0);
                MinKG440 = list.Where(p => p.CostPerSqmKG440 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG440 ?? 0);
                MinKG450 = list.Where(p => p.CostPerSqmKG450 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG450 ?? 0);
                MinKG460 = list.Where(p => p.CostPerSqmKG460 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG460 ?? 0);
                MinKG474 = list.Where(p => p.CostPerSqmKG474 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG474 ?? 0);
                MinKG475 = list.Where(p => p.CostPerSqmKG475 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG475 ?? 0);
                MinKG480 = list.Where(p => p.CostPerSqmKG480 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG480 ?? 0);
                MinKG550 = list.Where(p => p.CostPerSqmKG550 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG550 ?? 0);

                // Calculate Max values
                MaxKG220 = list.Max(p => p.CostPerSqmKG220);
                MaxKG230 = list.Max(p => p.CostPerSqmKG230);
                MaxKG410 = list.Max(p => p.CostPerSqmKG410);
                MaxKG420 = list.Max(p => p.CostPerSqmKG420);
                MaxKG434 = list.Max(p => p.CostPerSqmKG434);
                MaxKG430 = list.Max(p => p.CostPerSqmKG430);
                MaxKG440 = list.Max(p => p.CostPerSqmKG440);
                MaxKG450 = list.Max(p => p.CostPerSqmKG450);
                MaxKG460 = list.Max(p => p.CostPerSqmKG460);
                MaxKG474 = list.Max(p => p.CostPerSqmKG474);
                MaxKG475 = list.Max(p => p.CostPerSqmKG475);
                MaxKG480 = list.Max(p => p.CostPerSqmKG480);
                MaxKG550 = list.Max(p => p.CostPerSqmKG550);

                // Calculate Standard Deviations (excluding zeros)
                StdDevKG220 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG220 > 0).Select(p => (double)p.CostPerSqmKG220).ToList(), list.Where(p => p.CostPerSqmKG220 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG220 ?? 0));
                StdDevKG230 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG230 > 0).Select(p => (double)p.CostPerSqmKG230).ToList(), list.Where(p => p.CostPerSqmKG230 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG230 ?? 0));
                StdDevKG410 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG410 > 0).Select(p => (double)p.CostPerSqmKG410).ToList(), list.Where(p => p.CostPerSqmKG410 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG410 ?? 0));
                StdDevKG420 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG420 > 0).Select(p => (double)p.CostPerSqmKG420).ToList(), list.Where(p => p.CostPerSqmKG420 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG420 ?? 0));
                StdDevKG434 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG434 > 0).Select(p => (double)p.CostPerSqmKG434).ToList(), list.Where(p => p.CostPerSqmKG434 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG434 ?? 0));
                StdDevKG430 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG430 > 0).Select(p => (double)p.CostPerSqmKG430).ToList(), list.Where(p => p.CostPerSqmKG430 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG430 ?? 0));
                StdDevKG440 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG440 > 0).Select(p => (double)p.CostPerSqmKG440).ToList(), list.Where(p => p.CostPerSqmKG440 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG440 ?? 0));
                StdDevKG450 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG450 > 0).Select(p => (double)p.CostPerSqmKG450).ToList(), list.Where(p => p.CostPerSqmKG450 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG450 ?? 0));
                StdDevKG460 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG460 > 0).Select(p => (double)p.CostPerSqmKG460).ToList(), list.Where(p => p.CostPerSqmKG460 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG460 ?? 0));
                StdDevKG474 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG474 > 0).Select(p => (double)p.CostPerSqmKG474).ToList(), list.Where(p => p.CostPerSqmKG474 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG474 ?? 0));
                StdDevKG475 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG475 > 0).Select(p => (double)p.CostPerSqmKG475).ToList(), list.Where(p => p.CostPerSqmKG475 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG475 ?? 0));
                StdDevKG480 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG480 > 0).Select(p => (double)p.CostPerSqmKG480).ToList(), list.Where(p => p.CostPerSqmKG480 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG480 ?? 0));
                StdDevKG550 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG550 > 0).Select(p => (double)p.CostPerSqmKG550).ToList(), list.Where(p => p.CostPerSqmKG550 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG550 ?? 0));
                
                Console.WriteLine($"UpdateAverages: Calculated - Area: {AverageArea:F2}, KG220: Avg={AverageKG220:F2}, Min={MinKG220:F2}, Max={MaxKG220:F2}, StdDev={StdDevKG220:F2}");
            }

            OnPropertyChanged(nameof(AverageKG220));
            OnPropertyChanged(nameof(AverageKG230));
            OnPropertyChanged(nameof(AverageKG410));
            OnPropertyChanged(nameof(AverageKG420));
            OnPropertyChanged(nameof(AverageKG434));
            OnPropertyChanged(nameof(AverageKG430));
            OnPropertyChanged(nameof(AverageKG440));
            OnPropertyChanged(nameof(AverageKG450));
            OnPropertyChanged(nameof(AverageKG460));
            OnPropertyChanged(nameof(AverageKG474));
            OnPropertyChanged(nameof(AverageKG475));
            OnPropertyChanged(nameof(AverageKG480));
            OnPropertyChanged(nameof(AverageKG550));

            OnPropertyChanged(nameof(MinKG220));
            OnPropertyChanged(nameof(MinKG230));
            OnPropertyChanged(nameof(MinKG410));
            OnPropertyChanged(nameof(MinKG420));
            OnPropertyChanged(nameof(MinKG434));
            OnPropertyChanged(nameof(MinKG430));
            OnPropertyChanged(nameof(MinKG440));
            OnPropertyChanged(nameof(MinKG450));
            OnPropertyChanged(nameof(MinKG460));
            OnPropertyChanged(nameof(MinKG474));
            OnPropertyChanged(nameof(MinKG475));
            OnPropertyChanged(nameof(MinKG480));
            OnPropertyChanged(nameof(MinKG550));

            OnPropertyChanged(nameof(MaxKG220));
            OnPropertyChanged(nameof(MaxKG230));
            OnPropertyChanged(nameof(MaxKG410));
            OnPropertyChanged(nameof(MaxKG420));
            OnPropertyChanged(nameof(MaxKG434));
            OnPropertyChanged(nameof(MaxKG430));
            OnPropertyChanged(nameof(MaxKG440));
            OnPropertyChanged(nameof(MaxKG450));
            OnPropertyChanged(nameof(MaxKG460));
            OnPropertyChanged(nameof(MaxKG474));
            OnPropertyChanged(nameof(MaxKG475));
            OnPropertyChanged(nameof(MaxKG480));
            OnPropertyChanged(nameof(MaxKG550));

            OnPropertyChanged(nameof(StdDevKG220));
            OnPropertyChanged(nameof(StdDevKG230));
            OnPropertyChanged(nameof(StdDevKG410));
            OnPropertyChanged(nameof(StdDevKG420));
            OnPropertyChanged(nameof(StdDevKG434));
            OnPropertyChanged(nameof(StdDevKG430));
            OnPropertyChanged(nameof(StdDevKG440));
            OnPropertyChanged(nameof(StdDevKG450));
            OnPropertyChanged(nameof(StdDevKG460));
            OnPropertyChanged(nameof(StdDevKG474));
            OnPropertyChanged(nameof(StdDevKG475));
            OnPropertyChanged(nameof(StdDevKG480));
            OnPropertyChanged(nameof(StdDevKG550));
            
            UpdateCostGroupSummary();
            Console.WriteLine("UpdateAverages: Completed calculation and property notifications");
        }

        private void UpdateCostGroupSummary()
        {
            Console.WriteLine("=== UpdateCostGroupSummary: Starting summary calculation ===");
            var includedItems = FilteredIncluded.ToList();
            Console.WriteLine($"UpdateCostGroupSummary: Using {includedItems.Count} included items");

            CostGroupSummary.Clear();

            if (includedItems.Count == 0)
            {
                Console.WriteLine("UpdateCostGroupSummary: No items to summarize");
                return;
            }

            // Define cost groups with descriptions according to DIN 276
            var costGroups = new[]
            {
                new { Code = "KG220", Description = "Site Clearance & Preparation", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG220) },
                new { Code = "KG230", Description = "Earthworks & Foundations", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG230) },
                new { Code = "KG410", Description = "Sewage, Water & Gas Systems", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG410) },
                new { Code = "KG420", Description = "Heating Systems", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG420) },
                new { Code = "KG430", Description = "Ventilation & Air Conditioning", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG430) },
                new { Code = "KG434", Description = "Process-Specific Installations", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG434) },
                new { Code = "KG440", Description = "Electrical Systems", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG440) },
                new { Code = "KG450", Description = "Communication & Safety Systems", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG450) },
                new { Code = "KG460", Description = "Conveying Systems", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG460) },
                new { Code = "KG474", Description = "Fire Protection Systems", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG474) },
                new { Code = "KG475", Description = "Security & Access Control", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG475) },
                new { Code = "KG480", Description = "Building & System Automation", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG480) },
                new { Code = "KG550", Description = "Outdoor Technical Installations", GetValue = (Func<ProjectRecord, double>)(p => p.CostPerSqmKG550) }
            };

            foreach (var costGroup in costGroups)
            {
                var values = includedItems.Select(costGroup.GetValue).Where(v => v > 0).ToList();
                
                if (values.Count > 0)
                {
                    var average = values.Average();
                    var min = values.Min();
                    var max = values.Max();
                    var standardDeviation = CalculateStandardDeviation(values, average);

                    var summary = new CostGroupSummary
                    {
                        CostGroup = costGroup.Code,
                        Description = costGroup.Description,
                        AverageCost = average,
                        MinCost = min,
                        MaxCost = max,
                        StandardDeviation = standardDeviation
                    };

                    CostGroupSummary.Add(summary);
                    Console.WriteLine($"UpdateCostGroupSummary: {costGroup.Code} - Avg: {average:F2}, Min: {min:F2}, Max: {max:F2}, StdDev: {standardDeviation:F2}");
                }
                else
                {
                    Console.WriteLine($"UpdateCostGroupSummary: {costGroup.Code} - No valid data (all zeros)");
                }
            }

            Console.WriteLine($"UpdateCostGroupSummary: Created {CostGroupSummary.Count} cost group summaries");
        }

        private static double CalculateStandardDeviation(List<double> values, double mean)
        {
            if (values.Count <= 1) return 0;

            var sumOfSquaredDifferences = values.Sum(v => Math.Pow(v - mean, 2));
            var variance = sumOfSquaredDifferences / (values.Count - 1); // Sample standard deviation
            return Math.Sqrt(variance);
        }

        private void RefreshView()
        {
            Console.WriteLine("=== RefreshView: Starting view refresh ===");
            if (ProjectsView is IEditableCollectionView editable)
            {
                if (editable.IsAddingNew)
                {
                    Console.WriteLine("RefreshView: Committing new item");
                    try { editable.CommitNew(); } catch { }
                }
                if (editable.IsEditingItem)
                {
                    Console.WriteLine("RefreshView: Committing edit");
                    try { editable.CommitEdit(); } catch { }
                }
            }

            Console.WriteLine("RefreshView: Refreshing ProjectsView");
            ProjectsView.Refresh();
            UpdateAverages();
            OnPropertyChanged(nameof(IncludedCount));
            Console.WriteLine("RefreshView: Completed view refresh");
        }

        private void ExportExcel()
        {
            try 
            { 
                var exportItems = FilteredIncluded.ToList();
                var summaryItems = CostGroupSummary.ToList();
                Console.WriteLine($"ExportExcel: Exporting {exportItems.Count} included items and {summaryItems.Count} cost group summaries to Excel");
                ExportServices.ExcelExporter.Export(exportItems, summaryItems);
                Console.WriteLine("ExportExcel: Excel export completed successfully");
            } 
            catch (Exception ex)
            {
                Console.WriteLine($"ExportExcel: ERROR - {ex.Message}");
            }
        }

        private void ExportPdf()
        {
            try 
            { 
                var exportItems = FilteredIncluded.ToList();
                var summaryItems = CostGroupSummary.ToList();
                Console.WriteLine($"ExportPdf: Exporting {exportItems.Count} included items and {summaryItems.Count} cost group summaries to PDF");
                Console.WriteLine($"ExportPdf: Using averages - Area: {AverageArea:F2}, KG220: {AverageKG220:F2}");
                ExportServices.PdfExporter.Export(exportItems, summaryItems, AverageArea, AverageKG220, AverageKG410, AverageKG420, AverageKG434, AverageKG430, AverageKG440, AverageKG450, AverageKG460, AverageKG480, AverageKG550);
                Console.WriteLine("ExportPdf: PDF export completed successfully");
            } 
            catch (Exception ex)
            {
                Console.WriteLine($"ExportPdf: ERROR - {ex.Message}");
            }
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



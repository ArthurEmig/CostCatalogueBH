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
using CostsViewer.Views;
using CostsViewer.ViewModels;

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
        public ICommand CorrectionFactorSettingsCommand { get; }

        private CorrectionFactorSettings _correctionFactorSettings;

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

            // Load correction factor settings
            _correctionFactorSettings = CorrectionFactorService.LoadSettings();
            Console.WriteLine("MainViewModel: Correction factor settings loaded");

            LoadFileCommand = new RelayCommand(_ => LoadFile());
            ApplyFilterCommand = new RelayCommand(_ => RefreshView());
            ResetFilterCommand = new RelayCommand(_ => ResetFilters());
            IncludeMatchesCommand = new RelayCommand(_ => SetIncludeForMatches(true));
            ExcludeMatchesCommand = new RelayCommand(_ => SetIncludeForMatches(false));
            ExportExcelCommand = new RelayCommand(_ => ExportExcel());
            ExportPdfCommand = new RelayCommand(_ => ExportPdf());
            CorrectionFactorSettingsCommand = new RelayCommand(_ => OpenCorrectionFactorSettings());
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
            Console.WriteLine($"OnProjectPropertyChanged: PropertyName = {e.PropertyName}");
            if (e.PropertyName == nameof(ProjectRecord.Include))
            {
                var project = sender as ProjectRecord;
                Console.WriteLine($"OnProjectPropertyChanged: Project {project?.ProjectId} Include changed to {project?.Include}");
                Console.WriteLine("OnProjectPropertyChanged: Calling UpdateAverages()");
                UpdateAverages();
                OnPropertyChanged(nameof(IncludedCount));
                Console.WriteLine("OnProjectPropertyChanged: Completed");
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
        
        private double _averageKG220;
        public double AverageKG220 { get => _averageKG220; private set { 
            Console.WriteLine($"AverageKG220 setter: {_averageKG220} -> {value}"); 
            _averageKG220 = value; 
            OnPropertyChanged(); 
        } }
        
        private double _averageKG230;
        public double AverageKG230 { get => _averageKG230; private set { _averageKG230 = value; OnPropertyChanged(); } }
        
        private double _averageKG410;
        public double AverageKG410 { get => _averageKG410; private set { _averageKG410 = value; OnPropertyChanged(); } }
        
        private double _averageKG420;
        public double AverageKG420 { get => _averageKG420; private set { _averageKG420 = value; OnPropertyChanged(); } }
        
        private double _averageKG434;
        public double AverageKG434 { get => _averageKG434; private set { _averageKG434 = value; OnPropertyChanged(); } }
        
        private double _averageKG430;
        public double AverageKG430 { get => _averageKG430; private set { _averageKG430 = value; OnPropertyChanged(); } }
        
        private double _averageKG440;
        public double AverageKG440 { get => _averageKG440; private set { _averageKG440 = value; OnPropertyChanged(); } }
        
        private double _averageKG450;
        public double AverageKG450 { get => _averageKG450; private set { _averageKG450 = value; OnPropertyChanged(); } }
        
        private double _averageKG460;
        public double AverageKG460 { get => _averageKG460; private set { _averageKG460 = value; OnPropertyChanged(); } }
        
        private double _averageKG490;
        public double AverageKG490 { get => _averageKG490; private set { _averageKG490 = value; OnPropertyChanged(); } }
        
        private double _averageKG474;
        public double AverageKG474 { get => _averageKG474; private set { _averageKG474 = value; OnPropertyChanged(); } }
        
        private double _averageKG475;
        public double AverageKG475 { get => _averageKG475; private set { _averageKG475 = value; OnPropertyChanged(); } }
        
        private double _averageKG480;
        public double AverageKG480 { get => _averageKG480; private set { _averageKG480 = value; OnPropertyChanged(); } }
        
        private double _averageKG550;
        public double AverageKG550 { get => _averageKG550; private set { _averageKG550 = value; OnPropertyChanged(); } }

        // Min values for each cost group
        private double _minKG220;
        public double MinKG220 { get => _minKG220; private set { _minKG220 = value; OnPropertyChanged(); } }
        
        private double _minKG230;
        public double MinKG230 { get => _minKG230; private set { _minKG230 = value; OnPropertyChanged(); } }
        
        private double _minKG410;
        public double MinKG410 { get => _minKG410; private set { _minKG410 = value; OnPropertyChanged(); } }
        
        private double _minKG420;
        public double MinKG420 { get => _minKG420; private set { _minKG420 = value; OnPropertyChanged(); } }
        
        private double _minKG434;
        public double MinKG434 { get => _minKG434; private set { _minKG434 = value; OnPropertyChanged(); } }
        
        private double _minKG430;
        public double MinKG430 { get => _minKG430; private set { _minKG430 = value; OnPropertyChanged(); } }
        
        private double _minKG440;
        public double MinKG440 { get => _minKG440; private set { _minKG440 = value; OnPropertyChanged(); } }
        
        private double _minKG450;
        public double MinKG450 { get => _minKG450; private set { _minKG450 = value; OnPropertyChanged(); } }
        
        private double _minKG460;
        public double MinKG460 { get => _minKG460; private set { _minKG460 = value; OnPropertyChanged(); } }
        
        private double _minKG490;
        public double MinKG490 { get => _minKG490; private set { _minKG490 = value; OnPropertyChanged(); } }
        
        private double _minKG474;
        public double MinKG474 { get => _minKG474; private set { _minKG474 = value; OnPropertyChanged(); } }
        
        private double _minKG475;
        public double MinKG475 { get => _minKG475; private set { _minKG475 = value; OnPropertyChanged(); } }
        
        private double _minKG480;
        public double MinKG480 { get => _minKG480; private set { _minKG480 = value; OnPropertyChanged(); } }
        
        private double _minKG550;
        public double MinKG550 { get => _minKG550; private set { _minKG550 = value; OnPropertyChanged(); } }

        // Max values for each cost group
        private double _maxKG220;
        public double MaxKG220 { get => _maxKG220; private set { _maxKG220 = value; OnPropertyChanged(); } }
        
        private double _maxKG230;
        public double MaxKG230 { get => _maxKG230; private set { _maxKG230 = value; OnPropertyChanged(); } }
        
        private double _maxKG410;
        public double MaxKG410 { get => _maxKG410; private set { _maxKG410 = value; OnPropertyChanged(); } }
        
        private double _maxKG420;
        public double MaxKG420 { get => _maxKG420; private set { _maxKG420 = value; OnPropertyChanged(); } }
        
        private double _maxKG434;
        public double MaxKG434 { get => _maxKG434; private set { _maxKG434 = value; OnPropertyChanged(); } }
        
        private double _maxKG430;
        public double MaxKG430 { get => _maxKG430; private set { _maxKG430 = value; OnPropertyChanged(); } }
        
        private double _maxKG440;
        public double MaxKG440 { get => _maxKG440; private set { _maxKG440 = value; OnPropertyChanged(); } }
        
        private double _maxKG450;
        public double MaxKG450 { get => _maxKG450; private set { _maxKG450 = value; OnPropertyChanged(); } }
        
        private double _maxKG460;
        public double MaxKG460 { get => _maxKG460; private set { _maxKG460 = value; OnPropertyChanged(); } }
        
        private double _maxKG490;
        public double MaxKG490 { get => _maxKG490; private set { _maxKG490 = value; OnPropertyChanged(); } }
        
        private double _maxKG474;
        public double MaxKG474 { get => _maxKG474; private set { _maxKG474 = value; OnPropertyChanged(); } }
        
        private double _maxKG475;
        public double MaxKG475 { get => _maxKG475; private set { _maxKG475 = value; OnPropertyChanged(); } }
        
        private double _maxKG480;
        public double MaxKG480 { get => _maxKG480; private set { _maxKG480 = value; OnPropertyChanged(); } }
        
        private double _maxKG550;
        public double MaxKG550 { get => _maxKG550; private set { _maxKG550 = value; OnPropertyChanged(); } }

        // Standard Deviation values for each cost group
        private double _stdDevKG220;
        public double StdDevKG220 { get => _stdDevKG220; private set { _stdDevKG220 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG230;
        public double StdDevKG230 { get => _stdDevKG230; private set { _stdDevKG230 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG410;
        public double StdDevKG410 { get => _stdDevKG410; private set { _stdDevKG410 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG420;
        public double StdDevKG420 { get => _stdDevKG420; private set { _stdDevKG420 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG434;
        public double StdDevKG434 { get => _stdDevKG434; private set { _stdDevKG434 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG430;
        public double StdDevKG430 { get => _stdDevKG430; private set { _stdDevKG430 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG440;
        public double StdDevKG440 { get => _stdDevKG440; private set { _stdDevKG440 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG450;
        public double StdDevKG450 { get => _stdDevKG450; private set { _stdDevKG450 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG460;
        public double StdDevKG460 { get => _stdDevKG460; private set { _stdDevKG460 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG490;
        public double StdDevKG490 { get => _stdDevKG490; private set { _stdDevKG490 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG474;
        public double StdDevKG474 { get => _stdDevKG474; private set { _stdDevKG474 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG475;
        public double StdDevKG475 { get => _stdDevKG475; private set { _stdDevKG475 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG480;
        public double StdDevKG480 { get => _stdDevKG480; private set { _stdDevKG480 = value; OnPropertyChanged(); } }
        
        private double _stdDevKG550;
        public double StdDevKG550 { get => _stdDevKG550; private set { _stdDevKG550 = value; OnPropertyChanged(); } }

        // Corrected Average values (after applying correction factors)
        private double _correctedAverageKG220;
        public double CorrectedAverageKG220 { get => _correctedAverageKG220; private set { _correctedAverageKG220 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG230;
        public double CorrectedAverageKG230 { get => _correctedAverageKG230; private set { _correctedAverageKG230 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG410;
        public double CorrectedAverageKG410 { get => _correctedAverageKG410; private set { _correctedAverageKG410 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG420;
        public double CorrectedAverageKG420 { get => _correctedAverageKG420; private set { _correctedAverageKG420 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG434;
        public double CorrectedAverageKG434 { get => _correctedAverageKG434; private set { _correctedAverageKG434 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG430;
        public double CorrectedAverageKG430 { get => _correctedAverageKG430; private set { _correctedAverageKG430 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG440;
        public double CorrectedAverageKG440 { get => _correctedAverageKG440; private set { _correctedAverageKG440 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG450;
        public double CorrectedAverageKG450 { get => _correctedAverageKG450; private set { _correctedAverageKG450 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG460;
        public double CorrectedAverageKG460 { get => _correctedAverageKG460; private set { _correctedAverageKG460 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG490;
        public double CorrectedAverageKG490 { get => _correctedAverageKG490; private set { _correctedAverageKG490 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG474;
        public double CorrectedAverageKG474 { get => _correctedAverageKG474; private set { _correctedAverageKG474 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG475;
        public double CorrectedAverageKG475 { get => _correctedAverageKG475; private set { _correctedAverageKG475 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG480;
        public double CorrectedAverageKG480 { get => _correctedAverageKG480; private set { _correctedAverageKG480 = value; OnPropertyChanged(); } }
        
        private double _correctedAverageKG550;
        public double CorrectedAverageKG550 { get => _correctedAverageKG550; private set { _correctedAverageKG550 = value; OnPropertyChanged(); } }

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
                AverageKG490 = list.Average(p => p.CostPerSqmKG490);
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
                MinKG490 = list.Where(p => p.CostPerSqmKG490 > 0).DefaultIfEmpty().Min(p => p?.CostPerSqmKG490 ?? 0);
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
                MaxKG490 = list.Max(p => p.CostPerSqmKG490);
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
                StdDevKG490 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG490 > 0).Select(p => (double)p.CostPerSqmKG490).ToList(), list.Where(p => p.CostPerSqmKG490 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG490 ?? 0));
                StdDevKG550 = CalculateStandardDeviation(list.Where(p => p.CostPerSqmKG550 > 0).Select(p => (double)p.CostPerSqmKG550).ToList(), list.Where(p => p.CostPerSqmKG550 > 0).DefaultIfEmpty().Average(p => p?.CostPerSqmKG550 ?? 0));

                Console.WriteLine($"UpdateAverages: Calculated - Area: {AverageArea:F2}, KG220: Avg={AverageKG220:F2}, Min={MinKG220:F2}, Max={MaxKG220:F2}, StdDev={StdDevKG220:F2}");
                
                // Calculate corrected averages by applying correction factors based on project years
                CalculateCorrectedAverages(list);
            }

            // Property change notifications are now handled automatically by the property setters

            UpdateCostGroupSummary();
            Console.WriteLine("UpdateAverages: Completed calculation and property notifications");
        }

        private void CalculateCorrectedAverages(List<ProjectRecord> projects)
        {
            if (projects.Count == 0)
            {
                CorrectedAverageKG220 = CorrectedAverageKG230 = CorrectedAverageKG410 = CorrectedAverageKG420 = CorrectedAverageKG434 = CorrectedAverageKG430 = CorrectedAverageKG440 = CorrectedAverageKG450 = CorrectedAverageKG460 = CorrectedAverageKG474 = CorrectedAverageKG475 = CorrectedAverageKG480 = CorrectedAverageKG490 = CorrectedAverageKG550 = 0;
                return;
            }

            // Apply correction factors to each project's costs based on its year
            var correctedProjects = projects.Select(p => new
            {
                Project = p,
                Factor = _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG220 = p.CostPerSqmKG220 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG230 = p.CostPerSqmKG230 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG410 = p.CostPerSqmKG410 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG420 = p.CostPerSqmKG420 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG434 = p.CostPerSqmKG434 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG430 = p.CostPerSqmKG430 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG440 = p.CostPerSqmKG440 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG450 = p.CostPerSqmKG450 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG460 = p.CostPerSqmKG460 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG474 = p.CostPerSqmKG474 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG475 = p.CostPerSqmKG475 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG480 = p.CostPerSqmKG480 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG490 = p.CostPerSqmKG490 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG550 = p.CostPerSqmKG550 * _correctionFactorSettings.GetFactorForYear(p.Year)
            }).ToList();

            // Calculate averages of corrected values
            CorrectedAverageKG220 = correctedProjects.Average(cp => cp.CorrectedKG220);
            CorrectedAverageKG230 = correctedProjects.Average(cp => cp.CorrectedKG230);
            CorrectedAverageKG410 = correctedProjects.Average(cp => cp.CorrectedKG410);
            CorrectedAverageKG420 = correctedProjects.Average(cp => cp.CorrectedKG420);
            CorrectedAverageKG434 = correctedProjects.Average(cp => cp.CorrectedKG434);
            CorrectedAverageKG430 = correctedProjects.Average(cp => cp.CorrectedKG430);
            CorrectedAverageKG440 = correctedProjects.Average(cp => cp.CorrectedKG440);
            CorrectedAverageKG450 = correctedProjects.Average(cp => cp.CorrectedKG450);
            CorrectedAverageKG460 = correctedProjects.Average(cp => cp.CorrectedKG460);
            CorrectedAverageKG474 = correctedProjects.Average(cp => cp.CorrectedKG474);
            CorrectedAverageKG475 = correctedProjects.Average(cp => cp.CorrectedKG475);
            CorrectedAverageKG480 = correctedProjects.Average(cp => cp.CorrectedKG480);
            CorrectedAverageKG490 = correctedProjects.Average(cp => cp.CorrectedKG490);
            CorrectedAverageKG550 = correctedProjects.Average(cp => cp.CorrectedKG550);

            Console.WriteLine($"CalculateCorrectedAverages: Corrected KG220: {CorrectedAverageKG220:F2} (Original: {AverageKG220:F2})");
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

            // Apply correction factors to get corrected values for each project
            var correctedProjects = includedItems.Select(p => new
            {
                Project = p,
                Factor = _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG220 = p.CostPerSqmKG220 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG230 = p.CostPerSqmKG230 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG410 = p.CostPerSqmKG410 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG420 = p.CostPerSqmKG420 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG434 = p.CostPerSqmKG434 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG430 = p.CostPerSqmKG430 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG440 = p.CostPerSqmKG440 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG450 = p.CostPerSqmKG450 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG460 = p.CostPerSqmKG460 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG474 = p.CostPerSqmKG474 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG475 = p.CostPerSqmKG475 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG480 = p.CostPerSqmKG480 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG490 = p.CostPerSqmKG490 * _correctionFactorSettings.GetFactorForYear(p.Year),
                CorrectedKG550 = p.CostPerSqmKG550 * _correctionFactorSettings.GetFactorForYear(p.Year)
            }).ToList();

            // Define cost groups with descriptions according to DIN 276 - now using corrected values
            var costGroups = new[]
            {
                new { Code = "KG220", Description = "Public Infrastructure", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG220) },
                new { Code = "KG230", Description = "Private Infrastructure", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG230) },
                new { Code = "KG410", Description = "Wastewater, Water, and Gas", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG410) },
                new { Code = "KG420", Description = "Heating Systems", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG420) },
                new { Code = "KG430", Description = "Ventilation & Air Conditioning", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG430) },
                new { Code = "KG434", Description = "Cooling", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG434) },
                new { Code = "KG440", Description = "Electrical Systems", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG440) },
                new { Code = "KG450", Description = "Communication and Information Technology", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG450) },
                new { Code = "KG460", Description = "Conveying Systems", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG460) },
                new { Code = "KG490", Description = "Other Technical Systems", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG490) },
                new { Code = "KG474", Description = "Fire Extinguishing Systems", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG474) },
                new { Code = "KG475", Description = "Process Heat, Cooling, and Air Systems", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG475) },
                new { Code = "KG480", Description = "Building & System Automation", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG480) },
                new { Code = "KG550", Description = "Outdoor Technical Installations", GetValue = (Func<dynamic, double>)(cp => cp.CorrectedKG550) }
            };

            foreach (var costGroup in costGroups)
            {
                var values = correctedProjects.Select(costGroup.GetValue).Where(v => v > 0).ToList();

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
                    Console.WriteLine($"UpdateCostGroupSummary: {costGroup.Code} - Corrected Avg: {average:F2}, Min: {min:F2}, Max: {max:F2}, StdDev: {standardDeviation:F2}");
                }
                else
                {
                    Console.WriteLine($"UpdateCostGroupSummary: {costGroup.Code} - No valid data (all zeros)");
                }
            }

            Console.WriteLine($"UpdateCostGroupSummary: Created {CostGroupSummary.Count} cost group summaries with corrected values");
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
                ExportServices.PdfExporter.Export(exportItems, summaryItems, AverageArea, AverageKG220, AverageKG230, AverageKG410, AverageKG420, AverageKG434, AverageKG430, AverageKG440, AverageKG450, AverageKG460, AverageKG474, AverageKG475, AverageKG480, AverageKG550);
                Console.WriteLine("ExportPdf: PDF export completed successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ExportPdf: ERROR - {ex.Message}");
            }
        }

        private void OpenCorrectionFactorSettings()
        {
            try
            {
                var settingsViewModel = new CorrectionFactorSettingsViewModel(_correctionFactorSettings);
                var settingsWindow = new CorrectionFactorSettingsWindow(settingsViewModel);
                
                settingsViewModel.CloseRequested += () => settingsWindow.Close();
                
                if (settingsWindow.ShowDialog() == true || settingsViewModel.DialogResult)
                {
                    _correctionFactorSettings = settingsViewModel.GetSettings();
                    CorrectionFactorService.SaveSettings(_correctionFactorSettings);
                    
                    // Recalculate averages with new correction factors
                    UpdateAverages();
                    
                    Console.WriteLine("CorrectionFactorSettings: Settings updated and saved");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"OpenCorrectionFactorSettings: ERROR - {ex.Message}");
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

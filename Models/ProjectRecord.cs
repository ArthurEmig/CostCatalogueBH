using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CostsViewer.Models
{
    public class ProjectRecord : INotifyPropertyChanged
    {
        private bool _include = true;

        public bool Include
        {
            get => _include;
            set { _include = value; OnPropertyChanged(); }
        }

        public string ProjectId { get; set; } = string.Empty;
        public string ProjectTitle { get; set; } = string.Empty;
        public List<string> ProjectTypes { get; set; } = new();
        public int TotalArea { get; set; }

        public int CostPerSqmKG220 { get; set; }
        public int CostPerSqmKG230 { get; set; }
        public int CostPerSqmKG410 { get; set; }
        public int CostPerSqmKG420 { get; set; }
        public int CostPerSqmKG434 { get; set; }
        public int CostPerSqmKG430 { get; set; }
        public int CostPerSqmKG440 { get; set; }
        public int CostPerSqmKG450 { get; set; }
        public int CostPerSqmKG460 { get; set; }
        public int CostPerSqmKG474 { get; set; }
        public int CostPerSqmKG475 { get; set; }
        public int CostPerSqmKG480 { get; set; }
        public int CostPerSqmKG550 { get; set; }

        public string ProjectTypesDisplay => string.Join(", ", ProjectTypes);

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}



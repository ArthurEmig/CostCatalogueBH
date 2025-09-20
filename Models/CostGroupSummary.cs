using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CostsViewer.Models
{
    public class CostGroupSummary : INotifyPropertyChanged
    {
        private string _costGroup = string.Empty;
        private string _description = string.Empty;
        private double _averageCost;
        private double _minCost;
        private double _maxCost;
        private double _standardDeviation;

        public string CostGroup
        {
            get => _costGroup;
            set { _costGroup = value; OnPropertyChanged(); }
        }

        public string Description
        {
            get => _description;
            set { _description = value; OnPropertyChanged(); }
        }

        public double AverageCost
        {
            get => _averageCost;
            set { _averageCost = value; OnPropertyChanged(); }
        }

        public double MinCost
        {
            get => _minCost;
            set { _minCost = value; OnPropertyChanged(); }
        }

        public double MaxCost
        {
            get => _maxCost;
            set { _maxCost = value; OnPropertyChanged(); }
        }

        public double StandardDeviation
        {
            get => _standardDeviation;
            set { _standardDeviation = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}


using System.Windows;
using CostsViewer.ViewModels;

namespace CostsViewer.Views
{
    public partial class CorrectionFactorSettingsWindow : Window
    {
        public CorrectionFactorSettingsWindow(CorrectionFactorSettingsViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}

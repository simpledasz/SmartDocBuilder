using System.Windows.Input;
using CommunityToolkit.Mvvm.Input;
using SmartDocBuilder.Services;

namespace SmartDocBuilderGUI.ViewModels
{
    public partial class MainWindowViewModel : ViewModelBase
    {
        private ThemeManager.Theme _currentTheme = ThemeManager.Theme.Light;

        public ICommand ToggleThemeCommand { get; }

        public MainWindowViewModel()
        {
            ToggleThemeCommand = new RelayCommand(ToggleTheme);
        }

        private void ToggleTheme()
        {
            _currentTheme = _currentTheme == ThemeManager.Theme.Light ? ThemeManager.Theme.Dark : ThemeManager.Theme.Light;
            ThemeManager.SetTheme(_currentTheme);
        }
    }
}
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Input;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Aspose.Words;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Newtonsoft.Json;
using SmartDocBuilder.Services;
using SmartDocBuilderGUI.Models;

namespace SmartDocBuilderGUI.ViewModels
{
    public partial class MainWindowViewModel : ViewModelBase
    {
        private ThemeManager.Theme _currentTheme = ThemeManager.Theme.Light;

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(GenerateReportCommand))]
        private string? _jsonPath;

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(GenerateReportCommand))]
        private string? _templatePath;

        [ObservableProperty]
        private string? _statusText;

        public string? JsonFileName => Path.GetFileName(_jsonPath);
        public string? TemplateFileName => Path.GetFileName(_templatePath);

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

        [RelayCommand]
        private async Task LoadJsonAsync()
        {
            var dialog = new OpenFileDialog { Filters = { new FileDialogFilter { Name = "JSON", Extensions = { "json" } } } };
            if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                var result = await dialog.ShowAsync(desktop.MainWindow);
                if (result?.Length > 0)
                {
                    JsonPath = result[0];
                    StatusText = $"Loaded JSON file: {JsonFileName}";
                }
            }
        }

        [RelayCommand]
        private async Task LoadTemplateAsync()
        {
            var dialog = new OpenFileDialog
                { Filters = { new FileDialogFilter { Name = "Word Templates", Extensions = { "docx" } } } };
            if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                var result = await dialog.ShowAsync(desktop.MainWindow);
                if (result?.Length > 0)
                {
                    TemplatePath = result[0];
                    StatusText = $"Loaded template file: {TemplateFileName}";
                }
            }
        }

        [RelayCommand(CanExecute = nameof(CanGenerateReport))]
        private void GenerateReport()
        {
            if (string.IsNullOrEmpty(JsonPath) || string.IsNullOrEmpty(TemplatePath))
            {
                StatusText = "Please load JSON and template first";
                return;
            }

            try
            {
                StatusText = "Generating report...";
                string json = File.ReadAllText(JsonPath);
                ReportData data = JsonConvert.DeserializeObject<ReportData>(json);

                Document doc = new Document(TemplatePath);
                doc.MailMerge.Execute(
                    new string[] { "ClientName", "InvoiceDate", "AmountDue" },
                    new object[] { data.ClientName, data.InvoiceDate, data.AmountDue }
                );

                string outPath = $"Invoice_{data.ClientName.Replace(" ", "_")}.pdf";
                doc.Save(outPath, SaveFormat.Pdf);
                StatusText = $"PDF saved to {outPath}";
            }
            catch (Exception ex)
            {
                StatusText = "Error: " + ex.Message;
            }
        }

        private bool CanGenerateReport() => !string.IsNullOrEmpty(JsonPath) && !string.IsNullOrEmpty(TemplatePath);

        partial void OnJsonPathChanged(string? value)
        {
            OnPropertyChanged(nameof(JsonFileName));
        }

        partial void OnTemplatePathChanged(string? value)
        {
            OnPropertyChanged(nameof(TemplateFileName));
        }
    }
}

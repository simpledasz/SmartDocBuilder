using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Aspose.Words;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using CsvHelper;
using Newtonsoft.Json;
using SmartDocBuilder.Services;
using SmartDocBuilderGUI.Models;

namespace SmartDocBuilderGUI.ViewModels
{
    public enum OutputFormat
    {
        Pdf,
        Docx,
        Html,
        Txt
    }

    public enum InputFormat
    {
        Json,
        Xml,
        Csv
    }

    public partial class MainWindowViewModel : ViewModelBase
    {
        private ThemeManager.Theme _currentTheme = ThemeManager.Theme.Light;

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(GenerateReportCommand))]
        private string? _dataFilePath;

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(GenerateReportCommand))]
        private string? _templatePath;

        [ObservableProperty] private string? _statusText;

        [ObservableProperty] private OutputFormat _selectedOutputFormat = OutputFormat.Pdf;
        [ObservableProperty] private InputFormat _selectedInputFormat = InputFormat.Json;

        public IEnumerable<OutputFormat> OutputFormats => Enum.GetValues<OutputFormat>();
        public IEnumerable<InputFormat> InputFormats => Enum.GetValues<InputFormat>();

        public string? DataFileName => Path.GetFileName(_dataFilePath);
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
        private async Task LoadDataFileAsync()
        {
            var dialog = new OpenFileDialog
            {
                Filters =
                {
                    new FileDialogFilter { Name = "Data Files", Extensions = { "json", "xml", "csv" } },
                    new FileDialogFilter { Name = "JSON", Extensions = { "json" } },
                    new FileDialogFilter { Name = "XML", Extensions = { "xml" } },
                    new FileDialogFilter { Name = "CSV", Extensions = { "csv" } }
                }
            };

            if (Application.Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                var result = await dialog.ShowAsync(desktop.MainWindow);
                if (result?.Length > 0)
                {
                    DataFilePath = result[0];
                    StatusText = $"Loaded data file: {DataFileName}";
                    // Auto-detect input format based on file extension
                    SelectedInputFormat = Path.GetExtension(DataFilePath).ToLower() switch
                    {
                        ".xml" => InputFormat.Xml,
                        ".csv" => InputFormat.Csv,
                        _ => InputFormat.Json
                    };
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
            if (string.IsNullOrEmpty(DataFilePath) || string.IsNullOrEmpty(TemplatePath))
            {
                StatusText = "Please load data file and template first";
                return;
            }

            try
            {
                StatusText = "Generating report...";
                ReportData? data = null;

                switch (SelectedInputFormat)
                {
                    case InputFormat.Json:
                        string json = File.ReadAllText(DataFilePath);
                        data = JsonConvert.DeserializeObject<ReportData>(json);
                        break;
                    case InputFormat.Xml:
                        XDocument docXml = XDocument.Load(DataFilePath);
                        data = new ReportData
                        {
                            ClientName = docXml.Root?.Element("ClientName")?.Value,
                            InvoiceDate = docXml.Root?.Element("InvoiceDate")?.Value,
                            AmountDue = docXml.Root?.Element("AmountDue")?.Value
                        };
                        break;
                    case InputFormat.Csv:
                        using (var reader = new StreamReader(DataFilePath))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            data = csv.GetRecords<ReportData>().FirstOrDefault();
                        }
                        break;
                }

                if (data == null)
                {
                    StatusText = "Error: Could not parse data file.";
                    return;
                }

                Document doc = new Document(TemplatePath);
                doc.MailMerge.Execute(
                    new string[] { "ClientName", "InvoiceDate", "AmountDue" },
                    new object[] { data.ClientName, data.InvoiceDate, data.AmountDue }
                );

                string outExtension = SelectedOutputFormat.ToString().ToLower();
                string outPath = $"Invoice_{data.ClientName?.Replace(" ", "_")}.{outExtension}";
                
                SaveFormat saveFormat = SelectedOutputFormat switch
                {
                    OutputFormat.Docx => SaveFormat.Docx,
                    OutputFormat.Html => SaveFormat.Html,
                    OutputFormat.Txt => SaveFormat.Text,
                    _ => SaveFormat.Pdf
                };

                doc.Save(outPath, saveFormat);
                StatusText = $"{SelectedOutputFormat} saved to {outPath}";
            }
            catch (Exception ex)
            {
                StatusText = "Error: " + ex.Message;
            }
        }

        private bool CanGenerateReport() => !string.IsNullOrEmpty(DataFilePath) && !string.IsNullOrEmpty(TemplatePath);

        partial void OnDataFilePathChanged(string? value)
        {
            OnPropertyChanged(nameof(DataFileName));
        }

        partial void OnTemplatePathChanged(string? value)
        {
            OnPropertyChanged(nameof(TemplateFileName));
        }
    }
}
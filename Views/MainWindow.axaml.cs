using System;
using System.Data;
using Avalonia.Controls;
using Avalonia.Interactivity;
using SmartDocBuilderGUI.Models;
using System.IO;
using Newtonsoft.Json;
using Aspose.Words;
// desktop.MainWindow = new MainWindow();

namespace SmartDocBuilderGUI
{
    public partial class MainWindow : Window
    {
        private string jsonPath = "";
        private string templatePath = "";

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void OnLoadJson(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog { Filters = { new FileDialogFilter { Name = "JSON", Extensions = {"json"} } } };
            var result = await dialog.ShowAsync(this);
            if (result?.Length > 0)
            {
                jsonPath = result[0];
                StatusText.Text = "Loading JSON file...";
            }
        }

        private async void OnLoadTemplate(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
                { Filters = { new FileDialogFilter { Name = "Word Templates", Extensions = { "docx" } } } };
            var result = await dialog.ShowAsync(this);
            if (result?.Length > 0)
            {
                templatePath = result[0];
                StatusText.Text = "Loading template file...";
            }
        }

        private void OnGenerate(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(jsonPath) || string.IsNullOrEmpty(templatePath))
            {
                StatusText.Text = "Please load JSON and template first";
                return;
            }

            try
            {
                string json = File.ReadAllText(jsonPath);
                ReportData data = JsonConvert.DeserializeObject<ReportData>(json);

                Document doc = new Document(templatePath);
                doc.MailMerge.Execute(
                    new string[] { "ClientName", "InvoiceDate", "AmountDue" },
                    new object[] { data.ClientName, data.InvoiceDate, data.AmountDue }
                );

                string outPath = $"Invoice_{data.ClientName.Replace(" ", "_")}.pdf";
                doc.Save(outPath, SaveFormat.Pdf);
                StatusText.Text = $"PDF saved to {outPath}";
            }
            catch (Exception ex)
            {
                StatusText.Text = "Error: " + ex.Message;
            }
        }
    }
}
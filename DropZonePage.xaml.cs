using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace PdfConverter;

public partial class DropZonePage : Page
{
    private List<string> selectedFiles = new();
    private string operationType;
    private string fileFilter;
    private bool multiSelect;

    public DropZonePage(string operation, string filter, bool multi = true)
    {
        InitializeComponent();
        operationType = operation;
        fileFilter = filter;
        multiSelect = multi;
        TitleText.Text = operation;
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
        var mainWindow = (MainWindow)Window.GetWindow(this);
        mainWindow.ShowMainMenu();
    }

    private void DropZone_DragEnter(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
            e.Effects = System.Windows.DragDropEffects.Copy;
    }

    private void DropZone_DragLeave(object sender, System.Windows.DragEventArgs e)
    {
    }

    private void DropZone_Drop(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
            AddFiles(files);
        }
    }

    private void DropZone_Click(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = fileFilter, Multiselect = multiSelect };
        if (dialog.ShowDialog() == true)
        {
            AddFiles(dialog.FileNames);
        }
    }

    private void AddFiles(string[] files)
    {
        selectedFiles.Clear();
        selectedFiles.AddRange(files);
        FileCountText.Text = $"{selectedFiles.Count} dosya seçildi";
        ProcessButton.Visibility = Visibility.Visible;
    }

    private async void ProcessButton_Click(object sender, RoutedEventArgs e)
    {
        if (selectedFiles.Count == 0) return;

        ProcessButton.IsEnabled = false;
        ProgressBar.Visibility = Visibility.Visible;
        ProgressBar.IsIndeterminate = true;
        ProgressText.Text = "İşleniyor...";

        await Task.Run(() => System.Threading.Thread.Sleep(500));

        try
        {
            var saveDialog = new SaveFileDialog();
            
            switch (operationType)
            {
                case "PDF Birleştir":
                    saveDialog.Filter = "PDF Dosyası|*.pdf";
                    saveDialog.FileName = "birlestirilmis.pdf";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => PdfHelper.MergePdfs(selectedFiles, saveDialog.FileName));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "PDF Sıkıştır":
                    saveDialog.Filter = "PDF Dosyası|*.pdf";
                    saveDialog.FileName = "sikistirilmis.pdf";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => PdfHelper.CompressPdf(selectedFiles[0], saveDialog.FileName));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "Metin Çıkart":
                    saveDialog.Filter = "Metin Dosyası|*.txt";
                    saveDialog.FileName = "metin.txt";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => File.WriteAllText(saveDialog.FileName, PdfHelper.ExtractText(selectedFiles[0])));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "JPG → PDF":
                    saveDialog.Filter = "PDF Dosyası|*.pdf";
                    saveDialog.FileName = "resimler.pdf";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => PdfHelper.ImagesToPdf(selectedFiles, saveDialog.FileName));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "PDF → JPG":
                    saveDialog.Filter = "JPG Dosyası|*.jpg";
                    saveDialog.FileName = "sayfa_1.jpg";
                    if (saveDialog.ShowDialog() == true)
                    {
                        var folder = Path.GetDirectoryName(saveDialog.FileName);
                        await Task.Run(() => PdfHelper.PdfToImages(selectedFiles[0], folder!));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "Word → PDF":
                    saveDialog.Filter = "PDF Dosyası|*.pdf";
                    saveDialog.FileName = "word.pdf";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => PdfHelper.WordToPdf(selectedFiles[0], saveDialog.FileName));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "PDF → Word":
                    saveDialog.Filter = "Word Dosyası|*.docx";
                    saveDialog.FileName = "word.docx";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => PdfHelper.PdfToWord(selectedFiles[0], saveDialog.FileName));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "Excel → PDF":
                    saveDialog.Filter = "PDF Dosyası|*.pdf";
                    saveDialog.FileName = "excel.pdf";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => PdfHelper.ExcelToPdf(selectedFiles[0], saveDialog.FileName));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "PDF → Excel":
                    saveDialog.Filter = "Excel Dosyası|*.xlsx";
                    saveDialog.FileName = "excel.xlsx";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => PdfHelper.PdfToExcel(selectedFiles[0], saveDialog.FileName));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "PowerPoint → PDF":
                    saveDialog.Filter = "PDF Dosyası|*.pdf";
                    saveDialog.FileName = "powerpoint.pdf";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => PdfHelper.PowerPointToPdf(selectedFiles[0], saveDialog.FileName));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;

                case "PDF → PowerPoint":
                    saveDialog.Filter = "PowerPoint Dosyası|*.pptx";
                    saveDialog.FileName = "powerpoint.pptx";
                    if (saveDialog.ShowDialog() == true)
                    {
                        await Task.Run(() => PdfHelper.PdfToPowerPoint(selectedFiles[0], saveDialog.FileName));
                        ProgressText.Text = "Başarıyla tamamlandı!";
                    }
                    break;
            }
        }
        catch (Exception ex)
        {
            ProgressText.Text = $"Hata: {ex.Message}";
        }
        finally
        {
            ProgressBar.IsIndeterminate = false;
            ProgressBar.Visibility = Visibility.Collapsed;
            ProcessButton.IsEnabled = true;
        }
    }
}

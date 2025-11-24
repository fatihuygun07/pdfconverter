using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace PdfConverter;

public partial class OcrPage : Page
{
    private string? selectedFile;

    public OcrPage()
    {
        InitializeComponent();
    }

    private void Border_Drop(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var pdfFile = files.FirstOrDefault(f => f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase));
            if (pdfFile != null)
            {
                selectedFile = pdfFile;
                SelectedFileText.Text = System.IO.Path.GetFileName(selectedFile);
            }
        }
    }

    private void SelectFile_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "PDF Dosyası|*.pdf"
        };

        if (dialog.ShowDialog() == true)
        {
            selectedFile = dialog.FileName;
            SelectedFileText.Text = System.IO.Path.GetFileName(selectedFile);
        }
    }

    private void Ocr_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrEmpty(selectedFile))
        {
            MessageBox.Show("Lütfen bir PDF dosyası seçin!", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var saveDialog = new SaveFileDialog
        {
            Filter = "Metin Dosyası|*.txt",
            FileName = "cikartilan_metin.txt"
        };

        if (saveDialog.ShowDialog() == true)
        {
            try
            {
                var text = PdfHelper.ExtractText(selectedFile);
                File.WriteAllText(saveDialog.FileName, text);
                MessageBox.Show("Metin başarıyla çıkartıldı!", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information);
                selectedFile = null;
                SelectedFileText.Text = "Dosya seçilmedi";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hata: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}

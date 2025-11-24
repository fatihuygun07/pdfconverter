using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace PdfConverter;

public partial class JpgToPdfPage : Page
{
    private List<string> selectedFiles = new();

    public JpgToPdfPage()
    {
        InitializeComponent();
    }

    private void Border_Drop(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (var file in files.Where(f => f.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) || 
                                                   f.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) || 
                                                   f.EndsWith(".png", StringComparison.OrdinalIgnoreCase)))
            {
                if (!selectedFiles.Contains(file))
                {
                    selectedFiles.Add(file);
                    FileListBox.Items.Add(System.IO.Path.GetFileName(file));
                }
            }
        }
    }

    private void SelectFiles_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Resim Dosyaları|*.jpg;*.jpeg;*.png",
            Multiselect = true
        };

        if (dialog.ShowDialog() == true)
        {
            selectedFiles.AddRange(dialog.FileNames);
            FileListBox.Items.Clear();
            foreach (var file in selectedFiles)
            {
                FileListBox.Items.Add(System.IO.Path.GetFileName(file));
            }
        }
    }

    private void Convert_Click(object sender, RoutedEventArgs e)
    {
        if (selectedFiles.Count == 0)
        {
            MessageBox.Show("En az 1 resim dosyası seçmelisiniz!", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var saveDialog = new SaveFileDialog
        {
            Filter = "PDF Dosyası|*.pdf",
            FileName = "donusturulmus.pdf"
        };

        if (saveDialog.ShowDialog() == true)
        {
            try
            {
                PdfHelper.ImagesToPdf(selectedFiles, saveDialog.FileName);
                MessageBox.Show("Resimler başarıyla PDF'e dönüştürüldü!", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information);
                selectedFiles.Clear();
                FileListBox.Items.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hata: {ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}

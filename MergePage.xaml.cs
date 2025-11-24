using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace PdfConverter;

public partial class MergePage : Page
{
    private List<string> selectedFiles = new();

    public MergePage()
    {
        InitializeComponent();
    }

    private void Border_Drop(object sender, System.Windows.DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (var file in files.Where(f => f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)))
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
            Filter = "PDF Dosyaları|*.pdf",
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

    private void Merge_Click(object sender, RoutedEventArgs e)
    {
        if (selectedFiles.Count < 2)
        {
            MessageBox.Show("En az 2 PDF dosyası seçmelisiniz!", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var saveDialog = new SaveFileDialog
        {
            Filter = "PDF Dosyası|*.pdf",
            FileName = "birlestirilmis.pdf"
        };

        if (saveDialog.ShowDialog() == true)
        {
            try
            {
                PdfHelper.MergePdfs(selectedFiles, saveDialog.FileName);
                MessageBox.Show("PDF dosyaları başarıyla birleştirildi!", "Başarılı", MessageBoxButton.OK, MessageBoxImage.Information);
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

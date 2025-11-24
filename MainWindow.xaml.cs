using System.Windows;

namespace PdfConverter;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void ShowDropZonePage(string operation, string filter, bool multi)
    {
        MainGrid.Visibility = System.Windows.Visibility.Collapsed;
        ContentFrame.Visibility = System.Windows.Visibility.Visible;
        ContentFrame.Navigate(new DropZonePage(operation, filter, multi));
    }

    public void ShowMainMenu()
    {
        ContentFrame.Visibility = System.Windows.Visibility.Collapsed;
        MainGrid.Visibility = System.Windows.Visibility.Visible;
    }

    private void BtnMerge_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("PDF Birleştir", "PDF Dosyaları|*.pdf", true);
    private void BtnCompress_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("PDF Sıkıştır", "PDF Dosyası|*.pdf", false);
    private void BtnOcr_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("Metin Çıkart", "PDF Dosyası|*.pdf", false);
    private void BtnJpgToPdf_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("JPG → PDF", "Resim Dosyaları|*.jpg;*.jpeg;*.png", true);
    private void BtnPdfToJpg_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("PDF → JPG", "PDF Dosyası|*.pdf", false);
    private void BtnWordToPdf_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("Word → PDF", "Word Dosyası|*.docx;*.doc;*.txt", false);
    private void BtnPdfToWord_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("PDF → Word", "PDF Dosyası|*.pdf", false);
    private void BtnExcelToPdf_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("Excel → PDF", "Excel Dosyası|*.xlsx;*.xls;*.csv", false);
    private void BtnPdfToExcel_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("PDF → Excel", "PDF Dosyası|*.pdf", false);
    private void BtnPptToPdf_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("PowerPoint → PDF", "PowerPoint Dosyası|*.pptx;*.ppt", false);
    private void BtnPdfToPpt_Click(object sender, RoutedEventArgs e) => ShowDropZonePage("PDF → PowerPoint", "PDF Dosyası|*.pdf", false);
}

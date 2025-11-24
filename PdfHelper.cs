using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using iText.IO.Image;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Layout;
using iText.Layout.Element;
using ExcelInterop = Microsoft.Office.Interop.Excel;
using PowerPointInterop = Microsoft.Office.Interop.PowerPoint;
using WordInterop = Microsoft.Office.Interop.Word;
using DrawingImage = System.Drawing.Image;
using PdfImage = iText.Layout.Element.Image;
using WordProcessing = DocumentFormat.OpenXml.Wordprocessing;

namespace PdfConverter;

public static class PdfHelper
{
    public static void MergePdfs(List<string> inputFiles, string outputFile)
    {
        using var writer = new PdfWriter(outputFile, new WriterProperties().SetFullCompressionMode(true));
        using var pdf = new PdfDocument(writer);

        foreach (var file in inputFiles)
        {
            using var reader = new PdfReader(file);
            using var sourcePdf = new PdfDocument(reader);
            sourcePdf.CopyPagesTo(1, sourcePdf.GetNumberOfPages(), pdf);
        }
    }

    public static void ImagesToPdf(List<string> imageFiles, string outputFile)
    {
        using var writer = new PdfWriter(outputFile);
        using var pdf = new PdfDocument(writer);
        using var document = new Document(pdf);

        foreach (var imageFile in imageFiles)
        {
            var imageData = ImageDataFactory.Create(imageFile);
            var image = new PdfImage(imageData).SetAutoScale(true);
            document.Add(image);
            document.Add(new AreaBreak());
        }
    }

    public static void CompressPdf(string inputFile, string outputFile)
    {
        var writerProperties = new WriterProperties()
            .SetFullCompressionMode(true)
            .SetCompressionLevel(9);

        using var reader = new PdfReader(inputFile);
        using var writer = new PdfWriter(outputFile, writerProperties);
        using var pdf = new PdfDocument(reader, writer);

        for (var i = 1; i <= pdf.GetNumberOfPages(); i++)
        {
            var page = pdf.GetPage(i);
            page.Flush();
        }
    }

    public static string ExtractText(string inputFile)
    {
        using var reader = new PdfReader(inputFile);
        using var pdf = new PdfDocument(reader);
        var builder = new StringBuilder();

        for (var i = 1; i <= pdf.GetNumberOfPages(); i++)
        {
            var strategy = new SimpleTextExtractionStrategy();
            builder.AppendLine(PdfTextExtractor.GetTextFromPage(pdf.GetPage(i), strategy));
        }

        return builder.ToString();
    }

    public static void PdfToImages(string inputFile, string outputFolder, int dpi = 200)
    {
        Directory.CreateDirectory(outputFolder);

        using var docReader = DocLib.Instance.GetDocReader(File.ReadAllBytes(inputFile), new PageDimensions(dpi, dpi));
        var pageCount = docReader.GetPageCount();

        for (var i = 0; i < pageCount; i++)
        {
            using var bitmap = RenderPage(docReader, i);
            var outputPath = Path.Combine(outputFolder, $"page_{i + 1}.jpg");
            bitmap.Save(outputPath, ImageFormat.Jpeg);
        }
    }

    public static void WordToPdf(string inputFile, string outputFile)
    {
        if (inputFile.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
        {
            var text = File.ReadAllText(inputFile);
            using var writer = new PdfWriter(outputFile);
            using var pdf = new PdfDocument(writer);
            using var document = new Document(pdf);
            document.Add(new Paragraph(text));
            return;
        }

        RunSta(() =>
        {
            WordInterop.Application? app = null;
            WordInterop.Document? document = null;
            try
            {
                app = new WordInterop.Application { Visible = false };
                document = app.Documents.Open(inputFile, ReadOnly: true, Visible: false);
                document.ExportAsFixedFormat(outputFile, WordInterop.WdExportFormat.wdExportFormatPDF);
            }
            finally
            {
                if (document is not null)
                {
                    document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                    ReleaseComObject(document);
                }

                if (app is not null)
                {
                    app.Quit();
                    ReleaseComObject(app);
                }
            }
        });
    }

    public static void PdfToWord(string inputFile, string outputFile)
    {
        Exception? interopError = null;

        try
        {
            RunSta(() =>
            {
                WordInterop.Application? app = null;
                WordInterop.Document? document = null;
                try
                {
                    app = new WordInterop.Application { Visible = false };
                    app.DisplayAlerts = WordInterop.WdAlertLevel.wdAlertsNone;
                    document = app.Documents.Open(inputFile, ReadOnly: true, AddToRecentFiles: false, ConfirmConversions: false, Visible: false);
                    document.SaveAs2(outputFile, WordInterop.WdSaveFormat.wdFormatXMLDocument);
                }
                finally
                {
                    if (document is not null)
                    {
                        document.Close(WordInterop.WdSaveOptions.wdDoNotSaveChanges);
                        ReleaseComObject(document);
                    }

                    if (app is not null)
                    {
                        app.Quit();
                        ReleaseComObject(app);
                    }
                }
            });

            return;
        }
        catch (Exception ex)
        {
            interopError = ex;
        }

        var fallbackText = ExtractText(inputFile);
        if (string.IsNullOrWhiteSpace(fallbackText))
        {
            fallbackText = "[Boş içerik bulunamadı]";
        }

        CreateDocxFromText(outputFile, fallbackText);
    }

    public static void ExcelToPdf(string inputFile, string outputFile)
    {
        RunSta(() =>
        {
            ExcelInterop.Application? app = null;
            ExcelInterop.Workbook? workbook = null;
            try
            {
                app = new ExcelInterop.Application { Visible = false, DisplayAlerts = false };
                workbook = app.Workbooks.Open(inputFile, ReadOnly: true);
                workbook.ExportAsFixedFormat(ExcelInterop.XlFixedFormatType.xlTypePDF, outputFile);
            }
            finally
            {
                if (workbook is not null)
                {
                    workbook.Close(false);
                    ReleaseComObject(workbook);
                }

                if (app is not null)
                {
                    app.Quit();
                    ReleaseComObject(app);
                }
            }
        });
    }

    public static void PdfToExcel(string inputFile, string outputFile)
    {
        var text = ExtractText(inputFile);
        var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

        RunSta(() =>
        {
            ExcelInterop.Application? app = null;
            ExcelInterop.Workbook? workbook = null;
            ExcelInterop.Worksheet? sheet = null;

            try
            {
                app = new ExcelInterop.Application { Visible = false, DisplayAlerts = false };
                workbook = app.Workbooks.Add();
                sheet = (ExcelInterop.Worksheet)workbook.ActiveSheet;
                sheet.Name = "PDF Metni";

                for (var i = 0; i < lines.Length; i++)
                {
                    sheet.Cells[i + 1, 1] = lines[i];
                }

                sheet.Columns.AutoFit();
                workbook.SaveAs(outputFile);
            }
            finally
            {
                if (sheet is not null)
                {
                    ReleaseComObject(sheet);
                }

                if (workbook is not null)
                {
                    workbook.Close();
                    ReleaseComObject(workbook);
                }

                if (app is not null)
                {
                    app.Quit();
                    ReleaseComObject(app);
                }
            }
        });
    }

    public static void PowerPointToPdf(string inputFile, string outputFile)
    {
        RunSta(() =>
        {
            PowerPointInterop.Application? app = null;
            PowerPointInterop.Presentation? presentation = null;

            try
            {
                app = new PowerPointInterop.Application();
                ((dynamic)app).Visible = 0;
                presentation = app.Presentations.Open(inputFile);
                presentation.ExportAsFixedFormat(outputFile, PowerPointInterop.PpFixedFormatType.ppFixedFormatTypePDF);
            }
            finally
            {
                if (presentation is not null)
                {
                    presentation.Close();
                    ReleaseComObject(presentation);
                }

                if (app is not null)
                {
                    app.Quit();
                    ReleaseComObject(app);
                }
            }
        });
    }

    public static void PdfToPowerPoint(string inputFile, string outputFile)
    {
        var (tempFolder, images) = RenderPdfToTempImages(inputFile);

        try
        {
            RunSta(() =>
            {
                PowerPointInterop.Application? app = null;
                PowerPointInterop.Presentation? presentation = null;

                try
                {
                    app = new PowerPointInterop.Application();
                    ((dynamic)app).Visible = 0;
                    presentation = app.Presentations.Add();

                    var slideWidth = presentation.PageSetup.SlideWidth;
                    var slideHeight = presentation.PageSetup.SlideHeight;

                    for (var i = 0; i < images.Count; i++)
                    {
                        var slide = presentation.Slides.Add(i + 1, PowerPointInterop.PpSlideLayout.ppLayoutBlank);
                        using var img = DrawingImage.FromFile(images[i]);
                        var width = (float)PixelsToPoints(img.Width, img.HorizontalResolution);
                        var height = (float)PixelsToPoints(img.Height, img.VerticalResolution);
                        var left = (float)((slideWidth - width) / 2);
                        var top = (float)((slideHeight - height) / 2);
                        ((dynamic)slide.Shapes).AddPicture(images[i], false, true, left, top, width, height);
                    }

                    presentation.SaveAs(outputFile, PowerPointInterop.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                }
                finally
                {
                    if (presentation is not null)
                    {
                        presentation.Close();
                        ReleaseComObject(presentation);
                    }

                    if (app is not null)
                    {
                        app.Quit();
                        ReleaseComObject(app);
                    }
                }
            });
        }
        finally
        {
            SafeDeleteDirectory(tempFolder);
        }
    }

    private static Bitmap RenderPage(IDocReader docReader, int pageIndex)
    {
        using var pageReader = docReader.GetPageReader(pageIndex);
        var rawBytes = pageReader.GetImage();
        var width = pageReader.GetPageWidth();
        var height = pageReader.GetPageHeight();

        var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        var data = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
        Marshal.Copy(rawBytes, 0, data.Scan0, rawBytes.Length);
        bitmap.UnlockBits(data);

        return bitmap;
    }

    private static (string TempFolder, List<string> Images) RenderPdfToTempImages(string inputFile, int dpi = 200)
    {
        var tempFolder = Path.Combine(Path.GetTempPath(), "PdfConverter", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempFolder);
        var images = new List<string>();

        using var docReader = DocLib.Instance.GetDocReader(File.ReadAllBytes(inputFile), new PageDimensions(dpi, dpi));
        var pageCount = docReader.GetPageCount();

        for (var i = 0; i < pageCount; i++)
        {
            using var bitmap = RenderPage(docReader, i);
            var path = Path.Combine(tempFolder, $"page_{i + 1}.png");
            bitmap.Save(path, ImageFormat.Png);
            images.Add(path);
        }

        return (tempFolder, images);
    }

    private static void RunSta(Action action)
    {
        Exception? captured = null;
        using var resetEvent = new ManualResetEvent(false);

        var thread = new Thread(() =>
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                captured = ex;
            }
            finally
            {
                resetEvent.Set();
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        resetEvent.WaitOne();

        if (captured is not null)
        {
            throw captured;
        }
    }

    private static void ReleaseComObject(object comObject)
    {
        try
        {
            if (Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch
        {
            // ignored
        }
    }

    private static void SafeDeleteDirectory(string? path)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
        }
        catch
        {
            // ignored
        }
    }

    private static void CreateDocxFromText(string outputFile, string text)
    {
        using var wordDoc = WordprocessingDocument.Create(outputFile, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new WordProcessing.Document(new WordProcessing.Body());
        var body = mainPart.Document.Body!;

        var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        foreach (var line in lines)
        {
            var paragraph = new WordProcessing.Paragraph(
                new WordProcessing.Run(
                    new WordProcessing.Text(line ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve }));
            body.AppendChild(paragraph);
        }
    }

    private static double PixelsToPoints(int pixels, double dpi) => pixels * 72.0 / dpi;
}

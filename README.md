# PDF Converter

Modern WPF desktop app for PDF and Office conversions. Uses iText7 for PDF manipulation, Docnet.Core for rendering pages to images, and Microsoft Office interop for Word/Excel/PowerPoint conversions.

## Features
- Merge multiple PDFs
- Convert images (JPG/PNG) to PDF and PDF to images
- Compress PDF (full compression, level 9)
- Extract text from PDF
- Word/Excel/PowerPoint to PDF, and PDF to Word/Excel/PowerPoint (requires installed Office)

## Requirements
- Windows
- .NET 8 SDK (for build/run) or the published exe
- Microsoft Office installed for Word/Excel/PowerPoint conversions

## Build & Run
```bash
dotnet restore
dotnet run -c Release
```

### Publish single exe (self-contained)
```bash
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true --self-contained true
```
Resulting exe: `bin/Release/net8.0-windows/win-x64/publish/PdfConverter.exe`

## Notes
- Office conversions rely on locally installed Office; the app cannot convert those formats without it.
- PDF-to-Word uses Word's built-in import for best layout fidelity; Word may show an informational prompt on first use.

---

# PDF Dönüştürücü

PDF ve Office dönüşümleri için modern WPF masaüstü uygulaması. PDF işlemleri iText7, sayfa render için Docnet.Core ve Word/Excel/PowerPoint dönüşümleri için Microsoft Office interop kullanır.

## Özellikler
- PDF birleştirme
- Resim (JPG/PNG) → PDF ve PDF → resim
- PDF sıkıştırma (tam sıkıştırma, seviye 9)
- PDF metin çıkarma
- Word/Excel/PowerPoint ↔ PDF (Office kurulu olmalı)

## Gereksinimler
- Windows
- .NET 8 SDK (build/çalıştırma) veya yayınlanmış exe
- Word/Excel/PowerPoint dönüşümleri için yüklü Microsoft Office

## Derleme & Çalıştırma
```bash
dotnet restore
dotnet run -c Release
```

### Tek dosya exe yayınlama (self-contained)
```bash
dotnet publish -c Release -r win-x64 -p:PublishSingleFile=true --self-contained true
```
Oluşan exe: `bin/Release/net8.0-windows/win-x64/publish/PdfConverter.exe`

## Notlar
- Office dönüşümleri, yüklü Office'e bağlıdır; Office yoksa bu formatlar çevrilemez.
- PDF→Word, en iyi düzen için Word'ün kendi içe aktarmasını kullanır; ilk çalıştırmada bilgilendirme uyarısı görebilirsiniz.

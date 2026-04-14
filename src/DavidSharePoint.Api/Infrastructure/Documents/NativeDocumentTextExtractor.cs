using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using UglyToad.PdfPig;

namespace DavidSharePoint.Api.Infrastructure.Documents;

public sealed class NativeDocumentTextExtractor : IDocumentTextExtractor
{
    public Task<DocumentTextExtractionResult> ExtractAsync(string fileName, byte[] content, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var extension = Path.GetExtension(fileName).ToLowerInvariant();

        return Task.FromResult(extension switch
        {
            ".txt" or ".csv" or ".json" or ".xml" => ExtractTextFile(content),
            ".docx" => ExtractDocx(content),
            ".pdf" => ExtractPdf(content),
            ".png" or ".jpg" or ".jpeg" or ".tif" or ".tiff" or ".bmp" or ".gif" or ".webp" =>
                DocumentTextExtractionResult.OcrRequired("The source file is an image and needs OCR."),
            _ => DocumentTextExtractionResult.Unsupported($"The file extension '{extension}' is not supported without OCR.")
        });
    }

    private static DocumentTextExtractionResult ExtractTextFile(byte[] content)
    {
        var text = Encoding.UTF8.GetString(content).Trim();
        return string.IsNullOrWhiteSpace(text)
            ? DocumentTextExtractionResult.Unsupported("The source text file is empty.")
            : DocumentTextExtractionResult.Extracted(text);
    }

    private static DocumentTextExtractionResult ExtractDocx(byte[] content)
    {
        using var stream = new MemoryStream(content, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);

        if (document.MainDocumentPart?.Document is null)
        {
            return DocumentTextExtractionResult.Unsupported("The DOCX file does not contain a main document part.");
        }

        var text = string.Join(' ', document.MainDocumentPart
            .Document
            .Descendants<Text>()
            .Select(textNode => textNode.Text)
            .Where(static value => !string.IsNullOrWhiteSpace(value)));

        return string.IsNullOrWhiteSpace(text)
            ? DocumentTextExtractionResult.Unsupported("The DOCX file does not contain any readable text.")
            : DocumentTextExtractionResult.Extracted(text.Trim());
    }

    private static DocumentTextExtractionResult ExtractPdf(byte[] content)
    {
        using var stream = new MemoryStream(content, writable: false);
        using var document = PdfDocument.Open(stream);

        var builder = new StringBuilder();
        foreach (var page in document.GetPages())
        {
            builder.AppendLine(page.Text);
        }

        var text = builder.ToString().Trim();

        return string.IsNullOrWhiteSpace(text)
            ? DocumentTextExtractionResult.OcrRequired("The PDF does not expose any native text and needs OCR.")
            : DocumentTextExtractionResult.Extracted(text);
    }
}
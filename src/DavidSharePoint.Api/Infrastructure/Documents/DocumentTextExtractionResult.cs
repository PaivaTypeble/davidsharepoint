namespace DavidSharePoint.Api.Infrastructure.Documents;

public sealed record DocumentTextExtractionResult(
    string Status,
    string? Text,
    string? Message,
    bool RequiresOcr)
{
    public static DocumentTextExtractionResult Extracted(string? text) =>
        new("native", text, null, false);

    public static DocumentTextExtractionResult OcrRequired(string message) =>
        new("ocr_required", null, message, true);

    public static DocumentTextExtractionResult Unsupported(string message) =>
        new("unsupported", null, message, false);
}
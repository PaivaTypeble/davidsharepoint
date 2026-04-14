namespace DavidSharePoint.Api.Infrastructure.Documents;

public interface IDocumentTextExtractor
{
    Task<DocumentTextExtractionResult> ExtractAsync(string fileName, byte[] content, CancellationToken cancellationToken);
}
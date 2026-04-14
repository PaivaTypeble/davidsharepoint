namespace DavidSharePoint.Api.Features.SharePoint.RouteDocument;

public sealed record RouteSharePointDocumentResponse(
    string Status,
    string SourceFileUrl,
    string SourceFileName,
    string SourceExtension,
    bool DryRun,
    string ExtractionStatus,
    string? ExtractionMessage,
    string? ExtractedTextPreview,
    bool OcrPending,
    string? MatchType,
    string? MatchValue,
    string? MatchMessage,
    RouteSharePointDocumentCompanyMatch? Company,
    RouteSharePointDocumentDestination? Destination);

public sealed record RouteSharePointDocumentCompanyMatch(
    string Acronym,
    string ClientName,
    string Nipc,
    bool HasFolder,
    string? Email,
    string? FolderName);

public sealed record RouteSharePointDocumentDestination(
    string RootFolderUrl,
    string FolderName,
    string FolderUrl,
    string TargetFileName,
    string? CreatedFileUrl);
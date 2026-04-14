namespace DavidSharePoint.Api.Features.SharePoint.RouteDocument;

public sealed record RouteSharePointDocumentRequest
{
    public string SourceFileUrl { get; init; } = string.Empty;

    public string? MappingWorkbookUrl { get; init; }

    public string? DestinationRootFolderUrl { get; init; }

    public string? TargetFileName { get; init; }

    public bool DryRun { get; init; } = true;
}
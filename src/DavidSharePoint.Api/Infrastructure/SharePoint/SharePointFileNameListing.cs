namespace DavidSharePoint.Api.Infrastructure.SharePoint;

public sealed record SharePointFileNameListing(
    string SourceUrl,
    string SiteId,
    string? SiteDisplayName,
    string DriveId,
    string DriveName,
    string? TargetPath,
    IReadOnlyList<string> FileNames);
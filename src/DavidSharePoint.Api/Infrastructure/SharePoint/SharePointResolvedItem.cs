namespace DavidSharePoint.Api.Infrastructure.SharePoint;

public sealed record SharePointResolvedItem(
    string SourceUrl,
    string SiteId,
    string? SiteDisplayName,
    string DriveId,
    string DriveName,
    string? TargetPath,
    SharePointDriveItem Item);
namespace DavidSharePoint.Api.Infrastructure.SharePoint;

public sealed record SharePointDriveItem(
    string Id,
    string Name,
    string? WebUrl,
    bool IsFile,
    bool IsFolder,
    long? Size,
    string? DownloadUrl,
    string? ParentPath);
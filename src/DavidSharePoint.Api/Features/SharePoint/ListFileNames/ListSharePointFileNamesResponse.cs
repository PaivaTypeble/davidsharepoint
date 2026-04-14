namespace DavidSharePoint.Api.Features.SharePoint.ListFileNames;

public sealed record ListSharePointFileNamesResponse(
    string SourceUrl,
    string SiteId,
    string? SiteDisplayName,
    string DriveId,
    string DriveName,
    string? TargetPath,
    int FileCount,
    IReadOnlyList<string> FileNames);
namespace DavidSharePoint.Api.Features.SharePoint.ListFileNames;

public sealed record ListSharePointFileNamesRequest
{
    public string SharePointUrl { get; init; } = string.Empty;
}
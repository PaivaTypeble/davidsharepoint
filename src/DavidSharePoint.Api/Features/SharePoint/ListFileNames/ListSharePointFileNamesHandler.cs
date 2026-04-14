using DavidSharePoint.Api.Infrastructure.SharePoint;

namespace DavidSharePoint.Api.Features.SharePoint.ListFileNames;

public sealed class ListSharePointFileNamesHandler
{
    private readonly ISharePointFileNameService _sharePointFileNameService;

    public ListSharePointFileNamesHandler(ISharePointFileNameService sharePointFileNameService)
    {
        _sharePointFileNameService = sharePointFileNameService;
    }

    public async Task<ListSharePointFileNamesResponse> HandleAsync(
        ListSharePointFileNamesRequest request,
        CancellationToken cancellationToken)
    {
        var listing = await _sharePointFileNameService.ListFileNamesAsync(request.SharePointUrl, cancellationToken);

        return new ListSharePointFileNamesResponse(
            listing.SourceUrl,
            listing.SiteId,
            listing.SiteDisplayName,
            listing.DriveId,
            listing.DriveName,
            listing.TargetPath,
            listing.FileNames.Count,
            listing.FileNames);
    }
}
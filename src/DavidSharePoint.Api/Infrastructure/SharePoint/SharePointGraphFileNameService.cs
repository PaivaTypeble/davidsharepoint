namespace DavidSharePoint.Api.Infrastructure.SharePoint;

public sealed class SharePointGraphFileNameService : ISharePointFileNameService
{
    private readonly ISharePointGraphService _sharePointGraphService;
    private readonly ILogger<SharePointGraphFileNameService> _logger;

    public SharePointGraphFileNameService(
        ISharePointGraphService sharePointGraphService,
        ILogger<SharePointGraphFileNameService> logger)
    {
        _sharePointGraphService = sharePointGraphService;
        _logger = logger;
    }

    public async Task<SharePointFileNameListing> ListFileNamesAsync(string sharePointUrl, CancellationToken cancellationToken)
    {
        var resolvedItem = await _sharePointGraphService.ResolveItemAsync(sharePointUrl, cancellationToken);

        IReadOnlyList<string> fileNames = resolvedItem.Item.IsFile
            ? [resolvedItem.Item.Name]
            : await TraverseFileNamesAsync(resolvedItem.DriveId, resolvedItem.Item.Id, cancellationToken);

        _logger.LogInformation(
            "Resolved SharePoint URL {SharePointUrl} to site {SiteId}, drive {DriveId}, path {TargetPath}. Returned {FileCount} files.",
            sharePointUrl,
            resolvedItem.SiteId,
            resolvedItem.DriveId,
            resolvedItem.TargetPath ?? "/",
            fileNames.Count);

        return new SharePointFileNameListing(
            sharePointUrl,
            resolvedItem.SiteId,
            resolvedItem.SiteDisplayName,
            resolvedItem.DriveId,
            resolvedItem.DriveName,
            resolvedItem.TargetPath,
            fileNames);
    }

    private async Task<IReadOnlyList<string>> TraverseFileNamesAsync(string driveId, string rootItemId, CancellationToken cancellationToken)
    {
        var pendingFolderIds = new Queue<string>();
        var fileNames = new List<string>();

        pendingFolderIds.Enqueue(rootItemId);

        while (pendingFolderIds.Count > 0)
        {
            var folderId = pendingFolderIds.Dequeue();
            var children = await _sharePointGraphService.ListChildrenAsync(driveId, folderId, cancellationToken);

            foreach (var item in children)
            {
                if (item.IsFile)
                {
                    fileNames.Add(item.Name);
                }
                else if (item.IsFolder)
                {
                    pendingFolderIds.Enqueue(item.Id);
                }
            }
        }

        fileNames.Sort(StringComparer.OrdinalIgnoreCase);

        return fileNames;
    }
}
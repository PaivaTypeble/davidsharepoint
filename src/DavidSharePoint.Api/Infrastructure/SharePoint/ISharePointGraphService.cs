namespace DavidSharePoint.Api.Infrastructure.SharePoint;

public interface ISharePointGraphService
{
    Task<SharePointResolvedItem> ResolveItemAsync(string sharePointUrl, CancellationToken cancellationToken);

    Task<IReadOnlyList<SharePointDriveItem>> ListChildrenAsync(string driveId, string folderItemId, CancellationToken cancellationToken);

    Task<byte[]> DownloadFileContentAsync(string driveId, string itemId, CancellationToken cancellationToken);

    Task<SharePointDriveItem> UploadFileAsync(
        string driveId,
        string parentFolderItemId,
        string fileName,
        byte[] content,
        string? contentType,
        CancellationToken cancellationToken);
}
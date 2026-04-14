namespace DavidSharePoint.Api.Infrastructure.SharePoint;

public interface ISharePointFileNameService
{
    Task<SharePointFileNameListing> ListFileNamesAsync(string sharePointUrl, CancellationToken cancellationToken);
}
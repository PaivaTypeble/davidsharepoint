using System.ComponentModel;
using System.Text;
using DavidSharePoint.Api.Infrastructure.SharePoint;
using ModelContextProtocol.Server;

namespace DavidSharePoint.Api.Features.SharePoint.ListFileNames;

[McpServerToolType]
public sealed class ListSharePointFileNamesMcpTools
{
    [McpServerTool, Description("Lists all file names under a SharePoint URL without downloading the files.")]
    public async Task<string> ListSharePointFileNames(
        [Description("SharePoint site, library, folder, or file URL to inspect.")] string sharePointUrl,
        ISharePointFileNameService sharePointFileNameService,
        CancellationToken cancellationToken)
    {
        var listing = await sharePointFileNameService.ListFileNamesAsync(sharePointUrl, cancellationToken);
        var builder = new StringBuilder();

        builder.AppendLine($"Site: {listing.SiteDisplayName ?? listing.SiteId}");
        builder.AppendLine($"Drive: {listing.DriveName}");
        builder.AppendLine($"Path: {listing.TargetPath ?? "/"}");
        builder.AppendLine($"Files: {listing.FileNames.Count}");
        builder.AppendLine();

        foreach (var fileName in listing.FileNames)
        {
            builder.AppendLine(fileName);
        }

        return builder.ToString().TrimEnd();
    }
}
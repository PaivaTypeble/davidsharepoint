using System.Net;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.Json.Serialization;
using DavidSharePoint.Api.Infrastructure.Graph;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Primitives;

namespace DavidSharePoint.Api.Infrastructure.SharePoint;

public sealed class SharePointGraphFileNameService : ISharePointFileNameService
{
    private static readonly JsonSerializerOptions JsonSerializerOptions = new(JsonSerializerDefaults.Web);

    private readonly HttpClient _httpClient;
    private readonly IGraphAccessTokenProvider _accessTokenProvider;
    private readonly ILogger<SharePointGraphFileNameService> _logger;

    public SharePointGraphFileNameService(
        HttpClient httpClient,
        IGraphAccessTokenProvider accessTokenProvider,
        ILogger<SharePointGraphFileNameService> logger)
    {
        _httpClient = httpClient;
        _accessTokenProvider = accessTokenProvider;
        _logger = logger;
    }

    public async Task<SharePointFileNameListing> ListFileNamesAsync(string sharePointUrl, CancellationToken cancellationToken)
    {
        if (!Uri.TryCreate(sharePointUrl, UriKind.Absolute, out var sharePointUri))
        {
            throw new ArgumentException("sharePointUrl must be an absolute SharePoint URL.", nameof(sharePointUrl));
        }

        if (!sharePointUri.Host.EndsWith(".sharepoint.com", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("sharePointUrl must point to a SharePoint Online host.", nameof(sharePointUrl));
        }

        var serverRelativeSegments = ExtractServerRelativeSegments(sharePointUri);
        var resolvedSite = await ResolveSiteAsync(sharePointUri.Host, serverRelativeSegments, cancellationToken);
        var drives = await GetDrivesAsync(resolvedSite.SiteId, cancellationToken);
        var resolvedDrive = ResolveDrive(drives, resolvedSite.RemainingSegments);
        var targetItem = await GetTargetItemAsync(resolvedDrive.DriveId, resolvedDrive.ItemSegments, cancellationToken);

        IReadOnlyList<string> fileNames = targetItem.IsFile
            ? [targetItem.Name]
            : await TraverseFileNamesAsync(resolvedDrive.DriveId, targetItem.Id, cancellationToken);

        _logger.LogInformation(
            "Resolved SharePoint URL {SharePointUrl} to site {SiteId}, drive {DriveId}, path {TargetPath}. Returned {FileCount} files.",
            sharePointUrl,
            resolvedSite.SiteId,
            resolvedDrive.DriveId,
            resolvedDrive.TargetPath ?? "/",
            fileNames.Count);

        return new SharePointFileNameListing(
            sharePointUrl,
            resolvedSite.SiteId,
            resolvedSite.SiteDisplayName,
            resolvedDrive.DriveId,
            resolvedDrive.DriveName,
            resolvedDrive.TargetPath,
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
            var nextRequest = $"drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(folderId)}/children?$select=id,name,file,folder";

            while (!string.IsNullOrWhiteSpace(nextRequest))
            {
                var page = await GetFromGraphAsync<GraphCollectionResponse<GraphItem>>(nextRequest, cancellationToken);

                foreach (var item in page.Value)
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

                nextRequest = page.NextLink;
            }
        }

        fileNames.Sort(StringComparer.OrdinalIgnoreCase);

        return fileNames;
    }

    private async Task<GraphItem> GetTargetItemAsync(string driveId, string[] itemSegments, CancellationToken cancellationToken)
    {
        if (itemSegments.Length == 0)
        {
            return await GetFromGraphAsync<GraphItem>(
                $"drives/{Uri.EscapeDataString(driveId)}/root?$select=id,name,file,folder",
                cancellationToken);
        }

        var encodedItemPath = JoinEncodedPath(itemSegments);

        return await GetFromGraphAsync<GraphItem>(
            $"drives/{Uri.EscapeDataString(driveId)}/root:/{encodedItemPath}?$select=id,name,file,folder",
            cancellationToken);
    }

    private async Task<ResolvedSite> ResolveSiteAsync(string host, string[] serverRelativeSegments, CancellationToken cancellationToken)
    {
        foreach (var candidate in BuildSiteCandidates(serverRelativeSegments))
        {
            var site = await TryGetSiteAsync(host, candidate.SiteSegments, cancellationToken);
            if (site is null)
            {
                continue;
            }

            return new ResolvedSite(
                site.Id,
                site.DisplayName,
                serverRelativeSegments.Skip(candidate.ConsumedSegments).ToArray());
        }

        throw new InvalidOperationException("Could not resolve the SharePoint site from the supplied URL.");
    }

    private static IReadOnlyList<SiteCandidate> BuildSiteCandidates(string[] segments)
    {
        var candidates = new List<SiteCandidate>();

        if (segments.Length >= 2 && IsSiteCollectionSegment(segments[0]))
        {
            for (var consumed = segments.Length; consumed >= 2; consumed--)
            {
                candidates.Add(new SiteCandidate(segments.Take(consumed).ToArray(), consumed));
            }
        }

        candidates.Add(new SiteCandidate([], 0));

        return candidates;
    }

    private async Task<GraphSite?> TryGetSiteAsync(string host, string[] siteSegments, CancellationToken cancellationToken)
    {
        var requestUri = siteSegments.Length == 0
            ? $"sites/{host}"
            : $"sites/{host}:/{JoinEncodedPath(siteSegments)}";

        return await TryGetFromGraphAsync<GraphSite>(requestUri, cancellationToken);
    }

    private async Task<IReadOnlyList<GraphDrive>> GetDrivesAsync(string siteId, CancellationToken cancellationToken)
    {
        var response = await GetFromGraphAsync<GraphCollectionResponse<GraphDrive>>(
            $"sites/{Uri.EscapeDataString(siteId)}/drives?$select=id,name,webUrl,driveType",
            cancellationToken);

        return response.Value;
    }

    private static ResolvedDrive ResolveDrive(IReadOnlyList<GraphDrive> drives, string[] remainingSegments)
    {
        if (drives.Count == 0)
        {
            throw new InvalidOperationException("The resolved SharePoint site does not expose any document libraries.");
        }

        if (remainingSegments.Length == 0)
        {
            var defaultDrive = drives.FirstOrDefault(d => d.MatchesAlias("Documents") || d.MatchesAlias("Shared Documents"))
                ?? (drives.Count == 1 ? drives[0] : null);

            if (defaultDrive is null)
            {
                throw new InvalidOperationException(
                    "Could not infer the document library from the supplied URL. Use a library or folder URL.");
            }

            return new ResolvedDrive(defaultDrive.Id, defaultDrive.Name, [], null);
        }

        var drive = drives.FirstOrDefault(d => d.MatchesAlias(remainingSegments[0]));
        if (drive is null)
        {
            throw new InvalidOperationException(
                $"Could not match the document library '{remainingSegments[0]}' on the resolved site.");
        }

        var itemSegments = CleanItemSegments(remainingSegments.Skip(1).ToArray());
        var targetPath = itemSegments.Length == 0 ? null : string.Join('/', itemSegments);

        return new ResolvedDrive(drive.Id, drive.Name, itemSegments, targetPath);
    }

    private static string[] ExtractServerRelativeSegments(Uri sharePointUri)
    {
        var path = ExtractServerRelativePath(sharePointUri);

        return path.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(Uri.UnescapeDataString)
            .ToArray();
    }

    private static string ExtractServerRelativePath(Uri sharePointUri)
    {
        var query = QueryHelpers.ParseQuery(sharePointUri.Query);

        if (TryGetQueryPath(query, "id", out var idPath))
        {
            return idPath;
        }

        if (TryGetQueryPath(query, "RootFolder", out var rootFolderPath))
        {
            return rootFolderPath;
        }

        return StripSpecialShareLinkPrefix(sharePointUri.AbsolutePath);
    }

    private static bool TryGetQueryPath(Dictionary<string, StringValues> query, string key, out string path)
    {
        if (query.TryGetValue(key, out var values) && !StringValues.IsNullOrEmpty(values))
        {
            path = StripSpecialShareLinkPrefix(values[0]!);
            return true;
        }

        path = string.Empty;
        return false;
    }

    private static string StripSpecialShareLinkPrefix(string path)
    {
        var segments = path.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(Uri.UnescapeDataString)
            .ToArray();

        if (segments.Length >= 3 &&
            segments[0].StartsWith(":", StringComparison.Ordinal) &&
            segments[1].Equals("r", StringComparison.OrdinalIgnoreCase))
        {
            return "/" + string.Join('/', segments.Skip(2));
        }

        return path;
    }

    private static bool IsSiteCollectionSegment(string segment) =>
        segment.Equals("sites", StringComparison.OrdinalIgnoreCase) ||
        segment.Equals("teams", StringComparison.OrdinalIgnoreCase);

    private static string[] CleanItemSegments(string[] segments)
    {
        if (segments.Length == 0)
        {
            return [];
        }

        var aspxIndex = Array.FindIndex(segments, segment => segment.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase));
        if (aspxIndex >= 0)
        {
            segments = segments[..aspxIndex];
        }

        if (segments.Length > 0 && segments[^1].Equals("Forms", StringComparison.OrdinalIgnoreCase))
        {
            segments = segments[..^1];
        }

        if (segments.Length >= 1 && segments[0].Equals("Forms", StringComparison.OrdinalIgnoreCase))
        {
            return [];
        }

        return segments;
    }

    private async Task<T> GetFromGraphAsync<T>(string requestUri, CancellationToken cancellationToken)
    {
        using var response = await SendAsync(requestUri, cancellationToken);

        return await DeserializeAsync<T>(response, cancellationToken);
    }

    private async Task<T?> TryGetFromGraphAsync<T>(string requestUri, CancellationToken cancellationToken)
        where T : class
    {
        using var response = await SendAsync(requestUri, cancellationToken, allowNotFound: true);
        if (response.StatusCode == HttpStatusCode.NotFound)
        {
            return null;
        }

        return await DeserializeAsync<T>(response, cancellationToken);
    }

    private async Task<HttpResponseMessage> SendAsync(
        string requestUri,
        CancellationToken cancellationToken,
        bool allowNotFound = false)
    {
        var token = await _accessTokenProvider.GetAccessTokenAsync(cancellationToken);

        using var request = new HttpRequestMessage(HttpMethod.Get, requestUri);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);

        if (allowNotFound && response.StatusCode == HttpStatusCode.NotFound)
        {
            return response;
        }

        if (response.IsSuccessStatusCode)
        {
            return response;
        }

        throw await CreateGraphExceptionAsync(response, cancellationToken);
    }

    private static async Task<T> DeserializeAsync<T>(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        var result = await response.Content.ReadFromJsonAsync<T>(JsonSerializerOptions, cancellationToken);

        return result ?? throw new InvalidOperationException("Microsoft Graph returned an empty response.");
    }

    private static async Task<InvalidOperationException> CreateGraphExceptionAsync(
        HttpResponseMessage response,
        CancellationToken cancellationToken)
    {
        var content = await response.Content.ReadAsStringAsync(cancellationToken);
        var detail = TryGetGraphErrorMessage(content) ?? response.ReasonPhrase ?? "Unexpected Graph error.";

        return new InvalidOperationException(
            $"Microsoft Graph request failed with status {(int)response.StatusCode}: {detail}");
    }

    private static string? TryGetGraphErrorMessage(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return null;
        }

        try
        {
            using var document = JsonDocument.Parse(content);
            if (document.RootElement.TryGetProperty("error", out var error) &&
                error.TryGetProperty("message", out var message) &&
                message.ValueKind == JsonValueKind.String)
            {
                return message.GetString();
            }
        }
        catch (JsonException)
        {
        }

        return null;
    }

    private static string JoinEncodedPath(IEnumerable<string> segments) =>
        string.Join('/', segments.Select(Uri.EscapeDataString));

    private static string NormalizePathToken(string value) =>
        value.Trim().Replace("+", " ", StringComparison.Ordinal).ToUpperInvariant();

    private sealed record SiteCandidate(string[] SiteSegments, int ConsumedSegments);

    private sealed record ResolvedSite(string SiteId, string? SiteDisplayName, string[] RemainingSegments);

    private sealed record ResolvedDrive(string DriveId, string DriveName, string[] ItemSegments, string? TargetPath);

    private sealed record GraphSite(string Id, string? DisplayName);

    private sealed record GraphDrive(string Id, string Name, string? WebUrl, string? DriveType)
    {
        public bool MatchesAlias(string candidate)
        {
            var normalizedCandidate = NormalizePathToken(candidate);
            if (NormalizePathToken(Name) == normalizedCandidate)
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(WebUrl) || !Uri.TryCreate(WebUrl, UriKind.Absolute, out var uri))
            {
                return false;
            }

            var lastSegment = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .LastOrDefault();

            return lastSegment is not null && NormalizePathToken(Uri.UnescapeDataString(lastSegment)) == normalizedCandidate;
        }
    }

    private sealed record GraphItem(string Id, string Name, JsonElement File, JsonElement Folder)
    {
        public bool IsFile => File.ValueKind != JsonValueKind.Undefined && File.ValueKind != JsonValueKind.Null;

        public bool IsFolder => Folder.ValueKind != JsonValueKind.Undefined && Folder.ValueKind != JsonValueKind.Null;
    }

    private sealed record GraphCollectionResponse<T>
    {
        [JsonPropertyName("value")]
        public List<T> Value { get; init; } = [];

        [JsonPropertyName("@odata.nextLink")]
        public string? NextLink { get; init; }
    }
}
using System.Text;
using DavidSharePoint.Api.Infrastructure.Configuration;
using DavidSharePoint.Api.Infrastructure.Documents;
using DavidSharePoint.Api.Infrastructure.SharePoint;
using Microsoft.Extensions.Options;

namespace DavidSharePoint.Api.Features.SharePoint.RouteDocument;

public sealed class RouteSharePointDocumentHandler
{
    private readonly ISharePointGraphService _sharePointGraphService;
    private readonly IDocumentTextExtractor _documentTextExtractor;
    private readonly ICompanyWorkbookReader _companyWorkbookReader;
    private readonly ICompanyMatcher _companyMatcher;
    private readonly IOptions<DocumentRoutingOptions> _options;

    public RouteSharePointDocumentHandler(
        ISharePointGraphService sharePointGraphService,
        IDocumentTextExtractor documentTextExtractor,
        ICompanyWorkbookReader companyWorkbookReader,
        ICompanyMatcher companyMatcher,
        IOptions<DocumentRoutingOptions> options)
    {
        _sharePointGraphService = sharePointGraphService;
        _documentTextExtractor = documentTextExtractor;
        _companyWorkbookReader = companyWorkbookReader;
        _companyMatcher = companyMatcher;
        _options = options;
    }

    public async Task<RouteSharePointDocumentResponse> HandleAsync(
        RouteSharePointDocumentRequest request,
        CancellationToken cancellationToken)
    {
        var source = await _sharePointGraphService.ResolveItemAsync(request.SourceFileUrl, cancellationToken);
        if (!source.Item.IsFile)
        {
            throw new ArgumentException("sourceFileUrl must resolve to a file.", nameof(request.SourceFileUrl));
        }

        var destinationRootFolderUrl = ResolveDestinationRootFolderUrl(request);
        var mappingWorkbookUrl = ResolveMappingWorkbookUrl(request, destinationRootFolderUrl);

        var sourceContent = await _sharePointGraphService.DownloadFileContentAsync(source.DriveId, source.Item.Id, cancellationToken);
        var extraction = await _documentTextExtractor.ExtractAsync(source.Item.Name, sourceContent, cancellationToken);

        if (extraction.RequiresOcr)
        {
            return BuildResponse(
                status: "ocr_required",
                source: source,
                dryRun: request.DryRun,
                extraction: extraction,
                match: null,
                destination: null,
                createdFileUrl: null,
                targetFileName: null);
        }

        var workbook = await _sharePointGraphService.ResolveItemAsync(mappingWorkbookUrl, cancellationToken);
        if (!workbook.Item.IsFile)
        {
            throw new InvalidOperationException("The mapping workbook URL must resolve to a file.");
        }

        var workbookContent = await _sharePointGraphService.DownloadFileContentAsync(workbook.DriveId, workbook.Item.Id, cancellationToken);
        var mappingEntries = _companyWorkbookReader.Read(workbookContent);

        var candidateText = BuildCandidateText(extraction.Text, source.Item.Name);
        var match = _companyMatcher.Match(candidateText, mappingEntries);
        if (!match.IsMatch)
        {
            return BuildResponse(
                status: "unmatched",
                source: source,
                dryRun: request.DryRun,
                extraction: extraction,
                match: match,
                destination: null,
                createdFileUrl: null,
                targetFileName: null);
        }

        if (!match.Entry!.HasFolder)
        {
            return BuildResponse(
                status: "company_has_no_folder",
                source: source,
                dryRun: request.DryRun,
                extraction: extraction,
                match: match,
                destination: null,
                createdFileUrl: null,
                targetFileName: null);
        }

        var destinationRoot = await _sharePointGraphService.ResolveItemAsync(destinationRootFolderUrl, cancellationToken);
        if (!destinationRoot.Item.IsFolder)
        {
            throw new InvalidOperationException("The destination root folder URL must resolve to a folder.");
        }

        var folder = await ResolveDestinationFolderAsync(destinationRoot, match.Entry, cancellationToken);
        if (folder is null)
        {
            return BuildResponse(
                status: "folder_not_found",
                source: source,
                dryRun: request.DryRun,
                extraction: extraction,
                match: match,
                destination: null,
                createdFileUrl: null,
                targetFileName: null);
        }

        var targetFileName = BuildTargetFileName(match.Entry, source.Item.Name, request.TargetFileName);
        string? createdFileUrl = null;

        if (!request.DryRun)
        {
            var uploadedFile = await _sharePointGraphService.UploadFileAsync(
                destinationRoot.DriveId,
                folder.Id,
                targetFileName,
                sourceContent,
                GetContentType(source.Item.Name),
                cancellationToken);

            createdFileUrl = uploadedFile.WebUrl;
        }

        return BuildResponse(
            status: request.DryRun ? "preview_ready" : "copied",
            source: source,
            dryRun: request.DryRun,
            extraction: extraction,
            match: match,
            destination: folder,
            createdFileUrl: createdFileUrl,
            targetFileName: targetFileName,
            destinationRootFolderUrl: destinationRootFolderUrl);
    }

    private string ResolveDestinationRootFolderUrl(RouteSharePointDocumentRequest request)
    {
        var destinationRootFolderUrl = FirstNonEmpty(request.DestinationRootFolderUrl, _options.Value.DestinationRootFolderUrl);
        if (string.IsNullOrWhiteSpace(destinationRootFolderUrl))
        {
            throw new InvalidOperationException(
                "Configure DocumentRouting:DestinationRootFolderUrl or send destinationRootFolderUrl in the request.");
        }

        return destinationRootFolderUrl;
    }

    private string ResolveMappingWorkbookUrl(RouteSharePointDocumentRequest request, string destinationRootFolderUrl)
    {
        var configuredWorkbookUrl = FirstNonEmpty(request.MappingWorkbookUrl, _options.Value.MappingWorkbookUrl);
        if (!string.IsNullOrWhiteSpace(configuredWorkbookUrl))
        {
            return configuredWorkbookUrl;
        }

        var workbookFileName = FirstNonEmpty(_options.Value.MappingWorkbookFileName, "Company Acronyms and NIPC.xlsx");
        return CombineUrl(destinationRootFolderUrl, workbookFileName);
    }

    private async Task<SharePointDriveItem?> ResolveDestinationFolderAsync(
        SharePointResolvedItem destinationRoot,
        CompanyMappingEntry company,
        CancellationToken cancellationToken)
    {
        var children = await _sharePointGraphService.ListChildrenAsync(destinationRoot.DriveId, destinationRoot.Item.Id, cancellationToken);
        var expectedPrefix = $"{company.Acronym}_{company.Nipc}_";

        return children
            .Where(child => child.IsFolder)
            .FirstOrDefault(child => child.Name.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase));
    }

    private static string BuildCandidateText(string? extractedText, string sourceFileName)
    {
        var builder = new StringBuilder();

        if (!string.IsNullOrWhiteSpace(extractedText))
        {
            builder.AppendLine(extractedText);
        }

        builder.AppendLine(Path.GetFileNameWithoutExtension(sourceFileName));

        return builder.ToString();
    }

    private static string BuildTargetFileName(CompanyMappingEntry company, string sourceFileName, string? overrideFileName)
    {
        var requestedFileName = string.IsNullOrWhiteSpace(overrideFileName)
            ? $"{BuildDefaultTargetPrefix(company)}_{Path.GetFileName(sourceFileName)}"
            : overrideFileName.Trim();

        var extension = Path.GetExtension(requestedFileName);
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(requestedFileName);
        var sanitizedBaseName = string.Concat(fileNameWithoutExtension.Select(SanitizeCharacter)).Trim();

        if (string.IsNullOrWhiteSpace(sanitizedBaseName))
        {
            sanitizedBaseName = BuildDefaultTargetPrefix(company);
        }

        return sanitizedBaseName + extension;
    }

    private static string BuildDefaultTargetPrefix(CompanyMappingEntry company)
    {
        var preferredPrefix = FirstNonEmpty(company.FolderName, company.ClientName, $"{company.Acronym}_{company.Nipc}");
        var withUnderscores = string.Join('_', preferredPrefix
            .Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries));

        var sanitizedPrefix = string.Concat(withUnderscores.Select(SanitizeCharacter)).Trim('_');

        return string.IsNullOrWhiteSpace(sanitizedPrefix)
            ? $"{company.Acronym}_{company.Nipc}"
            : sanitizedPrefix;
    }

    private static char SanitizeCharacter(char character) =>
        character switch
        {
            '~' or '"' or '#' or '%' or '&' or '*' or ':' or '<' or '>' or '?' or '/' or '\\' or '{' or '|' or '}' => '_',
            _ => character
        };

    private static string GetContentType(string fileName)
    {
        var extension = Path.GetExtension(fileName).ToLowerInvariant();
        return extension switch
        {
            ".pdf" => "application/pdf",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".txt" => "text/plain",
            ".csv" => "text/csv",
            _ => "application/octet-stream"
        };
    }

    private static string CombineUrl(string baseUrl, string fileName)
    {
        if (baseUrl.EndsWith('/'))
        {
            return baseUrl + Uri.EscapeDataString(fileName);
        }

        return baseUrl + "/" + Uri.EscapeDataString(fileName);
    }

    private static string FirstNonEmpty(params string?[] values) =>
        values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))?.Trim() ?? string.Empty;

    private static RouteSharePointDocumentResponse BuildResponse(
        string status,
        SharePointResolvedItem source,
        bool dryRun,
        DocumentTextExtractionResult extraction,
        CompanyMatchResult? match,
        SharePointDriveItem? destination,
        string? createdFileUrl,
        string? targetFileName,
        string? destinationRootFolderUrl = null)
    {
        var extractedTextPreview = string.IsNullOrWhiteSpace(extraction.Text)
            ? null
            : extraction.Text.Length <= 500
                ? extraction.Text
                : extraction.Text[..500];

        var company = match?.Entry is null
            ? null
            : new RouteSharePointDocumentCompanyMatch(
                match.Entry.Acronym,
                match.Entry.ClientName,
                match.Entry.Nipc,
                match.Entry.HasFolder,
                match.Entry.Email,
                match.Entry.FolderName);

        var responseDestination = destination is null || string.IsNullOrWhiteSpace(targetFileName)
            ? null
            : new RouteSharePointDocumentDestination(
                destinationRootFolderUrl ?? string.Empty,
                destination.Name,
                destination.WebUrl ?? string.Empty,
                targetFileName,
                createdFileUrl);

        return new RouteSharePointDocumentResponse(
            status,
            source.SourceUrl,
            source.Item.Name,
            Path.GetExtension(source.Item.Name),
            dryRun,
            extraction.Status,
            extraction.Message,
            extractedTextPreview,
            extraction.RequiresOcr,
            match?.MatchType,
            match?.MatchValue,
            match?.Message,
            company,
            responseDestination);
    }
}
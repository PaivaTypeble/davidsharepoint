namespace DavidSharePoint.Api.Infrastructure.Configuration;

public sealed class DocumentRoutingOptions
{
    public const string SectionName = "DocumentRouting";

    public string DestinationRootFolderUrl { get; init; } = string.Empty;

    public string MappingWorkbookUrl { get; init; } = string.Empty;

    public string MappingWorkbookFileName { get; init; } = "Company Acronyms and NIPC.xlsx";
}
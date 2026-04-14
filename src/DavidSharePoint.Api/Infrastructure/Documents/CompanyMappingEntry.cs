namespace DavidSharePoint.Api.Infrastructure.Documents;

public sealed record CompanyMappingEntry(
    string Acronym,
    string ClientName,
    string Nipc,
    bool HasFolder,
    string? Email,
    string? FolderName);
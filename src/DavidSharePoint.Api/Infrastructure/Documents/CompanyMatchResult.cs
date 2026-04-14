namespace DavidSharePoint.Api.Infrastructure.Documents;

public sealed record CompanyMatchResult(
    CompanyMappingEntry? Entry,
    string MatchType,
    string? MatchValue,
    string? Message)
{
    public bool IsMatch => Entry is not null;
}
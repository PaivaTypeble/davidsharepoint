namespace DavidSharePoint.Api.Infrastructure.Documents;

public interface ICompanyMatcher
{
    CompanyMatchResult Match(string candidateText, IReadOnlyList<CompanyMappingEntry> entries);
}
namespace DavidSharePoint.Api.Infrastructure.Documents;

public interface ICompanyWorkbookReader
{
    IReadOnlyList<CompanyMappingEntry> Read(byte[] workbookContent);
}
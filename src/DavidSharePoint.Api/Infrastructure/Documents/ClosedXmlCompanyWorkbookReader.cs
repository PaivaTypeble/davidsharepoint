using ClosedXML.Excel;
using System.Globalization;
using System.Text;

namespace DavidSharePoint.Api.Infrastructure.Documents;

public sealed class ClosedXmlCompanyWorkbookReader : ICompanyWorkbookReader
{
    public IReadOnlyList<CompanyMappingEntry> Read(byte[] workbookContent)
    {
        using var stream = new MemoryStream(workbookContent, writable: false);
        using var workbook = new XLWorkbook(stream);

        var worksheet = workbook.Worksheets.FirstOrDefault(ws => NormalizeHeader(ws.Name) == NormalizeHeader("Folha1"))
            ?? workbook.Worksheet(1);

        var usedRange = worksheet.RangeUsed() ?? throw new InvalidOperationException("The mapping workbook does not contain any data.");
        var rows = usedRange.RowsUsed().ToList();

        if (rows.Count < 2)
        {
            return [];
        }

        var headerMap = rows[0]
            .CellsUsed()
            .ToDictionary(cell => NormalizeHeader(cell.GetString()), cell => cell.Address.ColumnNumber);

        var acronymColumn = GetRequiredColumn(headerMap, "Sigla");
        var clientNameColumn = GetRequiredColumn(headerMap, "Nome do cliente");
        var nipcColumn = GetRequiredColumn(headerMap, "NIPC nosso cliente");
        var hasFolderColumn = GetOptionalColumn(headerMap, "Cliente tem Pasta");
        var emailColumn = GetOptionalColumn(headerMap, "Email");
        var folderNameColumn = GetOptionalColumn(headerMap,
            "Nome da pasta",
            "Nome pasta",
            "Pasta",
            "Pasta destino",
            "Nome da pasta SharePoint",
            "Nome da pasta no SharePoint",
            "Folder name");

        var entries = new List<CompanyMappingEntry>();

        foreach (var row in rows.Skip(1))
        {
            var acronym = row.Cell(acronymColumn).GetString().Trim();
            var clientName = row.Cell(clientNameColumn).GetString().Trim();
            var nipc = row.Cell(nipcColumn).GetString().Trim();

            if (string.IsNullOrWhiteSpace(acronym) || string.IsNullOrWhiteSpace(clientName) || string.IsNullOrWhiteSpace(nipc))
            {
                continue;
            }

            var hasFolder = hasFolderColumn is not null && ParseBoolean(row.Cell(hasFolderColumn.Value).GetString());
            var email = emailColumn is not null ? NullIfWhiteSpace(row.Cell(emailColumn.Value).GetString()) : null;
            var folderName = folderNameColumn is not null ? NullIfWhiteSpace(row.Cell(folderNameColumn.Value).GetString()) : null;

            entries.Add(new CompanyMappingEntry(acronym, clientName, nipc, hasFolder, email, folderName));
        }

        return entries;
    }

    private static int GetRequiredColumn(IReadOnlyDictionary<string, int> headerMap, string header)
    {
        var normalizedHeader = NormalizeHeader(header);
        if (!headerMap.TryGetValue(normalizedHeader, out var columnNumber))
        {
            throw new InvalidOperationException($"The mapping workbook is missing the required column '{header}'.");
        }

        return columnNumber;
    }

    private static int? GetOptionalColumn(IReadOnlyDictionary<string, int> headerMap, params string[] headers)
    {
        foreach (var header in headers)
        {
            var normalizedHeader = NormalizeHeader(header);
            if (headerMap.TryGetValue(normalizedHeader, out var columnNumber))
            {
                return columnNumber;
            }
        }

        return null;
    }

    private static bool ParseBoolean(string value)
    {
        var normalizedValue = NormalizeHeader(value);
        return normalizedValue is "SIM" or "YES" or "TRUE" or "1";
    }

    private static string NormalizeHeader(string value)
    {
        var normalized = value.Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder(normalized.Length);

        foreach (var character in normalized)
        {
            var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(character);
            if (unicodeCategory == UnicodeCategory.NonSpacingMark)
            {
                continue;
            }

            if (char.IsLetterOrDigit(character))
            {
                builder.Append(char.ToUpperInvariant(character));
                continue;
            }

            if (char.IsWhiteSpace(character))
            {
                builder.Append(' ');
            }
        }

        return string.Join(' ', builder.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }

    private static string? NullIfWhiteSpace(string value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}
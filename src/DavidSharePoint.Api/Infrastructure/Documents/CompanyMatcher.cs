using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DavidSharePoint.Api.Infrastructure.Documents;

public sealed partial class CompanyMatcher : ICompanyMatcher
{
    public CompanyMatchResult Match(string candidateText, IReadOnlyList<CompanyMappingEntry> entries)
    {
        if (entries.Count == 0)
        {
            return new CompanyMatchResult(null, "none", null, "The mapping workbook does not contain any company rows.");
        }

        var normalizedText = Normalize(candidateText);
        if (string.IsNullOrWhiteSpace(normalizedText))
        {
            return new CompanyMatchResult(null, "none", null, "The source document did not produce any searchable text.");
        }

        var nipcMatches = entries.Where(entry => NipcRegex(entry.Nipc).IsMatch(normalizedText)).ToList();
        if (nipcMatches.Count == 1)
        {
            return new CompanyMatchResult(nipcMatches[0], "nipc", nipcMatches[0].Nipc, null);
        }

        if (nipcMatches.Count > 1)
        {
            return new CompanyMatchResult(null, "nipc", null, "The document matched multiple NIPC values.");
        }

        var nameMatches = entries
            .Where(entry => normalizedText.Contains(Normalize(entry.ClientName), StringComparison.Ordinal))
            .OrderByDescending(entry => entry.ClientName.Length)
            .ToList();

        if (nameMatches.Count == 1)
        {
            return new CompanyMatchResult(nameMatches[0], "client_name", nameMatches[0].ClientName, null);
        }

        if (nameMatches.Count > 1 && nameMatches[0].ClientName.Length != nameMatches[1].ClientName.Length)
        {
            return new CompanyMatchResult(nameMatches[0], "client_name", nameMatches[0].ClientName, null);
        }

        if (nameMatches.Count > 1)
        {
            return new CompanyMatchResult(null, "client_name", null, "The document matched multiple client names.");
        }

        var acronymMatches = entries.Where(entry => AcronymRegex(entry.Acronym).IsMatch(normalizedText)).ToList();
        if (acronymMatches.Count == 1)
        {
            return new CompanyMatchResult(acronymMatches[0], "acronym", acronymMatches[0].Acronym, null);
        }

        if (acronymMatches.Count > 1)
        {
            return new CompanyMatchResult(null, "acronym", null, "The document matched multiple acronyms.");
        }

        return new CompanyMatchResult(null, "none", null, "The document text did not match any acronym, client name, or NIPC from the workbook.");
    }

    private static string Normalize(string value)
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

            builder.Append(char.IsWhiteSpace(character) ? ' ' : char.ToUpperInvariant(character));
        }

        return string.Join(' ', builder.ToString().Split(' ', StringSplitOptions.RemoveEmptyEntries));
    }

    private static Regex NipcRegex(string nipc) =>
        new($"(?<!\\d){Regex.Escape(nipc)}(?!\\d)", RegexOptions.CultureInvariant);

    private static Regex AcronymRegex(string acronym) =>
        new($"(?<![A-Z0-9]){Regex.Escape(Normalize(acronym))}(?![A-Z0-9])", RegexOptions.CultureInvariant);
}
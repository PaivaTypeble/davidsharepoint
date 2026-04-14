using DavidSharePoint.Api.Infrastructure.Configuration;
using DavidSharePoint.Api.Infrastructure.Documents;
using DavidSharePoint.Api.Infrastructure.Graph;
using DavidSharePoint.Api.Infrastructure.SharePoint;

namespace DavidSharePoint.Api.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddSharePointInfrastructure(this IServiceCollection services, IConfiguration configuration)
    {
        services.AddOptions<MicrosoftGraphOptions>()
            .Bind(configuration.GetSection(MicrosoftGraphOptions.SectionName));
        services.AddOptions<DocumentRoutingOptions>()
            .Bind(configuration.GetSection(DocumentRoutingOptions.SectionName));

        services.AddSingleton<IGraphAccessTokenProvider, GraphAccessTokenProvider>();
        services.AddSingleton<ICompanyWorkbookReader, ClosedXmlCompanyWorkbookReader>();
        services.AddSingleton<ICompanyMatcher, CompanyMatcher>();
        services.AddSingleton<IDocumentTextExtractor, NativeDocumentTextExtractor>();
        services.AddHttpClient<ISharePointGraphService, SharePointGraphService>(client =>
        {
            client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");
        });
        services.AddTransient<ISharePointFileNameService, SharePointGraphFileNameService>();

        return services;
    }
}
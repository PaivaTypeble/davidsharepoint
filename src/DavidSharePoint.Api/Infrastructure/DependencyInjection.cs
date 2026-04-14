using DavidSharePoint.Api.Infrastructure.Configuration;
using DavidSharePoint.Api.Infrastructure.Graph;
using DavidSharePoint.Api.Infrastructure.SharePoint;

namespace DavidSharePoint.Api.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddSharePointInfrastructure(this IServiceCollection services, IConfiguration configuration)
    {
        services.AddOptions<MicrosoftGraphOptions>()
            .Bind(configuration.GetSection(MicrosoftGraphOptions.SectionName));

        services.AddSingleton<IGraphAccessTokenProvider, GraphAccessTokenProvider>();
        services.AddHttpClient<ISharePointFileNameService, SharePointGraphFileNameService>(client =>
        {
            client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");
        });

        return services;
    }
}
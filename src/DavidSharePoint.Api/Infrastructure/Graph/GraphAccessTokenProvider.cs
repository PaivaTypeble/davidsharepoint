using Azure.Core;
using Azure.Identity;
using DavidSharePoint.Api.Infrastructure.Configuration;
using Microsoft.Extensions.Options;

namespace DavidSharePoint.Api.Infrastructure.Graph;

public interface IGraphAccessTokenProvider
{
    ValueTask<string> GetAccessTokenAsync(CancellationToken cancellationToken);
}

public sealed class GraphAccessTokenProvider : IGraphAccessTokenProvider
{
    private static readonly string[] Scopes = ["https://graph.microsoft.com/.default"];

    private readonly IOptions<MicrosoftGraphOptions> _options;

    public GraphAccessTokenProvider(IOptions<MicrosoftGraphOptions> options)
    {
        _options = options;
    }

    public async ValueTask<string> GetAccessTokenAsync(CancellationToken cancellationToken)
    {
        var options = _options.Value;

        if (string.IsNullOrWhiteSpace(options.TenantId) ||
            string.IsNullOrWhiteSpace(options.ClientId) ||
            string.IsNullOrWhiteSpace(options.ClientSecret))
        {
            throw new InvalidOperationException(
                "Configure MicrosoftGraph:TenantId, MicrosoftGraph:ClientId and MicrosoftGraph:ClientSecret before calling the API or MCP tool.");
        }

        var credential = new ClientSecretCredential(options.TenantId, options.ClientId, options.ClientSecret);
        var token = await credential.GetTokenAsync(new TokenRequestContext(Scopes), cancellationToken);

        return token.Token;
    }
}
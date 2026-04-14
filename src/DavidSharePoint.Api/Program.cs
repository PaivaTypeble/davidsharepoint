using DavidSharePoint.Api.Features.SharePoint.ListFileNames;
using DavidSharePoint.Api.Infrastructure;
using Microsoft.AspNetCore.HttpOverrides;
using ModelContextProtocol.Server;
using Scalar.AspNetCore;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddProblemDetails();
builder.Services.AddHealthChecks();
builder.Services.AddOpenApi();
builder.Services.AddSharePointInfrastructure(builder.Configuration);
builder.Services.AddTransient<ListSharePointFileNamesHandler>();
builder.Services.Configure<ForwardedHeadersOptions>(options =>
{
	options.ForwardedHeaders = ForwardedHeaders.XForwardedFor | ForwardedHeaders.XForwardedProto;
	options.KnownIPNetworks.Clear();
	options.KnownProxies.Clear();
});
builder.Services.AddMcpServer()
	.WithHttpTransport(options => options.Stateless = true)
	.WithToolsFromAssembly();

var app = builder.Build();

app.UseForwardedHeaders();
app.UseExceptionHandler();

if (!app.Environment.IsDevelopment())
{
	app.UseHttpsRedirection();
}

app.MapGet("/", () => TypedResults.Ok(new
{
	service = "DavidSharePoint",
	openApi = "/openapi/v1.json",
	scalar = "/scalar",
	mcp = "/mcp"
}))
.ExcludeFromDescription();

app.MapHealthChecks("/health")
	.ExcludeFromDescription();

app.MapOpenApi();
app.MapScalarApiReference("/scalar");
app.MapListSharePointFileNamesEndpoint();
app.MapMcp("/mcp");

app.Run();

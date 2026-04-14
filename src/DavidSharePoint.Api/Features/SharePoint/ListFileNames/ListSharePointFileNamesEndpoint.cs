using Microsoft.AspNetCore.Http.HttpResults;

namespace DavidSharePoint.Api.Features.SharePoint.ListFileNames;

public static class ListSharePointFileNamesEndpoint
{
    public static IEndpointRouteBuilder MapListSharePointFileNamesEndpoint(this IEndpointRouteBuilder endpoints)
    {
        var group = endpoints.MapGroup("/api/sharepoint")
            .WithTags("SharePoint");

        group.MapPost("/file-names", HandleAsync)
            .WithName("ListSharePointFileNames")
            .WithSummary("List SharePoint file names without downloading files.")
            .WithDescription("Resolves a SharePoint URL, walks the document library or folder recursively, and returns only the file names.");

        return endpoints;
    }

    private static async Task<Results<Ok<ListSharePointFileNamesResponse>, ValidationProblem, ProblemHttpResult>> HandleAsync(
        ListSharePointFileNamesRequest request,
        ListSharePointFileNamesHandler handler,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.SharePointUrl))
        {
            return TypedResults.ValidationProblem(new Dictionary<string, string[]>
            {
                ["sharePointUrl"] = ["sharePointUrl is required."]
            });
        }

        try
        {
            var response = await handler.HandleAsync(request, cancellationToken);
            return TypedResults.Ok(response);
        }
        catch (ArgumentException ex)
        {
            return TypedResults.ValidationProblem(new Dictionary<string, string[]>
            {
                ["sharePointUrl"] = [ex.Message]
            });
        }
        catch (InvalidOperationException ex)
        {
            return TypedResults.Problem(
                title: "Unable to list SharePoint file names",
                detail: ex.Message,
                statusCode: StatusCodes.Status400BadRequest);
        }
    }
}
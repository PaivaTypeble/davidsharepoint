using Microsoft.AspNetCore.Http.HttpResults;

namespace DavidSharePoint.Api.Features.SharePoint.RouteDocument;

public static class RouteSharePointDocumentEndpoint
{
    public static IEndpointRouteBuilder MapRouteSharePointDocumentEndpoint(this IEndpointRouteBuilder endpoints)
    {
        var group = endpoints.MapGroup("/api/sharepoint")
            .WithTags("SharePoint");

        group.MapPost("/route-document", HandleAsync)
            .WithName("RouteSharePointDocument")
            .WithSummary("Preview or copy a SharePoint document into the matched client folder.")
            .WithDescription("Reads a SharePoint file, extracts native text without OCR, matches the company using the workbook, resolves the target folder under the configured Matriz root, and optionally uploads the renamed copy.");

        return endpoints;
    }

    private static async Task<Results<Ok<RouteSharePointDocumentResponse>, ValidationProblem, ProblemHttpResult>> HandleAsync(
        RouteSharePointDocumentRequest request,
        RouteSharePointDocumentHandler handler,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.SourceFileUrl))
        {
            return TypedResults.ValidationProblem(new Dictionary<string, string[]>
            {
                ["sourceFileUrl"] = ["sourceFileUrl is required."]
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
                ["sourceFileUrl"] = [ex.Message]
            });
        }
        catch (InvalidOperationException ex)
        {
            return TypedResults.Problem(
                title: "Unable to route the SharePoint document",
                detail: ex.Message,
                statusCode: StatusCodes.Status400BadRequest);
        }
    }
}
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src

COPY src/DavidSharePoint.Api/DavidSharePoint.Api.csproj src/DavidSharePoint.Api/
RUN dotnet restore src/DavidSharePoint.Api/DavidSharePoint.Api.csproj

COPY . .
RUN dotnet publish src/DavidSharePoint.Api/DavidSharePoint.Api.csproj -c Release -o /app/publish /p:UseAppHost=false

FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS final
WORKDIR /app

ENV ASPNETCORE_ENVIRONMENT=Production
ENV ASPNETCORE_URLS=http://+:8080

EXPOSE 8080

COPY --from=build /app/publish .

ENTRYPOINT ["dotnet", "DavidSharePoint.Api.dll"]
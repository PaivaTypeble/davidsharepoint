param(
    [string]$Url = "http://localhost:5058"
)

$projectPath = ".\src\DavidSharePoint.Api\DavidSharePoint.Api.csproj"

$existingProcesses = Get-CimInstance Win32_Process |
    Where-Object {
        $_.Name -eq "dotnet.exe" -and
        $_.CommandLine -like "*DavidSharePoint.Api.csproj*"
    }

foreach ($process in $existingProcesses) {
    Stop-Process -Id $process.ProcessId -Force
}

$env:ASPNETCORE_URLS = $Url
$env:ASPNETCORE_ENVIRONMENT = "Development"

dotnet run --no-launch-profile --project $projectPath -- --urls $Url
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SharePointUrl,

    [string]$DestinationPath,

    [string]$TenantId,

    [string]$ClientId,

    [switch]$UseDeviceCode,

    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Join-UrlSegments {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Segments
    )

    $encodedSegments = foreach ($segment in $Segments) {
        if ([string]::IsNullOrWhiteSpace($segment)) {
            continue
        }

        [uri]::EscapeDataString($segment)
    }

    return ($encodedSegments -join '/')
}

function Join-RelativePath {
    param(
        [string]$Left,
        [string]$Right
    )

    if ([string]::IsNullOrWhiteSpace($Left)) {
        return $Right
    }

    if ([string]::IsNullOrWhiteSpace($Right)) {
        return $Left
    }

    return ('{0}/{1}' -f $Left.Trim('/'), $Right.Trim('/'))
}

function Get-QueryParameters {
    param(
        [Parameter(Mandatory = $true)]
        [uri]$Uri
    )

    $result = @{}
    $query = $Uri.Query.TrimStart('?')

    if ([string]::IsNullOrWhiteSpace($query)) {
        return $result
    }

    foreach ($pair in $query -split '&') {
        if ([string]::IsNullOrWhiteSpace($pair)) {
            continue
        }

        $parts = $pair -split '=', 2
        $name = [uri]::UnescapeDataString($parts[0].Replace('+', ' '))
        $value = ''

        if ($parts.Count -gt 1) {
            $value = [uri]::UnescapeDataString($parts[1].Replace('+', ' '))
        }

        $result[$name] = $value
    }

    return $result
}

function Get-ServerRelativeTargetPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputUrl
    )

    $uri = [uri]$InputUrl
    $queryParameters = Get-QueryParameters -Uri $uri

    if ($queryParameters.ContainsKey('id') -and -not [string]::IsNullOrWhiteSpace($queryParameters['id'])) {
        return $queryParameters['id']
    }

    return [uri]::UnescapeDataString($uri.AbsolutePath)
}

function Ensure-GraphAuthenticationModule {
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) {
        return
    }

    Write-Host 'Installing Microsoft.Graph.Authentication from PSGallery...'
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force
}

function Connect-ToMicrosoftGraph {
    param(
        [string]$Tenant,
        [string]$AppId,
        [switch]$DeviceCode
    )

    $connectParams = @{
        Scopes       = @('Sites.Read.All')
        ContextScope = 'Process'
        NoWelcome    = $true
    }

    if ($Tenant) {
        $connectParams.TenantId = $Tenant
    }

    if ($AppId) {
        $connectParams.ClientId = $AppId
    }

    if ($DeviceCode) {
        $connectParams.UseDeviceCode = $true
    }

    Connect-MgGraph @connectParams | Out-Null
}

function Invoke-GraphGet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )

    return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
}

function Resolve-SharePointSite {
    param(
        [Parameter(Mandatory = $true)]
        [string]$HostName,

        [Parameter(Mandatory = $true)]
        [string]$ServerRelativePath
    )

    $segments = $ServerRelativePath.Trim('/') -split '/'

    if ($segments.Count -lt 2) {
        throw "Could not infer a SharePoint site path from '$ServerRelativePath'."
    }

    for ($length = $segments.Count; $length -ge 2; $length--) {
        $candidateSegments = $segments[0..($length - 1)]
        $candidatePath = '/' + ($candidateSegments -join '/')
        $encodedCandidate = Join-UrlSegments -Segments $candidateSegments
        $uri = "https://graph.microsoft.com/v1.0/sites/${HostName}:/$encodedCandidate"

        try {
            $site = Invoke-GraphGet -Uri $uri
            return [pscustomobject]@{
                Site = $site
                SiteServerRelativePath = $candidatePath
            }
        }
        catch {
            continue
        }
    }

    throw "No SharePoint site could be resolved from '$ServerRelativePath'."
}

function Get-DriveForFolderPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteId,

        [Parameter(Mandatory = $true)]
        [string]$PathInsideSite
    )

    $pathSegments = $PathInsideSite.Trim('/') -split '/'

    if ($pathSegments.Count -lt 1 -or [string]::IsNullOrWhiteSpace($pathSegments[0])) {
        throw 'The path inside the site does not include a document library segment.'
    }

    $libraryName = $pathSegments[0]
    $drivesResponse = Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/drives?`$select=id,name,webUrl,driveType"
    $drives = @($drivesResponse.value)

    $drive = $drives | Where-Object { $_.name -ieq $libraryName } | Select-Object -First 1

    if (-not $drive) {
        $encodedLibraryName = [uri]::EscapeDataString($libraryName)
        $drive = $drives | Where-Object { $_.webUrl -match "/$encodedLibraryName(?:/|$)" } | Select-Object -First 1
    }

    if (-not $drive -and $drives.Count -eq 1) {
        $drive = $drives[0]
    }

    if (-not $drive) {
        $available = ($drives | ForEach-Object { $_.name }) -join ', '
        throw "Could not match document library '$libraryName'. Available drives: $available"
    }

    $folderSegments = @()
    if ($pathSegments.Count -gt 1) {
        $folderSegments = $pathSegments[1..($pathSegments.Count - 1)]
    }

    return [pscustomobject]@{
        Drive = $drive
        LibraryName = $libraryName
        FolderPath = ($folderSegments -join '/')
    }
}

function Get-ChildrenForRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [string]$RelativePath
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath)) {
        $nextUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root/children?`$top=200"
    }
    else {
        $encodedPath = Join-UrlSegments -Segments ($RelativePath -split '/')
        $nextUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/${encodedPath}:/children?`$top=200"
    }

    do {
        $response = Invoke-GraphGet -Uri $nextUri
        foreach ($item in @($response.value)) {
            $item
        }

        $nextUri = $response.'@odata.nextLink'
    }
    while ($nextUri)
}

function Get-DownloadUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter(Mandatory = $true)]
        [string]$ItemId
    )

    $encodedSelect = [uri]::EscapeDataString('id,name,@microsoft.graph.downloadUrl')
    $item = Invoke-GraphGet -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId?`$select=$encodedSelect"
    $downloadUrl = $item.'@microsoft.graph.downloadUrl'

    if ([string]::IsNullOrWhiteSpace($downloadUrl)) {
        throw "Graph did not return @microsoft.graph.downloadUrl for item '$ItemId'."
    }

    return $downloadUrl
}

function Download-DriveFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [string]$RelativePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationRoot,

        [ref]$FileCount,

        [ref]$FolderCount
    )

    foreach ($item in Get-ChildrenForRelativePath -DriveId $DriveId -RelativePath $RelativePath) {
        $childRelativePath = Join-RelativePath -Left $RelativePath -Right $item.name
        $localPath = Join-Path -Path $DestinationRoot -ChildPath ($childRelativePath -replace '/', [IO.Path]::DirectorySeparatorChar)

        if ($item.folder) {
            if (-not (Test-Path -LiteralPath $localPath)) {
                New-Item -ItemType Directory -Path $localPath | Out-Null
            }

            $FolderCount.Value++
            Write-Host "[DIR ] $childRelativePath"
            Download-DriveFolder -DriveId $DriveId -RelativePath $childRelativePath -DestinationRoot $DestinationRoot -FileCount $FileCount -FolderCount $FolderCount
            continue
        }

        $parentDirectory = Split-Path -Parent $localPath
        if (-not (Test-Path -LiteralPath $parentDirectory)) {
            New-Item -ItemType Directory -Path $parentDirectory -Force | Out-Null
        }

        Write-Host "[FILE] $childRelativePath"

        try {
            Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($item.id)/content" -OutputFilePath $localPath | Out-Null
        }
        catch {
            $downloadUrl = Get-DownloadUrl -DriveId $DriveId -ItemId $item.id
            Invoke-WebRequest -Uri $downloadUrl -OutFile $localPath | Out-Null
        }

        $FileCount.Value++
    }
}

$sharePointUri = [uri]$SharePointUrl
$targetServerRelativePath = Get-ServerRelativeTargetPath -InputUrl $SharePointUrl

if (-not $targetServerRelativePath.StartsWith('/')) {
    $targetServerRelativePath = '/' + $targetServerRelativePath
}

Ensure-GraphAuthenticationModule
Import-Module Microsoft.Graph.Authentication

if (-not $SkipConnect) {
    Connect-ToMicrosoftGraph -Tenant $TenantId -AppId $ClientId -DeviceCode:$UseDeviceCode
}

$resolvedSite = Resolve-SharePointSite -HostName $sharePointUri.Host -ServerRelativePath $targetServerRelativePath
$site = $resolvedSite.Site

$pathInsideSite = $targetServerRelativePath.Substring($resolvedSite.SiteServerRelativePath.Length).Trim('/')

if ([string]::IsNullOrWhiteSpace($pathInsideSite)) {
    throw 'The supplied link points to the site root. Provide a folder URL or an AllItems.aspx link with an id parameter for the target folder.'
}

$driveResolution = Get-DriveForFolderPath -SiteId $site.id -PathInsideSite $pathInsideSite
$folderLeafName = if ([string]::IsNullOrWhiteSpace($driveResolution.FolderPath)) {
    $driveResolution.LibraryName
}
else {
    Split-Path -Path $driveResolution.FolderPath -Leaf
}

if ([string]::IsNullOrWhiteSpace($DestinationPath)) {
    $DestinationPath = Join-Path -Path (Get-Location) -ChildPath $folderLeafName
}

if (-not (Test-Path -LiteralPath $DestinationPath)) {
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
}

Write-Host ''
Write-Host 'SharePoint resolution summary'
Write-Host ('  Site...............: {0}' -f $site.webUrl)
Write-Host ('  Library............: {0}' -f $driveResolution.LibraryName)
Write-Host ('  Folder in library..: {0}' -f $(if ([string]::IsNullOrWhiteSpace($driveResolution.FolderPath)) { '<root>' } else { $driveResolution.FolderPath }))
Write-Host ('  Destination........: {0}' -f (Resolve-Path -LiteralPath $DestinationPath))
Write-Host ''

$fileCount = 0
$folderCount = 0
Download-DriveFolder -DriveId $driveResolution.Drive.id -RelativePath $driveResolution.FolderPath -DestinationRoot $DestinationPath -FileCount ([ref]$fileCount) -FolderCount ([ref]$folderCount)

Write-Host ''
Write-Host ('Downloaded {0} files and created {1} folders.' -f $fileCount, $folderCount)
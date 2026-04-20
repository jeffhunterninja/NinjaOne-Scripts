[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$NinjaOneClientID,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$NinjaOneClientSecret,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$NinjaOneInstance = 'ca.ninjarmm.com',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutPath = 'c:\temp\InstallerTokens.csv',

    [Parameter()]
    [ValidateRange(1, 1000)]
    [int]$PageSize = 1000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Invoke-NinjaApiRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST')]
        [string]$Method,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [AllowNull()]
        [object]$Body,

        [Parameter()]
        [AllowNull()]
        [hashtable]$Headers
    )

    $requestParams = @{
        Uri = "https://$NinjaOneInstance$Path"
        Method = $Method
        ContentType = 'application/json'
    }

    if ($Headers) {
        $requestParams.Headers = $Headers
    }

    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $requestParams.Body = ($Body | ConvertTo-Json -Depth 100)
    }

    $response = Invoke-WebRequest @requestParams
    if ([string]::IsNullOrWhiteSpace($response.Content)) {
        return $null
    }

    return $response.Content | ConvertFrom-Json -Depth 100
}

function Get-NinjaPagedResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    $after = 0
    $results = @()
    do {
        $page = @(Invoke-NinjaApiRequest -Method 'GET' -Path "$Path?pageSize=$PageSize&after=$after" -Headers $script:NinjaAuthHeader)
        if (-not $page -or $page.Count -eq 0) {
            break
        }

        $results += $page
        $maxId = ($page | Measure-Object -Property id -Maximum).Maximum
        if ($null -eq $maxId) {
            break
        }

        $after = $maxId
    } while ($page.Count -eq $PageSize)

    return $results
}

function Get-NinjaAdvancedInstaller {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [int]$OrgID,

        [Parameter(Mandatory = $true)]
        [int]$LocID,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Type,

        [Parameter(Mandatory = $true)]
        [int]$NodeRoleId
    )

    $installerBody = @{
        organizationId = $OrgID
        locationId = $LocID
        installerType = $Type
        content = @{
            nodeRoleId = $NodeRoleId
        }
    }

    return Invoke-NinjaApiRequest -Method 'POST' -Path '/api/v2/organization/generate-installer' -Body $installerBody -Headers $script:NinjaAuthHeader
}

function Get-InstallerTokenFromUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Url
    )

    $fileName = [System.IO.Path]::GetFileName(([Uri]$Url).AbsolutePath)
    if ($fileName -match '^NinjaOneAgent_(?<token>[^.]+)\.(msi|pkg)$') {
        return $Matches.token
    }

    throw "Unable to parse token from installer URL: $Url"
}

$generatedCount = 0
$failedCount = 0

try {
    if (-not (Test-Path -Path (Split-Path -Path $OutPath -Parent))) {
        throw "Output directory does not exist: $(Split-Path -Path $OutPath -Parent)"
    }

    Write-Host "Authenticating to NinjaOne instance $NinjaOneInstance..."
    $authBody = @{
        grant_type = 'client_credentials'
        client_id = $NinjaOneClientID
        client_secret = $NinjaOneClientSecret
        scope = 'monitoring management'
    }

    $authResponse = Invoke-WebRequest -Uri "https://$NinjaOneInstance/ws/oauth/token" -Method POST -Body $authBody -ContentType 'application/x-www-form-urlencoded'
    $accessToken = ($authResponse.Content | ConvertFrom-Json).access_token
    if ([string]::IsNullOrWhiteSpace($accessToken)) {
        throw 'Authentication did not return an access token.'
    }

    $script:NinjaAuthHeader = @{
        Authorization = "Bearer $accessToken"
    }

    Write-Host "Fetching organizations..."
    $ninjaOrgs = @(Get-NinjaPagedResults -Path '/api/v2/organizations')
    Write-Host "Fetching locations..."
    $ninjaLocs = @(Get-NinjaPagedResults -Path '/api/v2/locations')
    Write-Host "Fetching roles..."
    $roles = @(Invoke-NinjaApiRequest -Method 'GET' -Path '/api/v2/roles' -Headers $script:NinjaAuthHeader)
    $windowsRoles = @($roles | Where-Object { $_.nodeClass -match 'WINDOWS' })

    Write-Host "Organizations: $($ninjaOrgs.Count) | Locations: $($ninjaLocs.Count) | Windows roles: $($windowsRoles.Count)"

    $orgById = @{}
    foreach ($org in $ninjaOrgs) {
        $orgById[[string]$org.id] = $org
    }

    $tokens = foreach ($location in $ninjaLocs) {
        $org = $orgById[[string]$location.organizationId]
        $orgName = if ($null -ne $org) { $org.name } else { '' }

        foreach ($role in $windowsRoles) {
            try {
                $installer = Get-NinjaAdvancedInstaller -OrgID $location.organizationId -LocID $location.id -Type 'WINDOWS_MSI' -NodeRoleId $role.id
                if ($null -eq $installer -or [string]::IsNullOrWhiteSpace($installer.url)) {
                    throw "Installer response did not include a URL for location ID $($location.id), role ID $($role.id)."
                }

                $token = Get-InstallerTokenFromUrl -Url $installer.url
                $generatedCount++

                [PSCustomObject]@{
                    OrgName = $orgName
                    LocName = $location.name
                    OrgID = $location.organizationId
                    LocID = $location.id
                    RoleName = $role.name
                    RoleID = $role.id
                    Token = $token
                }
            } catch {
                $failedCount++
                Write-Warning "Failed to generate token for location ID $($location.id), role ID $($role.id): $($_.Exception.Message)"
            }
        }
    }

    $tokens | Export-Csv -Path $OutPath -NoTypeInformation
    Write-Host "Exported $generatedCount tokens to $OutPath"
    Write-Host "Failures: $failedCount"
} catch {
    throw "Installer token export failed: $($_.Exception.Message)"
}

<#
.SYNOPSIS
    Creates NinjaOne ITAM unmanaged device node roles from a CSV file.

.DESCRIPTION
    Reads a CSV of node roles (Name, ParentName, Icon), then ensures those unmanaged device roles
    exist in NinjaOne. Uses the NinjaOne API to list existing roles and create any that are missing.
    Root roles (empty or ROOT ParentName) use nodeRoleParentId 901; child roles use nodeRoleParentName.
    All logic is standalone (no dot-sourcing). Run this before importing devices from CSV.

.PARAMETER CsvPath
    Path to the CSV file. Required columns: Name. Optional: ParentName (empty or ROOT = root under
    UNMANAGED_DEVICE), Icon (e.g. faWrench, faMousePointer; default faTag if blank).

.PARAMETER BaseUrl
    NinjaOne instance base URL (e.g. app.ninjarmm.com or https://app.ninjarmm.com).

.PARAMETER ClientId
    NinjaOne API application Client ID. Can use $env:NinjaOneClientId if not provided.

.PARAMETER ClientSecret
    NinjaOne API application Client Secret. Can use $env:NinjaOneClientSecret if not provided.

.EXAMPLE
    .\Create-NinjaNodeRoles.ps1 -CsvPath ".\NodeRoles-Example.csv" -BaseUrl app.ninjarmm.com
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$CsvPath = '.\NodeRoles-Example.csv',

    [Parameter()]
    [string]$BaseUrl = 'app.ninjarmm.com',

    [Parameter()]
    [string]$ClientId,

    [Parameter()]
    [string]$ClientSecret
)

$ErrorActionPreference = 'Stop'

# Resolve CSV path when default was used (run from script directory)
if ($CsvPath -eq '.\NodeRoles-Example.csv' -and $PSScriptRoot) {
    $CsvPath = Join-Path $PSScriptRoot 'NodeRoles-Example.csv'
}
if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "CSV file not found: $CsvPath"
}

function Get-NinjaOAuthTokenInline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$ClientId,
        [Parameter(Mandatory)] [string]$ClientSecret,
        [Parameter(Mandatory)] [string]$BaseUrl,
        [int]$TimeoutSec = 30
    )
    $uri = "$($BaseUrl.TrimEnd('/'))/ws/oauth/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = 'client_credentials'
        scope         = 'monitoring management'
    }
    $headers = @{
        'Accept'       = 'application/json'
        'Content-Type' = 'application/x-www-form-urlencoded'
    }
    $resp = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $body -TimeoutSec $TimeoutSec -ErrorAction Stop
    return $resp
}

function Invoke-NinjaOneApiInline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')] [string]$Method,
        [Parameter(Mandatory)] [string]$Endpoint,
        [string]$Query,
        $Body,
        [int]$TimeoutSec = 60,
        [Parameter(Mandatory)] [PSCustomObject]$Session
    )
    if (-not $Session.ExpiresAt -or (Get-Date) -ge $Session.ExpiresAt) {
        throw 'No active session or token expired. Re-authenticate.'
    }
    $uri = if ($Query -and $Query.Length -gt 0) {
        "$($Session.BaseUrl)/$($Endpoint.TrimStart('/'))?df=$Query"
    } else {
        "$($Session.BaseUrl)/$($Endpoint.TrimStart('/'))"
    }
    $headers = @{ Authorization = $Session.AuthHeader; 'Accept' = 'application/json' }
    $bodyParam = $null
    if ($Body) {
        $bodyParam = $Body | ConvertTo-Json -Depth 100
        $headers['Content-Type'] = 'application/json'
    }
    return Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers -Body $bodyParam -TimeoutSec $TimeoutSec -ErrorAction Stop
}

# Resolve credentials
$ClientId = if (-not $ClientId) { $env:NinjaOneClientId } else { $ClientId }
$ClientSecret = if (-not $ClientSecret) { $env:NinjaOneClientSecret } else { $ClientSecret }
if ([string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($ClientSecret)) {
    if ($PSVersionTable.PSVersion -lt [version]'7.0') {
        $ClientId     = Read-Host -Prompt 'ClientId'
        $ClientSecret = Read-Host -Prompt 'ClientSecret'
    } else {
        $ClientId     = Read-Host -MaskInput 'ClientId'
        $ClientSecret = Read-Host -MaskInput 'ClientSecret'
    }
}

# Normalize BaseUrl to https, no trailing slash
$base = $BaseUrl.Trim()
if (-not $base.StartsWith('http://', [StringComparison]::OrdinalIgnoreCase) -and -not $base.StartsWith('https://', [StringComparison]::OrdinalIgnoreCase)) {
    $base = "https://$base"
}
$base = $base.TrimEnd('/')

# OAuth and session
$tokenResp = Get-NinjaOAuthTokenInline -ClientId $ClientId -ClientSecret $ClientSecret -BaseUrl $base
$Session = [PSCustomObject]@{
    BaseUrl    = $base
    AuthHeader = "Bearer $($tokenResp.access_token)"
    ExpiresAt  = if ($tokenResp.expires_in) { (Get-Date).AddSeconds([int]$tokenResp.expires_in) } else { (Get-Date).AddHours(1) }
}

# Get existing unmanaged device roles
$rolesRaw = Invoke-NinjaOneApiInline -Method GET -Endpoint 'v2/noderole/list' -Session $Session
$existingRoles = $rolesRaw | Where-Object { $_.nodeClass -eq 'UNMANAGED_DEVICE' }
$existingNames = @($existingRoles | ForEach-Object { $_.name })

# Parse CSV: root roles (ParentName empty or ROOT) and child roles (nodeRoleParentName). Dedupe by Name (first wins); skip blank Name.
$defaultIcon = 'faTag'
$seenNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$rootRoles = [System.Collections.Generic.List[hashtable]]::new()
$childRoles = [System.Collections.Generic.List[hashtable]]::new()

foreach ($row in (Import-Csv -LiteralPath $CsvPath)) {
    $name = if ($row.PSObject.Properties['Name']) { ($row.Name -as [string]).Trim() } else { '' }
    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Warning "Skipping row with blank Name."
        continue
    }
    if ($seenNames.Contains($name)) {
        Write-Warning "Duplicate role name '$name' in CSV; using first occurrence."
        continue
    }
    $seenNames.Add($name) | Out-Null

    $parentName = if ($row.PSObject.Properties['ParentName']) { ($row.ParentName -as [string]).Trim() } else { '' }
    $icon = if ($row.PSObject.Properties['Icon']) { ($row.Icon -as [string]).Trim() } else { '' }
    if ([string]::IsNullOrWhiteSpace($icon)) { $icon = $defaultIcon }

    $isRoot = [string]::IsNullOrWhiteSpace($parentName) -or [string]::Equals($parentName, 'ROOT', [StringComparison]::OrdinalIgnoreCase)
    if ($isRoot) {
        $rootRoles.Add(@{
                nodeClass        = 'UNMANAGED_DEVICE'
                nodeRoleParentId = 901
                name             = $name
                icon             = $icon
            })
    } else {
        $childRoles.Add(@{
                nodeClass         = 'UNMANAGED_DEVICE'
                nodeRoleParentName = $parentName
                name              = $name
                icon              = $icon
            })
    }
}

$createdCount = 0
$createdNames = [System.Collections.Generic.List[string]]::new()

# Pass 1: create root roles
foreach ($role in $rootRoles) {
    if ($role.name -in $existingNames) { continue }
    try {
        Invoke-NinjaOneApiInline -Method POST -Endpoint 'v2/noderole' -Body $role -Session $Session | Out-Null
        $createdCount++
        $createdNames.Add($role.name)
        $existingNames = @($existingNames + $role.name)
    } catch {
        Write-Error "Failed to create role '$($role.name)': $($_.Exception.Message)"
        throw
    }
}

# Pass 2: create child roles
foreach ($role in $childRoles) {
    if ($role.name -in $existingNames) { continue }
    try {
        Invoke-NinjaOneApiInline -Method POST -Endpoint 'v2/noderole' -Body $role -Session $Session | Out-Null
        $createdCount++
        $createdNames.Add($role.name)
    } catch {
        Write-Error "Failed to create role '$($role.name)': $($_.Exception.Message)"
        throw
    }
}

if ($createdCount -eq 0) {
    Write-Host 'All roles already exist. No roles created.'
} else {
    Write-Host "Created $createdCount role(s): $($createdNames -join ', ')."
}

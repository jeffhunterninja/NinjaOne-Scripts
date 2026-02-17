#Requires -Version 5.1
<#
.SYNOPSIS
  Updates NinjaOne device display names from a CSV file via the API.

.DESCRIPTION
  Authenticates to NinjaOne using client credentials, imports a CSV of systemName and
  displayName pairs, matches devices by system name, and PATCHes the display name via
  the API. Intended as an educational example of NinjaOne API interaction.

  This is provided as an educational example. Evaluate and test in a controlled
  setting before use in production. Further improvements may be necessary for larger datasets.

.EXIT CODES
  0 = Success (all matched devices updated)
  1 = Auth failure or device fetch failure
  2 = Validation error (CSV missing, invalid, or duplicate systemName in CSV)

.PARAMETER CsvPath
  Path to the CSV file with columns systemName and displayName.

.PARAMETER NinjaOneInstance
  NinjaOne instance hostname (e.g. ca.ninjarmm.com, app.ninjarmm.com).

.PARAMETER NinjaOneClientId
  OAuth client ID. Can be set via NINJA_CLIENT_ID environment variable.

.PARAMETER NinjaOneClientSecret
  OAuth client secret. Can be set via NINJA_CLIENT_SECRET environment variable.

.PARAMETER WhatIf
  Preview changes without applying them. Shows which devices would be updated.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$CsvPath,

    [string]$NinjaOneInstance = 'ca.ninjarmm.com',

    [string]$NinjaOneClientId = $env:NINJA_CLIENT_ID,

    [string]$NinjaOneClientSecret = $env:NINJA_CLIENT_SECRET,

    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# Validate CSV path
if (-not (Test-Path -Path $CsvPath -PathType Leaf)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 2
}

# Validate credentials
if ([string]::IsNullOrWhiteSpace($NinjaOneClientId) -or [string]::IsNullOrWhiteSpace($NinjaOneClientSecret)) {
    Write-Error "NinjaOneClientId and NinjaOneClientSecret are required. Pass -NinjaOneClientId and -NinjaOneClientSecret, or set NINJA_CLIENT_ID and NINJA_CLIENT_SECRET environment variables."
    exit 2
}

$NinjaOneInstance = $NinjaOneInstance.Trim()
$baseUrl = "https://$NinjaOneInstance"

# Prepare the body for authentication
$body = @{
    grant_type    = "client_credentials"
    client_id     = $NinjaOneClientId.Trim()
    client_secret = $NinjaOneClientSecret.Trim()
    scope         = "monitoring management"
}

$API_AuthHeaders = @{
    'accept'       = 'application/json'
    'Content-Type' = 'application/x-www-form-urlencoded'
}

# Obtain the authentication token
try {
    $auth_token = Invoke-RestMethod -Uri "$baseUrl/ws/oauth/token" -Method POST -Headers $API_AuthHeaders -Body $body
    $access_token = $auth_token | Select-Object -ExpandProperty 'access_token' -EA 0
    if (-not $access_token) {
        Write-Error "Token response did not include access_token."
        exit 1
    }
} catch {
    Write-Error "Failed to obtain authentication token. $_"
    exit 1
}

$headers = @{
    'accept'        = 'application/json'
    'Authorization' = "Bearer $access_token"
}

# Import device data from CSV
try {
    $deviceimports = Import-Csv -Path $CsvPath
} catch {
    Write-Error "Could not import CSV from $CsvPath : $_"
    exit 2
}

# Validate required columns
$sample = $deviceimports | Select-Object -First 1
if (-not $sample.PSObject.Properties['systemName'] -or -not $sample.PSObject.Properties['displayName']) {
    Write-Error "CSV must have columns 'systemName' and 'displayName'."
    exit 2
}

# Process each device import entry
$assets = foreach ($deviceimport in $deviceimports) {
    $name = ($deviceimport.systemName -as [string]).Trim()
    $display = ($deviceimport.displayName -as [string]).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    [PSCustomObject]@{
        Name        = $name
        DisplayName = $display
        ID          = $null
    }
}

# Detect duplicate systemName in CSV (same name appearing more than once)
$duplicateNames = $assets | Group-Object -Property Name | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Name
if ($duplicateNames) {
    Write-Error "Duplicate systemName in CSV: $($duplicateNames -join ', '). Each systemName must appear only once."
    exit 2
}

# Fetch the detailed list of devices from NinjaOne
$devices_url = "$baseUrl/ws/api/v2/devices-detailed"
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
} catch {
    Write-Error "Failed to fetch devices. $_"
    exit 1
}

# Match devices and add their IDs to the assets (devices-detailed returns systemName)
$multipleMatches = @{}
foreach ($device in $devices) {
    $sysName = ($device.systemName -as [string])
    if ([string]::IsNullOrWhiteSpace($sysName)) { continue }
    $matchingAssets = @($assets | Where-Object { $_.Name -eq $sysName })
    if ($matchingAssets.Count -gt 1) {
        # Should not happen given we validated CSV duplicates, but device list could have dupes
        $multipleMatches[$sysName] = $matchingAssets.Count
    }
    foreach ($a in $matchingAssets) {
        $a.ID = $device.id
    }
}
if ($multipleMatches.Count -gt 0) {
    Write-Warning "Multiple devices in NinjaOne share the same systemName. Last match wins: $($multipleMatches.Keys -join ', ')"
}

# Update the display names for each asset
$updated = 0
$wouldUpdate = 0
$failed = 0

foreach ($asset in $assets) {
    if ($null -ne $asset.ID) {
        $displayname_url = "$baseUrl/ws/api/v2/device/$($asset.ID)"
        $request_body = @{ displayName = $asset.DisplayName }
        $json = $request_body | ConvertTo-Json

        if ($PSCmdlet.ShouldProcess($asset.Name, "Update display name to '$($asset.DisplayName)'")) {
            try {
                Invoke-RestMethod -Method Patch -Uri $displayname_url -Headers $headers -Body $json -ContentType "application/json"
                Write-Host "Changed display name for: $($asset.Name) to $($asset.DisplayName)"
                $updated++
            } catch {
                Write-Error "Failed to update display name for $($asset.Name). $_"
                $failed++
            }
        } else {
            Write-Host "[WhatIf] Would change display name for: $($asset.Name) to $($asset.DisplayName)"
            $wouldUpdate++
        }
    } else {
        Write-Warning "Skipping $($asset.Name) - no matching device found in NinjaOne."
    }
}

if ($WhatIf -and $wouldUpdate -gt 0) {
    Write-Host "[WhatIf] $wouldUpdate device(s) would be updated. Run without -WhatIf to apply."
}

exit 0

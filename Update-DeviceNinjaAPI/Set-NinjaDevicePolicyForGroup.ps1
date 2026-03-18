#Requires -Version 5.1
<#
.SYNOPSIS
  Sets the device policy for all devices in a NinjaOne group via the API.

.DESCRIPTION
  Authenticates to NinjaOne using client credentials, retrieves all device IDs
  in the specified group, and PATCHes each device to set the given policyId.
  Intended as an educational example of NinjaOne API interaction.

  Evaluate and test in a controlled setting before use in production. Further
  improvements may be necessary for larger datasets.

.EXIT CODES
  0 = Success (all devices in group updated)
  1 = Auth failure, API failure, or one or more device updates failed
  2 = Validation error (missing/invalid parameters or credentials)

.PARAMETER GroupId
  NinjaOne group ID. All devices in this group will have their policy updated.

.PARAMETER PolicyId
  Policy ID to assign to each device (integer).

.PARAMETER NinjaOneInstance
  NinjaOne instance hostname or base URL (e.g. app.ninjarmm.com). Optional; defaults to env:NINJA_BASE_URL or https://app.ninjarmm.com.

.PARAMETER NinjaOneClientId
  OAuth client ID. Optional; defaults to env:NINJA_CLIENT_ID.

.PARAMETER NinjaOneClientSecret
  OAuth client secret. Optional; defaults to env:NINJA_CLIENT_SECRET.

.PARAMETER WhatIf
  Preview which devices would be updated without applying changes.

.PARAMETER UseWsPaths
  Use /ws/oauth/token instead of /oauth/token (required for some instances). Default: $false.

.PARAMETER ThrottleMs
  Delay in milliseconds between each device PATCH. Use 0 for no delay (default). For large groups, a value such as 200-500 can help avoid rate limits.

.EXAMPLE
  .\Set-NinjaDevicePolicyForGroup.ps1 -GroupId 41 -PolicyId 66

.EXAMPLE
  $env:NINJA_CLIENT_ID = '...'; $env:NINJA_CLIENT_SECRET = '...'
  .\Set-NinjaDevicePolicyForGroup.ps1 -GroupId 41 -PolicyId 66 -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$GroupId,

    [Parameter(Mandatory)]
    [int]$PolicyId,

    [string]$NinjaOneInstance = $env:NINJA_BASE_URL,

    [string]$NinjaOneClientId = $env:NINJA_CLIENT_ID,

    [string]$NinjaOneClientSecret = $env:NINJA_CLIENT_SECRET,

    [switch]$WhatIf,

    [switch]$UseWsPaths,

    [int]$ThrottleMs = 0
)

$ErrorActionPreference = 'Stop'

# Validate credentials
if ([string]::IsNullOrWhiteSpace($NinjaOneClientId) -or [string]::IsNullOrWhiteSpace($NinjaOneClientSecret)) {
    Write-Error "NinjaOneClientId and NinjaOneClientSecret are required. Pass -NinjaOneClientId and -NinjaOneClientSecret, or set NINJA_CLIENT_ID and NINJA_CLIENT_SECRET environment variables."
    exit 2
}

if ([string]::IsNullOrWhiteSpace($GroupId)) {
    Write-Error "GroupId is required and must be non-empty."
    exit 2
}
if ($GroupId.Trim() -notmatch '^\d+$' -or [int]$GroupId -le 0) {
    Write-Error "GroupId must be a positive integer."
    exit 2
}

# Resolve base URL
$NinjaBaseUrl = $NinjaOneInstance
if ([string]::IsNullOrWhiteSpace($NinjaBaseUrl)) { $NinjaBaseUrl = 'https://app.ninjarmm.com' }
$NinjaBaseUrl = $NinjaBaseUrl.Trim()
if ($NinjaBaseUrl -notmatch '^https?://') { $NinjaBaseUrl = "https://$NinjaBaseUrl" }

# Standalone auth: no shared helpers
$baseUrl = $NinjaBaseUrl
$oauthPath = if ($UseWsPaths) { 'ws/oauth/token' } else { 'oauth/token' }
$tokenUri = "$baseUrl/$oauthPath"
$body = @{
    grant_type    = 'client_credentials'
    client_id     = $NinjaOneClientId.Trim()
    client_secret = $NinjaOneClientSecret.Trim()
    scope         = 'monitoring management'
}
$authHeaders = @{
    'accept'       = 'application/json'
    'Content-Type' = 'application/x-www-form-urlencoded'
}
try {
    $authResp = Invoke-RestMethod -Uri $tokenUri -Method POST -Headers $authHeaders -Body $body
    $accessToken = $authResp | Select-Object -ExpandProperty 'access_token' -ErrorAction SilentlyContinue
    if (-not $accessToken) { throw "Token response did not include access_token." }
    Write-Verbose "Obtained NinjaOne access token."
} catch {
    Write-Error "Failed to obtain NinjaOne access token. $($_.Exception.Message)"
    exit 1
}
$headers = @{
    'accept'        = 'application/json'
    'Authorization' = "Bearer $accessToken"
}

# Fetch device IDs in group
$groupDeviceIdsUrl = "$baseUrl/v2/group/$GroupId/device-ids"
try {
    $groupDevices = Invoke-RestMethod -Uri $groupDeviceIdsUrl -Method GET -Headers $headers
} catch {
    Write-Error "Failed to fetch device IDs for group $GroupId. $_"
    exit 1
}

$deviceIds = @($groupDevices)
if (-not $deviceIds -or $deviceIds.Count -eq 0) {
    Write-Host "No devices found in group $GroupId."
    $result = [PSCustomObject]@{
        TotalDevices   = 0
        UpdatedCount   = 0
        FailedCount    = 0
        FailedDeviceIds = @()
    }
    Write-Output $result
    exit 0
}

Write-Verbose "Found $($deviceIds.Count) device(s) in group $GroupId."
$updated = 0
$failed = 0
$failedDeviceIds = [System.Collections.Generic.List[object]]::new()
$total = $deviceIds.Count
$current = 0

foreach ($deviceId in $deviceIds) {
    $current++
    $deviceUpdateUrl = "$baseUrl/api/v2/device/$deviceId"
    $requestBody = @{ policyId = $PolicyId }
    $json = $requestBody | ConvertTo-Json

    if ($PSCmdlet.ShouldProcess($deviceId, "Assign policy $PolicyId")) {
        try {
            Write-Progress -Activity "Updating devices in group $GroupId" -Status "Assigning policy to device $current of $total" -PercentComplete ([int](100 * $current / $total))
            Invoke-RestMethod -Method Patch -Uri $deviceUpdateUrl -Headers $headers -Body $json -ContentType 'application/json'
            Write-Host "Assigning policy to $deviceId"
            $updated++
        } catch {
            Write-Error "Failed to update device $deviceId. $_"
            $failed++
            $failedDeviceIds.Add($deviceId)
        }
    } else {
        Write-Host "[WhatIf] Would assign policy $PolicyId to $deviceId"
    }
    if ($ThrottleMs -gt 0 -and $current -lt $total) {
        Start-Sleep -Milliseconds $ThrottleMs
    }
}

Write-Progress -Activity "Updating devices in group $GroupId" -Completed

$result = [PSCustomObject]@{
    TotalDevices    = $total
    UpdatedCount   = $updated
    FailedCount    = $failed
    FailedDeviceIds = @($failedDeviceIds)
}
Write-Output $result

if ($failed -gt 0) {
    Write-Error "$failed device(s) failed to update. Failed device IDs: $($failedDeviceIds -join ', ')"
    exit 1
}

exit 0

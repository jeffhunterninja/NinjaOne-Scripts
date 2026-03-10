#Requires -Version 5.1
<#
.SYNOPSIS
  Uploads device QR code images to NinjaOne as related-item attachments.

.DESCRIPTION
  Discovers Device_<deviceId>.png files in a directory (e.g. output from
  New-NinjaDeviceQRCode.ps1), parses the device ID from each filename, and
  uploads each image to the matching NinjaOne device as a related-item
  attachment via the NinjaOne API.

  All logic is standalone (no dot-sourcing). Run New-NinjaDeviceQRCode.ps1
  first to generate the images, or point -ImageDirectory to an existing
  folder containing Device_*.png files.

  By default, skips uploading if the device already has a related item
  with the same AttachmentDescription (avoids duplicates). Use -Replace to
  delete any existing matching related item and upload again.

.PARAMETER ImageDirectory
  Folder containing Device_*.png files. Default matches New-NinjaDeviceQRCode.ps1 output.

.PARAMETER NinjaOneInstance
  NinjaOne instance hostname or base URL. Default: $env:NINJA_BASE_URL or ca.ninjarmm.com.

.PARAMETER ClientId
  NinjaOne API application Client ID. Default: $env:NinjaOneClientId.

.PARAMETER ClientSecret
  NinjaOne API application Client Secret. Default: $env:NinjaOneClientSecret.

.PARAMETER AttachmentDescription
  Description stored with the related item in NinjaOne.

.PARAMETER Replace
  If the device already has a related item with the same AttachmentDescription,
  delete it and upload the new image. Default is to skip when one already exists.

.PARAMETER WhatIf
  List discovered files and device IDs only; do not obtain a token or upload.

.EXAMPLE
  .\Send-NinjaDeviceQRCodeToDevice.ps1 -ImageDirectory .\DeviceQRCodes

.EXAMPLE
  .\Send-NinjaDeviceQRCodeToDevice.ps1 -Replace

.EXAMPLE
  .\Send-NinjaDeviceQRCodeToDevice.ps1 -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]
    $ImageDirectory = '.\DeviceQRCodes',

    [Parameter()]
    [string]
    $NinjaOneInstance,

    [string]
    $ClientId = $env:NinjaOneClientId,

    [string]
    $ClientSecret = $env:NinjaOneClientSecret,

    [Parameter()]
    [string]
    $AttachmentDescription = 'Device dashboard QR code',

    [Parameter()]
    [switch]
    $Replace,

    [Parameter()]
    [switch]
    $WhatIf
)

$ErrorActionPreference = 'Stop'

# --- Resolve image directory to full path ---
$resolvedImageDir = $ImageDirectory
if (-not [System.IO.Path]::IsPathRooted($resolvedImageDir)) {
    $resolvedImageDir = Join-Path -Path (Get-Location).Path -ChildPath $resolvedImageDir
}

if (-not (Test-Path -LiteralPath $resolvedImageDir -PathType Container)) {
    Write-Error "Image directory does not exist or is not a folder: $resolvedImageDir"
    exit 2
}

# --- Discover Device_*.png and parse device IDs ---
$imageFiles = Get-ChildItem -Path $resolvedImageDir -Filter 'Device_*.png' -File -ErrorAction Stop
$fileToDeviceId = [System.Collections.Generic.List[object]]::new()
foreach ($f in $imageFiles) {
    if ($f.Name -match '^Device_(\d+)\.png$') {
        $fileToDeviceId.Add([PSCustomObject]@{ File = $f; DeviceId = [int]$Matches[1] })
    }
}

if ($fileToDeviceId.Count -eq 0) {
    Write-Error "No Device_*.png files found in: $resolvedImageDir"
    exit 2
}

if ($WhatIf) {
    Write-Host "WhatIf: Would upload the following to NinjaOne (no token or upload performed):"
    foreach ($item in $fileToDeviceId) {
        Write-Host "  Device $($item.DeviceId) <- $($item.File.Name)"
    }
    exit 0
}

# --- Resolve credentials and base URL ---
$clientIdVal = if ($ClientId) { $ClientId } else { $env:NinjaOneClientId }
$clientSecretVal = if ($ClientSecret) { $ClientSecret } else { $env:NinjaOneClientSecret }
if ([string]::IsNullOrWhiteSpace($clientIdVal) -or [string]::IsNullOrWhiteSpace($clientSecretVal)) {
    Write-Error "ClientId and ClientSecret are required. Set -ClientId/-ClientSecret or env:NinjaOneClientId and env:NinjaOneClientSecret."
    exit 2
}

$NinjaBaseUrl = if ($NinjaOneInstance) { $NinjaOneInstance.Trim() } else { $env:NINJA_BASE_URL }
if ([string]::IsNullOrWhiteSpace($NinjaBaseUrl)) { $NinjaBaseUrl = 'ca.ninjarmm.com' }
$NinjaBaseUrl = $NinjaBaseUrl.Trim()
if ($NinjaBaseUrl -notmatch '^https?://') { $NinjaBaseUrl = "https://$NinjaBaseUrl" }
$NinjaBaseUrl = $NinjaBaseUrl.TrimEnd('/')

# --- OAuth: obtain Bearer token (in-line, no dot-sourcing) ---
$tokenUri = "$NinjaBaseUrl/ws/oauth/token"
$authBody = @{
    grant_type    = 'client_credentials'
    client_id     = $clientIdVal
    client_secret = $clientSecretVal
    scope         = 'monitoring management'
}
$authHeaders = @{
    'Accept'       = 'application/json'
    'Content-Type' = 'application/x-www-form-urlencoded'
}
try {
    $authResp = Invoke-RestMethod -Uri $tokenUri -Method POST -Headers $authHeaders -Body $authBody -UseBasicParsing -ErrorAction Stop
    $accessToken = $authResp | Select-Object -ExpandProperty access_token -ErrorAction SilentlyContinue
    if (-not $accessToken) { throw "Token response did not include access_token." }
} catch {
    Write-Error "Failed to obtain NinjaOne access token. $($_.Exception.Message)"
    exit 1
}

$authHeader = @{ 'Authorization' = "Bearer $accessToken" }

# --- Upload each file as related-item attachment (multipart) ---
$listUriTemplate = "$NinjaBaseUrl/ws/api/v2/related-items/with-entity/NODE/{0}"
$deleteUriTemplate = "$NinjaBaseUrl/ws/api/v2/related-items/{0}"
$lf = "`r`n"

foreach ($item in $fileToDeviceId) {
    $deviceId = $item.DeviceId
    $fileInfo = $item.File
    $filePath = $fileInfo.FullName

    # Check for existing related item with same description (skip or delete)
    $existingIds = [System.Collections.Generic.List[int]]::new()
    try {
        $listUri = $listUriTemplate -f $deviceId
        $listResp = Invoke-RestMethod -Uri $listUri -Method GET -Headers $authHeader -UseBasicParsing -ErrorAction Stop
        $items = if ($listResp -is [Array]) { @($listResp) } elseif ($listResp.PSObject.Properties['data']) { @($listResp.data) } elseif ($listResp.PSObject.Properties['items']) { @($listResp.items) } else { @($listResp) }
        $targetName = [System.IO.Path]::GetFileNameWithoutExtension($fileInfo.Name)
        foreach ($ri in $items) {
            if ($null -eq $ri.id) { continue }
            if ($ri.relEntityType -ne 'ATTACHMENT') { continue }
            $meta = $null
            if ($ri.value -and $ri.value.PSObject.Properties['metadata']) { $meta = $ri.value.metadata }
            if (-not $meta) { continue }
            $metaName = if ($meta.PSObject.Properties['name']) { $meta.name } else { $null }
            if ($metaName -and [string]::Equals($metaName, $targetName, [StringComparison]::OrdinalIgnoreCase)) {
                $existingIds.Add([int]$ri.id)
            }
        }
    } catch {
        if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 404) { }
        else { Write-Warning "Could not list related items for device $deviceId : $($_.Exception.Message)" }
    }

    if ($existingIds.Count -gt 0) {
        if (-not $Replace) {
            Write-Host "Skipped device $deviceId <- $($fileInfo.Name) (QR already exists; use -Replace to overwrite)"
            continue
        }
        foreach ($rid in $existingIds) {
            try {
                $delUri = $deleteUriTemplate -f $rid
                Invoke-RestMethod -Uri $delUri -Method DELETE -Headers $authHeader -UseBasicParsing -ErrorAction Stop | Out-Null
                Write-Verbose "Deleted existing related item $rid for device $deviceId"
            } catch {
                Write-Warning "Could not delete related item $rid for device $deviceId : $($_.Exception.Message)"
            }
        }
    }

    try {
        $boundary = [System.Guid]::NewGuid().ToString()
        $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
        $fileContentEncoded = [System.Text.Encoding]::GetEncoding('iso-8859-1').GetString($fileBytes)

        $bodyLines = (
            "--$boundary",
            "content-disposition: form-data; name=`"description`"$lf",
            $AttachmentDescription,
            "--$boundary",
            "content-disposition: form-data; name=`"file`"; filename=`"$($fileInfo.Name)`"",
            "content-type: image/png$lf",
            $fileContentEncoded,
            "--$boundary--$lf"
        ) -join $lf

        $uploadUri = "$NinjaBaseUrl/ws/api/v2/related-items/entity/NODE/$deviceId/attachment"
        $contentType = "multipart/form-data; boundary=`"$boundary`""

        Invoke-RestMethod -Uri $uploadUri -Method POST -Headers $authHeader -ContentType $contentType -Body $bodyLines -UseBasicParsing -ErrorAction Stop | Out-Null
        Write-Host "Uploaded device $deviceId <- $($fileInfo.Name)"
    } catch {
        Write-Warning "Failed to upload $($fileInfo.Name) to device $deviceId : $($_.Exception.Message)"
    }
}

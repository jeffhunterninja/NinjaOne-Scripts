#Requires -Version 5.1
<#
.SYNOPSIS
  Uploads user QR code images to NinjaOne end users as related-item attachments.

.DESCRIPTION
  Discovers User_<userId>.png files in a directory (e.g. output from
  New-NinjaUserQRCode.ps1), parses the user ID from each filename, and
  uploads each image to the matching NinjaOne end user as a related-item
  attachment via the NinjaOne API.

  Related items are only supported for end users. Technicians and contacts
  are skipped (no upload). Only user IDs that appear in the end-users list
  are processed.

  All logic is standalone (no dot-sourcing). Run New-NinjaUserQRCode.ps1
  first to generate the images, or point -ImageDirectory to an existing
  folder containing User_*.png files.

  By default, skips uploading if the end user already has a related item
  with the same AttachmentDescription (avoids duplicates). Use -Replace to
  delete any existing matching related item and upload again.

.PARAMETER ImageDirectory
  Folder containing User_*.png files. Default matches New-NinjaUserQRCode.ps1 output.

.PARAMETER NinjaOneInstance
  NinjaOne instance hostname or base URL. Default: $env:NINJA_BASE_URL or ca.ninjarmm.com.

.PARAMETER ClientId
  NinjaOne API application Client ID. Default: $env:NinjaOneClientId.

.PARAMETER ClientSecret
  NinjaOne API application Client Secret. Default: $env:NinjaOneClientSecret.

.PARAMETER AttachmentDescription
  Description stored with the related item in NinjaOne.

.PARAMETER Replace
  If the end user already has a related item with the same AttachmentDescription,
  delete it and upload the new image. Default is to skip when one already exists.

.EXAMPLE
  .\Send-NinjaUserQRCodeToEndUser.ps1 -ImageDirectory .\UserQRCodes

.EXAMPLE
  .\Send-NinjaUserQRCodeToEndUser.ps1 -Replace
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]
    $ImageDirectory = '.\AEUserQRCodes',

    [Parameter()]
    [string]
    $NinjaOneInstance = 'rcs-sales.rmmservice.ca',

    [string]
    $ClientId = $env:NinjaOneClientId,

    [string]
    $ClientSecret = $env:NinjaOneClientSecret,

    [Parameter()]
    [string]
    $AttachmentDescription = 'User dashboard QR code',

    [Parameter()]
    [switch]
    $Replace = $true
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

# --- Discover User_*.png and parse user IDs ---
$imageFiles = Get-ChildItem -Path $resolvedImageDir -Filter 'User_*.png' -File -ErrorAction Stop
$fileToUserId = [System.Collections.Generic.List[object]]::new()
foreach ($f in $imageFiles) {
    if ($f.Name -match '^User_(\d+)\.png$') {
        $fileToUserId.Add([PSCustomObject]@{ File = $f; UserId = [int]$Matches[1] })
    }
}

if ($fileToUserId.Count -eq 0) {
    Write-Error "No User_*.png files found in: $resolvedImageDir"
    exit 2
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

# --- OAuth: obtain Bearer token ---
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

# --- Fetch end users (only end users support related items; technicians/contacts are skipped) ---
$endUserIds = @{}
try {
    $endUsersUri = "$NinjaBaseUrl/ws/api/v2/user/end-users"
    $endUsersResp = Invoke-RestMethod -Uri $endUsersUri -Method GET -Headers $authHeader -UseBasicParsing -ErrorAction Stop
    $endUsersList = if ($endUsersResp -is [Array]) { @($endUsersResp) } elseif ($endUsersResp.PSObject.Properties['data']) { @($endUsersResp.data) } elseif ($endUsersResp.PSObject.Properties['items']) { @($endUsersResp.items) } else { @($endUsersResp) }
    foreach ($eu in $endUsersList) {
        $id = $eu.id
        if ($null -ne $id) {
            $n = 0
            if ([int]::TryParse($id.ToString(), [ref]$n)) { $endUserIds[$n] = $true }
        }
    }
} catch {
    Write-Error "Failed to fetch end users from NinjaOne API. $($_.Exception.Message)"
    exit 1
}

# --- Related-items list: try with-entity first, then with-related-entity if 404 ---
$listUriTemplateWithEntity = "$NinjaBaseUrl/ws/api/v2/related-items/with-entity/END_USER/{0}"
$listUriTemplateWithRelatedEntity = "$NinjaBaseUrl/ws/api/v2/related-items/with-related-entity/END_USER/{0}"
$deleteUriTemplate = "$NinjaBaseUrl/ws/api/v2/related-items/{0}"

foreach ($item in $fileToUserId) {
    $userId = $item.UserId
    $fileInfo = $item.File
    $filePath = $fileInfo.FullName

    if (-not $endUserIds.ContainsKey($userId)) {
        Write-Host "Skipped $($fileInfo.Name) - ID $userId is not an end user (technician/contact)"
        continue
    }

    # List existing related items for this end user; match by name (filename) for -Replace; optional description filter
    $existingIds = [System.Collections.Generic.List[int]]::new()
    $listUri = $listUriTemplateWithEntity -f $userId
    $listResp = $null
    try {
        $listResp = Invoke-RestMethod -Uri $listUri -Method GET -Headers $authHeader -UseBasicParsing -ErrorAction Stop
    } catch {
        if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 404) {
            try {
                $listUri = $listUriTemplateWithRelatedEntity -f $userId
                $listResp = Invoke-RestMethod -Uri $listUri -Method GET -Headers $authHeader -UseBasicParsing -ErrorAction Stop
            } catch {
                if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 404) { }
                else { Write-Warning "Could not list related items for end user $userId : $($_.Exception.Message)" }
            }
        } else {
            Write-Warning "Could not list related items for end user $userId : $($_.Exception.Message)"
        }
    }
    if ($listResp) {
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
    }

    if ($existingIds.Count -gt 0) {
        if (-not $Replace) {
            Write-Host "Skipped $($fileInfo.Name) (already exists; use -Replace to overwrite)"
            continue
        }
        foreach ($rid in $existingIds) {
            try {
                $delUri = $deleteUriTemplate -f $rid
                Invoke-RestMethod -Uri $delUri -Method DELETE -Headers $authHeader -UseBasicParsing -ErrorAction Stop | Out-Null
                Write-Verbose "Deleted existing related item $rid ($($fileInfo.Name), '$AttachmentDescription') for end user $userId"
            } catch {
                Write-Warning "Could not delete related item $rid for end user $userId : $($_.Exception.Message)"
            }
        }
    }

    try {
        $boundary = [System.Guid]::NewGuid().ToString()
        $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
        $enc = [System.Text.Encoding]::UTF8
        $bodyParts = [System.Collections.Generic.List[byte]]::new()

        $preamble = "--$boundary`r`nContent-Disposition: form-data; name=`"description`"`r`n`r`n$AttachmentDescription`r`n"
        $bodyParts.AddRange([byte[]]$enc.GetBytes($preamble))
        $filePartHeaders = "--$boundary`r`nContent-Disposition: form-data; name=`"file`"; filename=`"$($fileInfo.Name)`"`r`nContent-Type: image/png`r`n`r`n"
        $bodyParts.AddRange([byte[]]$enc.GetBytes($filePartHeaders))
        $bodyParts.AddRange([byte[]]$fileBytes)
        $closing = "`r`n--$boundary--`r`n"
        $bodyParts.AddRange([byte[]]$enc.GetBytes($closing))

        $bodyBytes = $bodyParts.ToArray()
        $uploadUri = "$NinjaBaseUrl/ws/api/v2/related-items/entity/END_USER/$userId/attachment"
        $contentType = "multipart/form-data; boundary=`"$boundary`""

        Invoke-RestMethod -Uri $uploadUri -Method POST -Headers $authHeader -ContentType $contentType -Body $bodyBytes -UseBasicParsing -ErrorAction Stop | Out-Null
        Write-Host "Uploaded end user $userId <- $($fileInfo.Name)"
    } catch {
        Write-Warning "Failed to upload $($fileInfo.Name) to end user $userId : $($_.Exception.Message)"
    }
}

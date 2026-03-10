#Requires -Version 5.1
<#
.SYNOPSIS
  Creates QR code images that link to NinjaOne device dashboards.

.DESCRIPTION
  Accepts a list of NinjaOne device IDs and generates one PNG QR code per device.
  Each QR code encodes the NinjaOne device dashboard URL for that device so that
  scanning the code opens the device in the NinjaOne portal.

  By default, only the instance hostname/URL is used to build the dashboard links.
  Optionally, use -PullAllFromApi to fetch all devices from the NinjaOne API
  (requires OAuth ClientId and ClientSecret) and generate QR codes for every device.

  By default, QR images are generated via a public API (api.qrserver.com). The
  device URL is sent to that service. For air-gapped or high-sensitivity
  environments, consider using a local QR generator or a different script.

.PARAMETER DeviceIds
  One or more NinjaOne device IDs (integers). Can be combined with -Path and pipeline.

.PARAMETER Path
  Path to a file containing device IDs: one ID per line, or a CSV with a column
  named "DeviceId" or "Device ID". Lines that are not a positive integer are skipped.

.PARAMETER NinjaOneInstance
  NinjaOne instance hostname or base URL (e.g. app.ninjarmm.com or https://app.ninjarmm.com).
  Defaults to env:NINJA_BASE_URL or https://app.ninjarmm.com.

.PARAMETER OutputDirectory
  Directory where PNG files will be saved. Created if it does not exist.
  Default: .\DeviceQRCodes

.PARAMETER Size
  QR code image size in pixels (width and height). Default: 200.

.PARAMETER UseQrApi
  Use the public QR API to generate images (default). When $true, the device
  dashboard URL is sent to the API. Set $false only if you implement an
  alternative (e.g. local generator) in the script.

.PARAMETER PullAllFromApi
  When set, fetch all devices from the NinjaOne API and generate a QR code for
  each. Requires -ClientId and -ClientSecret (or env NinjaOneClientId and
  NinjaOneClientSecret). When used, -DeviceIds and -Path are ignored.

.PARAMETER ClientId
  OAuth client ID. Required when -PullAllFromApi is set. Default: env:NinjaOneClientId.

.PARAMETER ClientSecret
  OAuth client secret. Required when -PullAllFromApi is set. Default: env:NinjaOneClientSecret. 

.EXAMPLE
  .\New-NinjaDeviceQRCode.ps1 -DeviceIds 101,102,103 -NinjaOneInstance app.ninjarmm.com

.EXAMPLE
  .\New-NinjaDeviceQRCode.ps1 -Path .\device-ids.txt -OutputDirectory C:\QR

.EXAMPLE
  Get-Content .\ids.txt | .\New-NinjaDeviceQRCode.ps1 -OutputDirectory .\QR

.EXAMPLE
  .\New-NinjaDeviceQRCode.ps1 -PullAllFromApi -NinjaOneInstance ca.ninjarmm.com -ClientId "..." -ClientSecret "..."
#>

[CmdletBinding()]
param(
    [Parameter(ValueFromPipeline = $true)]
    [object[]]
    $DeviceIds,

    [Parameter()]
    [ValidateScript({ [string]::IsNullOrWhiteSpace($_) -or (Test-Path -LiteralPath $_ -PathType Leaf) })]
    [string]
    $Path,

    [string]
    $NinjaOneInstance = "ca.ninjarmm.com",

    [string]
    $OutputDirectory = '.\DeviceQRCodes',

    [ValidateRange(100, 600)]
    [int]
    $Size = 200,

    [switch]
    $UseQrApi = $true,

    [switch]
    $PullAllFromApi,

    [string]
    $ClientId = $env:NinjaOneClientId,

    [string]
    $ClientSecret = $env:NinjaOneClientSecret
)

$ErrorActionPreference = 'Stop'

# --- Normalize NinjaOne instance URL (needed for both API and QR URL building) ---
$NinjaBaseUrl = $NinjaOneInstance
if ([string]::IsNullOrWhiteSpace($NinjaBaseUrl)) { $NinjaBaseUrl = 'https://app.ninjarmm.com' }
$NinjaBaseUrl = $NinjaBaseUrl.Trim()
if ($NinjaBaseUrl -notmatch '^https?://') { $NinjaBaseUrl = "https://$NinjaBaseUrl" }
$NinjaBaseUrl = $NinjaBaseUrl.TrimEnd('/')

# --- Resolve all device IDs: either from API (PullAllFromApi) or from parameters/pipeline/file ---
$deviceIdsToUse = @()
if ($PullAllFromApi) {
    $clientIdVal = if ($ClientId) { $ClientId.Trim() } else { '' }
    $clientSecretVal = if ($ClientSecret) { $ClientSecret.Trim() } else { '' }
    if ([string]::IsNullOrWhiteSpace($clientIdVal) -or [string]::IsNullOrWhiteSpace($clientSecretVal)) {
        Write-Error "When -PullAllFromApi is set, ClientId and ClientSecret are required. Set -ClientId and -ClientSecret or env:NinjaOneClientId and env:NinjaOneClientSecret."
        exit 2
    }
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
        $authResp = Invoke-RestMethod -Uri "$NinjaBaseUrl/ws/oauth/token" -Method POST -Headers $authHeaders -Body $authBody -UseBasicParsing -ErrorAction Stop
        $accessToken = $authResp | Select-Object -ExpandProperty access_token -ErrorAction SilentlyContinue
        if (-not $accessToken) { throw "Token response did not include access_token." }
    } catch {
        Write-Error "Failed to obtain NinjaOne access token. $($_.Exception.Message)"
        exit 1
    }
    $apiHeaders = @{ 'Authorization' = "Bearer $accessToken"; 'Accept' = 'application/json' }
    try {
        $devices = Invoke-RestMethod -Uri "$NinjaBaseUrl/ws/api/v2/devices-detailed" -Method GET -Headers $apiHeaders -UseBasicParsing -ErrorAction Stop
        $devicesList = @($devices)
        $collectedIds = [System.Collections.Generic.List[int]]::new()
        foreach ($d in $devicesList) {
            $idVal = $d.id
            if ($null -ne $idVal) {
                $n = 0
                if ([int]::TryParse($idVal.ToString(), [ref]$n) -and $n -gt 0) { [void]$collectedIds.Add($n) }
            }
        }
        $deviceIdsToUse = @($collectedIds | Sort-Object -Unique)
    } catch {
        Write-Error "Failed to fetch devices from NinjaOne API. $($_.Exception.Message)"
        exit 1
    }
} else {
    # --- Collect pipeline / -DeviceIds (allow strings from pipeline, coerce to int) ---
    $pipelineIds = [System.Collections.Generic.List[int]]::new()
    if ($DeviceIds) {
        foreach ($id in $DeviceIds) {
            if ($id -eq $null) { continue }
            $s = $id.ToString().Trim()
            $n = 0
            if ([int]::TryParse($s, [ref]$n) -and $n -gt 0) { $pipelineIds.Add($n) }
        }
    }

    # --- Resolve all device IDs from file if -Path was provided ---
    function Get-DeviceIdsFromFile {
    param([string]$FilePath)
    $resolved = [System.Collections.Generic.List[int]]::new()
    $content = Get-Content -LiteralPath $FilePath -ErrorAction Stop
    if (-not $content -or $content.Count -eq 0) { return @($resolved) }
    # Try CSV: first line has headers, look for DeviceId or "Device ID"
    $firstLine = $content[0].Trim()
    if ($firstLine -match '[,;\t]') {
        try {
            $csv = $content | ConvertFrom-Csv
            if ($csv) {
                $cols = @($csv[0].PSObject.Properties.Name)
                $idCol = $cols | Where-Object { $_ -eq 'DeviceId' -or $_ -eq 'Device ID' } | Select-Object -First 1
                if ($idCol) {
                    foreach ($row in $csv) {
                        $val = $row.$idCol
                        if ($val -and $val.ToString().Trim() -match '^\d+$' -and [int]$val -gt 0) {
                            $resolved.Add([int]$val)
                        }
                    }
                    return @($resolved)
                }
            }
        } catch { }
    }
    # Plain list: one ID per line
    foreach ($line in $content) {
        $line = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line -match '^\d+$' -and [int]$line -gt 0) {
            $resolved.Add([int]$line)
        }
    }
    return @($resolved)
}

    # Read IDs from file if -Path was provided
    $fileIds = @()
    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        $fileIds = Get-DeviceIdsFromFile -FilePath $Path
    }

    # Combined unique list: pipeline/DeviceIds first, then file
    $allIds = [System.Collections.Generic.List[int]]::new()
    foreach ($id in $pipelineIds) { $allIds.Add($id) }
    if ($Path) {
        foreach ($id in $fileIds) { $allIds.Add($id) }
    }
    $deviceIdsToUse = @($allIds | Sort-Object -Unique)
}

if ($deviceIdsToUse.Count -eq 0) {
    Write-Error "No device IDs provided. Use -DeviceIds, -Path, pipeline input, or -PullAllFromApi (e.g. Get-Content ids.txt | .\New-NinjaDeviceQRCode.ps1)."
    exit 2
}

# --- Ensure output directory exists ---
$outDir = $OutputDirectory
if (-not [System.IO.Path]::IsPathRooted($outDir)) {
    $outDir = Join-Path -Path (Get-Location).Path -ChildPath $outDir
}
if (-not (Test-Path -LiteralPath $outDir -PathType Container)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    Write-Verbose "Created output directory: $outDir"
}

# --- QR generation: public API (device URL is sent to the API) ---
$qrApiBase = 'https://api.qrserver.com/v1/create-qr-code/'

foreach ($deviceId in $deviceIdsToUse) {
    $deviceUrl = "$NinjaBaseUrl/#/deviceDashboard/$deviceId/overview"
    $fileName = "Device_$deviceId.png"
    $filePath = Join-Path -Path $outDir -ChildPath $fileName

    if ($UseQrApi) {
        $encodedData = [uri]::EscapeDataString($deviceUrl)
        $requestUrl = "${qrApiBase}?size=${Size}x${Size}&data=$encodedData&format=png"
        try {
            $response = Invoke-WebRequest -Uri $requestUrl -Method GET -UseBasicParsing
            [System.IO.File]::WriteAllBytes($filePath, $response.Content)
        } catch {
            Write-Error "Failed to generate QR for device $deviceId : $($_.Exception.Message)"
            continue
        }
    } else {
        Write-Error "Only -UseQrApi is implemented. No local QR generator is configured for device $deviceId."
        continue
    }

    Write-Output (Get-Item -LiteralPath $filePath)
}

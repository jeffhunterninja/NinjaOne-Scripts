<#
.SYNOPSIS
    Retrieves automation data from the Activities SQLite DB, filters by time frame, and generates structured HTML reports suitable for NinjaOne KB articles. Requires NinjaOne API credentials (custom fields when run in NinjaOne).
.DESCRIPTION
    Reads from the same SQLite database populated by Get-AutomationActivities.ps1. Creates a folder structure under BaseOutputFolder with a single subfolder
    "Past 10 Days". In that folder it generates:
      - AutomationHeadsUp.html (global success/failure summary at top, then tile-based view showing most recent execution per script per device)
      - AutomationDetails/ (per-automation pages for all automations, showing most recent execution per script per device)
      - DeviceDetails/ (per-device pages for all devices, showing full history of all executions)
      - Organizations/ subfolders (structure only; org-specific pages can be extended later).
    Uses sqlite3.exe (no DLL). Output uses inline styles only for NinjaOne KB compatibility.
    Credentials are required. When run in the NinjaOne API server framework, credentials are read from custom fields (ninjaoneInstance, ninjaoneClientId, ninjaoneClientSecret). Otherwise use env vars NINJAONE_INSTANCE, NINJAONE_CLIENT_ID, NINJAONE_CLIENT_SECRET or script parameters. Resolution order: Ninja-Property-Get -> env -> params. Uses NinjaOneDocs module and Connect-NinjaOne. All report links (KB article links and device dashboard links) can target a different base URL (e.g. branded/partner portals like rcs-sales.rmmservice.ca) via -KbLinkBaseUrl or env NINJAONE_KB_LINK_BASE_URL; when not set, links use the API instance.
.PARAMETER DbPath
    Path to the SQLite database file. Defaults to C:\RMM\Activities.db (same as Get-AutomationActivities).
.PARAMETER SqliteExePath
    Full path to sqlite3.exe. If not set, script directory, PATH, then C:\RMM\sqlite3.exe are tried.
.PARAMETER BaseOutputFolder
    Top-level folder for report output. Defaults to C:\RMM\Reports\Script Tracking.
.PARAMETER NinjaOneInstance
    NinjaOne instance host (e.g. app.ninjaone.com). Used to build device dashboard links. When run in NinjaOne, credentials are expected from custom fields (API server framework); this parameter is a fallback.
.PARAMETER NinjaOneClientId
    NinjaOne API client ID. When run in NinjaOne, use custom property ninjaoneClientId; else use env NINJAONE_CLIENT_ID or this parameter.
.PARAMETER NinjaOneClientSecret
    NinjaOne API client secret. When run in NinjaOne, use custom property ninjaoneClientSecret; else use env NINJAONE_CLIENT_SECRET or this parameter.
.PARAMETER KbArticleIdMapPath
    Path to the KB article ID mapping file (.kb-article-ids.json). When present, links to other report articles (e.g. Automation Detail pages) are emitted as NinjaOne deep links. Default: BaseOutputFolder\.kb-article-ids.json. Generate or refresh the file with Update-ScriptTrackerKbArticleMap.ps1.
.PARAMETER KbLinkBaseUrl
    Base URL or host for all report links: KB article links and device dashboard links (e.g. rcs-sales.rmmservice.ca or https://rcs-sales.rmmservice.ca). When empty (default), links use the same instance as NinjaOneInstance. Optional: set via env NINJAONE_KB_LINK_BASE_URL or this parameter. Use for branded/partner portals so links open on the correct site. Parameter takes precedence when non-empty.
.PARAMETER MaxDetailRows
    Maximum number of table rows to emit in each automation or device detail page. When exceeded, only the first N rows are shown and a truncation message is appended. Helps keep HTML under NinjaOne KB article size limits. Default: 5000.
.PARAMETER MaxDetailHtmlChars
    Optional safety cap (characters) on detail page body size. If accumulated content exceeds this while building the table, no more rows are added and a truncation message is appended. Default: 18000000 (18M) to stay under the 20M API limit.
.LINK
    https://www.sqlite.org/download.html
.LINK
    Update-ScriptTrackerKbArticleMap.ps1
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$DbPath = 'C:\RMM\Activities.db',
    [Parameter()]
    [string]$SqliteExePath = 'C:\ProgramData\chocolatey\bin\sqlite3.exe',
    [Parameter()]
    [string]$BaseOutputFolder = 'C:\RMM\Reports\Script Tracking',
    [Parameter()]
    [string]$NinjaOneInstance = '',
    [Parameter()]
    [string]$NinjaOneClientId = '',
    [Parameter()]
    [string]$NinjaOneClientSecret = '',
    [Parameter()]
    [string]$KbArticleIdMapPath = 'c:\RMM\Reports\Script Tracking\.kb-article-ids.json',
    [Parameter()]
    [string]$KbLinkBaseUrl = '',
    [Parameter()]
    [int]$MaxDetailRows = 5000,
    [Parameter()]
    [int]$MaxDetailHtmlChars = 18000000
)

$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrWhiteSpace($DbPath)) { $DbPath = 'C:\RMM\Activities.db' }
$dbFile = $DbPath
$baseOutputFolder = $BaseOutputFolder

$timeframeDays = 10
$timeframeName = "Past 10 Days"

$Start = Get-Date

# --- Resolve sqlite3.exe (same as Get-AutomationActivities / Invoke-ScriptStatusSync) ---
$sqliteExe = $null
if (-not [string]::IsNullOrWhiteSpace($SqliteExePath)) {
    if ((Test-Path -LiteralPath $SqliteExePath -PathType Leaf)) { $sqliteExe = $SqliteExePath }
    else { throw "SqliteExePath specified but file not found: $SqliteExePath. Download sqlite3.exe from https://www.sqlite.org/download.html and pass a valid -SqliteExePath." }
}
if (-not $sqliteExe -and $scriptDir) {
    $candidate = Join-Path $scriptDir 'sqlite3.exe'
    if (Test-Path -LiteralPath $candidate -PathType Leaf) { $sqliteExe = $candidate }
}
if (-not $sqliteExe) {
    $cmd = Get-Command sqlite3 -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -and (Test-Path -LiteralPath $cmd.Source -PathType Leaf)) { $sqliteExe = $cmd.Source }
}
if (-not $sqliteExe -and (Test-Path -LiteralPath 'C:\RMM\sqlite3.exe' -PathType Leaf)) { $sqliteExe = 'C:\RMM\sqlite3.exe' }
if (-not $sqliteExe) {
    throw "sqlite3.exe not found. Place it in the script directory, add to PATH, set -SqliteExePath, or install to C:\RMM\sqlite3.exe. Download from https://www.sqlite.org/download.html (Precompiled Binaries for Windows)."
}

if (-not (Test-Path -LiteralPath $dbFile -PathType Leaf)) {
    throw "SQLite database not found: $dbFile. Run Get-AutomationActivities.ps1 first to create and populate the Activities database."
}

# --- NinjaOneDocs module (must load before Ninja-Property-Get) ---
try {
    $moduleName = "NinjaOneDocs"
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Install-Module -Name $moduleName -Force -AllowClobber
    }
    Import-Module $moduleName -ErrorAction Stop
} catch {
    Write-Error "Failed to import NinjaOneDocs module. Error: $_"
    exit 1
}

# --- Credentials: Ninja-Property-Get (when in NinjaOne) -> env vars -> parameters ---
$resolvedInstance = $null
$resolvedClientId = $null
$resolvedClientSecret = $null
try {
    $fromNinja = Ninja-Property-Get ninjaoneInstance
    if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $resolvedInstance = $fromNinja }
} catch { }
if ([string]::IsNullOrWhiteSpace($resolvedInstance)) { $resolvedInstance = $env:NINJAONE_INSTANCE }
if ([string]::IsNullOrWhiteSpace($resolvedInstance) -and $PSBoundParameters.ContainsKey('NinjaOneInstance')) { $resolvedInstance = $NinjaOneInstance }

try {
    $fromNinja = Ninja-Property-Get ninjaoneClientId
    if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $resolvedClientId = $fromNinja }
} catch { }
if ([string]::IsNullOrWhiteSpace($resolvedClientId)) { $resolvedClientId = $env:NINJAONE_CLIENT_ID }
if ([string]::IsNullOrWhiteSpace($resolvedClientId) -and $PSBoundParameters.ContainsKey('NinjaOneClientId')) { $resolvedClientId = $NinjaOneClientId }

try {
    $fromNinja = Ninja-Property-Get ninjaoneClientSecret
    if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $resolvedClientSecret = $fromNinja }
} catch { }
if ([string]::IsNullOrWhiteSpace($resolvedClientSecret)) { $resolvedClientSecret = $env:NINJAONE_CLIENT_SECRET }
if ([string]::IsNullOrWhiteSpace($resolvedClientSecret) -and $PSBoundParameters.ContainsKey('NinjaOneClientSecret')) { $resolvedClientSecret = $NinjaOneClientSecret }

if ([string]::IsNullOrWhiteSpace($resolvedInstance) -or [string]::IsNullOrWhiteSpace($resolvedClientId) -or [string]::IsNullOrWhiteSpace($resolvedClientSecret)) {
    Write-Error "Missing required API credentials. Set ninjaoneInstance, ninjaoneClientId, ninjaoneClientSecret in NinjaOne custom properties, or use env vars NINJAONE_INSTANCE, NINJAONE_CLIENT_ID, NINJAONE_CLIENT_SECRET, or -NinjaOneInstance, -NinjaOneClientId, -NinjaOneClientSecret."
    exit 1
}

try {
    Connect-NinjaOne -NinjaOneInstance $resolvedInstance -NinjaOneClientID $resolvedClientId -NinjaOneClientSecret $resolvedClientSecret
} catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit 1
}

# --- Resolve KB link base URL (for links between KB articles); default to API instance when not set ---
# Parameter and env only (no Ninja custom field). Parameter takes precedence when non-empty so -KbLinkBaseUrl or script default is used.
$kbLinkBase = $null
if (-not [string]::IsNullOrWhiteSpace($KbLinkBaseUrl)) { $kbLinkBase = $KbLinkBaseUrl.Trim() }
if ([string]::IsNullOrWhiteSpace($kbLinkBase)) { $kbLinkBase = $env:NINJAONE_KB_LINK_BASE_URL }
if ([string]::IsNullOrWhiteSpace($kbLinkBase)) { $kbLinkBase = $resolvedInstance }
$script:KbLinkBaseUrlResolved = $kbLinkBase

# --- Resolve KB article ID map path and load map if present ---
$kbArticleIdMapPathResolved = $KbArticleIdMapPath
if ([string]::IsNullOrWhiteSpace($kbArticleIdMapPathResolved)) { $kbArticleIdMapPathResolved = Join-Path $baseOutputFolder '.kb-article-ids.json' }
$script:KbArticleIdMap = @{}
if (Test-Path -LiteralPath $kbArticleIdMapPathResolved -PathType Leaf) {
    try {
        $json = Get-Content -LiteralPath $kbArticleIdMapPathResolved -Raw -Encoding UTF8
        if (-not [string]::IsNullOrWhiteSpace($json)) {
            $parsed = $json | ConvertFrom-Json
            $parsed.PSObject.Properties | ForEach-Object {
                $key = $_.Name
                $val = $_.Value
                if ($val -is [PSCustomObject] -and $val.PSObject.Properties['id'] -and $val.PSObject.Properties['parentFolderId']) {
                    $script:KbArticleIdMap[$key] = @{ id = [long]$val.id; parentFolderId = [long]$val.parentFolderId }
                }
            }
        }
    } catch {
        Write-Warning "Could not load KB article ID map from '$kbArticleIdMapPathResolved': $_. Links will use fallback URLs."
    }
}

# --- In-line: build KB deep link or search URL for an article name (for use in generated HTML) ---
function Get-KBLinkForArticle {
    param([Parameter(Mandatory)][string]$ArticleName, [Parameter(Mandatory)][string]$BaseUrl)
    $base = if ([string]::IsNullOrWhiteSpace($BaseUrl)) { '' } elseif ($BaseUrl -match '^https?://') { $BaseUrl } else { "https://$BaseUrl" }
    $base = $base -replace '/+$', ''
    if ([string]::IsNullOrWhiteSpace($base)) { return '#' }
    $trimName = if ([string]::IsNullOrWhiteSpace($ArticleName)) { '' } else { $ArticleName.Trim() }
    if ($script:KbArticleIdMap -and $script:KbArticleIdMap.ContainsKey($trimName)) {
        $entry = $script:KbArticleIdMap[$trimName]
        return "$base/#/systemDashboard/knowledgeBase/$($entry.parentFolderId)/$($entry.id)/file"
    }
    return "$base/#/knowledgebase/global/articles?articleName=$([uri]::EscapeDataString($trimName))"
}

# --- SQLite helpers (in-line, no dot-sourcing) ---
function Invoke-SqliteQuery {
    param([Parameter(Mandatory)] [string]$SqliteExe, [Parameter(Mandatory)] [string]$DataSource, [Parameter(Mandatory)] [string]$Sql)
    $out = & $SqliteExe -csv -header $DataSource $Sql 2>$null
    $text = if ($null -eq $out) { '' } elseif ($out -is [string]) { $out } else { $out -join "`n" }
    $text = $text.TrimStart([char]0xFEFF)
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    $lines = $text -split "`r?`n"
    $lines = $lines | Where-Object { $_.Length -gt 0 }
    if ($lines.Count -lt 2) { return @() }
    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText($tempFile, $text, $utf8NoBom)
        $result = Import-Csv -Path $tempFile -Encoding UTF8
        return @($result)
    } finally {
        if (Test-Path -LiteralPath $tempFile) { Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue }
    }
}

# --- Epoch conversion for date range (supports both seconds and milliseconds in DB) ---
$epoch0 = [datetime]'1970-01-01T00:00:00Z'
function Get-EpochSeconds { param([Parameter(Mandatory)][datetime]$Date) $utc = $Date.ToUniversalTime(); [int64][math]::Floor(($utc - $epoch0).TotalSeconds) }

# --- Safe path segment for file/folder names (illegal chars, truncation, empty placeholder) ---
function Get-SafePathSegment {
    param([string]$Value, [string]$Placeholder = 'Unknown', [int]$MaxLength = 80)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace($Value)) { return $Placeholder }
    $s = [string]$Value
    $s = $s -replace '[\\/:*?"<>|]', '_'
    if ($s.Length -gt $MaxLength) { $s = $s.Substring(0, $MaxLength) }
    return $s
}

# --- HTML page writer: inline styles only for KB compatibility ---
function Write-HTMLPage {
    param([string]$FilePath, [string]$Title, [string]$BodyContent)
    $parentDir = [System.IO.Path]::GetDirectoryName($FilePath)
    if (-not [string]::IsNullOrWhiteSpace($parentDir) -and -not (Test-Path -LiteralPath $parentDir -PathType Container)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>$Title</title>
</head>
<body style="margin: 0; padding: 20px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height: 1.5;">
  <h1 style="margin: 0 0 20px 0; color: #333; font-size: 1.5rem;">$Title</h1>
  $BodyContent
</body>
</html>
"@
    $html | Out-File -FilePath $FilePath -Encoding utf8
}

# --- Convert activityTime (epoch sec or ms) to local time string ---
function Convert-ActivityTimeToLocalString {
    param($TimeValue)
    if ($null -eq $TimeValue -or [string]::IsNullOrWhiteSpace([string]$TimeValue)) { return 'N/A' }
    $num = $null
    if (-not [double]::TryParse([string]$TimeValue, [ref]$num)) { return 'N/A' }
    $sec = if ($num -ge 1000000000000) { [long][math]::Floor($num / 1000) } else { [long][math]::Floor($num) }
    try { return [DateTimeOffset]::FromUnixTimeSeconds($sec).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") }
    catch { return 'N/A' }
}

# --- Ensure base output folder exists ---
if (-not (Test-Path $baseOutputFolder)) {
    New-Item -ItemType Directory -Path $baseOutputFolder | Out-Null
}

# --- COMPLETED + ACTION filter (same as Invoke-ScriptStatusSync / Dashboard) ---
$statusCondition = "( (statusCode IS NOT NULL AND UPPER(TRIM(statusCode)) = 'COMPLETED') OR (status IS NOT NULL AND UPPER(TRIM(status)) = 'COMPLETED') )"
$typeCondition   = "( (activityType IS NOT NULL AND UPPER(TRIM(activityType)) = 'ACTION') OR (type IS NOT NULL AND UPPER(TRIM(type)) = 'ACTION') )"

$folderName = $timeframeName
$days = $timeframeDays
$outputFolder = Join-Path $baseOutputFolder $folderName
if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

# --- Epoch-based date filter (supports both sec and ms in DB) ---
    $rangeStart = (Get-Date).AddDays(-$days)
    $rangeEndEx = (Get-Date).AddSeconds(1)
    $afterEpoch = Get-EpochSeconds -Date $rangeStart
    $beforeEpochEx = Get-EpochSeconds -Date $rangeEndEx
    $afterEpochMs = $afterEpoch * 1000
    $beforeEpochExMs = $beforeEpochEx * 1000
    $timeCondition = " AND ( (activityTime >= $afterEpoch AND activityTime < $beforeEpochEx) OR (activityTime >= $afterEpochMs AND activityTime < $beforeEpochExMs) )"

    $whereClause = "WHERE $statusCondition AND $typeCondition $timeCondition"

    # --- Queries: full detail (for Device Detail pages) and latest per script per device (for metrics, Heads Up, Automation Detail) ---
    $queryDetails = "SELECT DeviceName, deviceId, sourceName, message, activityResult, activityTime, OrgName FROM Activities $whereClause ORDER BY OrgName, sourceName, activityTime DESC;"
    $queryLatest = "WITH Ranked AS ( SELECT DeviceName, deviceId, sourceName, message, activityResult, activityTime, OrgName, ROW_NUMBER() OVER (PARTITION BY sourceName, COALESCE(deviceId, -1) ORDER BY activityTime DESC) AS rn FROM Activities $whereClause ) SELECT DeviceName, deviceId, sourceName, message, activityResult, activityTime, OrgName FROM Ranked WHERE rn = 1;"

    $detailRows = Invoke-SqliteQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql $queryDetails
    $latestRows = Invoke-SqliteQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql $queryLatest

    # --- Build full details list (for Device Detail pages only) ---
    $automationDetails = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($r in $detailRows) {
        $deviceId = if ($null -ne $r.deviceId) { [string]$r.deviceId } else { '' }
        $rawDeviceName = if ($null -ne $r.DeviceName) { [string]$r.DeviceName } else { 'N/A' }
        $deviceName = $rawDeviceName
        if (-not [string]::IsNullOrWhiteSpace($script:KbLinkBaseUrlResolved) -and -not [string]::IsNullOrWhiteSpace($deviceId)) {
            $baseUrl = if ($script:KbLinkBaseUrlResolved -match '^https?://') { $script:KbLinkBaseUrlResolved } else { "https://$($script:KbLinkBaseUrlResolved)" }
            $baseUrl = $baseUrl -replace '/+$', ''
            $deviceName = "<a href='$baseUrl/#/deviceDashboard/$deviceId/overview' target='_blank' style='color: #0066cc; text-decoration: none; font-weight: 500;' title='Open device dashboard'>$deviceName</a>"
        }
        $displayTime = Convert-ActivityTimeToLocalString -TimeValue $r.activityTime
        $num = $null
        $epochSec = 0L
        if ($null -ne $r.activityTime -and [double]::TryParse([string]$r.activityTime, [ref]$num)) {
            $epochSec = if ($num -ge 1000000000000) { [long][math]::Floor($num / 1000) } else { [long][math]::Floor($num) }
        }
        [void]$automationDetails.Add([pscustomobject]@{
            DeviceName    = $deviceName
            SystemName    = $rawDeviceName
            deviceId      = $r.deviceId
            sourceName    = $r.sourceName
            message       = $r.message
            activityResult = $r.activityResult
            activityTime  = $displayTime
            activityTimeEpochSeconds = $epochSec
            OrgName       = $r.OrgName
        })
    }

    # --- Build latest-per-script list (for overall metrics, Heads Up, Automation Detail pages) ---
    $automationDetailsLatest = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($r in $latestRows) {
        $deviceId = if ($null -ne $r.deviceId) { [string]$r.deviceId } else { '' }
        $rawDeviceName = if ($null -ne $r.DeviceName) { [string]$r.DeviceName } else { 'N/A' }
        $deviceName = $rawDeviceName
        if (-not [string]::IsNullOrWhiteSpace($script:KbLinkBaseUrlResolved) -and -not [string]::IsNullOrWhiteSpace($deviceId)) {
            $baseUrl = if ($script:KbLinkBaseUrlResolved -match '^https?://') { $script:KbLinkBaseUrlResolved } else { "https://$($script:KbLinkBaseUrlResolved)" }
            $baseUrl = $baseUrl -replace '/+$', ''
            $deviceName = "<a href='$baseUrl/#/deviceDashboard/$deviceId/overview' target='_blank' style='color: #0066cc; text-decoration: none; font-weight: 500;' title='Open device dashboard'>$deviceName</a>"
        }
        $displayTime = Convert-ActivityTimeToLocalString -TimeValue $r.activityTime
        $num = $null
        $epochSec = 0L
        if ($null -ne $r.activityTime -and [double]::TryParse([string]$r.activityTime, [ref]$num)) {
            $epochSec = if ($num -ge 1000000000000) { [long][math]::Floor($num / 1000) } else { [long][math]::Floor($num) }
        }
        [void]$automationDetailsLatest.Add([pscustomobject]@{
            DeviceName    = $deviceName
            SystemName    = $rawDeviceName
            deviceId      = $r.deviceId
            sourceName    = $r.sourceName
            message       = $r.message
            activityResult = $r.activityResult
            activityTime  = $displayTime
            activityTimeEpochSeconds = $epochSec
            OrgName       = $r.OrgName
        })
    }

    # --- Overall metrics from latest execution per script (one success/failure per script) ---
    # Normalize activityResult (trim + case-insensitive) so summary and table agree; count FAILURE explicitly so unknown/empty are not shown as failures.
    $totalCount = $automationDetailsLatest.Count
    $successCount = @($automationDetailsLatest | Where-Object { ([string]$_.activityResult).Trim() -ieq 'SUCCESS' }).Count
    $failureCount = @($automationDetailsLatest | Where-Object { ([string]$_.activityResult).Trim() -ieq 'FAILURE' }).Count
    $successPercentage = if ($totalCount -gt 0) { [math]::Round(($successCount / $totalCount * 100), 2) } else { 0 }
    $failurePercentage = if ($totalCount -gt 0) { [math]::Round(($failureCount / $totalCount * 100), 2) } else { 0 }

    # --- Automation Heads-Up (global summary at top, then cards from latest execution per script) ---
    $headsUpPath = Join-Path $outputFolder "AutomationHeadsUp.html"
    $automations = $automationDetailsLatest | Group-Object -Property sourceName
    $globalSummaryHtml = "<div style='margin: 15px 0 20px 0; padding: 15px; border: 1px solid #e0e0e0; border-radius: 8px; background: #fafafa;'>"
    $globalSummaryHtml += "<p style='margin: 0 0 10px 0; font-size: 0.9rem; color: #555;'>Based on most recent execution per script per device.</p>"
    $globalSummaryHtml += "<div style='display: flex; height: 24px; border-radius: 4px; overflow: hidden; background: #f0f0f0;'>"
    $globalSummaryHtml += "<div style='width: $successPercentage%; background-color: #22c55e;'></div>"
    $globalSummaryHtml += "<div style='width: $failurePercentage%; background-color: #ef4444;'></div>"
    $globalSummaryHtml += "</div>"
    $globalSummaryHtml += "<div style='margin-top: 10px; display: flex; justify-content: space-between; flex-wrap: wrap; gap: 10px;'>"
    $globalSummaryHtml += "<span><span style='display: inline-block; width: 12px; height: 12px; margin-right: 5px; background-color: #22c55e; vertical-align: middle;'></span>Success ($successCount)</span>"
    $globalSummaryHtml += "<span><span style='display: inline-block; width: 12px; height: 12px; margin-right: 5px; background-color: #ef4444; vertical-align: middle;'></span>Failure ($failureCount)</span>"
    $globalSummaryHtml += "</div>"
    $globalSummaryHtml += "<p style='margin: 15px 0 0 0;'>Total Automations: $totalCount</p>"
    $globalSummaryHtml += "</div>"
    $introParagraph = "<p style='margin: 0 0 16px 0; color: #666; font-size: 0.9rem;'>Cards show success/failure at a glance; green top border = mostly success, red = mostly failures.</p>"
    $cardHtmls = [System.Collections.Generic.List[string]]::new()
    $automationSummaries = [System.Collections.Generic.List[pscustomobject]]::new()
    foreach ($group in $automations) {
        $success_count = @($group.Group | Where-Object { ([string]$_.activityResult).Trim() -ieq 'SUCCESS' }).Count
        $total = $group.Group.Count
        $success_percentage = if ($total -gt 0) { [math]::Round(($success_count / $total * 100), 2) } else { 0 }
        $maxEpoch = 0L
        foreach ($row in $group.Group) {
            $e = [long]$row.activityTimeEpochSeconds
            if ($e -gt $maxEpoch) { $maxEpoch = $e }
        }
        [void]$automationSummaries.Add([pscustomobject]@{ Group = $group; SuccessPercentage = $success_percentage; MaxActivityEpoch = $maxEpoch })
    }
    $sortedSummaries = $automationSummaries | Sort-Object -Property SuccessPercentage, @{ Expression = { 0 - [long]$_.MaxActivityEpoch }; Ascending = $true }
    foreach ($summary in $sortedSummaries) {
        $group = $summary.Group
        $automationName = [string]$group.Name
        $safeName = Get-SafePathSegment -Value $automationName -Placeholder 'UnnamedAutomation' -MaxLength 100
        $success_count = @($group.Group | Where-Object { ([string]$_.activityResult).Trim() -ieq 'SUCCESS' }).Count
        $failure_count = @($group.Group | Where-Object { ([string]$_.activityResult).Trim() -ieq 'FAILURE' }).Count
        $total = $group.Group.Count
        $success_percentage = if ($total -gt 0) { [math]::Round(($success_count / $total * 100), 2) } else { 0 }
        $failure_percentage = if ($total -gt 0) { [math]::Round(($failure_count / $total * 100), 2) } else { 0 }
        $borderTopColor = if ($failure_count -gt $success_count) { '#ef4444' } else { '#22c55e' }
        $targetArticleName = "$folderName - Automation Detail: $automationName"
        $automationDetailUrl = Get-KBLinkForArticle -ArticleName $targetArticleName -BaseUrl $script:KbLinkBaseUrlResolved
        $cardHtml = "<div style='width: 100%; box-sizing: border-box; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); border-top: 3px solid $borderTopColor; background: white; overflow-wrap: break-word;'>"
        $cardHtml += "<a href='$([System.Net.WebUtility]::HtmlEncode($automationDetailUrl))' target='_blank' style='color: #0066cc; text-decoration: none; font-weight: 500;'>$([System.Net.WebUtility]::HtmlEncode($automationName))</a><br/>"
        $cardHtml += "<table style='width: 100%; margin-top: 10px; border-collapse: collapse; border: 0;' role='presentation'><tr><td style='width: $success_percentage%; background-color: #22c55e; height: 20px; border: 0; padding: 0; vertical-align: top;'></td><td style='width: $failure_percentage%; background-color: #ef4444; height: 20px; border: 0; padding: 0; vertical-align: top;'></td></tr></table>"
        $cardHtml += "<div style='margin-top: 8px; font-size: 0.85rem;'><span style='display: inline-block; width: 10px; height: 10px; margin-right: 4px; background-color: #22c55e; vertical-align: middle;'></span>Success ($success_count) &nbsp; <span style='display: inline-block; width: 10px; height: 10px; margin-right: 4px; background-color: #ef4444; vertical-align: middle;'></span>Failure ($failure_count)</div></div>"
        [void]$cardHtmls.Add($cardHtml)
    }
    $tileContent = ""
    if ($cardHtmls.Count -eq 0) {
        $tileContent = $globalSummaryHtml + $introParagraph + "<div style='padding: 20px; color: #666;'>No automation data for this period.</div>"
    } else {
        $cols = 4
        $tileContent = $globalSummaryHtml + $introParagraph + "<table style='width: 100%; border-collapse: collapse; border-spacing: 0; margin: 0 0 25px 0;'><tbody>"
        for ($i = 0; $i -lt $cardHtmls.Count; $i += $cols) {
            $tileContent += "<tr>"
            for ($j = 0; $j -lt $cols; $j++) {
                $idx = $i + $j
                $cellStyle = "vertical-align: top; padding: 8px; width: 25%;"
                if ($idx -lt $cardHtmls.Count) {
                    $tileContent += "<td style='$cellStyle'>" + $cardHtmls[$idx] + "</td>"
                } else {
                    $tileContent += "<td style='$cellStyle'></td>"
                }
            }
            $tileContent += "</tr>"
        }
        $tileContent += "</tbody></table>"
    }
    $tileContent += "<div style='margin-top: 25px; padding: 15px 20px; background: #e9ecef; border-radius: 8px; color: #6c757d; font-size: 0.9rem; text-align: center;'>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>"
    Write-HTMLPage -FilePath $headsUpPath -Title "Automation Heads-Up Display ($folderName)" -BodyContent $tileContent

    # --- Folders for detail pages (clear stale content from previous run so only current window is reflected) ---
    $automationDetailsFolder = Join-Path $outputFolder "AutomationDetails"
    if (Test-Path -LiteralPath $automationDetailsFolder -PathType Container) {
        Get-ChildItem -LiteralPath $automationDetailsFolder -Filter '*.html' -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    } else {
        New-Item -ItemType Directory -Path $automationDetailsFolder | Out-Null
    }
    $deviceDetailsFolder = Join-Path $outputFolder "DeviceDetails"
    if (Test-Path -LiteralPath $deviceDetailsFolder -PathType Container) {
        Get-ChildItem -LiteralPath $deviceDetailsFolder -Filter '*.html' -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    } else {
        New-Item -ItemType Directory -Path $deviceDetailsFolder | Out-Null
    }
    $orgFolder = Join-Path $outputFolder "Organizations"
    if (Test-Path -LiteralPath $orgFolder -PathType Container) {
        Get-ChildItem -LiteralPath $orgFolder -Directory -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Load failure clusters (from Invoke-FailureClustering.ps1) when present ---
    $failureClustersByAutomation = @{}
    $failureClustersPath = Join-Path $outputFolder "failure-clusters.json"
    if (Test-Path -LiteralPath $failureClustersPath -PathType Leaf) {
        try {
            $fcJson = Get-Content -LiteralPath $failureClustersPath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($fcJson)) {
                $fcParsed = $fcJson | ConvertFrom-Json
                if ($fcParsed -and $fcParsed.byAutomation) {
                    $fcParsed.byAutomation.PSObject.Properties | ForEach-Object {
                        $failureClustersByAutomation[$_.Name] = $_.Value
                    }
                }
            }
        } catch {
            Write-Warning "Could not load failure-clusters.json: $_. Failure groupings will be omitted."
        }
    }

    # --- Per-automation detail pages (all automations) ---
    foreach ($group in $automations) {
        $automationName = [string]$group.Name
        $safeName = Get-SafePathSegment -Value $automationName -Placeholder 'UnnamedAutomation' -MaxLength 100
        $detailPath = Join-Path $automationDetailsFolder "Automation_$safeName.html"
        $content = "<style>tr.success { background-color: #dcfce7; } tr.danger { background-color: #fee2e2; }</style>"
        $content += "<h2 style='margin: 0 0 15px 0; color: #333; font-size: 1.2rem;'>Details for Automation: $([System.Net.WebUtility]::HtmlEncode($automationName))</h2>"
        $success_count = @($group.Group | Where-Object { ([string]$_.activityResult).Trim() -ieq 'SUCCESS' }).Count
        $failure_count = @($group.Group | Where-Object { ([string]$_.activityResult).Trim() -ieq 'FAILURE' }).Count
        $total = $group.Group.Count
        $success_percentage = if ($total -gt 0) { [math]::Round(($success_count / $total * 100), 2) } else { 0 }
        $failure_percentage = if ($total -gt 0) { [math]::Round(($failure_count / $total * 100), 2) } else { 0 }
        $content += "<div style='margin: 15px 0; padding: 15px; border: 1px solid #ddd; border-radius: 8px; background: #fafafa;'>"
        $content += "<div style='display: flex; height: 20px; border-radius: 4px; overflow: hidden; background: #f0f0f0;'>"
        $content += "<div style='width: $success_percentage%; background-color: #22c55e;'></div>"
        $content += "<div style='width: $failure_percentage%; background-color: #ef4444;'></div>"
        $content += "</div>"
        $content += "<div style='margin-top: 8px; display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px; font-size: 0.85rem;'>"
        $content += "<span><span style='display: inline-block; width: 10px; height: 10px; margin-right: 4px; background-color: #22c55e; vertical-align: middle;'></span>Success ($success_count)</span>"
        $content += "<span><span style='display: inline-block; width: 10px; height: 10px; margin-right: 4px; background-color: #ef4444; vertical-align: middle;'></span>Failure ($failure_count)</span>"
        $content += "</div></div>"
        # --- Failure message groupings (from Invoke-FailureClustering.ps1) when present ---
        $clustersForAutomation = $failureClustersByAutomation[$automationName]
        if ($clustersForAutomation -and @($clustersForAutomation).Count -gt 0) {
            $content += "<div style='margin: 15px 0 20px 0; padding: 12px 15px; border: 1px solid #fecaca; border-radius: 8px; background: #fef2f2;'>"
            $content += "<h3 style='margin: 0 0 10px 0; font-size: 1rem; color: #991b1b;'>Common failure message groupings</h3>"
            $content += "<p style='margin: 0 0 12px 0; font-size: 0.85rem; color: #7f1d1d;'>Clustered by similarity (TF-IDF + KMeans).</p>"
            $clusterIndex = 0
            foreach ($cluster in @($clustersForAutomation)) {
                $clusterIndex++
                $cCount = if ($cluster.count) { [int]$cluster.count } else { 0 }
                $cLabel = if ($cluster.label) { [System.Net.WebUtility]::HtmlEncode([string]$cluster.label) } else { $null }
                $cDevices = if ($cluster.affectedDeviceCount -ne $null) { [int]$cluster.affectedDeviceCount } else { $null }
                $cSample = if ($cluster.sampleMessage) { [System.Net.WebUtility]::HtmlEncode([string]$cluster.sampleMessage) } else { '' }
                if ($cSample.Length -gt 400) { $cSample = $cSample.Substring(0, 400) + "..." }
                $content += "<div style='margin-bottom: 10px; padding: 10px; border-left: 3px solid #ef4444; background: #fff; border-radius: 4px;'>"
                $headerText = "Group $clusterIndex &mdash; $cCount occurrence$(if ($cCount -ne 1) { 's' })"
                if ($cDevices -ge 0) { $headerText += " ($cDevices device$(if ($cDevices -ne 1) { 's' }))" }
                $content += "<div style='font-size: 0.85rem; font-weight: 600; color: #b91c1c; margin-bottom: 4px;'>$headerText</div>"
                if ($cLabel) { $content += "<div style='font-size: 0.8rem; color: #6b7280; margin-bottom: 4px;'>$cLabel</div>" }
                $content += "<div style='font-size: 0.85rem; color: #374151; white-space: pre-wrap; word-break: break-word;'>$cSample</div>"
                $topMsgs = $cluster.topMessages
                if ($topMsgs -and @($topMsgs).Count -gt 0) {
                    $content += "<ul style='margin: 6px 0 0 0; padding-left: 20px; font-size: 0.8rem; color: #6b7280;'>"
                    foreach ($tm in @($topMsgs) | Select-Object -First 5) {
                        $tmEnc = [System.Net.WebUtility]::HtmlEncode([string]$tm)
                        if ($tmEnc.Length -gt 200) { $tmEnc = $tmEnc.Substring(0, 200) + "..." }
                        $content += "<li>$tmEnc</li>"
                    }
                    $content += "</ul>"
                }
                $content += "</div>"
            }
            $content += "</div>"
        }
        $content += "<div style='overflow-x: auto;'><table style='width: 100%; border-collapse: collapse; font-size: 0.9rem;'>"
        $content += "<thead><tr style='background: #667eea; color: white;'><th style='padding: 8px; text-align: left;'>DeviceName</th><th style='padding: 8px; text-align: left;'>sourceName</th><th style='padding: 8px; text-align: left;'>message</th><th style='padding: 8px; text-align: left;'>activityResult</th><th style='padding: 8px; text-align: left;'>activityTime</th><th style='padding: 8px; text-align: left;'>OrgName</th></tr></thead><tbody>"
        $totalRows = $group.Group.Count
        $rowsForTable = $group.Group | Sort-Object { if (([string]$_.activityResult).Trim() -ieq 'FAILURE') { 0 } else { 1 } }
        $rowsAdded = 0
        foreach ($r in $rowsForTable) {
            if ($rowsAdded -ge $MaxDetailRows) { break }
            if ($MaxDetailHtmlChars -gt 0 -and $content.Length -ge $MaxDetailHtmlChars) { break }
            $ar = ([string]$r.activityResult).Trim()
            $rowClass = if ($ar -ieq 'SUCCESS') { 'success' } elseif ($ar -ieq 'FAILURE') { 'danger' } else { 'danger' }
            $content += "<tr class='$rowClass'><td style='padding: 8px;'>$($r.DeviceName)</td><td style='padding: 8px;'>$([System.Net.WebUtility]::HtmlEncode($r.sourceName))</td><td style='padding: 8px;'>$([System.Net.WebUtility]::HtmlEncode($r.message))</td><td style='padding: 8px;'>$([System.Net.WebUtility]::HtmlEncode($r.activityResult))</td><td style='padding: 8px;'>$([System.Net.WebUtility]::HtmlEncode($r.activityTime))</td><td style='padding: 8px;'>$([System.Net.WebUtility]::HtmlEncode($r.OrgName))</td></tr>"
            $rowsAdded++
        }
        $content += "</tbody></table></div>"
        if ($rowsAdded -lt $totalRows) {
            $content += "<p style='margin-top: 15px; padding: 10px; background: #fff3cd; border-radius: 6px; color: #856404; font-size: 0.9rem;'>Showing first $rowsAdded of $totalRows activities; content truncated to meet Knowledge Base size limits.</p>"
        }
        $content += "<div style='margin-top: 15px; color: #666; font-size: 0.85rem;'>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>"
        Write-HTMLPage -FilePath $detailPath -Title "Automation Detail: $automationName" -BodyContent $content
    }

    # --- Per-device detail pages (all devices) ---
    $devices = $automationDetails | Group-Object -Property deviceId
    foreach ($group in $devices) {
        $deviceId = [string]$group.Name
        $safeDeviceId = Get-SafePathSegment -Value $deviceId -Placeholder 'NoDeviceId'
        $systemName = if ($group.Group[0].SystemName) { [string]$group.Group[0].SystemName } else { 'Unknown' }
        $safeSystemName = Get-SafePathSegment -Value $systemName -Placeholder 'Unknown' -MaxLength 80
        $detailPath = Join-Path $deviceDetailsFolder "$safeSystemName - device $safeDeviceId.html"
        $pageTitle = "Device Detail: $systemName ($deviceId)"
        $content = "<style>tr.success { background-color: #dcfce7; } tr.danger { background-color: #fee2e2; }</style>"
        $content += "<h2 style='margin: 0 0 15px 0; color: #333; font-size: 1.2rem;'>Details for Device: $([System.Net.WebUtility]::HtmlEncode($systemName)) ($([System.Net.WebUtility]::HtmlEncode($deviceId)))</h2>"
        $content += "<div style='overflow-x: auto;'><table style='width: 100%; border-collapse: collapse; font-size: 0.9rem;'>"
        $content += "<thead><tr style='background: #667eea; color: white;'><th style='padding: 8px; text-align: left;'>DeviceName</th><th style='padding: 8px; text-align: left;'>sourceName</th><th style='padding: 8px; text-align: left;'>activityResult</th><th style='padding: 8px; text-align: left;'>activityTime</th><th style='padding: 8px; text-align: left;'>OrgName</th></tr></thead><tbody>"
        $totalRows = $group.Group.Count
        $rowsAdded = 0
        foreach ($r in ($group.Group | Sort-Object -Property activityTimeEpochSeconds -Descending)) {
            if ($rowsAdded -ge $MaxDetailRows) { break }
            if ($MaxDetailHtmlChars -gt 0 -and $content.Length -ge $MaxDetailHtmlChars) { break }
            $ar = ([string]$r.activityResult).Trim()
            $rowClass = if ($ar -ieq 'SUCCESS') { 'success' } elseif ($ar -ieq 'FAILURE') { 'danger' } else { 'danger' }
            $content += "<tr class='$rowClass'><td style='padding: 8px;'>$($r.DeviceName)</td><td style='padding: 8px;'>$([System.Net.WebUtility]::HtmlEncode($r.sourceName))</td><td style='padding: 8px;'>$([System.Net.WebUtility]::HtmlEncode($r.activityResult))</td><td style='padding: 8px;'>$([System.Net.WebUtility]::HtmlEncode($r.activityTime))</td><td style='padding: 8px;'>$([System.Net.WebUtility]::HtmlEncode($r.OrgName))</td></tr>"
            $rowsAdded++
        }
        $content += "</tbody></table></div>"
        if ($rowsAdded -lt $totalRows) {
            $content += "<p style='margin-top: 15px; padding: 10px; background: #fff3cd; border-radius: 6px; color: #856404; font-size: 0.9rem;'>Showing first $rowsAdded of $totalRows activities; content truncated to meet Knowledge Base size limits.</p>"
        }
        $content += "<div style='margin-top: 15px; color: #666; font-size: 0.85rem;'>Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</div>"
        Write-HTMLPage -FilePath $detailPath -Title $pageTitle -BodyContent $content
    }

    # --- Organizations folder (structure only) ---
    $orgFolder = Join-Path $outputFolder "Organizations"
    if (-not (Test-Path $orgFolder)) { New-Item -ItemType Directory -Path $orgFolder | Out-Null }
    $orgGroups = $automationDetails | Group-Object -Property OrgName
    foreach ($org in $orgGroups) {
        $orgName = [string]$org.Name
        $safeOrgName = Get-SafePathSegment -Value $orgName -Placeholder 'UnknownOrg' -MaxLength 80
        $orgOutputFolder = Join-Path $orgFolder $safeOrgName
        if (-not (Test-Path $orgOutputFolder)) { New-Item -ItemType Directory -Path $orgOutputFolder | Out-Null }
    }

Write-Host "Reports for $folderName generated in $outputFolder"

Write-Host "Total runtime: $((New-TimeSpan -Start $Start -End (Get-Date)).TotalSeconds) seconds"

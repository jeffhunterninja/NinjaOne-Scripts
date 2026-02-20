#Requires -Version 5.1
<#
.SYNOPSIS
  Compares NinjaOne API activities with webhook.site requests to identify delivery discrepancies.

.DESCRIPTION
  Compares NinjaOne API activities with webhook.site requests (webhook-first, date-based alignment).
  First fetches webhook.site requests in the given time window, parses NinjaOne activity payloads,
  and uses the date of the oldest activity in that dataset as the starting point. Then fetches from
  NinjaOne using the after parameter (afterUnixEpoch) so the NinjaOne set is aligned by time with
  what webhook received. If there are no webhook activities in the window, falls back to the -After
  parameter for the NinjaOne date range. Use this to identify webhook delivery gaps (in NinjaOne
  but not at webhook.site) or unexpected webhook requests (in webhook.site but not in NinjaOne API).

.PARAMETER After
  Start of the time window (DateTime). Default: 7 days ago. End of window is always now.
  Used for webhook.site request list (date_from/date_to). When falling back (no webhook activities),
  also used as the start of the NinjaOne date range.

.PARAMETER WebhookTokenId
  webhook.site token UUID (required). Found in your webhook.site URL.

.PARAMETER WebhookApiKey
  Optional API key for webhook.site (if token requires authentication).

.PARAMETER NinjaInstance
  NinjaOne instance base URL (e.g. https://app.ninjarmm.com). Uses Get-NinjaProperty if not passed.

.PARAMETER NinjaClientId
  NinjaOne OAuth client ID. Uses Get-NinjaProperty if not passed.

.PARAMETER NinjaClientSecret
  NinjaOne OAuth client secret. Uses Get-NinjaProperty if not passed.

.PARAMETER OutputPath
  Directory for CSV and JSON exports. Default: current directory.

.PARAMETER IncludeMatched
  Include matched activities in CSV export.

.PARAMETER SortOrder
  Sort order for output by activityTime. 'ascending' (oldest first) or 'descending' (newest first). Default: ascending.

.EXAMPLE
  .\Compare-NinjaActivityDiscrepancies.ps1 -WebhookTokenId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

.EXAMPLE
  .\Compare-NinjaActivityDiscrepancies.ps1 -WebhookTokenId "xxx" -After (Get-Date).AddDays(-1) -OutputPath C:\Reports

.EXAMPLE
  For local testing, set the $Config block (WebhookTokenId, After, NinjaInstance, etc.) and run without parameters:
  .\Compare-NinjaActivityDiscrepancies.ps1
#>

[CmdletBinding()]
param(
    [datetime]$After = (Get-Date).AddDays(-7),
    [string]$WebhookTokenId,
    [string]$WebhookApiKey,
    [string]$NinjaInstance,
    [string]$NinjaClientId,
    [string]$NinjaClientSecret,
    [string]$OutputPath = (Get-Location).Path,
    [switch]$IncludeMatched,
    [ValidateSet('ascending', 'descending')]
    [string]$SortOrder = 'ascending'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

<# -----------------------------
# Config block (for easier testing)
# Set values here to avoid passing parameters every run. CLI parameters take precedence when provided.
# -----------------------------
$Config = @{
    WebhookTokenId    = ''   # e.g. 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
    WebhookApiKey     = ''
    After             = (Get-Date).AddHours(-10)
    NinjaInstance     = ''   # e.g. 'https://app.ninjarmm.com'
    NinjaClientId     = ''
    NinjaClientSecret = ''
    OutputPath        = "C:\temp\"
    IncludeMatched    = $false
    SortOrder         = 'ascending'
}
#>
if (-not $PSBoundParameters.ContainsKey('After'))            { $After = $Config.After }
$Before = Get-Date
if (-not $PSBoundParameters.ContainsKey('WebhookTokenId'))   { $WebhookTokenId = $Config.WebhookTokenId }
if (-not $PSBoundParameters.ContainsKey('WebhookApiKey'))   { $WebhookApiKey = $Config.WebhookApiKey }
if (-not $PSBoundParameters.ContainsKey('NinjaInstance'))    { $NinjaInstance = $Config.NinjaInstance }
if (-not $PSBoundParameters.ContainsKey('NinjaClientId'))   { $NinjaClientId = $Config.NinjaClientId }
if (-not $PSBoundParameters.ContainsKey('NinjaClientSecret')){ $NinjaClientSecret = $Config.NinjaClientSecret }
if (-not $PSBoundParameters.ContainsKey('OutputPath'))      { $OutputPath = $Config.OutputPath }
if (-not $PSBoundParameters.ContainsKey('IncludeMatched'))  { $IncludeMatched = $Config.IncludeMatched }
if (-not $PSBoundParameters.ContainsKey('SortOrder'))      { $SortOrder = $Config.SortOrder }

if ([string]::IsNullOrWhiteSpace($WebhookTokenId)) {
    throw "WebhookTokenId is required. Pass -WebhookTokenId or set Config.WebhookTokenId in the script."
}

# -----------------------------
# Constants
# -----------------------------
$ApiPaths = @{
    OAuthToken = '/ws/oauth/token'   # NinjaOne requires /ws/ prefix; /oauth/token returns 405
    Activities = '/api/v2/activities'
}

$ResponseProps = @{
    Activities = 'activities'
}

$WebhookSiteBase = 'https://webhook.site'

# -----------------------------
# Config: NinjaOne credentials from params or Get-NinjaProperty
# -----------------------------
function Get-NinjaConfig {
    $instance = $NinjaInstance
    $clientId = $NinjaClientId
    $clientSecret = $NinjaClientSecret

    if (Get-Command Ninja-Property-Get -ErrorAction SilentlyContinue) {
        if ([string]::IsNullOrWhiteSpace($instance)) { $instance = Ninja-Property-Get ninjaoneInstance }
        if ([string]::IsNullOrWhiteSpace($clientId)) { $clientId = Ninja-Property-Get ninjaoneClientId }
        if ([string]::IsNullOrWhiteSpace($clientSecret)) { $clientSecret = Ninja-Property-Get ninjaoneClientSecret }
    }

    if ([string]::IsNullOrWhiteSpace($instance) -or [string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($clientSecret)) {
        throw "NinjaOne credentials required. Pass -NinjaInstance, -NinjaClientId, -NinjaClientSecret or run in NinjaOne context with Get-NinjaProperty."
    }

    $instance = $instance.Trim()
    if (-not $instance -match '^https?://') {
        $instance = "https://$instance"
    }

    return @{
        Instance     = $instance
        ClientId     = $clientId.Trim()
        ClientSecret = $clientSecret.Trim()
        Scope        = "monitoring management"
    }
}

# -----------------------------
# API helpers
# -----------------------------
function New-QueryString {
    param([hashtable]$Params)

    $pairs = foreach ($k in $Params.Keys) {
        $v = $Params[$k]
        if ($null -eq $v) { continue }
        if ($v -is [string] -and [string]::IsNullOrWhiteSpace($v)) { continue }

        if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) {
            foreach ($item in $v) {
                if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace([string]$item)) {
                    "{0}={1}" -f [uri]::EscapeDataString($k), [uri]::EscapeDataString([string]$item)
                }
            }
        } else {
            "{0}={1}" -f [uri]::EscapeDataString($k), [uri]::EscapeDataString([string]$v)
        }
    }

    if (-not $pairs) { return "" }
    return "?" + ($pairs -join "&")
}

function Get-NinjaOneAccessToken {
    param(
        [Parameter(Mandatory)][string]$Instance,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$ClientSecret,
        [Parameter()][string]$Scope = ""
    )

    $tokenUri = "$Instance$($ApiPaths.OAuthToken)"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $ClientId
        client_secret = $ClientSecret
    }
    if (-not [string]::IsNullOrWhiteSpace($Scope)) { $body.scope = $Scope }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded"
        if (-not $resp.access_token) { throw "Token response did not include access_token." }
        return $resp.access_token
    } catch {
        throw "Failed to obtain access token from $tokenUri. $($_.Exception.Message)"
    }
}

function Invoke-NinjaOneGet {
    param(
        [Parameter(Mandatory)][string]$Instance,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter()][hashtable]$QueryParams
    )

    $qs = New-QueryString -Params $(if ($null -ne $QueryParams) { $QueryParams } else { @{} })
    $uri = "$Instance$Path$qs"
    $headers = @{
        Authorization = "Bearer $AccessToken"
        Accept        = "application/json"
    }
    return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
}

# -----------------------------
# NinjaOne activity fetching (paginated)
# -----------------------------
function Get-NinjaOneActivitiesInRange {
    param(
        [Parameter(Mandatory)][string]$Instance,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][datetime]$After,
        [Parameter(Mandatory)][datetime]$Before,
        [ValidateRange(1, 5000)][int]$pageSize = 500
    )

    $afterUtc = if ($After.Kind -eq [DateTimeKind]::Utc) { $After } else { [DateTimeOffset]::new($After).UtcDateTime }
    $beforeUtc = if ($Before.Kind -eq [DateTimeKind]::Utc) { $Before } else { [DateTimeOffset]::new($Before).UtcDateTime }
    if ($beforeUtc -lt $afterUtc) { throw "Before must be >= After." }

    $afterVal = [long][DateTimeOffset]::new($afterUtc).ToUnixTimeSeconds()
    $beforeVal = [long][DateTimeOffset]::new($beforeUtc).ToUnixTimeSeconds()

    $all = New-Object System.Collections.Generic.List[object]
    $olderThan = $null

    while ($true) {
        $qp = @{
            after    = $afterVal
            before   = $beforeVal
            pageSize          = $pageSize
        }
        if ($null -ne $olderThan) { $qp.olderThan = $olderThan }

        $resp = Invoke-NinjaOneGet -Instance $Instance -Path $ApiPaths.Activities -AccessToken $AccessToken -QueryParams $qp

        $raw = if ($null -ne $resp -and $resp.PSObject.Properties.Name -contains $ResponseProps.Activities) {
            $resp.$($ResponseProps.Activities)
        } else {
            $resp
        }
        # Force to array so .Count and indexer always exist (API may return types without .Count)
        [array]$items = @($raw)

        if ($items.Count -eq 0) { break }

        foreach ($item in $items) { $all.Add($item) }

        $last = $items[-1]

        if ($null -eq $last.id) { break }

        $olderThan = $last.id
        Start-Sleep -Milliseconds 200
    }

    return $all
}

# -----------------------------
# NinjaOne activity fetching by newerThan (activity id); paginates with olderThan
# -----------------------------
function Get-NinjaOneActivitiesNewerThan {
    param(
        [Parameter(Mandatory)][string]$Instance,
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter(Mandatory)][int]$NewerThan,
        [ValidateRange(1, 5000)][int]$pageSize = 500
    )

    $all = New-Object System.Collections.Generic.List[object]
    $olderThan = $null

    while ($true) {
        # API allows only one of before/after-style params (newerThan and olderThan are mutually exclusive).
        $qp = @{
            pageSize = $pageSize
        }
        if ($null -ne $olderThan) {
            $qp.olderThan = $olderThan
        } else {
            $qp.newerThan = $NewerThan
        }

        $resp = Invoke-NinjaOneGet -Instance $Instance -Path $ApiPaths.Activities -AccessToken $AccessToken -QueryParams $qp

        $raw = if ($null -ne $resp -and $resp.PSObject.Properties.Name -contains $ResponseProps.Activities) {
            $resp.$($ResponseProps.Activities)
        } else {
            $resp
        }
        [array]$items = @($raw)

        if ($items.Count -eq 0) { break }

        foreach ($item in $items) { $all.Add($item) }

        $last = $items[-1]
        if ($null -eq $last.id) { break }

        $olderThan = $last.id
        Start-Sleep -Milliseconds 200
    }

    return $all
}

# -----------------------------
# Webhook.site request fetching and parsing
# -----------------------------
function Get-WebhookSiteRequests {
    param(
        [Parameter(Mandatory)][string]$TokenId,
        [string]$ApiKey,
        [Parameter(Mandatory)][datetime]$DateFrom,
        [Parameter(Mandatory)][datetime]$DateTo,
        [ValidateRange(1, 100)][int]$PerPage = 100
    )

    $dateFromStr = $DateFrom.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
    $dateToStr = $DateTo.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')

    $headers = @{ Accept = "application/json" }
    if (-not [string]::IsNullOrWhiteSpace($ApiKey)) {
        $headers['api-key'] = $ApiKey
    }

    $allRequests = New-Object System.Collections.Generic.List[object]
    $page = 1
    $isLastPage = $false

    while (-not $isLastPage) {
        $uri = "$WebhookSiteBase/token/$TokenId/requests"
        $query = "?date_from=$([uri]::EscapeDataString($dateFromStr))&date_to=$([uri]::EscapeDataString($dateToStr))&per_page=$PerPage&page=$page&sorting=oldest"

        try {
            $resp = Invoke-RestMethod -Method Get -Uri "$uri$query" -Headers $headers
        } catch {
            throw "Failed to fetch webhook.site requests: $($_.Exception.Message)"
        }

        $data = if ($resp.data) { @($resp.data) } else { @() }
        foreach ($r in $data) { $allRequests.Add($r) }

        $isLastPage = $true
        if ($resp.PSObject.Properties['is_last_page']) { $isLastPage = $resp.is_last_page }
        if ($data.Count -lt $PerPage) { $isLastPage = $true }

        if (-not $isLastPage) {
            $page++
            Start-Sleep -Milliseconds 550
        }
    }

    return $allRequests
}

function Get-ActivityFromWebhookContent {
    param([string]$Content)

    if ([string]::IsNullOrWhiteSpace($Content)) { return @() }

    try {
        $obj = $Content | ConvertFrom-Json
    } catch {
        Write-Verbose "Skipping non-JSON webhook content: $($Content.Substring(0, [Math]::Min(100, $Content.Length)))..."
        return @()
    }

    $activities = @()

    if ($null -ne $obj.PSObject.Properties['id'] -and $null -ne $obj.id) {
        $activities += $obj
    } elseif ($obj.PSObject.Properties['activity']) {
        $activities += $obj.activity
    } elseif ($obj.PSObject.Properties['activities']) {
        $arr = $obj.activities
        if ($arr -is [array]) {
            $activities += $arr
        } else {
            $activities += @($arr)
        }
    }

    return $activities
}

function Get-WebhookActivitiesMap {
    param(
        [Parameter(Mandatory)]$Requests,
        [switch]$DeduplicateById
    )

    $map = @{}
    $parseErrors = 0

    foreach ($req in $Requests) {
        $content = if ($req.content) { $req.content } else { "" }
        $activities = @(Get-ActivityFromWebhookContent -Content $content)

        foreach ($act in $activities) {
            $id = $null
            if ($act.PSObject.Properties['id']) { $id = $act.id }
            if ($null -eq $id) { continue }

            $key = [string]$id
            # When DeduplicateById: last occurrence wins (always overwrite). Otherwise: first occurrence wins.
            if ($DeduplicateById -or -not $map.ContainsKey($key)) {
                $map[$key] = [PSCustomObject]@{
                    Activity   = $act
                    WebhookUuid = $req.uuid
                    CreatedAt  = $req.created_at
                }
            }
        }

        if ($activities.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($content)) {
            $parseErrors++
        }
    }

    if ($parseErrors -gt 0) {
        Write-Warning "Could not parse $parseErrors webhook request(s) as NinjaOne activity JSON. Skipped."
    }

    return $map
}

# -----------------------------
# Convert Unix timestamp (seconds, optional fractional) to human-readable string (UTC)
# -----------------------------
function ConvertFrom-UnixActivityTime {
    param([object]$UnixTime)
    if ($null -eq $UnixTime) { return $null }
    $sec = $UnixTime -as [double]
    if ($null -eq $sec -or ([double]::IsNaN($sec))) { return $UnixTime.ToString() }
    $secWhole = [long][Math]::Floor($sec)
    try {
        $dt = [DateTimeOffset]::FromUnixTimeSeconds($secWhole)
        return $dt.UtcDateTime.ToString('yyyy-MM-dd HH:mm:ss') + ' UTC'
    } catch {
        return $UnixTime.ToString()
    }
}

# -----------------------------
# Flatten activity for CSV/display
# -----------------------------
function Get-ActivityFlatRow {
    param($Activity, [string]$Source = "NinjaOne")

    $id = $null
    $activityTime = $null
    $deviceId = $null
    $type = $null
    $status = $null
    $message = $null
    $subject = $null

    if ($Activity.PSObject.Properties['id']) { $id = $Activity.id }
    if ($Activity.PSObject.Properties['activityTime']) { $activityTime = $Activity.activityTime }
    if ($Activity.PSObject.Properties['deviceId']) { $deviceId = $Activity.deviceId }
    if ($Activity.PSObject.Properties['activityType']) { $type = $Activity.activityType }
    if ($Activity.PSObject.Properties['statusCode']) { $status = $Activity.statusCode }
    if ($Activity.PSObject.Properties['message']) { $message = $Activity.message }
    if ($Activity.PSObject.Properties['subject']) { $subject = $Activity.subject }

    $msgTruncated = ""
    if ($message -is [string] -and $message.Length -gt 0) {
        $replaced = $message -replace "`r`n", " "
        $msgTruncated = $replaced.Substring(0, [Math]::Min(200, $replaced.Length))
    }

    return [PSCustomObject]@{
        Id           = $id
        ActivityTime = ConvertFrom-UnixActivityTime $activityTime
        DeviceId     = $deviceId
        Type         = $type
        Status       = $status
        Message      = $msgTruncated
        Subject      = $subject
        Source       = $Source
    }
}

# -----------------------------
# Activity type for grouping (null/empty -> sentinel)
# -----------------------------
function Get-ActivityType {
    param($Activity)
    if ($null -eq $Activity) { return '(unknown)' }
    if ($Activity.PSObject.Properties['activityType'] -and $null -ne $Activity.activityType -and -not [string]::IsNullOrWhiteSpace([string]$Activity.activityType)) {
        return [string]$Activity.activityType.Trim()
    }
    return '(unknown)'
}

# -----------------------------
# Main execution
# -----------------------------
$ninjaConfig = Get-NinjaConfig
$OutputPath = $OutputPath.Trim()
if (-not (Test-Path -Path $OutputPath -PathType Container)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# 1) Fetch webhook.site first (date window defines what we consider from webhook)
Write-Host "Fetching webhook.site requests ($After to $Before)..."
$webhookRequests = Get-WebhookSiteRequests -TokenId $WebhookTokenId -ApiKey $WebhookApiKey -DateFrom $After -DateTo $Before -PerPage 100
$webhookMap = Get-WebhookActivitiesMap -Requests $webhookRequests -DeduplicateById
Write-Host "Parsed $($webhookMap.Count) NinjaOne activity payload(s) from $($webhookRequests.Count) webhook request(s)."

# 2) Derive date of oldest activity from webhook dataset as starting point for NinjaOne (after parameter)
$ninjaAfter = $After
if ($webhookMap.Count -gt 0) {
    $activityTimes = @($webhookMap.Values | ForEach-Object {
        $t = $_.Activity.PSObject.Properties['activityTime']
        if ($null -ne $t -and $null -ne $t.Value) { $t.Value -as [double] } else { $null }
    } | Where-Object { $null -ne $_ -and -not [double]::IsNaN($_) })
    if ($activityTimes.Count -gt 0) {
        $oldestEpochSeconds = [long][Math]::Floor(($activityTimes | Measure-Object -Minimum).Minimum)
        try {
            $ninjaAfter = [DateTimeOffset]::FromUnixTimeSeconds($oldestEpochSeconds).UtcDateTime
            Write-Host "Oldest webhook activity date: $($ninjaAfter.ToString('yyyy-MM-dd HH:mm:ss')) UTC"
        } catch {
            Write-Warning "Could not convert oldest webhook activityTime to date; using -After for NinjaOne."
        }
    } else {
        Write-Warning "No webhook activities with valid activityTime in window; using -After for NinjaOne."
    }
} else {
    Write-Warning "No webhook activities in window; using -After for NinjaOne ($After to $Before)."
}

$BeforeUtc = [DateTimeOffset]::new($Before).UtcDateTime
if ($ninjaAfter -gt $BeforeUtc) {
    $ninjaAfter = [DateTimeOffset]::new($After).UtcDateTime
    Write-Host "Oldest webhook activity is after end of window (timezone); using -After as NinjaOne lower bound."
}

# 3) Fetch NinjaOne: after (oldest webhook activity date) to Before
$token = Get-NinjaOneAccessToken -Instance $ninjaConfig.Instance -ClientId $ninjaConfig.ClientId -ClientSecret $ninjaConfig.ClientSecret -Scope $ninjaConfig.Scope
Write-Host "Fetching NinjaOne activities (after $($ninjaAfter.ToString('yyyy-MM-dd HH:mm:ss')) UTC to $($BeforeUtc.ToString('yyyy-MM-dd HH:mm:ss')) UTC)..."
$ninjaActivities = @(Get-NinjaOneActivitiesInRange -Instance $ninjaConfig.Instance -AccessToken $token -After $ninjaAfter -Before $BeforeUtc -pageSize 500)
Write-Host "Fetched $($ninjaActivities.Count) activities from NinjaOne."

$ninjaMap = @{}
foreach ($a in $ninjaActivities) {
    if ($null -ne $a.id) { $ninjaMap[[string]$a.id] = $a }
}

# -----------------------------
# Comparison
# -----------------------------
$ninjaIds = [System.Collections.Generic.HashSet[string]]::new([string[]]$ninjaMap.Keys)
$webhookIds = [System.Collections.Generic.HashSet[string]]::new([string[]]$webhookMap.Keys)

$inNinjaOnly = @()
foreach ($id in $ninjaIds) {
    if (-not $webhookIds.Contains($id)) {
        $inNinjaOnly += $ninjaMap[$id]
    }
}

$inWebhookOnly = @()
foreach ($id in $webhookIds) {
    if (-not $ninjaIds.Contains($id)) {
        $inWebhookOnly += $webhookMap[$id].Activity
    }
}

$matchedCount = @($ninjaIds | Where-Object { $webhookIds.Contains($_) }).Count

# Sort by activityTime (ascending = oldest first; null treated as 0)
$sortByTime = {
    $t = $_.activityTime
    if ($null -eq $t) { 0 } else { $t -as [double] }
}
$sortDescending = ($SortOrder.Trim().ToLowerInvariant() -eq 'descending')
$inNinjaOnly   = @($inNinjaOnly   | Sort-Object -Property @{ Expression = $sortByTime } -Descending:$sortDescending)
$inWebhookOnly = @($inWebhookOnly | Sort-Object -Property @{ Expression = $sortByTime } -Descending:$sortDescending)

# Full datasets for "all" exports (same sort)
$allWebhookActivities = @($webhookMap.Values | ForEach-Object { $_.Activity } | Sort-Object -Property @{ Expression = $sortByTime } -Descending:$sortDescending)
$allNinjaActivities   = @($ninjaActivities | Sort-Object -Property @{ Expression = $sortByTime } -Descending:$sortDescending)

# -----------------------------
# Type breakdowns (for summaries and discrepancy-by-type)
# -----------------------------
$inNinjaOnlyByType = @($inNinjaOnly | ForEach-Object { Get-ActivityType -Activity $_ } | Group-Object | ForEach-Object { [PSCustomObject]@{ Type = $_.Name; Count = $_.Count } } | Sort-Object Count -Descending)
$allWebhookByType  = @($allWebhookActivities | ForEach-Object { Get-ActivityType -Activity $_ } | Group-Object | ForEach-Object { [PSCustomObject]@{ Type = $_.Name; Count = $_.Count } } | Sort-Object Count -Descending)

$allTypes = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($o in $inNinjaOnlyByType) { $null = $allTypes.Add($o.Type) }
foreach ($o in $allWebhookByType)  { $null = $allTypes.Add($o.Type) }
foreach ($a in $ninjaActivities)  { $null = $allTypes.Add((Get-ActivityType -Activity $a)) }
foreach ($a in $allWebhookActivities) { $null = $allTypes.Add((Get-ActivityType -Activity $a)) }

$typeBreakdown = [System.Collections.Generic.List[object]]::new()
foreach ($typeName in $allTypes) {
    $allNinja = @($ninjaActivities | Where-Object { (Get-ActivityType -Activity $_) -eq $typeName }).Count
    $allWebhook = @($allWebhookActivities | Where-Object { (Get-ActivityType -Activity $_) -eq $typeName }).Count
    $inNinjaOnlyCnt = @($inNinjaOnly | Where-Object { (Get-ActivityType -Activity $_) -eq $typeName }).Count
    $inWebhookOnlyCnt = @($inWebhookOnly | Where-Object { (Get-ActivityType -Activity $_) -eq $typeName }).Count
    $matchedCnt = 0
    foreach ($id in $ninjaIds) {
        if ($webhookIds.Contains($id)) {
            $act = $ninjaMap[$id]
            if ((Get-ActivityType -Activity $act) -eq $typeName) { $matchedCnt++ }
        }
    }
    $typeBreakdown.Add([PSCustomObject]@{
        Type        = $typeName
        AllNinja    = $allNinja
        AllWebhook  = $allWebhook
        Matched     = $matchedCnt
        InNinjaOnly = $inNinjaOnlyCnt
        InWebhookOnly = $inWebhookOnlyCnt
    })
}
$typeBreakdown = @($typeBreakdown | Sort-Object InNinjaOnly -Descending)

# -----------------------------
# Console output
# -----------------------------
$ninjaWindowDescription = "after $($ninjaAfter.ToString('yyyy-MM-dd HH:mm:ss')) UTC to $($Before.ToString('yyyy-MM-dd HH:mm:ss')) UTC"
Write-Host ""
Write-Host "=== Summary ==="
Write-Host "NinjaOne window:   $ninjaWindowDescription"
Write-Host "NinjaOne total:    $($ninjaActivities.Count)"
Write-Host "Webhook total:     $($webhookMap.Count)"
Write-Host "Matched:           $matchedCount"
Write-Host "In NinjaOnly:      $($inNinjaOnly.Count)"
Write-Host "In WebhookOnly:    $($inWebhookOnly.Count)"

if ($inNinjaOnly.Count -gt 0) {
    Write-Host ""
    Write-Host "=== Sample: In NinjaOne Only (first 10) ==="
    $inNinjaOnly | Select-Object -First 10 | ForEach-Object {
        $row = Get-ActivityFlatRow -Activity $_ -Source "NinjaOnly"
        Write-Host "  id=$($row.Id) activityTime=$($row.ActivityTime) deviceId=$($row.DeviceId) type=$($row.Type) status=$($row.Status)"
    }
}

if ($inWebhookOnly.Count -gt 0) {
    Write-Host ""
    Write-Host "=== Sample: In Webhook Only (first 10) ==="
    $inWebhookOnly | Select-Object -First 10 | ForEach-Object {
        $row = Get-ActivityFlatRow -Activity $_ -Source "WebhookOnly"
        Write-Host "  id=$($row.Id) activityTime=$($row.ActivityTime) deviceId=$($row.DeviceId) type=$($row.Type) status=$($row.Status)"
    }
}

# -----------------------------
# CSV export
# -----------------------------
$inNinjaOnlyCsv = Join-Path $OutputPath "InNinjaOnly.csv"
$inNinjaOnlyFlat = @($inNinjaOnly | ForEach-Object { Get-ActivityFlatRow -Activity $_ -Source "NinjaOnly" })
if ($inNinjaOnlyFlat.Count -gt 0) {
    $inNinjaOnlyFlat | Export-Csv -Path $inNinjaOnlyCsv -NoTypeInformation
    Write-Host ""
    Write-Host "Exported InNinjaOnly to $inNinjaOnlyCsv"
} else {
    Write-Host ""
    Write-Host "No InNinjaOnly activities to export."
}

$inWebhookOnlyCsv = Join-Path $OutputPath "InWebhookOnly.csv"
$inWebhookOnlyFlat = @($inWebhookOnly | ForEach-Object { Get-ActivityFlatRow -Activity $_ -Source "WebhookOnly" })
if ($inWebhookOnlyFlat.Count -gt 0) {
    $inWebhookOnlyFlat | Export-Csv -Path $inWebhookOnlyCsv -NoTypeInformation
    Write-Host "Exported InWebhookOnly to $inWebhookOnlyCsv"
} else {
    Write-Host "No InWebhookOnly activities to export."
}

# All webhook activities (full set parsed from webhook)
$allWebhookCsv = Join-Path $OutputPath "AllWebhookActivities.csv"
$allWebhookFlat = @($allWebhookActivities | ForEach-Object { Get-ActivityFlatRow -Activity $_ -Source "Webhook" })
if ($allWebhookFlat.Count -gt 0) {
    $allWebhookFlat | Export-Csv -Path $allWebhookCsv -NoTypeInformation
    Write-Host "Exported AllWebhookActivities ($($allWebhookFlat.Count) rows) to $allWebhookCsv"
} else {
    Write-Host "No webhook activities to export for AllWebhookActivities."
}

# All NinjaOne activities (full set from API)
$allNinjaCsv = Join-Path $OutputPath "AllNinjaOneActivities.csv"
$allNinjaFlat = @($allNinjaActivities | ForEach-Object { Get-ActivityFlatRow -Activity $_ -Source "NinjaOne" })
if ($allNinjaFlat.Count -gt 0) {
    $allNinjaFlat | Export-Csv -Path $allNinjaCsv -NoTypeInformation
    Write-Host "Exported AllNinjaOneActivities ($($allNinjaFlat.Count) rows) to $allNinjaCsv"
} else {
    Write-Host "No NinjaOne activities to export for AllNinjaOneActivities."
}

if ($IncludeMatched -and $matchedCount -gt 0) {
    $matchedCsv = Join-Path $OutputPath "Matched.csv"
    $matchedActivities = $ninjaIds | Where-Object { $webhookIds.Contains($_) } | ForEach-Object { $ninjaMap[$_] }
    $matchedActivities = @($matchedActivities | Sort-Object -Property @{ Expression = $sortByTime } -Descending:$sortDescending)
    $matchedFlat = $matchedActivities | ForEach-Object { Get-ActivityFlatRow -Activity $_ -Source "Matched" }
    $matchedFlat | Export-Csv -Path $matchedCsv -NoTypeInformation
    Write-Host "Exported Matched to $matchedCsv"
}

# -----------------------------
# Type breakdown summaries (console + CSV)
# -----------------------------
Write-Host ""
Write-Host "=== Breakdown: In Ninja Only by type ==="
if ($inNinjaOnlyByType.Count -gt 0) {
    $inNinjaOnlyByType | Format-Table -AutoSize Type, Count
} else {
    Write-Host "  (none)"
}
Write-Host ""
Write-Host "=== Breakdown: All Webhook activities by type ==="
if ($allWebhookByType.Count -gt 0) {
    $allWebhookByType | Format-Table -AutoSize Type, Count
} else {
    Write-Host "  (none)"
}
Write-Host ""
Write-Host "=== Discrepancy by type ==="
if ($typeBreakdown.Count -gt 0) {
    $typeBreakdown | Format-Table -AutoSize Type, AllNinja, AllWebhook, Matched, InNinjaOnly, InWebhookOnly
} else {
    Write-Host "  (none)"
}

$typeBreakdownCsv = Join-Path $OutputPath "TypeBreakdown.csv"
$typeBreakdown | Export-Csv -Path $typeBreakdownCsv -NoTypeInformation
Write-Host ""
Write-Host "Exported TypeBreakdown ($($typeBreakdown.Count) rows) to $typeBreakdownCsv"

# -----------------------------
# JSON export
# -----------------------------
$report = [PSCustomObject]@{
    GeneratedAt     = (Get-Date).ToString('o')
    TimeWindow      = @{
        After  = $After.ToString('o')
        Before = $Before.ToString('o')
    }
    NinjaOneWindow  = @{
        Description = $ninjaWindowDescription
        After       = $ninjaAfter.ToString('o')
        Before      = $Before.ToString('o')
    }
    Summary         = @{
        NinjaOneTotal  = $ninjaActivities.Count
        WebhookTotal   = $webhookMap.Count
        Matched        = $matchedCount
        InNinjaOnly    = $inNinjaOnly.Count
        InWebhookOnly  = $inWebhookOnly.Count
    }
    InNinjaOnly           = $inNinjaOnly | ForEach-Object { $_ }
    InWebhookOnly         = $inWebhookOnly | ForEach-Object { $_ }
    AllWebhookActivities  = $allWebhookActivities | ForEach-Object { $_ }
    AllNinjaOneActivities = $allNinjaActivities | ForEach-Object { $_ }
    TypeBreakdown         = $typeBreakdown
}

$jsonPath = Join-Path $OutputPath "discrepancy-report.json"
$report | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding UTF8
Write-Host "Exported discrepancy-report.json to $jsonPath"
Write-Host ""

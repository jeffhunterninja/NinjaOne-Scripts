<#
.SYNOPSIS
  Retrieve NinjaOne Activities from the NinjaOne Public API over a specified time window.

.REQUIREMENTS
  - PowerShell 5.1+ (PS7 recommended)
  - NinjaOne OAuth app configured for Client Credentials flow with appropriate scopes.

.AUTH
  Uses client_credentials to POST {Instance}/ws/oauth/token and then calls GET {Instance}/api/v2/activities

.PARAMETERS (high level)
  -After / -Before : time window (DateTime)
  -Filters: deviceId, class, type, activityType, status, user, deviceFilter(df), sourceConfigUid, expandActivities
#>

# -----------------------------
# CONFIG (edit these)
# -----------------------------
$Config = [ordered]@{
    # Instance base URL:
    #   NA: https://app.ninjarmm.com
    #   US2: https://us2.ninjarmm.com
    #   EU: https://eu.ninjarmm.com
    #   OC: https://oc.ninjarmm.com
    Instance     = (Get-NinjaProperty "ninjaoneInstance") -replace '^(?!https?://)', 'https://'

    ClientId     = Get-NinjaProperty "ninjaoneClientId"
    ClientSecret = Get-NinjaProperty "ninjaoneClientSecret"

    # Scopes depend on what you're querying; "monitoring" is commonly needed for read access.
    # You can space-separate them for client_credentials.
    Scope        = "monitoring management"
}

# -----------------------------
# Constants (field names, API paths – edit to match your NinjaOne custom field API names)
# -----------------------------
$CustomFieldNames = [ordered]@{
    TotalConditionsTriggered = 'totalConditionsTriggered'
    DeviceAlertRank         = 'deviceAlertRank'
    TotalAlertingDevices    = 'totalAlertingDevices'
    AlertHeatMap            = 'alertHeatMap'
    MostFrequentAlerts      = 'mostFrequentAlerts'
}

$ActivityProps = [ordered]@{
    ActivityType    = 'activityType'
    StatusCode      = 'statusCode'
    ActivityTime    = 'activityTime'
    DeviceId        = 'deviceId'
    Message         = 'message'
    Subject         = 'subject'
    SourceConfigUid = 'sourceConfigUid'
}

$ActivityFilter = @{
    Type   = 'CONDITION'
    Status = 'TRIGGERED'
}

$ApiPaths = @{
    OAuthToken   = '/ws/oauth/token'   # NinjaOne requires /ws/ prefix; /oauth/token returns 405
    Activities   = '/api/v2/activities'
    CustomFields = '/api/v2/device/{0}/custom-fields'
}

$ResponseProps = @{
    Activities = 'activities'
}

$HtmlTableExcludeProps   = @('RowColour')
$HtmlTableRowColourProp  = 'RowColour'
$NinjaHtmlCharLimit     = 200000

Set-StrictMode -Version Latest

# ===== Defaults =====
$OverwriteEmptyValues = $false

# -----------------------------
# API helpers
# -----------------------------
function New-QueryString {
    param([hashtable]$Params)

    $pairs = foreach ($k in $Params.Keys) {
        $v = $Params[$k]
        if ($null -eq $v) { continue }
        if ($v -is [string] -and [string]::IsNullOrWhiteSpace($v)) { continue }

        # Arrays -> repeat parameter
        if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) {
            foreach ($item in $v) {
                if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace([string]$item)) {
                    "{0}={1}" -f [uri]::EscapeDataString($k), [uri]::EscapeDataString([string]$item)
                }
            }
        }
        else {
            "{0}={1}" -f [uri]::EscapeDataString($k), [uri]::EscapeDataString([string]$v)
        }
    }

    if (-not $pairs) { return "" }
    return "?" + ($pairs -join "&")
}

function ConvertTo-UnixEpochSeconds {
    param([Parameter(Mandatory)][datetime]$DateTime)

    # Unix epoch seconds (UTC)
    return [int64]([DateTimeOffset]$DateTime.ToUniversalTime()).ToUnixTimeSeconds()
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
    }
    catch {
        throw "Failed to obtain access token from $tokenUri. $($_.Exception.Message)"
    }
}

function Invoke-NinjaOneGet {
    param(
        [Parameter(Mandatory)][string]$Instance,
        [Parameter(Mandatory)][string]$Path,           # e.g. /api/v2/activities
        [Parameter(Mandatory)][string]$AccessToken,
        [Parameter()][hashtable]$QueryParams
    )

    $qs  = New-QueryString -Params $(if ($null -ne $QueryParams) { $QueryParams } else { @{} })
    $uri = "$Instance$Path$qs"

    $headers = @{
        Authorization = "Bearer $AccessToken"
        Accept        = "application/json"
    }
    try {
        return Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
    }
    catch {
        throw "GET failed: $uri. $($_.Exception.Message)"
    }
}

function Invoke-NinjaAPIRequest {
  param(
    [Parameter(Mandatory=$true)][string]$Uri,
    [ValidateSet('GET','POST','PATCH','PUT','DELETE')][string]$Method = 'GET',
    [Parameter(Mandatory=$true)][hashtable]$Headers,
    [string]$Body = $null
  )

  $maxRetries = 3
  for ($i = 1; $i -le $maxRetries; $i++) {
    try {
      return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -Body $Body -ContentType "application/json"
    } catch {
      Write-Warning "API request failed (attempt $i/$maxRetries): $Uri :: $($_.Exception.Message)"
      Start-Sleep -Seconds 2
    }
  }
  return $null
}

# -----------------------------
# Activity fetching (auto-pages with olderThan)
# -----------------------------
function Get-NinjaOneActivitiesInRange {
    param(
        [Parameter(Mandatory)][string]$Instance,
        [Parameter(Mandatory)][string]$AccessToken,

        [Parameter(Mandatory)][datetime]$After,
        [Parameter(Mandatory)][datetime]$Before,

        # This maps to the query param "type=" (which corresponds to activityType enums in your tenant)
        [string]$type,

        [string]$class,
        [string[]]$activityType,
        [string]$status,
        [string]$user,
        [string]$seriesUid,
        [string]$deviceFilter,        # df
        [string]$sourceConfigUid,
        [switch]$expandActivities,

        [ValidateRange(1, 5000)]
        [int]$pageSize = 500
    )

    if ($Before -lt $After) { throw "Before must be >= After." }

    # You’re using epoch seconds already:
    $afterVal  = [int64]([DateTimeOffset]$After.ToUniversalTime()).ToUnixTimeSeconds()
    $beforeVal = [int64]([DateTimeOffset]$Before.ToUniversalTime()).ToUnixTimeSeconds()

    $all = New-Object System.Collections.Generic.List[object]
    $olderThan = $null

    while ($true) {
        $qp = @{
            after    = $afterVal
            before   = $beforeVal
            pageSize = $pageSize
            type     = $type
            class    = $class
            activityType = $activityType
            status   = $status
            user     = $user
            seriesUid= $seriesUid
            df       = $deviceFilter
            sourceConfigUid = $sourceConfigUid
        }

        if ($expandActivities.IsPresent) { $qp.expand = "activities" }
        if ($null -ne $olderThan) { $qp.olderThan = $olderThan }

        $resp = Invoke-NinjaOneGet -Instance $Instance -Path $ApiPaths.Activities -AccessToken $AccessToken -QueryParams $qp

        # ✅ Normalize: the real records are usually in resp.activities
        $items =
            if ($null -ne $resp -and $resp.PSObject.Properties.Name -contains $ResponseProps.Activities) { @($resp.$($ResponseProps.Activities)) }
            else { @($resp) }

        if ($items.Count -eq 0) { break }

        foreach ($item in $items) { $all.Add($item) }

        $last = $items[-1]
        if ($null -eq $last.id) { break }

        $olderThan = $last.id

        if ($items.Count -lt $pageSize) { break }
    }

    return $all  # ✅ returns a flat array of activity objects (not a wrapper)
}

function Get-NinjaOneActivitiesInRangeMultiType {
    param(
        [Parameter(Mandatory)][string]$Instance,
        [Parameter(Mandatory)][string]$AccessToken,

        [Parameter(Mandatory)][datetime]$After,
        [Parameter(Mandatory)][datetime]$Before,

        [Parameter(Mandatory)][string[]]$Types,

        [string]$class,
        [string[]]$activityType,
        [string]$status,
        [string]$user,
        [string]$seriesUid,
        [string]$deviceFilter,
        [string]$sourceConfigUid,
        [switch]$expandActivities,

        [ValidateRange(1, 5000)]
        [int]$pageSize = 500,

        [int]$SleepMsBetweenCalls = 0
    )

    $all = New-Object System.Collections.Generic.List[object]

    foreach ($t in $Types) {
        if ($SleepMsBetweenCalls -gt 0) { Start-Sleep -Milliseconds $SleepMsBetweenCalls }

        $batch = Get-NinjaOneActivitiesInRange `
            -Instance $Instance `
            -AccessToken $AccessToken `
            -After $After `
            -Before $Before `
            -pageSize $pageSize `
            -type $t `
            -class $class `
            -activityType $activityType `
            -status $status `
            -user $user `
            -seriesUid $seriesUid `
            -deviceFilter $deviceFilter `
            -sourceConfigUid $sourceConfigUid `
            -expandActivities:$expandActivities

        foreach ($item in $batch) { $all.Add($item) }
    }

    # De-dupe by id
    $deduped =
        $all |
        Group-Object -Property id |
        ForEach-Object { $_.Group | Select-Object -First 1 }

    return $deduped
}

# -----------------------------
# Main execution
# -----------------------------
if (-not $Config.Instance -or -not $Config.ClientId -or -not $Config.ClientSecret) {
    throw "Config requires Instance, ClientId, and ClientSecret."
}

$token   = Get-NinjaOneAccessToken -Instance $Config.Instance -ClientId $Config.ClientId -ClientSecret $Config.ClientSecret -Scope $Config.Scope
$headers = @{ Authorization = "Bearer $token"; Accept = "application/json" }
$After   = (Get-Date).AddDays(-30)
$Before  = Get-Date

$typesWanted = @($ActivityFilter.Type)

$activities = Get-NinjaOneActivitiesInRangeMultiType `
  -Instance $Config.Instance `
  -AccessToken $token `
  -After $After `
  -Before $Before `
  -pageSize 1000 `
  -Types $typesWanted `
  -SleepMsBetweenCalls 150

$activities.Count

<#
Device Alert-Fatigue Report (from $activities)

Creates per-device reporting:
1) Total number of CONDITION TRIGGERED events in the time period
2) Top 10 most common conditions (trigger count) formatted as a table (text + HTML)
3) Device rank by trigger volume (among devices that triggered at least one condition)
4) Heatmap: counts by Day-of-Month (rows 1..31) x Month (columns yyyy-MM) for CONDITION TRIGGERED events
   - Outputs both a data matrix and an HTML heatmap table (inline styles)

INPUT:
  $activities : collection of activity objects (your dataset)

OUTPUT:
  $DeviceReports : array of per-device report objects
  $GlobalSummary : overall summary incl. per-device ranks

Notes:
- Only counts CONDITION + statusCode TRIGGERED as “conditions triggered”
- Uses UTC timestamps derived from activityTime (epoch seconds)
- conditionKey = sourceConfigUid|subject (subject may be blank)
#>

# -----------------------------
# Report helpers
# -----------------------------
function Convert-FromUnixSecondsUtc {
  param([Parameter(Mandatory)][double]$UnixSeconds)
  (Get-Date -Date '1970-01-01T00:00:00Z').AddSeconds($UnixSeconds)
}

function Get-PropValue {
  param(
    [Parameter()]$Obj,
    [Parameter(Mandatory)][string]$Name
  )
  if ($null -eq $Obj) { return $null }
  $p = $Obj.PSObject.Properties[$Name]
  if ($null -ne $p) { return $p.Value }
  return $null
}

function Get-ConditionKey {
  param([Parameter(Mandatory)]$a)
  $cfg = Get-PropValue $a $ActivityProps.SourceConfigUid
  $sub = Get-PropValue $a $ActivityProps.Subject
  if (-not $cfg) { $cfg = '<noConfigUid>' }
  if ([string]::IsNullOrWhiteSpace($sub)) { $sub = '<noSubject>' }
  "$cfg|$sub"
}

function ConvertTo-ObjectToHtmlTable {
    param (
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[Object]]$Objects
    )
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table><thead><tr>')
    $Objects[0].PSObject.Properties.Name |
    Where-Object { $_ -notin $HtmlTableExcludeProps } |
    ForEach-Object { [void]$sb.Append("<th>$_</th>") }

    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($obj in $Objects) {
        $rowClass = if ($null -ne $obj.PSObject.Properties[$HtmlTableRowColourProp]) { $obj.$HtmlTableRowColourProp } else { "" }
        [void]$sb.Append("<tr class=`"$rowClass`">")
        foreach ($propName in $obj.PSObject.Properties.Name | Where-Object { $_ -notin $HtmlTableExcludeProps }) {
            [void]$sb.Append("<td>$($obj.$propName)</td>")
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table>')
    $OutputLength = $sb.ToString() | Measure-Object -Character -IgnoreWhiteSpace | Select-Object -ExpandProperty Characters
    if ($OutputLength -gt $NinjaHtmlCharLimit) {
        Write-Warning ('Output appears to be over the NinjaOne WYSIWYG field limit of 200,000 characters. Actual length was: {0}' -f $OutputLength)
    }
    return $sb.ToString()
}

function Get-MapColour {
  param(
    [Parameter(Mandatory)]$MapList,
    [Parameter(Mandatory)][int]$Count
  )

  $Maximum = (@($MapList | Measure-Object).Count) - 1
  if ($Count -eq 0 -or $Maximum -lt 0) { return "" }

  $Index = [array]::IndexOf(@($MapList), "$Count")
  if ($Index -lt 0) { return "$Count" }

  $Sixth = $Maximum / 6
  if ($Index -ge 0 -and $Index -le ($Sixth * 4)) { return "$Count" }
  if ($Index -gt ($Sixth * 4) -and $Index -le ($Sixth * 5)) { return "<strong>$Count</strong>" }
  if ($Index -gt ($Sixth * 5) -and $Index -lt $Maximum) { return "<h2>$Count</h2>" }
  return "<h1>$Count</h1>"
}

function Get-HeatMapTableHtml {
    <#
      Generic heatmap table builder using Luke’s table formatting + intensity logic.
      XValues: columns
      YValues: rows
      ValueLookup: scriptblock that returns the count for a given (y,x)
    #>
    param(
      [Parameter(Mandatory)][string[]]$XValues,
      [Parameter(Mandatory)]$YValues,
      [Parameter(Mandatory)][scriptblock]$ValueLookup
    )
  
    $CellStyle = 'padding: 5px;border-width: 1px;border-style: solid;border-color: #D1D0DA;word-break: break-word;box-sizing: border-box;text-align:left;'
  
    # Build a base map of all cells for distribution mapping (Luke’s approach)
    $BaseMap = [ordered]@{}
    foreach ($y in $YValues) {
      foreach ($x in $XValues) {
        $k = "$y|$x"
        $BaseMap[$k] = [int](& $ValueLookup $y $x)
      }
    }
  
    $MapValues = $BaseMap.Values | Where-Object { $_ -ne 0 } | Group-Object
    $MapList = @($MapValues.Name | Sort-Object {[int]$_}) # ensure ascending numeric order
  
    # Header row
    $TableHTML = '<table width="100%" style="border-collapse: collapse;"><tbody>'
    $TableHTML += "<tr><td style='$CellStyle'></td>"
    foreach ($x in $XValues) {
      $TableHTML += "<td style='$CellStyle'><strong>$x</strong></td>"
    }
    $TableHTML += '</tr>'
  
    # Rows
    foreach ($y in $YValues) {
      $RowHTML = ''
      foreach ($x in $XValues) {
        $v = $BaseMap["$y|$x"]
        $cell = Get-MapColour -MapList $MapList -Count $v
        $RowHTML += "<td style='$CellStyle'>$cell</td>"
      }
      $TableHTML += "<tr><td style='$CellStyle'>$y</td>$RowHTML</tr>"
    }
  
    $TableHTML += '</tbody></table>'
    return $TableHTML
  }
  
  function Get-HeatMap_DayOfMonthByMonth {
    <#
      Builds a Day-of-Month (1..31) x Month (yyyy-MM) heatmap using Luke Whitelock’s formatting.
      InputEvents should have properties: monthKey (yyyy-MM) and dayOfMonth (int)
    #>
    param(
      [Parameter(Mandatory)]$InputEvents,
      [Parameter()][string[]]$Months
    )
  
    $events = @($InputEvents)
  
    # Determine months from data if not provided
    if (-not $Months -or $Months.Count -eq 0) {
      $Months = @($events | Select-Object -ExpandProperty monthKey | Sort-Object -Unique)
    }
  
    $days = 1..31
  
    # Precompute counts: month -> day -> count
    $matrix = @{}
    foreach ($m in $Months) { $matrix[$m] = @{} }
  
    foreach ($e in $events) {
      $m = $e.monthKey
      $d = [int]$e.dayOfMonth
      if (-not $matrix.ContainsKey($m)) { $matrix[$m] = @{} }
      if (-not $matrix[$m].ContainsKey($d)) { $matrix[$m][$d] = 0 }
      $matrix[$m][$d] = [int]$matrix[$m][$d] + 1
    }
  
    # Value lookup used by generic table builder
    $lookup = {
      param($y,$x) # y = day, x = month
      $day = [int]$y
      $month = [string]$x
      if ($matrix.ContainsKey($month) -and $matrix[$month].ContainsKey($day)) { 
        return [int]$matrix[$month][$day] 
      }
      return 0
    }.GetNewClosure()
  
    # IMPORTANT: Luke’s code expects string-ish Y values; we’ll pass day numbers as strings
    $YValues = @($days | ForEach-Object { "$_" })
  
    Get-HeatMapTableHtml -XValues $Months -YValues $YValues -ValueLookup $lookup
  }
  

# ---------------------------
# 1) Normalize & filter triggers
# ---------------------------
if (-not $activities) { throw "Expected `$activities` to be populated." }

$triggers = foreach ($a in $activities) {
  $atype = Get-PropValue $a $ActivityProps.ActivityType
  $scode = Get-PropValue $a $ActivityProps.StatusCode
  if ($atype -ne $ActivityFilter.Type -or $scode -ne $ActivityFilter.Status) { continue }

  $t = Get-PropValue $a $ActivityProps.ActivityTime
  $dt = $null
  if ($t -ne $null) {
    try { $dt = Convert-FromUnixSecondsUtc -UnixSeconds ([double]$t) } catch { $dt = $null }
  }

  $deviceId = Get-PropValue $a $ActivityProps.DeviceId
  $key      = Get-ConditionKey -a $a
  $name     = Get-PropValue $a $ActivityProps.Message
  $subject  = Get-PropValue $a $ActivityProps.Subject
  $msg      = Get-PropValue $a $ActivityProps.Message

  # Friendly name fallback if sourceName is blank
  if ([string]::IsNullOrWhiteSpace($name)) {
    if (-not [string]::IsNullOrWhiteSpace($subject)) { $name = $subject }
    elseif (-not [string]::IsNullOrWhiteSpace($msg)) { $name = ($msg -replace '\s+',' ').Trim() }
    else { $name = '<unknown condition>' }
  }

  if ($null -eq $dt) { continue } # skip if time couldn't be parsed

  [pscustomobject]@{
    deviceId        = $deviceId
    conditionKey    = $key
    conditionName   = $name
    activityTimeUtc = $dt
    monthKey        = $dt.ToString('yyyy-MM')
    dayOfMonth      = [int]$dt.Day
  }
}

$triggers = @($triggers)
if ($triggers.Count -eq 0) {
  throw "No CONDITION/TRIGGERED events found in `$activities`."
}

# ---------------------------
# 2) Rank devices by total triggers
# ---------------------------
$deviceTotals = $triggers |
  Group-Object -Property deviceId |
  ForEach-Object {
    [pscustomobject]@{
      deviceId       = $_.Name
      triggerCount   = @($_.Group).Count
    }
  } |
  Sort-Object triggerCount -Descending

# Add rank (dense rank: ties share rank; next rank increments by 1)
$ranked = @()
$rank = 0
$prevCount = $null
foreach ($row in $deviceTotals) {
  if ($prevCount -eq $null -or $row.triggerCount -ne $prevCount) { $rank++; $prevCount = $row.triggerCount }
  $ranked += [pscustomobject]@{
    deviceId     = $row.deviceId
    triggerCount = $row.triggerCount
    rank         = $rank
  }
}

# Quick lookup
$rankLookup = @{}
foreach ($r in $ranked) { $rankLookup[$r.deviceId] = $r }

# ---------------------------
# 3) Build per-device report objects
# ---------------------------
$DeviceReports = foreach ($devGroup in ($triggers | Group-Object deviceId)) {
  $deviceId = $devGroup.Name
  $events = @($devGroup.Group)

  $totalTriggers = $events.Count

  # Top 10 conditions
  $top10 = $events |
    Group-Object -Property conditionKey |
    ForEach-Object {
      $g = @($_.Group)
      $rep = $g | Select-Object -First 1
      [pscustomobject]@{
        ConditionKey  = $_.Name
        ConditionName = $rep.conditionName
        Count         = $g.Count
      }
    } |
    Sort-Object Count -Descending |
    Select-Object -First 10

  $top10ForHtml = [System.Collections.Generic.List[Object]]::new([Object[]]@($top10 | Select-Object @{n='ConditionName';e={$_.ConditionName}}, @{n='Count';e={$_.Count}}))
  $top10Text = ConvertTo-ObjectToHtmlTable -Objects $top10ForHtml

  # Build heatmap matrix: Month -> Day -> Count
  $months = @($events | Select-Object -ExpandProperty monthKey | Sort-Object -Unique)

  $matrix = @{}
  foreach ($m in $months) { $matrix[$m] = @{} }

  foreach ($e in $events) {
    $m = $e.monthKey
    $d = [int]$e.dayOfMonth
    if (-not $matrix.ContainsKey($m)) { $matrix[$m] = @{} }
    if (-not $matrix[$m].ContainsKey($d)) { $matrix[$m][$d] = 0 }
    $matrix[$m][$d] = [int]$matrix[$m][$d] + 1
  }

  $heatmapHtml = Get-HeatMap_DayOfMonthByMonth -InputEvents $events -Months $months

  $rankInfo = $null
  if ($rankLookup.ContainsKey($deviceId)) { $rankInfo = $rankLookup[$deviceId] }

  [pscustomobject]@{
    deviceId               = $deviceId
    totalConditionsTriggered= $totalTriggers

    # top 10 as objects (great for exporting) + a preformatted table string
    top10Conditions         = $top10
    top10ConditionsTableText= $top10Text

    # rank relative to all devices that triggered at least one condition
    rankAmongAlertingDevices = if ($rankInfo) { $rankInfo.rank } else { $null }
    totalAlertingDevices     = $ranked.Count

    # heatmap outputs
    heatmapMonths           = $months
    heatmapMatrix           = $matrix
    heatmapHtml             = $heatmapHtml
  }
}

# ---------------------------
# 4) Global summary (optional)
# ---------------------------
$GlobalSummary = [pscustomobject]@{
  timeMinUtc          = ($triggers | Measure-Object activityTimeUtc -Minimum).Minimum
  timeMaxUtc          = ($triggers | Measure-Object activityTimeUtc -Maximum).Maximum
  totalTriggeredEvents= $triggers.Count
  totalAlertingDevices= $ranked.Count
  deviceRanks         = $ranked
}

# ---------------------------
# 5) Example console output
# ---------------------------
"== Device rank by CONDITION/TRIGGERED volume ==" | Write-Host
$GlobalSummary.deviceRanks | Select-Object -First 20 | Format-Table -AutoSize

"`n== Example: first device report (top10 table) ==" | Write-Host
$DeviceReports | Select-Object -First 1 -ExpandProperty top10ConditionsTableText | Write-Host

# ---------------------------
# Optional exports
# ---------------------------
# Export per-device summary (flattened top-level fields)
# $DeviceReports | Select-Object deviceId,totalConditionsTriggered,rankAmongAlertingDevices,totalAlertingDevices |
#   Export-Csv -NoTypeInformation .\DeviceConditionTriggerSummary.csv

# Export top10 conditions per device (one row per condition)
# $DeviceReports |
#   ForEach-Object {
#     $dev = $_.deviceId
#     $_.top10Conditions | ForEach-Object {
#       [pscustomobject]@{
#         deviceId      = $dev
#         conditionName = $_.ConditionName
#         conditionKey  = $_.ConditionKey
#         triggers      = $_.Count
#       }
#     }
#   } | Export-Csv -NoTypeInformation .\DeviceTop10Conditions.csv

# If you want to view the HTML heatmap for a specific device locally:
$d = $DeviceReports | Where-Object deviceId -eq 5 | Select-Object -First 1
$dir = "/Users/jeffhunter/Documents/NinjaReports"
New-Item -ItemType Directory -Path $dir -Force | Out-Null
$path = Join-Path $dir "heatmap_device_5.html"
$d.heatmapHtml | Out-File -Encoding utf8 $path
 
Write-Host "Wrote $path"

# -----------------------------
# Custom field mapping
# -----------------------------
function New-CustomFieldPayloadFromReport {
  param(
    [Parameter(Mandatory=$true)]$Report,
    [bool]$OverwriteEmptyValues
  )

  $cf = @{}

  function Set-CF([string]$Key, $Value) {
    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) {
      if ($OverwriteEmptyValues) { $cf[$Key] = $null }
      return
    }
    $cf[$Key] = $Value
  }

  function Set-CFHtml([string]$Key, [string]$Html) {
    if ([string]::IsNullOrWhiteSpace($Html)) {
      if ($OverwriteEmptyValues) { $cf[$Key] = @{ html = "" } }
      return
    }
    $cf[$Key] = @{ html = $Html }
  }

  Set-CF -Key $CustomFieldNames.TotalConditionsTriggered -Value $Report.totalConditionsTriggered
  Set-CF -Key $CustomFieldNames.DeviceAlertRank -Value $Report.rankAmongAlertingDevices
  Set-CF -Key $CustomFieldNames.TotalAlertingDevices -Value $Report.totalAlertingDevices
  Set-CFHtml -Key $CustomFieldNames.AlertHeatMap -Html $Report.heatmapHtml
  Set-CFHtml -Key $CustomFieldNames.MostFrequentAlerts -Html $Report.top10ConditionsTableText

  return $cf
}

# ========= UPDATE LOOP =========
foreach ($r in $DeviceReports) {
  if (-not $r.deviceId) {
    Write-Warning "Skipping report missing deviceId"
    continue
  }

  $customFields = New-CustomFieldPayloadFromReport -Report $r -OverwriteEmptyValues:$OverwriteEmptyValues

  if ($customFields.Count -eq 0) {
    Write-Host "Skipping deviceId $($r.deviceId) (no fields to update)."
    continue
  }

  $customfields_url = "$($Config.Instance)$($ApiPaths.CustomFields -f $r.deviceId)"
  $json = $customFields | ConvertTo-Json -Depth 10

  Write-Host "Patching deviceId $($r.deviceId) with:"
  Write-Host $json

  $result = Invoke-NinjaAPIRequest -Uri $customfields_url -Method 'PATCH' -Headers $headers -Body $json
  if ($null -eq $result) {
    Write-Error "Failed to update custom fields for deviceId $($r.deviceId)"
  }

  Start-Sleep -Seconds 1
}

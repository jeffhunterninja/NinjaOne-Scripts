# NinjaOne Automation & Scheduled Task Dashboard Script
# High-performance automation monitoring and reporting across all clients
#
# Activities are read from a SQLite database populated by Get-AutomationActivities.ps1.
# Run Get-AutomationActivities first to create/sync the DB; do not use the API for activity retrieval.
#
# Environment Variables (optional):
# - reportMonth : Month/year override (e.g., "December 2024")
# - URLOverride : Branded portal base URL for hyperlinks ONLY (e.g., "https://support.example.com").
#                  Intended for branded NinjaOne URLs that point to the SAME tenant/KB content as the
#                  API instance. Not for cross-tenant/region use. If omitted, links default to the
#                  discovered NinjaOne instance host used for API calls.

[CmdletBinding()]
param (
    [Parameter()]
    [string]$ReportMonth = [System.Convert]::ToString($env:reportMonth),

    [Parameter()]
    [string]$URLOverride = [System.Convert]::ToString($env:URLOverride),

    [Parameter()]
    [string]$NinjaOneInstance = '',

    [Parameter()]
    [string]$NinjaOneClientId = '',

    [Parameter()]
    [string]$NinjaOneClientSecret = '',

    [Parameter()]
    [string]$DbPath = 'C:\RMM\Activities.db',

    [Parameter()]
    [string]$SqliteExePath = 'C:\ProgramData\chocolatey\bin\sqlite3.exe'
)

#region Helper Functions
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error")]
        [string]$Level = "Info"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message"
}

function Initialize-PowerShell7 {
    if ($PSVersionTable.PSVersion.Major -ge 7) { return }
    $ps7Path = $null
    try {
        $cmd = Get-Command pwsh -ErrorAction SilentlyContinue
        if ($cmd -and $cmd.Source -and (Test-Path $cmd.Source)) { $ps7Path = $cmd.Source }
    } catch { }
    if (-not $ps7Path) {
        $candidates = @(
            (Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'),
            (Join-Path $env:ProgramFiles 'PowerShell\7-preview\pwsh.exe')
        )
        $ps7Path = ($candidates | Where-Object { Test-Path $_ } | Select-Object -First 1)
    }
    if (-not $ps7Path) {
        Write-Log 'PowerShell 7 not found; continuing under Windows PowerShell 5.1' 'Info'
        return
    }
    $arguments = @('-File', $PSCommandPath)
    if ($PSBoundParameters) {
        foreach ($param in $PSBoundParameters.GetEnumerator()) {
            $arguments += ('-' + $param.Key)
            if ($param.Value -is [switch] -and $param.Value) { continue }
            if ($param.Value) { $arguments += [string]$param.Value }
        }
    }
    & $ps7Path @arguments
    exit $LASTEXITCODE
}

function Initialize-NinjaOneModule {
    $moduleName = "NinjaOneDocs"
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Log "Installing $moduleName module..."
        try {
            Install-Module -Name $moduleName -Force -AllowClobber -ErrorAction Stop
        } catch {
            # Fallback to current user if AllUsers fails in restricted contexts
            Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
    }
    Import-Module -Name $moduleName -ErrorAction Stop
}


function Get-DateRange {
    param([string]$ReportMonth = "")
    if ($ReportMonth) {
        try { $parsedDate = [datetime]::ParseExact($ReportMonth, "MMMM yyyy", [cultureinfo]::InvariantCulture) }
        catch { throw "Invalid ReportMonth format. Use 'MMMM yyyy' (e.g., 'August 2025')." }
        $targetDate = $parsedDate
    } else { $targetDate = Get-Date }
    $firstDay = Get-Date -Year $targetDate.Year -Month $targetDate.Month -Day 1
    $lastDay  = $firstDay.AddMonths(1).AddDays(-1)
    return @{
        Current = @{
            FirstDay       = $firstDay
            LastDay        = $lastDay
            FirstDayString = $firstDay.ToString('yyyyMMdd')
            LastDayString  = $lastDay.ToString('yyyyMMdd')
            Month          = $firstDay.ToString("MMMM")
            Year           = $firstDay.ToString("yyyy")
        }
    }
}

function Get-KBFolderContent {
    [CmdletBinding()]
    param([string]$FolderPath, [int]$OrganizationId)
    $qs = @{}
    if ($FolderPath)     { $qs['folderPath']    = $FolderPath }
    if ($OrganizationId) { $qs['organizationId'] = $OrganizationId }
    $query = ($qs.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, [uri]::EscapeDataString([string]$_.Value) }) -join '&'
    return Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/folder' -QueryParams $query
}

function Get-GlobalKBArticleByName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$FolderPath = "Monthly Reports"
    )
    # Trim to avoid false misses
    $trimName = if ([string]::IsNullOrWhiteSpace($Name)) { '' } else { $Name.Trim() }
    # Direct search first (include archived)
    $qs = "articleName=$([uri]::EscapeDataString($trimName))&includeArchived=true"
    $list = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/global/articles' -QueryParams $qs
    $hit = $list | Where-Object { $_.name -eq $trimName } | Select-Object -First 1
    if ($hit -and $hit.PSObject.Properties['id'] -and ([long]$hit.id) -gt 0) { return $hit }
    # Fallback: enumerate the folder
    try {
        $folder = Get-KBFolderContent -FolderPath $FolderPath -OrganizationId 0
        $fromFolder = @($folder.files) | Where-Object { $_.name -eq $trimName } | Select-Object -First 1
        if ($fromFolder -and $fromFolder.PSObject.Properties['id'] -and ([long]$fromFolder.id) -gt 0) { return $fromFolder }
    } catch { }
    return $null
}

function Build-KBDeepLink {
    <#
      .SYNOPSIS
      Builds a deep link URL to a KB article using a provided base URL.
      .PARAMETER BaseUrl
      The definitive base URL for links (supports branded portal override). Scheme assumed https if missing.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Article, [Parameter(Mandatory)][string]$BaseUrl)
    if (-not $Article -or -not $Article.id -or -not $Article.parentFolderId) { return $null }
    $base = if ($BaseUrl -match '^https?://') { $BaseUrl } else { 'https://{0}' -f $BaseUrl }
    $base = $base -replace '/+$',''
    "$base/#/systemDashboard/knowledgeBase/$($Article.parentFolderId)/$($Article.id)/file"
}

function Get-KBSearchUrl {
    param([Parameter(Mandatory)][string]$BaseUrl, [Parameter(Mandatory)][string]$ArticleName)
    $base = if ($BaseUrl -match '^https?://') { $BaseUrl } else { "https://$BaseUrl" }
    $base = $base -replace '/+$',''
    "$base/#/knowledgebase/global/articles?articleName=$([uri]::EscapeDataString($ArticleName))"
}

function Convert-ActivityTime {
  param([Parameter(Mandatory = $false)]$TimeValue)
  if ($null -eq $TimeValue -or [string]::IsNullOrWhiteSpace([string]$TimeValue)) { return (Get-Date).ToUniversalTime() }
  try {
      switch ($TimeValue.GetType().Name) {
          "DateTime" { return ([DateTime]$TimeValue).ToUniversalTime() }
          "Int64"    { return [System.DateTimeOffset]::FromUnixTimeSeconds($TimeValue).UtcDateTime }
          "Int32"    { return [System.DateTimeOffset]::FromUnixTimeSeconds([int64]$TimeValue).UtcDateTime }
          "Double"   { $u=[long][math]::Floor($TimeValue); return [System.DateTimeOffset]::FromUnixTimeSeconds($u).UtcDateTime }
          "String"   {
              $dt=$null
              if ([DateTime]::TryParse($TimeValue,[ref]$dt)) { return $dt.ToUniversalTime() }
              $ux=0
              if ([long]::TryParse($TimeValue,[ref]$ux)) { return [System.DateTimeOffset]::FromUnixTimeSeconds($ux).UtcDateTime }
              return (Get-Date).ToUniversalTime()
          }
          default    { return (Get-Date).ToUniversalTime() }
      }
  } catch {
      return (Get-Date).ToUniversalTime()
  }
}

function Invoke-SqliteQuery {
    param([Parameter(Mandatory)][string]$SqliteExe, [Parameter(Mandatory)][string]$DataSource, [Parameter(Mandatory)][string]$Sql)
    $out = & $SqliteExe -csv -header $DataSource $Sql 2>$null
    $text = if ($null -eq $out) { '' } elseif ($out -is [string]) { $out } else { $out -join "`n" }
    $text = $text.TrimStart([char]0xFEFF)
    $lines = $text -split "`r?`n"
    $lines = $lines | Where-Object { $_.Length -gt 0 }
    if ($lines.Count -lt 2) { return @() }
    $headers = $lines[0] -split ',(?=(?:[^"]*"[^"]*")*[^"]*$)'
    $headers = $headers | ForEach-Object { $_.Trim().Trim('"').TrimStart([char]0xFEFF) }
    $result = [System.Collections.Generic.List[pscustomobject]]::new()
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $fields = $lines[$i] -split ',(?=(?:[^"]*"[^"]*")*[^"]*$)'
        $fields = $fields | ForEach-Object { $_.Trim().Trim('"').Replace('""', '"') }
        $row = [ordered]@{}
        for ($j = 0; $j -lt [Math]::Min($headers.Count, $fields.Count); $j++) { $row[$headers[$j]] = $fields[$j] }
        $result.Add([pscustomobject]$row)
    }
    return $result
}

function Get-ActivitiesForMonthFromDb {
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory = $true)][hashtable]$DateRange,
        [Parameter(Mandatory = $true)][string]$DbPath,
        [Parameter(Mandatory = $true)][string]$SqliteExe
    )
    $epoch0 = [datetime]'1970-01-01T00:00:00Z'
    function Convert-ToEpochSeconds { param([Parameter(Mandatory)][datetime]$Date) $utc = $Date.ToUniversalTime(); [int64][math]::Floor(($utc - $epoch0).TotalSeconds) }
    $monthStart = [datetime]$DateRange.Current.FirstDay
    $monthEndEx = ([datetime]$DateRange.Current.LastDay).AddDays(1).Date
    $afterEpoch = Convert-ToEpochSeconds -Date $monthStart
    $beforeEpochEx = Convert-ToEpochSeconds -Date $monthEndEx
    $afterEpochMs = $afterEpoch * 1000
    $beforeEpochExMs = $beforeEpochEx * 1000
    if ([string]::IsNullOrWhiteSpace($DbPath) -or -not (Test-Path -LiteralPath $DbPath -PathType Leaf)) {
        throw "SQLite database not found at '$DbPath'. Run Get-AutomationActivities.ps1 first to create and populate the Activities database."
    }
    $diagTotal = Invoke-SqliteQuery -SqliteExe $SqliteExe -DataSource $DbPath -Sql "SELECT COUNT(*) AS cnt FROM Activities;"
    $diagCompleted = Invoke-SqliteQuery -SqliteExe $SqliteExe -DataSource $DbPath -Sql "SELECT COUNT(*) AS cnt FROM Activities WHERE (statusCode IS NOT NULL AND UPPER(TRIM(statusCode)) = 'COMPLETED') OR (status IS NOT NULL AND UPPER(TRIM(status)) = 'COMPLETED');"
    $diagInRange = Invoke-SqliteQuery -SqliteExe $SqliteExe -DataSource $DbPath -Sql "SELECT COUNT(*) AS cnt FROM Activities WHERE ((statusCode IS NOT NULL AND UPPER(TRIM(statusCode)) = 'COMPLETED') OR (status IS NOT NULL AND UPPER(TRIM(status)) = 'COMPLETED')) AND ((activityTime >= $afterEpoch AND activityTime < $beforeEpochEx) OR (activityTime >= $afterEpochMs AND activityTime < $beforeEpochExMs));"
    $diagT = if ($diagTotal.Count -gt 0 -and $null -ne $diagTotal[0].cnt) { $diagTotal[0].cnt } else { '' }
    $diagC = if ($diagCompleted.Count -gt 0 -and $null -ne $diagCompleted[0].cnt) { $diagCompleted[0].cnt } else { '' }
    $diagR = if ($diagInRange.Count -gt 0 -and $null -ne $diagInRange[0].cnt) { $diagInRange[0].cnt } else { '' }
    Write-Log ("DB diagnostics: total={0}, COMPLETED={1}, COMPLETED in date range={2} (epoch {3}-{4}, ms {5}-{6})" -f $diagT, $diagC, $diagR, $afterEpoch, $beforeEpochEx, $afterEpochMs, $beforeEpochExMs)
    $sql = "SELECT id, created_at, activityTime, deviceId, seriesUid, activityType, statusCode, status, activityResult, sourceConfigUid, sourceName, subject, message, type, data, OrgID, OrgName, LocID, LocName, DeviceName FROM Activities WHERE ((statusCode IS NOT NULL AND UPPER(TRIM(statusCode)) = 'COMPLETED') OR (status IS NOT NULL AND UPPER(TRIM(status)) = 'COMPLETED')) AND ((activityTime >= $afterEpoch AND activityTime < $beforeEpochEx) OR (activityTime >= $afterEpochMs AND activityTime < $beforeEpochExMs));"
    $rows = Invoke-SqliteQuery -SqliteExe $SqliteExe -DataSource $DbPath -Sql $sql
    $results = [System.Collections.Generic.List[object]]::new()
    $fallbackUtc = $monthStart.ToUniversalTime()
    foreach ($r in $rows) {
        $act = [pscustomobject]@{
            id               = if ($r.id -match '^\d+$') { [int]$r.id } else { 0 }
            created_at       = $r.created_at
            activityTime     = $r.activityTime
            deviceId        = if ($r.deviceId -match '^\d+$') { [int]$r.deviceId } else { 0 }
            seriesUid        = $r.seriesUid
            activityType     = $r.activityType
            statusCode      = $r.statusCode
            status          = $r.status
            activityResult  = $r.activityResult
            sourceConfigUid  = $r.sourceConfigUid
            sourceName      = $r.sourceName
            subject         = $r.subject
            message         = $r.message
            type            = $r.type
            data            = $r.data
            OrgID           = if ($r.OrgID -match '^\d+$') { [int]$r.OrgID } else { 0 }
            OrgName         = $r.OrgName
            LocID           = if ($r.LocID -match '^\d+$') { [int]$r.LocID } else { 0 }
            LocName         = $r.LocName
            DeviceName      = $r.DeviceName
        }
        if ([string]::IsNullOrWhiteSpace([string]$act.statusCode) -and -not [string]::IsNullOrWhiteSpace([string]$act.status)) { $act | Add-Member -NotePropertyName statusCode -NotePropertyValue $act.status -Force }
        if ([string]::IsNullOrWhiteSpace([string]$act.activityType) -and -not [string]::IsNullOrWhiteSpace([string]$act.type)) { $act | Add-Member -NotePropertyName activityType -NotePropertyValue $act.type -Force }
        $activityTimeNum = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$r.activityTime)) {
            if ([double]::TryParse([string]$r.activityTime, [ref]$activityTimeNum)) { $act.activityTime = $activityTimeNum }
        }
        $timeForConvert = if ($null -ne $activityTimeNum -and $activityTimeNum -ge 1000000000000) { $activityTimeNum / 1000 } else { $act.activityTime }
        $activityTimeConverted = if ($null -ne $timeForConvert -and -not [string]::IsNullOrWhiteSpace([string]$timeForConvert)) { Convert-ActivityTime $timeForConvert } else { $fallbackUtc }
        $act | Add-Member -NotePropertyName ActivityTimeConverted -NotePropertyValue $activityTimeConverted -Force
        if ([string]::IsNullOrWhiteSpace($act.sourceName)) { $act.sourceName = 'Action ()' }
        if ($act.sourceName -eq 'Action ()' -and -not [string]::IsNullOrWhiteSpace([string]$act.data)) {
            try {
                $dataObj = $act.data | ConvertFrom-Json -ErrorAction Stop
                if ($dataObj.PSObject.Properties['message'] -and $dataObj.message.PSObject.Properties['params'] -and $dataObj.message.params.PSObject.Properties['action_name']) {
                    $act.sourceName = [string]$dataObj.message.params.action_name
                }
            } catch { }
        }
        [void]$results.Add($act)
    }
    return $results.ToArray()
}

function Get-NinjaOneData {
    param(
        [Parameter(Mandatory = $true)] $DateRange,
        [Parameter(Mandatory = $true)][string]$DbPath,
        [Parameter(Mandatory = $true)][string]$SqliteExe
    )
    Write-Log "Fetching core data from NinjaOne API..."
    $devices = Invoke-NinjaOneRequest -Method GET -Path 'devices-detailed'
    $organizations = Invoke-NinjaOneRequest -Method GET -Path 'organizations'
    $locations = Invoke-NinjaOneRequest -Method GET -Path 'locations'
    $deviceIndex = @{}; $orgIndex = @{}; $locationIndex = @{}
    foreach ($d in $devices) { $deviceIndex[$d.id] = $d }
    foreach ($o in $organizations) { $orgIndex[$o.id] = $o }
    foreach ($loc in $locations) { $locationIndex[$loc.id] = $loc }
    Write-Log "Found $($devices.Count) devices across $($organizations.Count) organizations, $($locations.Count) locations"
    Write-Log "Loading automation activities from SQLite database..."
    $allActivity = Get-ActivitiesForMonthFromDb -DateRange $DateRange -DbPath $DbPath -SqliteExe $SqliteExe
    Write-Log "Data collection complete: $($allActivity.Count) COMPLETED activities"
    foreach ($a in $allActivity) {
        if ([string]::IsNullOrWhiteSpace([string]$a.statusCode) -and $null -ne $a.PSObject.Properties['status']) { $a | Add-Member -NotePropertyName statusCode -NotePropertyValue $a.status -Force }
        if ([string]::IsNullOrWhiteSpace([string]$a.activityType) -and $null -ne $a.PSObject.Properties['type']) { $a | Add-Member -NotePropertyName activityType -NotePropertyValue $a.type -Force }
    }
    foreach ($a in $allActivity) {
        $deviceClass = ''
        if ($a.deviceId -and $deviceIndex.ContainsKey($a.deviceId)) {
            $deviceClass = $deviceIndex[$a.deviceId].nodeClass
        }
        $a | Add-Member -NotePropertyName DeviceClass -NotePropertyValue $deviceClass -Force
    }
    Write-Log "Data collection Enrichment complete: $($allActivity.Count) COMPLETED activities"
    return @{ Devices=@($devices); Organizations=@($organizations); DeviceIndex=$deviceIndex; OrgIndex=$orgIndex; Activities=@($allActivity) }
}

function Get-AutomationAnalysis {
    param([Parameter(Mandatory = $true)] $AutomationData, [bool]$UseCache = $true)
    $cacheKey = "analysis_" + ($AutomationData.Activities | Measure-Object).Count
    if ($UseCache -and $script:AnalysisCache -and $script:AnalysisCache.ContainsKey($cacheKey)) { Write-Log "Using cached analysis results"; return $script:AnalysisCache[$cacheKey] }
    if (-not $AutomationData.Activities -or $AutomationData.Activities.Count -eq 0) { return @{ SuccessfulActivities=@(); FailedActivities=@(); AutomationSummary=@(); DeviceFailures=@(); OrganizationStats=@(); OverallStats=@{ TotalActivities=0; SuccessRate=0; FailureRate=0 } } }

    Write-Log "Analyzing $($AutomationData.Activities.Count) automation activities..."
    $failedActivities     = $AutomationData.Activities | Where-Object { ($_.activityResult -eq "FAILURE") -or ($_.statusCode -in @("CANCELLED","BLOCKED","EVALUATION_FAILURE","FAILED")) }
    $successfulActivities = $AutomationData.Activities | Where-Object { ($_.activityResult -ne "FAILURE") -and ($_.statusCode -notin @("CANCELLED","BLOCKED","EVALUATION_FAILURE","FAILED")) }

    $automationSummary = @()
    if ($AutomationData.Activities.Count -gt 0) {
        $groups = $AutomationData.Activities | Group-Object sourceName
        foreach ($g in $groups) {
            $acts = $g.Group
            $failed = $acts | Where-Object { ($_.activityResult -eq "FAILURE") -or ($_.statusCode -in @("CANCELLED","BLOCKED","EVALUATION_FAILURE","FAILED")) }
            $succ   = $acts | Where-Object { ($_.activityResult -ne "FAILURE") -and ($_.statusCode -notin @("CANCELLED","BLOCKED","EVALUATION_FAILURE","FAILED")) }
            $automationSummary += [PSCustomObject]@{
                AutomationType  = if ($acts[0].activityType) { $acts[0].activityType } else { "Unknown" }
                AutomationName  = $g.Name
                SeriesUID       = ($acts | Select-Object -First 1).seriesUid
                TotalRuns       = $acts.Count
                Successful      = $succ.Count
                Failed          = $failed.Count
                SuccessRate     = if ($acts.Count -gt 0) { [math]::Round(($succ.Count / $acts.Count * 100), 1) } else { 0 }
                LastRun         = ($acts | Sort-Object ActivityTimeConverted -Descending | Select-Object -First 1).ActivityTimeConverted
                DevicesAffected = ($acts | Select-Object -ExpandProperty deviceId -Unique).Count
            }
        }
    }

    $deviceFailures = @()
    if ($failedActivities.Count -gt 0) {
        $dfGroups = $failedActivities | Group-Object deviceId
        $deviceFailures = foreach ($g in $dfGroups) {
            $fa = $g.Group[0]
            [PSCustomObject]@{
                DeviceName   = $fa.DeviceName
                DeviceClass  = $fa.DeviceClass
                OrgName      = $fa.OrgName
                LocName      = if ($fa.LocName) { $fa.LocName } else { '' }
                FailureCount = $g.Count
                LastFailure  = ($g.Group | Sort-Object ActivityTimeConverted -Descending | Select-Object -First 1).ActivityTimeConverted
                FailureTypes = ($g.Group | Select-Object -ExpandProperty sourceName -Unique | Select-Object -First 3) -join ', '
            }
        }
    }

    $organizationStats = @()
    if ($AutomationData.Activities.Count -gt 0) {
        $orgGroups = $AutomationData.Activities | Group-Object OrgID
        $organizationStats = foreach ($g in $orgGroups) {
            $orgActs = $g.Group

            # Use the same failure criteria used everywhere else
            $orgFailed = $orgActs | Where-Object {
                ($_.activityResult -eq "FAILURE") -or ($_.statusCode -in @("CANCELLED","BLOCKED","EVALUATION_FAILURE","FAILED"))
            }
            $orgSuccessful = $orgActs | Where-Object {
                ($_.activityResult -ne "FAILURE") -and ($_.statusCode -notin @("CANCELLED","BLOCKED","EVALUATION_FAILURE","FAILED"))
            }

            # NEW: most recent failure timestamp for the org (any device)
            $lastFailure = $null
            if ($orgFailed.Count -gt 0) {
                $lastFailure = ($orgFailed | Sort-Object ActivityTimeConverted -Descending | Select-Object -First 1).ActivityTimeConverted
            }

            [PSCustomObject]@{
                OrganizationName    = $orgActs[0].OrgName
                TotalActivities     = $orgActs.Count
                Successful          = $orgSuccessful.Count
                Failed              = $orgFailed.Count
                SuccessRate         = if ($orgActs.Count -gt 0) { [math]::Round(($orgSuccessful.Count / $orgActs.Count * 100), 1) } else { 0 }
                DevicesWithActivity = ($orgActs | Select-Object -ExpandProperty deviceId -Unique).Count
                AutomationTypes     = ($orgActs | Select-Object -ExpandProperty sourceName -Unique | Select-Object -First 5) -join ', '
                LastFailure         = $lastFailure   # NEW FIELD
            }
        }
    }


    $total = $AutomationData.Activities.Count
    $overall = @{ TotalActivities=$total; SuccessRate= if ($total -gt 0) { [math]::Round(($successfulActivities.Count / $total * 100), 1) } else { 0 }; FailureRate= if ($total -gt 0) { [math]::Round(($failedActivities.Count / $total * 100), 1) } else { 0 } }
    $result = @{ SuccessfulActivities=@($successfulActivities); FailedActivities=@($failedActivities); AutomationSummary=@($automationSummary | Sort-Object SuccessRate); DeviceFailures=@($deviceFailures | Sort-Object FailureCount -Descending); OrganizationStats=@($organizationStats | Sort-Object SuccessRate); OverallStats=$overall }
    if (-not $script:AnalysisCache) { $script:AnalysisCache = @{} }
    $script:AnalysisCache[$cacheKey] = $result
    Write-Log "Analysis complete: $total total, $($successfulActivities.Count) successful, $($failedActivities.Count) failed"
    return $result
}

function Get-DeviceSummary {
    param([Parameter(Mandatory = $true)][array]$Activities)
    if (-not $Activities -or $Activities.Count -eq 0) { return @() }
    $isFailure = { ($_.activityResult -eq "FAILURE") -or ($_.statusCode -in @("CANCELLED","BLOCKED","EVALUATION_FAILURE","FAILED")) }
    $isSuccess = { ($_.activityResult -ne "FAILURE") -and ($_.statusCode -notin @("CANCELLED","BLOCKED","EVALUATION_FAILURE","FAILED")) }
    $groups = $Activities | Group-Object deviceId
    $summaries = foreach ($g in $groups) {
        $acts = $g.Group; $first = $acts[0]
        $succ = @($acts | Where-Object $isSuccess)
        $fail = @($acts | Where-Object $isFailure)
        $last = ($acts | Sort-Object ActivityTimeConverted -Descending | Select-Object -First 1).ActivityTimeConverted
        [pscustomobject]@{
            DeviceId    = $g.Name
            DeviceName  = $first.DeviceName
            Type        = $first.DeviceClass
            LocName     = if ($first.LocName) { $first.LocName } else { '' }
            TotalRuns   = $acts.Count
            Success     = $succ.Count
            Failed      = $fail.Count
            SuccessRate = if ($acts.Count -gt 0) { [math]::Round(($succ.Count / $acts.Count * 100), 1) } else { 0 }
            LastRun     = $last
        }
    }
    return @($summaries | Sort-Object @{Expression='Failed';Descending=$true}, @{Expression='TotalRuns';Descending=$true}, @{Expression='DeviceName';Descending=$false})
}
#endregion

#region Dashboard Generation Functions
function New-DashboardKPICard {
    param([Parameter(Mandatory = $true)][string]$Title, [Parameter(Mandatory = $true)][string]$Value, [string]$Description = "", [ValidateSet("primary","success","warning","danger","info")][string]$Type = "primary")
    $colors = @{ primary="#0d6efd"; success="#198754"; warning="#ffc107"; danger="#dc3545"; info="#0dcaf0" }
    $color = $colors[$Type]
    $descriptionHtml = if ($Description) { "<div style='font-size: 0.75rem; color: #666; line-height: 1.4; margin-top: 5px;'>$Description</div>" } else { "" }
    return @"
<div style="display: inline-block; vertical-align: top; width: 180px; margin: 10px; padding: 20px 15px; background: white; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); border-left: 6px solid $color; text-align: center;">
  <div style="font-size: 2.5rem; font-weight: bold; color: $color; margin-bottom: 8px; line-height: 1;">$Value</div>
  <div style="font-size: 0.95rem; color: #333; font-weight: 600; margin-bottom: 5px; line-height: 1.3;">$Title</div>
  $descriptionHtml
</div>
"@
}

function New-PieChart {
    param([int]$SuccessfulCount, [int]$FailedCount, [string]$Title = "Automation Success Rate")
    $total = $SuccessfulCount + $FailedCount
    if ($total -eq 0) { return "<div style='text-align: center; color: #666; padding: 40px;'>No automation data available</div>" }
    $successRate    = [math]::Round(($SuccessfulCount / $total * 100), 1)
    $failureRate    = [math]::Round(($FailedCount / $total * 100), 1)
    $successFraction= [math]::Round(($SuccessfulCount / $total), 3)
    return @"
<div style="background: white; padding: 25px; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); margin: 0px auto 20px auto; max-width: 700px;">
  <div style="display: flex; justify-content: center; margin: 30px 0;">
    <div style="height: 300px; width: 300px;">
      <table class="charts-css pie show-heading" style="height: 100%; width: 100%; border: none !important; border-collapse: collapse;">
        <tbody>
          <tr style="color: #22c55e;">
            <th scope="row" style="border: none !important;">Successful</th>
            <td style="--start: 0; --end: $successFraction; --color: #22c55e; border: none !important;">
              <span class="data">$SuccessfulCount ($successRate%)</span>
            </td>
          </tr>
          <tr style="color: #ef4444;">
            <th scope="row" style="border: none !important;">Failed</th>
            <td style="--start: $successFraction; --end: 1; --color: #ef4444; border: none !important;">
              <span class="data">$FailedCount ($failureRate%)</span>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  </div>
  <div style="text-align: center; font-size: 1.1rem; font-weight: 600; color: #333; margin-top: 20px;">
    Successful · Failed
  </div>
</div>
"@
}

function New-AlertCard {
    param([string]$Title, [array]$Items, [string]$Type = "info", [int]$MaxItems = 3)
    if (-not $Items) { $Items = @() }
    $colors = @{ error=@{bg="#fee2e2"; border="#ef4444"}; warning=@{bg="#fef3c7"; border="#f59e0b"}; success=@{bg="#dcfce7"; border="#22c55e"}; info=@{bg="#e0f2fe"; border="#06b6d4"} }
    $colorSet = $colors[$Type]
    $displayItems = $Items | Select-Object -First $MaxItems
    $itemsHtml = ($displayItems | ForEach-Object { "<div style='margin: 6px 0; padding: 8px 12px; background: rgba(255,255,255,0.8); border-radius: 4px; border-left: 3px solid rgba(0,0,0,0.2); font-size: 0.85rem; line-height: 1.4;'>$_</div>" }) -join ''
    return @"
<div style="display: inline-block; vertical-align: top; width: 280px; margin: 10px; padding: 15px; background-color: $($colorSet.bg); border-radius: 12px; border-left: 6px solid $($colorSet.border); box-shadow: 0 2px 8px rgba(0,0,0,0.1); min-height: 140px;">
  <h4 style="margin: 0 0 12px 0; font-weight: bold; color: #333; font-size: 1rem; line-height: 1.3;">$Title</h4>
  <div style="max-height: 200px; overflow-y: auto;">$itemsHtml</div>
</div>
"@
}

function ConvertTo-FailedActivitiesTable {
    param([array]$FailedActivities, [string]$NinjaOneInstance = "")
    if (-not $FailedActivities -or $FailedActivities.Count -eq 0) { return "<div style='padding: 20px; text-align: center; color: #666;'>No failed activities found</div>" }
    $actualFailures = $FailedActivities | Where-Object { ($_.activityResult -eq "FAILURE") -or ($_.statusCode -in @("FAILED","CANCELLED","BLOCKED","EVALUATION_FAILURE")) } | Sort-Object ActivityTimeConverted -Descending | Select-Object -First 25
    if (-not $actualFailures -or $actualFailures.Count -eq 0) { return "<div style='padding: 20px; text-align: center; color: #666;'>No confirmed failures found</div>" }
    $tableRows = $actualFailures | ForEach-Object {
        $deviceUrl = if ($NinjaOneInstance -and $_.deviceId) { if ($NinjaOneInstance -match "^https?://") { "$NinjaOneInstance/#/deviceDashboard/$($_.deviceId)/overview" } else { "https://$NinjaOneInstance/#/deviceDashboard/$($_.deviceId)/overview" } } else { "#" }
        $deviceName = if ($_.DeviceName) { $_.DeviceName } else { "Unknown Device" }
        $automationName = if ($_.sourceName) { $_.sourceName } else { "Unknown Script" }
        $result = if ($_.activityResult) { $_.activityResult } else { $_.statusCode }
        $orgName = if ($_.OrgName) { $_.OrgName } else { "Unknown Org" }
        $locName = if ($_.LocName) { $_.LocName } else { "—" }
        $deviceLink = if ($deviceUrl -ne "#") { "<a href='$deviceUrl' target='_blank' style='color: #0066cc; text-decoration: none; font-weight: 500;' title='Open device dashboard'>$deviceName</a>" } else { $deviceName }
        $timestamp = if ($_.ActivityTimeConverted) { $_.ActivityTimeConverted.ToString("yyyy-MM-dd HH:mm 'UTC'") } else { "Unknown" }
        "<tr><td>$automationName</td><td>$deviceLink</td><td style='color: #ef4444; font-weight: 600;'>$result</td><td style='font-family: monospace; font-size: 0.85em;'>$timestamp</td><td>$orgName</td><td>$locName</td></tr>"
    }
    return @"
<div style='overflow-x: auto; margin: 15px 0; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);'>
  <table style='min-width: 600px; width: 100%; border-collapse: collapse; background: white; font-size: 0.9rem;'>
    <thead>
      <tr style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;'>
        <th style='padding: 12px 8px; text-align: left; font-weight: 600;'>Automation</th>
        <th style='padding: 12px 8px; text-align: left; font-weight: 600;'>Device</th>
        <th style='padding: 12px 8px; text-align: left; font-weight: 600;'>Result</th>
        <th style='padding: 12px 8px; text-align: left; font-weight: 600;'>Failed At (UTC)</th>
        <th style='padding: 12px 8px; text-align: left; font-weight: 600;'>Organization</th>
        <th style='padding: 12px 8px; text-align: left; font-weight: 600;'>Location</th>
      </tr>
    </thead>
    <tbody>
      $($tableRows -join '')
    </tbody>
  </table>
</div>
<div style='text-align:center; margin-top:10px; color:#666; font-size:0.8rem;'>
  Showing last 25 failures to occur
</div>
"@
}

function ConvertTo-OrganizationSummaryTable {
    param([array]$OrganizationStats, [string]$NinjaOneInstance = "", [hashtable]$OrgKBLinkIndex, [string]$MonthLabel = $(Get-Date -Format 'MMMM yyyy'))
    if (-not $OrganizationStats -or $OrganizationStats.Count -eq 0) { return "<div style='padding:20px; text-align:center; color:#666;'>No organization activity found</div>" }
    $top = $OrganizationStats | Sort-Object Failed -Descending | Select-Object -First 50
    $rows = $top | ForEach-Object {
      $last = if ($_.LastFailure) { $_.LastFailure.ToString("yyyy-MM-dd HH:mm 'UTC'") } else {"Never"}
      $successColor = if ($_.SuccessRate -ge 90) { "#22c55e" } elseif ($_.SuccessRate -ge 75) { "#f59e0b" } else { "#ef4444" }
        $orgName = $_.OrganizationName
        if ([string]::IsNullOrWhiteSpace([string]$orgName)) { $orgName = '(Unknown)' }
        $kbName  = "Automations - $orgName - $MonthLabel"
        $kbUrl = $null
        if ($OrgKBLinkIndex -and $null -ne $orgName -and $OrgKBLinkIndex.ContainsKey($orgName)) { $kbUrl = $OrgKBLinkIndex[$orgName] }
        elseif ($NinjaOneInstance) { $kbUrl = Get-KBSearchUrl -BaseUrl $NinjaOneInstance -ArticleName $kbName }
        else { $kbUrl = "#" }
        $orgLink = "<a href='$kbUrl' target='_blank' style='color:#0066cc; text-decoration:none; font-weight:500;' title='Open organization dashboard'>$orgName</a>"
@"
<tr>
  <td style='padding:8px;'>$orgLink</td>
  <td style='padding:8px; text-align:center;'>$($_.TotalActivities)</td>
  <td style='padding:8px; text-align:center;'>$($_.Successful)</td>
  <td style='padding:8px; text-align:center;'>$($_.Failed)</td>
  <td style='padding:8px; text-align:center; font-weight:600; color:$successColor;'>$($_.SuccessRate)%</td>
  <td style='padding:8px; font-family:monospace; font-size:0.85em;'>$last</td>
</tr>
"@
    }
    return @"
<div style='overflow-x:auto; margin:15px 0; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.1);'>
  <table style='min-width:800px; width:100%; border-collapse:collapse; background:white; font-size:0.9rem;'>
    <thead>
      <tr style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:white;'>
        <th style='padding:12px 8px; text-align:left; font-weight:600;'>Organization</th>
        <th style='padding:12px 8px; text-align:center; font-weight:600;'>Total Runs</th>
        <th style='padding:12px 8px; text-align:center; font-weight:600;'>Success</th>
        <th style='padding:12px 8px; text-align:center; font-weight:600;'>Failed</th>
        <th style='padding:12px 8px; text-align:center; font-weight:600;'>Success %</th>
        <th style='padding:12px 8px; text-align:left; font-weight:600;'>Last Failure (UTC)</th>
      </tr>
    </thead>
    <tbody>
      $($rows -join '')
    </tbody>
  </table>
</div>
<div style='text-align:center; margin-top:10px; color:#666; font-size:0.8rem;'>
  Showing top 50 organizations by failures, then total runs
</div>
"@
}

function ConvertTo-DeviceSummaryTable {
    param([array]$DeviceSummaries, [string]$NinjaOneInstance = "")
    if (-not $DeviceSummaries -or $DeviceSummaries.Count -eq 0) { return "<div style='padding:20px; text-align:center; color:#666;'>No device activity found</div>" }
    $top = $DeviceSummaries | Select-Object -First 50
    $rows = $top | ForEach-Object {
      $last = if ($_.LastRun) { $_.LastRun.ToString("yyyy-MM-dd HH:mm 'UTC'") } else {"Never"}
      $successColor = if ($_.SuccessRate -ge 90) { "#22c55e" } elseif ($_.SuccessRate -ge 75) { "#f59e0b" } else { "#ef4444" }
      $locName = if ($_.LocName) { $_.LocName } else { "—" }
        $deviceUrl = if ($NinjaOneInstance -and $_.DeviceId) { if ($NinjaOneInstance -match "^https?://") { "$NinjaOneInstance/#/deviceDashboard/$($_.DeviceId)/overview" } else { "https://$NinjaOneInstance/#/deviceDashboard/$($_.DeviceId)/overview" } } else { "#" }
        $deviceNameLink = if ($deviceUrl -ne "#") { "<a href='$deviceUrl' target='_blank' style='color:#0066cc; text-decoration:none; font-weight:500;' title='Open device dashboard'>$($_.DeviceName)</a>" } else { $_.DeviceName }
        "<tr>
  <td style='padding:8px;'>$deviceNameLink</td>
  <td style='padding:8px;'>$($_.Type)</td>
  <td style='padding:8px;'>$locName</td>
  <td style='padding:8px; text-align:center;'>$($_.TotalRuns)</td>
  <td style='padding:8px; text-align:center;'>$($_.Success)</td>
  <td style='padding:8px; text-align:center;'>$($_.Failed)</td>
  <td style='padding:8px; text-align:center; font-weight:600; color:$successColor;'>$($_.SuccessRate)%</td>
  <td style='padding:8px; font-family:monospace; font-size:0.85em;'>$last</td>
</tr>"
    }
    return @"
<div style='overflow-x:auto; margin:15px 0; border-radius:8px; box-shadow:0 2px 8px rgba(0,0,0,0.1);'>
  <table style='min-width:800px; width:100%; border-collapse:collapse; background:white; font-size:0.9rem;'>
    <thead>
      <tr style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color:white;'>
        <th style='padding:12px 8px; text-align:left; font-weight:600;'>Device Name</th>
        <th style='padding:12px 8px; text-align:left; font-weight:600;'>Type</th>
        <th style='padding:12px 8px; text-align:left; font-weight:600;'>Location</th>
        <th style='padding:12px 8px; text-align:center; font-weight:600;'>Total Runs</th>
        <th style='padding:12px 8px; text-align:center; font-weight:600;'>Success</th>
        <th style='padding:12px 8px; text-align:center; font-weight:600;'>Failed</th>
        <th style='padding:12px 8px; text-align:center; font-weight:600;'>Success %</th>
        <th style='padding:12px 8px; text-align:left; font-weight:600;'>Last Run (UTC)</th>
      </tr>
    </thead>
    <tbody>
      $($rows -join '')
    </tbody>
  </table>
</div>
<div style='text-align:center; margin-top:10px; color:#666; font-size:0.8rem;'>
  Showing top 50 devices by failures, then total runs
</div>
"@
}

function ConvertTo-AutomationChart { param([array]$AutomationSummary)
    if (-not $AutomationSummary -or $AutomationSummary.Count -eq 0) { return "<div style='padding: 20px; text-align: center; color: #666;'>No automation data available</div>" }
    $worst = $AutomationSummary | Sort-Object SuccessRate | Select-Object -First 5
    $best  = $AutomationSummary | Sort-Object SuccessRate -Descending | Select-Object -First 3
    $worstNames = $worst | Select-Object -ExpandProperty AutomationName
    $additionalBest = $best | Where-Object { $_.AutomationName -notin $worstNames }
    $topItems = @($worst) + @($additionalBest)
    $barsHtml = $topItems | ForEach-Object {
        $c = if ($_.SuccessRate -ge 90) { "#22c55e" } elseif ($_.SuccessRate -ge 75) { "#f59e0b" } else { "#ef4444" }
@"
  <div style="margin-bottom: 12px;">
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
      <span style="font-size: 0.85rem; font-weight: 500; color: #333;" title="$($_.AutomationName)">$($_.AutomationName)</span>
      <span style="font-size: 0.8rem; color: #666;">$($_.TotalRuns) runs ($($_.SuccessRate)%)</span>
    </div>
    <div style="width: 100%; height: 18px; background-color: #f0f0f0; border-radius: 9px; overflow: hidden;">
      <div style="width: $($_.SuccessRate)%; height: 100%; background-color: $c; display: flex; align-items: center; justify-content: center;">
        <span style="color: white; font-size: 0.65rem; font-weight: bold;">$($_.SuccessRate)%</span>
      </div>
    </div>
  </div>
"@
    }
    return @"
<div style="margin: 15px 0; padding: 20px; background: white; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">
$($barsHtml -join '')
  <div style="text-align: center; margin-top: 15px; padding-top: 10px; border-top: 1px solid #eee; color: #666; font-size: 0.8rem;">
    <strong>Automation performance (Top 5 worst and Top 3 best performers)</strong>
  </div>
</div>
"@
}

function ConvertTo-AutomationTable {
    param([array]$Objects, [string]$NinjaOneInstance = "")
    if (-not $Objects -or $Objects.Count -eq 0) { return "<div style='padding: 20px; text-align: center; color: #666;'>No automation data available</div>" }
    $sorted = $Objects | Sort-Object @{Expression='TotalRuns';Descending=$true}, @{Expression='LastRun';Descending=$true}
    $top    = $sorted | Select-Object -First 25
    $rows = $top | ForEach-Object {
        $c = if ($_.SuccessRate -ge 90) { "#22c55e" } elseif ($_.SuccessRate -ge 75) { "#f59e0b" } else { "#ef4444" }
$last = if ($_.LastRun) { $_.LastRun.ToString("yyyy-MM-dd HH:mm 'UTC'") } else { "Never" }
@"
<tr>
  <td style='font-weight: 500; padding: 8px;'>$($_.AutomationName)</td>
  <td style='text-align: center; padding: 8px;'>$($_.AutomationType)</td>
  <td style='text-align: center; padding: 8px;'>$($_.TotalRuns)</td>
  <td style='text-align: center; padding: 8px;'>$($_.Successful)</td>
  <td style='text-align: center; padding: 8px;'>$($_.Failed)</td>
  <td style='text-align: center; font-weight: 600; color: $c; padding: 8px;'>$($_.SuccessRate)%</td>
  <td style='font-family: monospace; font-size: 0.85em; padding: 8px;'>$last</td>
  <td style='text-align: center; padding: 8px;'>$($_.DevicesAffected)</td>
</tr>
"@
    }
    return @"
<div style='overflow-x: auto; margin: 15px 0; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);'>
  <table style='min-width: 800px; width: 100%; border-collapse: collapse; background: white; font-size: 0.9rem;'>
    <thead>
      <tr style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white;'>
        <th style='padding: 12px 8px; text-align: left; font-weight: 600;'>Automation</th>
        <th style='padding: 12px 8px; text-align: center; font-weight: 600;'>Type</th>
        <th style='padding: 12px 8px; text-align: center; font-weight: 600;'>Total Runs</th>
        <th style='padding: 12px 8px; text-align: center; font-weight: 600;'>Success</th>
        <th style='padding: 12px 8px; text-align: center; font-weight: 600;'>Failed</th>
        <th style='padding: 12px 8px; text-align: center; font-weight: 600;'>Success %</th>
        <th style='padding: 12px 8px; text-align: center; font-weight: 600;'>Last Run (UTC)</th>
        <th style='padding: 12px 8px; text-align: center; font-weight: 600;'>Devices</th>
      </tr>
    </thead>
    <tbody>
      $($rows -join '')
    </tbody>
  </table>
</div>
<div style='text-align: center; margin-top: 10px; color: #666; font-size: 0.8rem;'>
  Showing top 25 automations by execution count
</div>
"@
}

function New-AutomationDashboardHTML {
    param(
        [string]$Title, [string]$OrganizationName, [string]$KPICards, [string]$AlertCards,
        [array]$AutomationSummary, [array]$FailedActivities, [array]$SuccessfulActivities,
        [int]$UniqueDeviceCount, [string]$NinjaOneInstance = "", [array]$DeviceSummary,
        [hashtable]$OrgKBLinkIndex, $GlobalAnalysis, [string]$MonthLabel = $(Get-Date -Format 'MMMM yyyy')
    )
    $automationChart = ConvertTo-AutomationChart -AutomationSummary $AutomationSummary
    $pieChart        = New-PieChart -SuccessfulCount $SuccessfulActivities.Count -FailedCount $FailedActivities.Count
    $failedActivitiesTable = ConvertTo-FailedActivitiesTable -FailedActivities $FailedActivities -NinjaOneInstance $NinjaOneInstance
    $automationTable       = ConvertTo-AutomationTable   -Objects $AutomationSummary -NinjaOneInstance $NinjaOneInstance
    $organizationSummaryTable = ConvertTo-OrganizationSummaryTable -OrganizationStats $GlobalAnalysis.OrganizationStats -NinjaOneInstance $NinjaOneInstance -OrgKBLinkIndex $OrgKBLinkIndex -MonthLabel $MonthLabel
    $deviceSummaryTable    = ConvertTo-DeviceSummaryTable -DeviceSummaries $DeviceSummary -NinjaOneInstance $NinjaOneInstance

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>$Title - $OrganizationName</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .dashboard-container { max-width: 1400px; margin: 0 auto; background: #f8f9fa; padding: 20px; line-height: 1.5; }
    .section { background: white; padding: 25px; border-radius: 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); margin-bottom: 30px; }
    .kpi-grid { text-align: center; max-width: 800px; margin: 0 auto; }
    .alert-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px; justify-content: center; max-width: 900px; margin: 0 auto; }
    h2 { margin: 0 0 20px 0; color: #333; font-size: 1.5rem; }
    .footer { text-align: center; margin-top: 30px; padding: 20px; background: #e9ecef; border-radius: 10px; color: #6c757d; }
    @media (max-width: 768px) { .dashboard-container { padding: 10px; } .section { padding: 15px; } }
  </style>
</head>
<body>
  <div class="dashboard-container">
    <div class="section">
      <h2>Key Performance Indicators</h2>
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;" class="kpi-grid">$KPICards</div>
    </div>
    <div class="section">
      <h2>Overall Success Rate</h2>
      $pieChart
    </div>
    <div class="section">
      <h2>Attention Required</h2>
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;" class="alert-grid">$AlertCards</div>
    </div>
    <div class="section">
      <h2>Automation Performance</h2>
      $automationChart
    </div>
    <div class="section">
      <h2>Automation Summary</h2>
      $automationTable
    </div>
    $(if ($OrganizationName -eq 'All Organizations') {
      "<div class='section'>
        <h2>Organization Summary</h2>
        $organizationSummaryTable
      </div>"
     } else { '' })
    $(if ($OrganizationName -ne 'All Organizations') {
      "<div class='section'>
        <h2>Device Summary</h2>
        $deviceSummaryTable
      </div>"
     } else { '' })
    <div class="section">
      <h2>Recent Failed Activities</h2>
      $failedActivitiesTable
    </div>
    <div class="footer">
      <strong>Dashboard generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</strong><br>
      Monitoring $UniqueDeviceCount managed devices
    </div>
  </div>
</body>
</html>
"@
}

function Set-KnowledgeBaseArticles {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()] [array]$Articles,
        [Parameter(Mandatory = $true)][ValidateSet("create", "update")] [string]$Operation,
        [int]$BatchSize = 50,
        [switch]$Trace
    )
    $results = @()  # Always return collected results
    $isUpdate = $Operation -eq "update"
    $method   = if ($isUpdate) { "PATCH" } else { "POST" }

    if ($isUpdate) {
        foreach ($a in $Articles) {
            $idVal = 0
            if ($a -and $a.PSObject.Properties['id']) { try { $idVal = [long]$a.id } catch { $idVal = 0 } }
            if ($idVal -le 0) { throw "KB update requested but no valid 'id' was provided for article '$($a.name)'. The article likely does not exist yet; call with -Operation 'create' instead." }
        }
    }

    $payload = if ($isUpdate) {
        $Articles | ForEach-Object {
            $o = [PSCustomObject]@{ id = [long]$_.id; name = $_.name; content = $_.content }
            foreach ($p in @($o.PSObject.Properties)) { if ($null -eq $p.Value -or ($p.Value -is [string] -and [string]::IsNullOrWhiteSpace($p.Value))) { $o.PSObject.Properties.Remove($p.Name) } }
            if ($o.content -and $o.content.PSObject.Properties["html"]) { $o.content.html = [string]$o.content.html }
            $o
        }
    } else {
        $Articles | ForEach-Object {
            $o = [PSCustomObject]@{ name = $_.name; destinationFolderPath = $_.destinationFolderPath; content = $_.content; organizationId = $_.organizationId }
            foreach ($p in @($o.PSObject.Properties)) { if ($null -eq $p.Value -or ($p.Value -is [string] -and [string]::IsNullOrWhiteSpace($p.Value))) { $o.PSObject.Properties.Remove($p.Name) } }
            if ($o.content -and $o.content.PSObject.Properties["html"]) { $o.content.html = [string]$o.content.html }
            $o
        }
    }

    if ($payload.Count -eq 1) {
        $single = @($payload[0])
        try {
            $resp = Invoke-NinjaOneRequest -Path 'knowledgebase/articles' -Method $method -InputObject $single -AsArray
            if ($Trace) { $name = $single[0].name; $id = ($resp | Select-Object -First 1).id; $org = ($resp | Select-Object -First 1).organizationId; Write-Log ("KB {0}: name='{1}' id={2} orgId={3}" -f $Operation, $name, (if($id){$id}else{'<unknown>'}), (if($org){$org}else{'<global>'})) }
            if ($resp) { $results += $resp } else { $results += $single }
            Write-Log "Successfully ${Operation}d 1 KB article."
        } catch {
            Start-Sleep -Milliseconds 600
            try {
                $resp = Invoke-NinjaOneRequest -Path 'knowledgebase/articles' -Method $method -InputObject $single -AsArray
                if ($Trace) { $name = $single[0].name; $id = ($resp | Select-Object -First 1).id; $org = ($resp | Select-Object -First 1).organizationId; Write-Log ("KB {0} (retry): name='{1}' id={2} orgId={3}" -f $Operation, $name, (if($id){$id}else{'<unknown>'}), (if($org){$org}else{'<global>'})) }
                if ($resp) { $results += $resp } else { $results += $single }
                Write-Log "Recovered from transient API error; ${Operation} succeeded on retry."
            } catch {
                Write-Log ("KB $Operation failed for article id=$($single[0].id), name='$($single[0].name)': $($_.Exception.Message)") "Error"
                throw
            }
        }
        return $results
    }

    for ($i = 0; $i -lt $payload.Count; $i += $BatchSize) {
        $batchEnd = [Math]::Min($i + $BatchSize - 1, $payload.Count - 1)
        $batch = @($payload[$i..$batchEnd])
        $bulkOk = $true
        try {
            $resp = Invoke-NinjaOneRequest -Path 'knowledgebase/articles' -Method $method -InputObject $batch -AsArray
            if ($Trace) { foreach ($item in @($resp)) { Write-Log ("KB {0}: name='{1}' id={2} orgId={3}" -f $Operation, $item.name, $item.id, (if($item.organizationId){$item.organizationId}else{'<global>'})) } }
            if ($resp) { $results += $resp } else { $results += $batch }
        } catch {
            $bulkOk = $false
            Write-Log "Bulk $Operation encountered a transient API error for items $i..$batchEnd; retrying individually..." "Warning"
        }
        if (-not $bulkOk) {
            $failed = @()
            for ($j = $i; $j -le $batchEnd; $j++) {
                $single = @($payload[$j])
                try {
                    $resp = Invoke-NinjaOneRequest -Path 'knowledgebase/articles' -Method $method -InputObject $single -AsArray
                    if ($Trace) { $id = ($resp | Select-Object -First 1).id; $org = ($resp | Select-Object -First 1).organizationId; Write-Log ("KB {0} (fallback): name='{1}' id={2} orgId={3}" -f $Operation, $single[0].name, (if($id){$id}else{'<unknown>'}), (if($org){$org}else{'<global>'})) }
                    if ($resp) { $results += $resp } else { $results += $single }
                } catch { $failed += ,@{ id = $single[0].id; name = $single[0].name; err = $_.Exception.Message } }
            }
            if ($failed.Count -gt 0) {
                $msg = ($failed | ForEach-Object { "id=$($_.id), name='$($_.name)': $($_.err)" }) | Out-String
                Write-Log "Some KB items failed to ${Operation}:`n${msg}" "Error"
                throw "KB $Operation failed for $($failed.Count) item(s)."
            } else { Write-Log "Recovered from bulk API error; ${Operation} fallback succeeded for items $i..$batchEnd." }
        }
    }

    Write-Log "Successfully ${Operation}d $($Articles.Count) KB articles."
    if ($Trace -and $results.Count -gt 0) {
        $byScope = @($results | Group-Object { if ($_.organizationId) { "org:$($_.organizationId)" } else { 'global' } })
        foreach ($g in $byScope) { $names = ($g.Group | ForEach-Object { "${($_.name)} [id=$($_.id)]" }) -join '; '; Write-Log ("KB {0} summary → {1} item(s): {2}" -f $Operation, $g.Count, $names) }
    }
    return $results
}

function New-OrganizationDashboard {
    param(
        [Parameter(Mandatory = $true)] $Organization,
        [Parameter(Mandatory = $true)][array]$OrgActivities,
        [Parameter(Mandatory = $true)] $AutomationData,
        $GlobalAnalysis,
        [Parameter(Mandatory = $true)][string]$NinjaInstance,
        [Parameter(Mandatory = $true)] [string]$MonthLabel
    )
    if ($GlobalAnalysis -and $OrgActivities.Count -eq $AutomationData.Activities.Count) {
        Write-Log "Reusing global analysis for organization '$($Organization.name)' (contains all activities)"
        $analysis = $GlobalAnalysis
    } else {
        $orgAutomationData = @{ Devices=@($AutomationData.Devices | Where-Object { $_.organizationId -eq $Organization.id }); Organizations=@($Organization); DeviceIndex=$AutomationData.DeviceIndex; OrgIndex=$AutomationData.OrgIndex; Activities=$OrgActivities }
        $analysis = Get-AutomationAnalysis -AutomationData $orgAutomationData -UseCache $true
    }

    $kpiCards = @(
        (New-DashboardKPICard -Title "Total Automations" -Value $analysis.OverallStats.TotalActivities -Type "info"),
        (New-DashboardKPICard -Title "Success Rate" -Value "$($analysis.OverallStats.SuccessRate)%" -Type $(if ($analysis.OverallStats.SuccessRate -ge 90) { "success" } elseif ($analysis.OverallStats.SuccessRate -ge 75) { "warning" } else { "danger" })),
        (New-DashboardKPICard -Title "Failed Activities" -Value $analysis.FailedActivities.Count -Type $(if ($analysis.FailedActivities.Count -eq 0) { "success" } else { "danger" })),
        (New-DashboardKPICard -Title "Unique Automations" -Value $analysis.AutomationSummary.Count -Type "primary")
    ) -join ''

    $failedDeviceItems = if ($analysis.DeviceFailures.Count -gt 0) { ($analysis.DeviceFailures | Select-Object -First 5 | ForEach-Object { "$($_.DeviceName): $($_.FailureCount) failures" }) } else { @("No devices with failures") }
    $lowSuccessItems   = if ($analysis.AutomationSummary | Where-Object { $_.SuccessRate -lt 75 }) { ($analysis.AutomationSummary | Where-Object { $_.SuccessRate -lt 75 } | Sort-Object SuccessRate | Select-Object -First 5 | ForEach-Object { "$($_.AutomationName): $($_.SuccessRate)% success" }) } else { @("All automations performing well") }
    $topRunItems       = if ($analysis.AutomationSummary.Count -gt 0) { ($analysis.AutomationSummary | Sort-Object TotalRuns -Descending | Select-Object -First 5 | ForEach-Object { "$($_.AutomationName): $($_.TotalRuns) runs" }) } else { @("No automation data") }

    $alertCards = @(
        (New-AlertCard -Title "Devices with Failures" -Items $failedDeviceItems -Type "error" -MaxItems 3),
        (New-AlertCard -Title "Low Success Rate Scripts" -Items $lowSuccessItems -Type "warning" -MaxItems 3),
        (New-AlertCard -Title "Most Frequently Run Automations" -Items $topRunItems -Type "info" -MaxItems 3)
    ) -join ''

    # Page title (optional, but nice to align)
    $dashboardTitle = "Automations - $MonthLabel"

    $uniqueDeviceCount = if ($analysis.OverallStats.TotalActivities -gt 0) { ($OrgActivities | Select-Object -ExpandProperty deviceId -Unique).Count } else { 0 }
    $deviceSummary = Get-DeviceSummary -Activities $OrgActivities

    $dashboardHtml = New-AutomationDashboardHTML -Title $dashboardTitle -OrganizationName $Organization.name -KPICards $kpiCards -AlertCards $alertCards -AutomationSummary $analysis.AutomationSummary -FailedActivities $analysis.FailedActivities -SuccessfulActivities $analysis.SuccessfulActivities -UniqueDeviceCount $uniqueDeviceCount -NinjaOneInstance $NinjaInstance -DeviceSummary $deviceSummary -GlobalAnalysis $GlobalAnalysis -MonthLabel $MonthLabel

    $orgNameClean = if ($Organization -and $Organization.PSObject.Properties['name']) { [string]$Organization.name } else { "" }
    $orgNameClean = $orgNameClean.Trim()
    if ([string]::IsNullOrWhiteSpace($orgNameClean)) { throw "Organization name is empty; cannot build article name." }
    # Per-org article name (desired pattern)
    $articleName = "Automations - $($Organization.name) - $MonthLabel"

    return @{ name = $articleName; destinationFolderPath = "Monthly Reports"; content = @{ html = [string]$dashboardHtml } }
}
#endregion

#region Main Execution
function Start-AutomationDashboard {
    param(
        [string]$ReportMonth = "",
        [switch]$EnableKBTrace,
        [string]$URLOverride = [System.Convert]::ToString($env:URLOverride),
        [string]$DbPath = 'C:\RMM\Activities.db',
        [string]$SqliteExePath = ''
    )
    $startTime = Get-Date
    try {
        Write-Log "=== NinjaOne Automation Dashboard Starting ==="
        $VerbosePreference = 'SilentlyContinue'; $InformationPreference = 'SilentlyContinue'; $ProgressPreference = 'SilentlyContinue'
        if (-not $EnableKBTrace) { $EnableKBTrace = ($env:KB_TRACE -eq '1') }

        Initialize-PowerShell7
        Initialize-NinjaOneModule

        $scriptDir = $PSScriptRoot
        if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
        $sqliteExe = $null
        if (-not [string]::IsNullOrWhiteSpace($SqliteExePath) -and (Test-Path -LiteralPath $SqliteExePath -PathType Leaf)) { $sqliteExe = $SqliteExePath }
        if (-not $sqliteExe -and $scriptDir) { $candidate = Join-Path $scriptDir 'sqlite3.exe'; if (Test-Path -LiteralPath $candidate -PathType Leaf) { $sqliteExe = $candidate } }
        if (-not $sqliteExe) { $cmd = Get-Command sqlite3 -ErrorAction SilentlyContinue; if ($cmd -and $cmd.Source -and (Test-Path -LiteralPath $cmd.Source -PathType Leaf)) { $sqliteExe = $cmd.Source } }
        if (-not $sqliteExe -and (Test-Path -LiteralPath 'C:\RMM\sqlite3.exe' -PathType Leaf)) { $sqliteExe = 'C:\RMM\sqlite3.exe' }
        if (-not $sqliteExe) { throw "sqlite3.exe not found. Place it in the script directory, add to PATH, set -SqliteExePath, or install to C:\RMM\sqlite3.exe. See Get-AutomationActivities.ps1 for download link." }

        $resolvedDbPath = if ([string]::IsNullOrWhiteSpace($DbPath)) { 'C:\RMM\Activities.db' } else { $DbPath }

        # Credentials: Ninja-Property-Get (when in NinjaOne) -> env vars -> parameters
        try { $fromNinja = Ninja-Property-Get ninjaoneInstance; if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $global:NinjaOneInstance = $fromNinja } } catch { }
        if ([string]::IsNullOrWhiteSpace($global:NinjaOneInstance)) { $global:NinjaOneInstance = $env:NINJAONE_INSTANCE }
        if ([string]::IsNullOrWhiteSpace($global:NinjaOneInstance)) { $global:NinjaOneInstance = $NinjaOneInstance }
        try { $fromNinja = Ninja-Property-Get ninjaoneClientId; if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $global:NinjaOneClientId = $fromNinja } } catch { }
        if ([string]::IsNullOrWhiteSpace($global:NinjaOneClientId)) { $global:NinjaOneClientId = $env:NINJAONE_CLIENT_ID }
        if ([string]::IsNullOrWhiteSpace($global:NinjaOneClientId)) { $global:NinjaOneClientId = $NinjaOneClientId }
        try { $fromNinja = Ninja-Property-Get ninjaoneClientSecret; if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $global:NinjaOneClientSecret = $fromNinja } } catch { }
        if ([string]::IsNullOrWhiteSpace($global:NinjaOneClientSecret)) { $global:NinjaOneClientSecret = $env:NINJAONE_CLIENT_SECRET }
        if ([string]::IsNullOrWhiteSpace($global:NinjaOneClientSecret)) { $global:NinjaOneClientSecret = $NinjaOneClientSecret }
        if (!$global:NinjaOneInstance -or !$global:NinjaOneClientId -or !$global:NinjaOneClientSecret) {
            throw "Missing required API credentials. Set in NinjaOne custom properties, or use env vars NINJAONE_INSTANCE, NINJAONE_CLIENT_ID, NINJAONE_CLIENT_SECRET, or -NinjaOneInstance, -NinjaOneClientId, -NinjaOneClientSecret."
        }

        Connect-NinjaOne -NinjaOneInstance $global:NinjaOneInstance -NinjaOneClientID $global:NinjaOneClientId -NinjaOneClientSecret $global:NinjaOneClientSecret
        Write-Log "Connected to NinjaOne API successfully"
        Write-Log "URL $global:NinjaOneInstance"

        $ninjaInstance = [string]$global:NinjaOneInstance
        $effectiveBaseUrl = if ($URLOverride -and -not [string]::IsNullOrWhiteSpace($URLOverride)) { $URLOverride } else { $ninjaInstance }
        $effectiveBaseUrl = if ($effectiveBaseUrl -match '^https?://') { $effectiveBaseUrl } else { "https://$effectiveBaseUrl" }
        $effectiveBaseUrl = $effectiveBaseUrl -replace '/+$',''
        Write-Log ("Using base URL for hyperlinks: {0}" -f $effectiveBaseUrl)

        $dateRange = Get-DateRange -ReportMonth $ReportMonth
        $monthLabel = '{0} {1}' -f $dateRange.Current.Month, $dateRange.Current.Year
        Write-Log ("Generating dashboard for: {0} {1}" -f $dateRange.Current.Month, $dateRange.Current.Year)

        $data = Get-NinjaOneData -DateRange $dateRange -DbPath $resolvedDbPath -SqliteExe $sqliteExe
        if ($data.Activities.Count -eq 0) { Write-Log "No automation activities found for the specified period" "Warning"; return }

        $globalAnalysis = Get-AutomationAnalysis -AutomationData $data -UseCache $true

        $orgsProcessed = 0
        $script:OrgKBLinkIndex = @{}
        foreach ($organization in $data.Organizations) {
            try {
                $orgActivities = @($data.Activities | Where-Object { $_.OrgID -eq $organization.id })
                if ($orgActivities.Count -le 0) { Write-Log ("No activities found for organization '{0}'" -f $organization.name) "Warning"; continue }

                $orgDashboard = New-OrganizationDashboard -Organization $organization -OrgActivities $orgActivities -AutomationData $data -GlobalAnalysis $globalAnalysis -NinjaInstance $effectiveBaseUrl -MonthLabel $monthLabel

                $articleName = $orgDashboard.name
                $articlePayload = @{ name = $articleName; destinationFolderPath = $orgDashboard.destinationFolderPath; content = $orgDashboard.content }

                $existing = Get-GlobalKBArticleByName -Name $articleName
                $hasValidId = ($existing -and $existing.PSObject.Properties['id'] -and ([long]$existing.id) -gt 0)

                if ($hasValidId) {
                    $updatePayload = [pscustomobject]@{ id = [long]$existing.id; name = $articlePayload.name; content = $articlePayload.content }
                    $apiResult = Set-KnowledgeBaseArticles -Articles @($updatePayload) -Operation 'update' -Trace:$EnableKBTrace
                    Write-Log ("Updated GLOBAL KB article '{0}' (id={1})" -f $articleName, ($apiResult | Select-Object -First 1).id)
                } else {
                    $apiResult = Set-KnowledgeBaseArticles -Articles @($articlePayload) -Operation 'create' -Trace:$EnableKBTrace
                    Write-Log ("Created GLOBAL KB article '{0}' (id={1})" -f $articleName, ($apiResult | Select-Object -First 1).id)
                }

                $articleForLink = $apiResult | Select-Object -First 1
                if (-not $articleForLink -or -not $articleForLink.parentFolderId) { $articleForLink = Get-GlobalKBArticleByName -Name $articleName }
                if ($null -ne $articleForLink) { $deepLink = Build-KBDeepLink -Article $articleForLink -BaseUrl $effectiveBaseUrl }
                else { Write-Log ("Deep link resolution failed for article name '{0}' (no article found). Using search URL fallback." -f $articleName) 'Warning'; $deepLink = $null }

                if ($deepLink) { $script:OrgKBLinkIndex[[string]$organization.name] = [string]$deepLink }
                else { $script:OrgKBLinkIndex[[string]$organization.name] = Get-KBSearchUrl -BaseUrl $effectiveBaseUrl -ArticleName ($articleName.Trim()) }

                $orgsProcessed++
                Write-Log ("Generated dashboard for: {0}" -f $organization.name)
            } catch {
                Write-Log ("Failed to process organization '{0}': {1}" -f $organization.name, $_.Exception.Message) "Error"
            }
        }

        Write-Log "Generating global dashboard..."
        $failedDevicesGlobal = if ($globalAnalysis.DeviceFailures.Count -gt 0) { ($globalAnalysis.DeviceFailures | Select-Object -First 5 | ForEach-Object { "{0} ({1}): {2} failures" -f $_.DeviceName, $_.OrgName, $_.FailureCount }) } else { @("No device failures") }
        $lowSuccessGlobal    = if ($globalAnalysis.AutomationSummary | Where-Object { $_.SuccessRate -lt 75 }) { ($globalAnalysis.AutomationSummary | Where-Object { $_.SuccessRate -lt 75 } | Sort-Object SuccessRate | Select-Object -First 5 | ForEach-Object { "{0}: {1}% success" -f $_.AutomationName, $_.SuccessRate }) } else { @("All automations performing well") }
        $orgsNeedingAttention= if ($globalAnalysis.OrganizationStats | Where-Object { $_.SuccessRate -lt 85 }) { ($globalAnalysis.OrganizationStats | Where-Object { $_.SuccessRate -lt 85 } | Sort-Object SuccessRate | Select-Object -First 5 | ForEach-Object { "{0}: {1}% success" -f $_.OrganizationName, $_.SuccessRate }) } else { @("All organizations performing well") }

        $globalKpiCards = @(
            (New-DashboardKPICard -Title "Total Automations" -Value $globalAnalysis.OverallStats.TotalActivities -Type "info"),
            (New-DashboardKPICard -Title "Success Rate" -Value ("{0}%" -f $globalAnalysis.OverallStats.SuccessRate) -Type $(if ($globalAnalysis.OverallStats.SuccessRate -ge 90) { "success" } elseif ($globalAnalysis.OverallStats.SuccessRate -ge 75) { "warning" } else { "danger" })),
            (New-DashboardKPICard -Title "Failed Activities" -Value $globalAnalysis.FailedActivities.Count -Type $(if ($globalAnalysis.FailedActivities.Count -eq 0) { "success" } else { "danger" })),
            (New-DashboardKPICard -Title "Organizations" -Value $data.Organizations.Count -Type "primary")
        ) -join ''

        $globalAlertCards = @(
            (New-AlertCard -Title "Top Device Failures" -Items $failedDevicesGlobal -Type "error" -MaxItems 3),
            (New-AlertCard -Title "Low Success Scripts"   -Items $lowSuccessGlobal  -Type "warning" -MaxItems 3),
            (New-AlertCard -Title "Organizations Needing Attention" -Items $orgsNeedingAttention -Type "info" -MaxItems 3)
        ) -join ''

        $globalUniqueDeviceCount = if ($data.Activities.Count -gt 0) { ($data.Activities | Select-Object -ExpandProperty deviceId -Unique).Count } else { 0 }
        $globalDeviceSummary = Get-DeviceSummary -Activities $data.Activities
        $globalDashboardHtml = New-AutomationDashboardHTML -Title ("Automations - {0}" -f $monthLabel) -OrganizationName "All Organizations" -KPICards $globalKpiCards -AlertCards $globalAlertCards -AutomationSummary $globalAnalysis.AutomationSummary -FailedActivities $globalAnalysis.FailedActivities -SuccessfulActivities $globalAnalysis.SuccessfulActivities -UniqueDeviceCount $globalUniqueDeviceCount -NinjaOneInstance $effectiveBaseUrl -DeviceSummary $globalDeviceSummary -OrgKBLinkIndex $script:OrgKBLinkIndex -GlobalAnalysis $globalAnalysis -MonthLabel $monthLabel

        # Global KB article name (desired pattern)
        $globalArticleName = "Automations - All Organizations - $monthLabel"
        try {
            $globalExistingArticle = Get-GlobalKBArticleByName -Name $globalArticleName
            $globalArticleData = @{ name = $globalArticleName; destinationFolderPath = "Monthly Reports"; content = @{ html = [string]$globalDashboardHtml } }
            $hasValidId = ($globalExistingArticle -and $globalExistingArticle.PSObject.Properties['id'] -and [long]$globalExistingArticle.id -gt 0)
            if ($hasValidId) { $globalUpdateData = [pscustomobject]@{ id = [long]$globalExistingArticle.id; name = $globalArticleData.name; content = $globalArticleData.content }; $null = Set-KnowledgeBaseArticles -Articles @($globalUpdateData) -Operation "update" -Trace:$EnableKBTrace }
            else { $null = Set-KnowledgeBaseArticles -Articles @($globalArticleData) -Operation "create" -Trace:$EnableKBTrace }
        } catch { Write-Log ("Failed to process global KB article: {0}" -f $_.Exception.Message) "Error" }

        $runtime = (New-TimeSpan -Start $startTime -End (Get-Date)).TotalSeconds
        Write-Log "=== Dashboard Generation Complete ==="
        Write-Log ("Activities analyzed: {0}" -f $globalAnalysis.OverallStats.TotalActivities)
        Write-Log ("Success rate: {0}%" -f $globalAnalysis.OverallStats.SuccessRate)
        Write-Log ("Organizations processed: {0}" -f $orgsProcessed)
        Write-Log ("Runtime: {0} seconds" -f ([Math]::Round($runtime, 1)))
    } catch { Write-Log ("Dashboard generation failed: {0}" -f $_.Exception.Message) "Error"; exit 1 }
}

$script:AnalysisCache = @{}
$null = Start-AutomationDashboard -ReportMonth $ReportMonth -URLOverride $URLOverride -DbPath $DbPath -SqliteExePath $SqliteExePath
#endregion

<#
.SYNOPSIS
    Checks whether the device's current local time is within or outside a maintenance window (Daily, Weekly, or Monthly).

.DESCRIPTION
    Compares the device's current time to a recurring maintenance window. Maintenance window values are read from
    NinjaOne custom fields, set by Set-WeeklyMaintenanceWindow.ps1. NinjaOne injects device/org/location custom
    fields as environment variables. Supports Daily (every day at same time), Weekly (specific day of week), and
    Monthly (nth weekday of month, e.g. 2nd Tuesday). Supports HH:mm, seconds-from-midnight-UTC (0-86400), and
    Unix milliseconds. Handles overnight windows. Intended for use with NinjaOne compound conditions.

    Environment variables (from NinjaOne custom fields):
    - maintenanceRecurrence : Optional. Daily | Weekly | Monthly. Default: Weekly.
    - maintenanceDay        : Day of week (Weekly/Monthly). Not used for Daily.
    - maintenanceOccurrence : For Monthly only: 1, 2, 3, 4, or Last.
    - maintenanceStart      : Start time as HH:mm, seconds from midnight UTC (0-86400), or Unix milliseconds.
    - maintenanceEnd        : End time as HH:mm, seconds from midnight UTC (0-86400), or Unix milliseconds.
    - exitWhenInside        : Optional. "true"/"1" = exit 0 when inside window (default); "false"/"0" = exit 0 when outside.

.EXAMPLE
    # Run on device; NinjaOne supplies maintenance window custom fields.
    .\Check-WeeklyMaintenanceWindow.ps1
#>

$ErrorActionPreference = 'Stop'

# Read maintenance window from NinjaOne custom fields (injected as environment variables)
$MaintenanceRecurrence = $env:maintenanceRecurrence
$MaintenanceOccurrence  = $env:maintenanceOccurrence
try { $v = Get-NinjaProperty maintenanceRecurrence; if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) { $MaintenanceRecurrence = $v } } catch { }
try { $v = Get-NinjaProperty maintenanceOccurrence;  if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) { $MaintenanceOccurrence  = $v } } catch { }
$MaintenanceDay = Get-NinjaProperty maintenanceDay -type dropdown
$MaintenanceStart        = Get-NinjaProperty maintenanceStart
$MaintenanceEnd          = Get-NinjaProperty maintenanceEnd

$ExitWhenInside = $true
if ($env:exitWhenInside -ne $null) {
    $ev = $env:exitWhenInside -as [string]
    if ($ev -match '^(?i)(true|1|yes)$') { $ExitWhenInside = $true }
    elseif ($ev -match '^(?i)(false|0|no)$') { $ExitWhenInside = $false }
}

# Normalize recurrence: Daily | Weekly | Monthly (default Weekly)
if ([string]::IsNullOrWhiteSpace($MaintenanceRecurrence)) { $MaintenanceRecurrence = 'Weekly' }
else {
    $MaintenanceRecurrence = $MaintenanceRecurrence.Trim().ToLowerInvariant()
    if ($MaintenanceRecurrence -eq 'daily') { $MaintenanceRecurrence = 'Daily' }
    elseif ($MaintenanceRecurrence -eq 'weekly') { $MaintenanceRecurrence = 'Weekly' }
    elseif ($MaintenanceRecurrence -eq 'monthly') { $MaintenanceRecurrence = 'Monthly' }
    else { $MaintenanceRecurrence = 'Weekly' }
}

#region Validation

if ([string]::IsNullOrWhiteSpace($MaintenanceStart)) {
    Write-Error "MaintenanceStart is required. Set the maintenanceStart custom field in NinjaOne (via Set-WeeklyMaintenanceWindow.ps1 or the NinjaOne UI)."
    exit 2
}
if ([string]::IsNullOrWhiteSpace($MaintenanceEnd)) {
    Write-Error "MaintenanceEnd is required. Set the maintenanceEnd custom field in NinjaOne (via Set-WeeklyMaintenanceWindow.ps1 or the NinjaOne UI)."
    exit 2
}
if ($MaintenanceRecurrence -eq 'Weekly' -or $MaintenanceRecurrence -eq 'Monthly') {
    if ([string]::IsNullOrWhiteSpace($MaintenanceDay)) {
        Write-Error "MaintenanceDay is required for $MaintenanceRecurrence recurrence. Set the maintenanceDay custom field in NinjaOne."
        exit 2
    }
}
if ($MaintenanceRecurrence -eq 'Monthly') {
    if ([string]::IsNullOrWhiteSpace($MaintenanceOccurrence)) {
        Write-Error "MaintenanceOccurrence is required for Monthly recurrence. Use 1, 2, 3, 4, or Last."
        exit 2
    }
    $occNorm = $MaintenanceOccurrence.Trim().ToLowerInvariant()
    $validOcc = ($occNorm -eq 'last') -or ($occNorm -match '^[1-4]$')
    if (-not $validOcc) {
        Write-Error "Invalid MaintenanceOccurrence: '$MaintenanceOccurrence'. Use 1, 2, 3, 4, or Last."
        exit 2
    }
}

#endregion

#region Parse maintenance window (Format A: HH:mm, Format B: seconds from midnight UTC, Format C: Unix ms)

function Get-DayOfWeekFromName {
    param([string]$DayName)
    $d = ($DayName -as [string]).Trim()
    switch ($d) {
        'Sunday'    { return [System.DayOfWeek]::Sunday }
        'Monday'    { return [System.DayOfWeek]::Monday }
        'Tuesday'   { return [System.DayOfWeek]::Tuesday }
        'Wednesday' { return [System.DayOfWeek]::Wednesday }
        'Thursday'  { return [System.DayOfWeek]::Thursday }
        'Friday'    { return [System.DayOfWeek]::Friday }
        'Saturday'  { return [System.DayOfWeek]::Saturday }
        default     { return $null }
    }
}

$targetDayOfWeek = $null
$windowStartTS = $null
$windowEndTS = $null
$monthlyOccurrenceNorm = $null
if ($MaintenanceRecurrence -eq 'Monthly') {
    $mo = ($MaintenanceOccurrence -as [string]).Trim().ToLowerInvariant()
    $monthlyOccurrenceNorm = if ($mo -eq 'last') { 'last' } else { $mo }
}

# Format A: HH:mm
if ($MaintenanceStart -match '^(\d{1,2}):(\d{2})$' -and $MaintenanceEnd -match '^(\d{1,2}):(\d{2})$') {
    $null = $MaintenanceStart -match '^(\d{1,2}):(\d{2})$'
    $h1 = [int]$Matches[1]
    $m1 = [int]$Matches[2]
    if ($h1 -lt 0 -or $h1 -gt 23 -or $m1 -lt 0 -or $m1 -gt 59) {
        Write-Error "Invalid MaintenanceStart time: '$MaintenanceStart'. Use HH:mm (e.g. 02:00)."
        exit 2
    }
    $windowStartTS = [TimeSpan]::new($h1, $m1, 0)
    $null = $MaintenanceEnd -match '^(\d{1,2}):(\d{2})$'
    $h2 = [int]$Matches[1]
    $m2 = [int]$Matches[2]
    if ($h2 -lt 0 -or $h2 -gt 23 -or $m2 -lt 0 -or $m2 -gt 59) {
        Write-Error "Invalid MaintenanceEnd time: '$MaintenanceEnd'. Use HH:mm (e.g. 04:00)."
        exit 2
    }
    $windowEndTS = [TimeSpan]::new($h2, $m2, 0)
    if ($MaintenanceRecurrence -ne 'Daily') {
        $targetDayOfWeek = Get-DayOfWeekFromName -DayName $MaintenanceDay
        if ($null -eq $targetDayOfWeek) {
            Write-Error "Invalid MaintenanceDay: '$MaintenanceDay'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."
            exit 2
        }
    }
}
# Format B: Seconds from midnight UTC (0-86400). NinjaOne "Time" custom fields use this.
elseif ($MaintenanceStart -match '^\d+$' -and $MaintenanceEnd -match '^\d+$' `
    -and [long]$MaintenanceStart -ge 0 -and [long]$MaintenanceStart -le 86400 `
    -and [long]$MaintenanceEnd -ge 0 -and [long]$MaintenanceEnd -le 86400) {
    $utcDate = [DateTime]::UtcNow.Date
    $startLocal = [TimeZoneInfo]::ConvertTimeFromUtc($utcDate.AddSeconds([int]$MaintenanceStart), [TimeZoneInfo]::Local)
    $endLocal   = [TimeZoneInfo]::ConvertTimeFromUtc($utcDate.AddSeconds([int]$MaintenanceEnd), [TimeZoneInfo]::Local)
    $windowStartTS = $startLocal.TimeOfDay
    $windowEndTS   = $endLocal.TimeOfDay
    if ($MaintenanceRecurrence -ne 'Daily') {
        $targetDayOfWeek = Get-DayOfWeekFromName -DayName $MaintenanceDay
        if ($null -eq $targetDayOfWeek) {
            Write-Error "Invalid MaintenanceDay: '$MaintenanceDay'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."
            exit 2
        }
    }
}
# Format C: Unix milliseconds (e.g. from Set-WeeklyMaintenanceWindow.ps1 import)
elseif ($MaintenanceStart -match '^\d+$' -and $MaintenanceEnd -match '^\d+$') {
    try {
        $startDt = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$MaintenanceStart).LocalDateTime
        $endDt   = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$MaintenanceEnd).LocalDateTime
    } catch {
        Write-Error "Failed to parse MaintenanceStart or MaintenanceEnd as Unix milliseconds: $_"
        exit 2
    }
    $windowStartTS = $startDt.TimeOfDay
    $windowEndTS   = $endDt.TimeOfDay
    if ($MaintenanceRecurrence -ne 'Daily') {
        $targetDayOfWeek = Get-DayOfWeekFromName -DayName $MaintenanceDay
        if ($null -eq $targetDayOfWeek) {
            $targetDayOfWeek = $startDt.DayOfWeek
        }
    }
} else {
    Write-Error "MaintenanceStart and MaintenanceEnd must both be HH:mm (e.g. 02:00), seconds from midnight UTC (0-86400), or Unix milliseconds."
    exit 2
}

#endregion

#region Nth weekday in month (for Monthly recurrence)

function Get-NthWeekdayInMonth {
    param(
        [int]$Year,
        [int]$Month,
        [System.DayOfWeek]$DayOfWeek,
        [string]$Occurrence
    )
    $firstOfMonth = Get-Date -Year $Year -Month $Month -Day 1
    $dowFirst = [int]$firstOfMonth.DayOfWeek
    $targetDow = [int]$DayOfWeek
    $daysToFirst = ($targetDow - $dowFirst) % 7
    if ($daysToFirst -lt 0) { $daysToFirst += 7 }
    $firstOccurrence = $firstOfMonth.AddDays($daysToFirst)
    if ($Occurrence -eq 'last') {
        $lastDay = [DateTime]::DaysInMonth($Year, $Month)
        $cursor = Get-Date -Year $Year -Month $Month -Day $lastDay
        while ([int]$cursor.DayOfWeek -ne $targetDow) {
            $cursor = $cursor.AddDays(-1)
        }
        return $cursor.Date
    }
    $n = [int]$Occurrence
    $result = $firstOccurrence.AddDays(7 * ($n - 1))
    if ($result.Month -ne $Month) {
        $result = $result.AddDays(-7)
    }
    return $result.Date
}

#endregion

#region Core logic: is current time within the maintenance window?

$now = Get-Date
$isInsideWindow = $false
$occurrenceStart = $null
$occurrenceEnd = $null
$nextWindowStart = $null
$nextWindowEnd = $null

if ($MaintenanceRecurrence -eq 'Daily') {
    $todayStart = $now.Date + $windowStartTS
    $todayEnd   = $now.Date + $windowEndTS
    if ($windowEndTS -lt $windowStartTS) {
        $todayEnd = $todayEnd.AddDays(1)
    }
    $isInsideWindow = ($now -ge $todayStart) -and ($now -le $todayEnd)
    $occurrenceStart = $todayStart
    $occurrenceEnd   = $todayEnd
    $nextWindowStart = if ($now -gt $todayEnd) { $todayStart.AddDays(1) } else { $todayStart }
    $nextWindowEnd   = if ($now -gt $todayEnd) { $todayEnd.AddDays(1) } else { $todayEnd }
}
elseif ($MaintenanceRecurrence -eq 'Weekly') {
    $daysBack = ([int]$now.DayOfWeek - [int]$targetDayOfWeek + 7) % 7
    $mostRecentMaintenanceDate = $now.Date.AddDays(-$daysBack)
    $occurrenceStart = $mostRecentMaintenanceDate + $windowStartTS
    $occurrenceEnd   = $mostRecentMaintenanceDate + $windowEndTS
    if ($windowEndTS -lt $windowStartTS) {
        $occurrenceEnd = $occurrenceEnd.AddDays(1)
    }
    $isInsideWindow = ($now -ge $occurrenceStart) -and ($now -le $occurrenceEnd)
    $nextWindowStart = if ($now -gt $occurrenceEnd) { $occurrenceStart.AddDays(7) } else { $occurrenceStart }
    $nextWindowEnd   = if ($now -gt $occurrenceEnd) { $occurrenceEnd.AddDays(7) } else { $occurrenceEnd }
}
else {
    # Monthly: nth weekday of month
    $windowDate = Get-NthWeekdayInMonth -Year $now.Year -Month $now.Month -DayOfWeek $targetDayOfWeek -Occurrence $monthlyOccurrenceNorm
    $occurrenceStart = $windowDate + $windowStartTS
    $occurrenceEnd   = $windowDate + $windowEndTS
    if ($windowEndTS -lt $windowStartTS) {
        $occurrenceEnd = $occurrenceEnd.AddDays(1)
    }
    if ($now -ge $occurrenceStart -and $now -le $occurrenceEnd) {
        $isInsideWindow = $true
        $nextWindowStart = $occurrenceStart
        $nextWindowEnd   = $occurrenceEnd
    } elseif ($now -lt $occurrenceStart) {
        $nextWindowStart = $occurrenceStart
        $nextWindowEnd   = $occurrenceEnd
    } else {
        $nextMonth = $now.AddMonths(1)
        $windowDateNext = Get-NthWeekdayInMonth -Year $nextMonth.Year -Month $nextMonth.Month -DayOfWeek $targetDayOfWeek -Occurrence $monthlyOccurrenceNorm
        $nextWindowStart = $windowDateNext + $windowStartTS
        $nextWindowEnd   = $windowDateNext + $windowEndTS
        if ($windowEndTS -lt $windowStartTS) {
            $nextWindowEnd = $nextWindowEnd.AddDays(1)
        }
    }
}

#endregion

#region Output and exit

if ($isInsideWindow) {
    Write-Output "Current time is within maintenance window ($occurrenceStart - $occurrenceEnd)."
    if ($ExitWhenInside) { exit 0 } else { exit 1 }
} else {
    Write-Output "Current time is outside maintenance window. Next window: $nextWindowStart - $nextWindowEnd."
    if ($ExitWhenInside) { exit 1 } else { exit 0 }
}

#endregion

<#
.SYNOPSIS
    Checks whether the device should patch now based on a patching schedule (Daily, Weekly, or Monthly).

.DESCRIPTION
    Compares the device's current time to a recurring patching schedule. Patching schedule values are read
    from NinjaOne custom fields, set by Set-WeeklyPatchingSchedule.ps1. Supports Daily (every day at same time),
    Weekly (specific day of week), and Monthly (nth weekday of month). Supports HH:mm, seconds-from-midnight-UTC
    (0-86400), and Unix milliseconds. The script triggers only when run shortly before the patch time
    (within holdWindowMinutes): it waits until the exact start time, then exits 0 so patching runs at the precise time.
    If the script runs after the scheduled start time, it exits 1 (no trigger). Intended for NinjaOne compound conditions.

    Environment variables (from NinjaOne custom fields):
    - patchingRecurrence  : Optional. Daily | Weekly | Monthly. Default: Weekly.
    - patchingDay        : Day of week (Weekly/Monthly). Not used for Daily.
    - patchingOccurrence : For Monthly only: 1, 2, 3, 4, or Last.
    - patchingStart      : Start time as HH:mm, seconds from midnight UTC (0-86400), or Unix milliseconds.
    - holdWindowMinutes  : Optional. Max minutes before patching start to wait. Default: 15.
    - exitWhenShouldPatch: Optional. "true"/"1" = exit 0 when should patch (default); "false"/"0" = exit 0 when should not patch.

.EXAMPLE
    # Run on device; NinjaOne supplies patching schedule custom fields.
    .\Check-WeeklyPatchingSchedule.ps1
#>

$ErrorActionPreference = 'Stop'

# Read patching schedule from NinjaOne custom fields (injected as environment variables)
$PatchingRecurrence  = $env:patchingRecurrence
$PatchingOccurrence  = $env:patchingOccurrence
try { $v = Get-NinjaProperty patchingRecurrence; if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) { $PatchingRecurrence = $v } } catch { }
try { $v = Get-NinjaProperty patchingOccurrence;  if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) { $PatchingOccurrence  = $v } } catch { }
$PatchingDay   = Get-NinjaProperty patchingDay -type dropdown
$PatchingStart = Get-NinjaProperty patchingStart

$HoldWindowMinutes = 15
if ($env:holdWindowMinutes -ne $null) {
    $hwm = $env:holdWindowMinutes -as [string]
    if ($hwm -match '^\d+$' -and [int]$hwm -ge 0 -and [int]$hwm -le 1440) {
        $HoldWindowMinutes = [int]$hwm
    }
}

$ExitWhenShouldPatch = $true
if ($env:exitWhenShouldPatch -ne $null) {
    $ev = $env:exitWhenShouldPatch -as [string]
    if ($ev -match '^(?i)(true|1|yes)$') { $ExitWhenShouldPatch = $true }
    elseif ($ev -match '^(?i)(false|0|no)$') { $ExitWhenShouldPatch = $false }
}

# Normalize recurrence: Daily | Weekly | Monthly (default Weekly)
if ([string]::IsNullOrWhiteSpace($PatchingRecurrence)) { $PatchingRecurrence = 'Weekly' }
else {
    $PatchingRecurrence = $PatchingRecurrence.Trim().ToLowerInvariant()
    if ($PatchingRecurrence -eq 'daily') { $PatchingRecurrence = 'Daily' }
    elseif ($PatchingRecurrence -eq 'weekly') { $PatchingRecurrence = 'Weekly' }
    elseif ($PatchingRecurrence -eq 'monthly') { $PatchingRecurrence = 'Monthly' }
    else { $PatchingRecurrence = 'Weekly' }
}

#region Validation

if ([string]::IsNullOrWhiteSpace($PatchingStart)) {
    Write-Error "PatchingStart is required. Set the patchingStart custom field in NinjaOne (via Set-WeeklyPatchingSchedule.ps1 or the NinjaOne UI)."
    exit 2
}
if ($PatchingRecurrence -eq 'Weekly' -or $PatchingRecurrence -eq 'Monthly') {
    if ([string]::IsNullOrWhiteSpace($PatchingDay)) {
        Write-Error "PatchingDay is required for $PatchingRecurrence recurrence. Set the patchingDay custom field in NinjaOne."
        exit 2
    }
}
if ($PatchingRecurrence -eq 'Monthly') {
    if ([string]::IsNullOrWhiteSpace($PatchingOccurrence)) {
        Write-Error "PatchingOccurrence is required for Monthly recurrence. Use 1, 2, 3, 4, or Last."
        exit 2
    }
    $occNorm = $PatchingOccurrence.Trim().ToLowerInvariant()
    $validOcc = ($occNorm -eq 'last') -or ($occNorm -match '^[1-4]$')
    if (-not $validOcc) {
        Write-Error "Invalid PatchingOccurrence: '$PatchingOccurrence'. Use 1, 2, 3, 4, or Last."
        exit 2
    }
}

#endregion

#region Parse patching schedule (Format A: HH:mm, Format B: seconds from midnight, Format C: Unix ms)

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
$monthlyOccurrenceNorm = $null
if ($PatchingRecurrence -eq 'Monthly') {
    $mo = ($PatchingOccurrence -as [string]).Trim().ToLowerInvariant()
    $monthlyOccurrenceNorm = if ($mo -eq 'last') { 'last' } else { $mo }
}

# Format A: HH:mm
if ($PatchingStart -match '^(\d{1,2}):(\d{2})$') {
    if ($PatchingRecurrence -eq 'Weekly' -or $PatchingRecurrence -eq 'Monthly') {
        $targetDayOfWeek = Get-DayOfWeekFromName -DayName $PatchingDay
        if ($null -eq $targetDayOfWeek) {
            Write-Error "Invalid PatchingDay: '$PatchingDay'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."
            exit 2
        }
    }
    $null = $PatchingStart -match '^(\d{1,2}):(\d{2})$'
    $h1 = [int]$Matches[1]
    $m1 = [int]$Matches[2]
    if ($h1 -lt 0 -or $h1 -gt 23 -or $m1 -lt 0 -or $m1 -gt 59) {
        Write-Error "Invalid PatchingStart time: '$PatchingStart'. Use HH:mm (e.g. 02:00)."
        exit 2
    }
    $windowStartTS = [TimeSpan]::new($h1, $m1, 0)
}
# Format B: Seconds from midnight UTC (0-86400). NinjaOne "Time" custom fields use this.
elseif ($PatchingStart -match '^\d+$' -and [long]$PatchingStart -ge 0 -and [long]$PatchingStart -le 86400) {
    $secs = [int]$PatchingStart
    $utcDate = [DateTime]::UtcNow.Date
    $utcTime = $utcDate.AddSeconds($secs)
    $localTime = [TimeZoneInfo]::ConvertTimeFromUtc($utcTime, [TimeZoneInfo]::Local)
    $windowStartTS = $localTime.TimeOfDay
    if ($PatchingRecurrence -eq 'Weekly' -or $PatchingRecurrence -eq 'Monthly') {
        $targetDayOfWeek = Get-DayOfWeekFromName -DayName $PatchingDay
        if ($null -eq $targetDayOfWeek) {
            Write-Error "Invalid PatchingDay: '$PatchingDay'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."
            exit 2
        }
    }
}
# Format C: Unix milliseconds (e.g. from Set-WeeklyPatchingSchedule.ps1 import for Weekly)
elseif ($PatchingStart -match '^\d+$') {
    try {
        $startDt = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$PatchingStart).LocalDateTime
    } catch {
        Write-Error "Failed to parse PatchingStart as Unix milliseconds: $_"
        exit 2
    }
    $windowStartTS = $startDt.TimeOfDay
    if ($PatchingRecurrence -eq 'Weekly' -or $PatchingRecurrence -eq 'Monthly') {
        $targetDayOfWeek = Get-DayOfWeekFromName -DayName $PatchingDay
        if ($null -eq $targetDayOfWeek) {
            $targetDayOfWeek = $startDt.DayOfWeek
        }
    }
}
else {
    Write-Error "PatchingStart must be HH:mm (e.g. 02:00), seconds from midnight UTC (0-86400), or Unix milliseconds."
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

#region Core logic: compute next patching occurrence, then hold-window behavior

$now = Get-Date
$nextPatchingOccurrence = $null

if ($PatchingRecurrence -eq 'Daily') {
    $candidate = $now.Date + $windowStartTS
    if ($now -lt $candidate) {
        $nextPatchingOccurrence = $candidate
    } else {
        $nextPatchingOccurrence = $candidate.AddDays(1)
    }
}
elseif ($PatchingRecurrence -eq 'Weekly') {
    $daysBack = ([int]$now.DayOfWeek - [int]$targetDayOfWeek + 7) % 7
    $mostRecentThatDay = $now.Date.AddDays(-$daysBack)
    $candidate = $mostRecentThatDay + $windowStartTS
    if ($now -lt $candidate) {
        $nextPatchingOccurrence = $candidate
    } else {
        $nextPatchingOccurrence = $candidate.AddDays(7)
    }
}
else {
    # Monthly: nth weekday of month
    $windowDate = Get-NthWeekdayInMonth -Year $now.Year -Month $now.Month -DayOfWeek $targetDayOfWeek -Occurrence $monthlyOccurrenceNorm
    $candidate = $windowDate + $windowStartTS
    if ($now -lt $candidate) {
        $nextPatchingOccurrence = $candidate
    } else {
        $nextMonth = $now.AddMonths(1)
        $windowDateNext = Get-NthWeekdayInMonth -Year $nextMonth.Year -Month $nextMonth.Month -DayOfWeek $targetDayOfWeek -Occurrence $monthlyOccurrenceNorm
        $nextPatchingOccurrence = $windowDateNext + $windowStartTS
    }
}

# Case 1: Already at or past patching start time - do not trigger (script must run before time and hold)
if ($now -ge $nextPatchingOccurrence) {
    Write-Output "Past patching start (current time $now, next patching $nextPatchingOccurrence). Run before start time within hold window to trigger."
    if ($ExitWhenShouldPatch) { exit 1 } else { exit 0 }
}

# Case 2: Before patching start - check if within hold window
$minutesUntilStart = ($nextPatchingOccurrence - $now).TotalMinutes
if ($minutesUntilStart -le $HoldWindowMinutes) {
    $secondsToWait = [int][Math]::Ceiling(($nextPatchingOccurrence - $now).TotalSeconds)
    if ($secondsToWait -gt 0) {
        Write-Output "Within hold window. Waiting $secondsToWait seconds until patching start ($nextPatchingOccurrence)."
        Start-Sleep -Seconds $secondsToWait
    }
    Write-Output "Patching time reached."
    if ($ExitWhenShouldPatch) { exit 0 } else { exit 1 }
}

# Case 3: Before patching start, outside hold window
Write-Output "Before patching start and outside hold window ($([int]$minutesUntilStart) min until $nextPatchingOccurrence)."
if ($ExitWhenShouldPatch) { exit 1 } else { exit 0 }

#endregion

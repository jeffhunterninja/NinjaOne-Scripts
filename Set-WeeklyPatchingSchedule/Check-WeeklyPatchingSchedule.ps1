<#
.SYNOPSIS
    Checks whether the device should patch now based on a weekly patching schedule.

.DESCRIPTION
    Compares the device's current time to a recurring weekly patching schedule defined by day of week
    and exact start time. Patching schedule values are read from NinjaOne custom fields, which
    are set by Set-WeeklyPatchingSchedule.ps1. NinjaOne injects device/org/location custom fields
    (patchingDay, patchingStart) as environment variables when this script runs.
    Supports HH:mm, seconds-from-midnight-UTC (0-86400; NinjaOne Time fields), and Unix milliseconds (as stored by Set-WeeklyPatchingSchedule.ps1).
    The script triggers only when run shortly before the patch time (within holdWindowMinutes):
    it waits until the exact start time, then exits 0 so patching runs at the precise time.
    If the script runs after the scheduled start time, it exits 1 (no trigger). Intended for
    use with NinjaOne compound conditions.

    Environment variables (from NinjaOne custom fields):
    - patchingDay       : Day of week (e.g. Sunday, Monday).
    - patchingStart     : Start time as HH:mm (e.g. "02:00"), seconds from midnight UTC (0-86400), or Unix milliseconds.
    - holdWindowMinutes : Optional. Max minutes before patching start to wait. Match NinjaOne schedule interval. Default: 15.
    - exitWhenShouldPatch : Optional. "true"/"1" = exit 0 when should patch (default); "false"/"0" = exit 0 when should not patch.

.EXAMPLE
    # Run on device; NinjaOne supplies patchingDay, patchingStart from custom fields.
    .\Check-WeeklyPatchingSchedule.ps1
#>

$ErrorActionPreference = 'Stop'

# Read patching schedule from NinjaOne custom fields (injected as environment variables)
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

#region Validation

if ([string]::IsNullOrWhiteSpace($PatchingDay)) {
    Write-Error "PatchingDay is required. Set the patchingDay custom field in NinjaOne (via Set-WeeklyPatchingSchedule.ps1 or the NinjaOne UI)."
    exit 2
}
if ([string]::IsNullOrWhiteSpace($PatchingStart)) {
    Write-Error "PatchingStart is required. Set the patchingStart custom field in NinjaOne (via Set-WeeklyPatchingSchedule.ps1 or the NinjaOne UI)."
    exit 2
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

# Format A: HH:mm
if ($PatchingStart -match '^(\d{1,2}):(\d{2})$') {
    $targetDayOfWeek = Get-DayOfWeekFromName -DayName $PatchingDay
    if ($null -eq $targetDayOfWeek) {
        Write-Error "Invalid PatchingDay: '$PatchingDay'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."
        exit 2
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
    $targetDayOfWeek = Get-DayOfWeekFromName -DayName $PatchingDay
    if ($null -eq $targetDayOfWeek) {
        Write-Error "Invalid PatchingDay: '$PatchingDay'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."
        exit 2
    }
}
# Format C: Unix milliseconds (e.g. from Set-WeeklyPatchingSchedule.ps1 import)
elseif ($PatchingStart -match '^\d+$') {
    try {
        $startDt = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$PatchingStart).LocalDateTime
    } catch {
        Write-Error "Failed to parse PatchingStart as Unix milliseconds: $_"
        exit 2
    }
    # Use patchingDay for target day; Unix ms may come from different scope (org) than patchingDay (device)
    $targetDayOfWeek = Get-DayOfWeekFromName -DayName $PatchingDay
    if ($null -eq $targetDayOfWeek) {
        $targetDayOfWeek = $startDt.DayOfWeek
    }
    $windowStartTS = $startDt.TimeOfDay
}
else {
    Write-Error "PatchingStart must be HH:mm (e.g. 02:00), seconds from midnight UTC (0-86400), or Unix milliseconds."
    exit 2
}

#endregion

#region Core logic: is today patching day? Should we proceed or hold?

$now = Get-Date

# Check if today is patching day
if ($now.DayOfWeek -ne $targetDayOfWeek) {
    Write-Output "Not patching day (today is $($now.DayOfWeek), patching day is $targetDayOfWeek)."
    if ($ExitWhenShouldPatch) { exit 1 } else { exit 0 }
}

# Compute the patching start datetime for today
$patchingStartToday = $now.Date + $windowStartTS

# Case 1: Already at or past patching start time - do not trigger (script must run before time and hold)
if ($now -ge $patchingStartToday) {
    Write-Output "Past patching start (current time $now, patching start $patchingStartToday). Run before start time within hold window to trigger."
    if ($ExitWhenShouldPatch) { exit 1 } else { exit 0 }
}

# Case 2: Before patching start - check if within hold window
$minutesUntilStart = ($patchingStartToday - $now).TotalMinutes
if ($minutesUntilStart -le $HoldWindowMinutes) {
    $secondsToWait = [int][Math]::Ceiling(($patchingStartToday - $now).TotalSeconds)
    if ($secondsToWait -gt 0) {
        Write-Output "Within hold window. Waiting $secondsToWait seconds until patching start ($patchingStartToday)."
        Start-Sleep -Seconds $secondsToWait
    }
    Write-Output "Patching time reached."
    if ($ExitWhenShouldPatch) { exit 0 } else { exit 1 }
}

# Case 3: Before patching start, outside hold window
Write-Output "Before patching start and outside hold window ($minutesUntilStart min until $patchingStartToday)."
if ($ExitWhenShouldPatch) { exit 1 } else { exit 0 }

#endregion

<#
.SYNOPSIS
    Checks whether the device's current local time is within or outside a weekly maintenance window.

.DESCRIPTION
    Compares the device's current time to a recurring weekly maintenance window defined by day of week,
    start time, and end time. Maintenance window values are read from NinjaOne custom fields, which
    are set by Set-WeeklyMaintenanceWindow.ps1. NinjaOne injects device/org/location custom fields
    (maintenanceDay, maintenanceStart, maintenanceEnd) as environment variables when this script runs.
    Supports both HH:mm format and Unix milliseconds (as stored by Set-WeeklyMaintenanceWindow.ps1).
    Handles overnight windows (e.g. Sunday 22:00 - Monday 06:00). Intended for use with NinjaOne
    compound conditions.

    Environment variables (from NinjaOne custom fields):
    - maintenanceDay   : Day of week (e.g. Sunday, Monday).
    - maintenanceStart: Start time as HH:mm (e.g. "02:00") or Unix milliseconds.
    - maintenanceEnd  : End time as HH:mm (e.g. "04:00") or Unix milliseconds.
    - exitWhenInside  : Optional. "true"/"1" = exit 0 when inside window (default); "false"/"0" = exit 0 when outside.

.EXAMPLE
    # Run on device; NinjaOne supplies maintenanceDay, maintenanceStart, maintenanceEnd from custom fields.
    .\Check-WeeklyMaintenanceWindow.ps1
#>

$ErrorActionPreference = 'Stop'

# Read maintenance window from NinjaOne custom fields (injected as environment variables)
$MaintenanceDay   = Get-NinjaProperty maintenanceDay
$MaintenanceStart = Get-NinjaProperty maintenanceStart
$MaintenanceEnd   = Get-NinjaProperty maintenanceEnd

$ExitWhenInside = $true
if ($env:exitWhenInside -ne $null) {
    $ev = $env:exitWhenInside -as [string]
    if ($ev -match '^(?i)(true|1|yes)$') { $ExitWhenInside = $true }
    elseif ($ev -match '^(?i)(false|0|no)$') { $ExitWhenInside = $false }
}

#region Validation

if ([string]::IsNullOrWhiteSpace($MaintenanceDay)) {
    Write-Error "MaintenanceDay is required. Set the maintenanceDay custom field in NinjaOne (via Set-WeeklyMaintenanceWindow.ps1 or the NinjaOne UI)."
    exit 2
}
if ([string]::IsNullOrWhiteSpace($MaintenanceStart)) {
    Write-Error "MaintenanceStart is required. Set the maintenanceStart custom field in NinjaOne (via Set-WeeklyMaintenanceWindow.ps1 or the NinjaOne UI)."
    exit 2
}
if ([string]::IsNullOrWhiteSpace($MaintenanceEnd)) {
    Write-Error "MaintenanceEnd is required. Set the maintenanceEnd custom field in NinjaOne (via Set-WeeklyMaintenanceWindow.ps1 or the NinjaOne UI)."
    exit 2
}

#endregion

#region Parse maintenance window (Format A: HH:mm, Format B: Unix ms)

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

# Format A: HH:mm
if ($MaintenanceStart -match '^(\d{1,2}):(\d{2})$' -and $MaintenanceEnd -match '^(\d{1,2}):(\d{2})$') {
    $targetDayOfWeek = Get-DayOfWeekFromName -DayName $MaintenanceDay
    if ($null -eq $targetDayOfWeek) {
        Write-Error "Invalid MaintenanceDay: '$MaintenanceDay'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."
        exit 2
    }
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
}
# Format B: Unix milliseconds
elseif ($MaintenanceStart -match '^\d+$' -and $MaintenanceEnd -match '^\d+$') {
    try {
        $startDt = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$MaintenanceStart).LocalDateTime
        $endDt   = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$MaintenanceEnd).LocalDateTime
    } catch {
        Write-Error "Failed to parse MaintenanceStart or MaintenanceEnd as Unix milliseconds: $_"
        exit 2
    }
    $targetDayOfWeek = $startDt.DayOfWeek
    $windowStartTS = $startDt.TimeOfDay
    $windowEndTS   = $endDt.TimeOfDay
}
else {
    Write-Error "MaintenanceStart and MaintenanceEnd must both be HH:mm (e.g. 02:00) or both be Unix milliseconds."
    exit 2
}

#endregion

#region Core logic: is current time within the maintenance window?

$now = Get-Date

# Days back to the most recent occurrence of the maintenance day
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

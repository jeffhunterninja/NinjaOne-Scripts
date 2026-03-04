#Requires -Version 5.1
<#
.SYNOPSIS
  Runs the quiet-hours test and logs/alerting (e.g. for scheduled or manual runs).

.DESCRIPTION
  Checks whether the current local time falls within the user's Quiet Hours (as defined in the preference JSON).
  Optionally appends a one-line status to a log file and/or writes to the host.
  Use this script for scheduled checks or manual testing; use Set-QuietHours.ps1 to configure preferences via the UI.

.PARAMETER PreferencePath
  Full path to the quiet hours JSON file. Default: C:\RMM\NinjaOne-QuietHours\quiet_hours.json

.PARAMETER LogPath
  If provided, append a one-line status (timestamp and InQuietHours result) to this file. Directory is created if needed.

.PARAMETER Quiet
  Suppress host output; only log to file when LogPath is set.

.PARAMETER Mode
  When set, controls exit code for alerting (e.g. in NinjaOne): Alert Within = exit 1 when current time is within quiet hours; Alert Outside = exit 1 when current time is outside quiet hours. When not set, script always exits 0. Can be overridden by script variable $env:quietHoursMode in NinjaOne.

.EXIT CODES
  0 = No alert (or Mode not set).
  1 = Alert condition met (within quiet hours for Alert Within, outside quiet hours for Alert Outside). Only when Mode is set and prefs exist.

.EXAMPLE
  .\Test-QuietHours.ps1
  Runs the test and writes status to the host.

.EXAMPLE
  .\Test-QuietHours.ps1 -LogPath "C:\RMM\NinjaOne-QuietHours\quiet_hours_check.log"
  Runs the test, writes to host, and appends one line to the log file.

.EXAMPLE
  .\Test-QuietHours.ps1 -LogPath "C:\Logs\quiet.log" -Quiet
  Runs the test and appends to the log file only (no host output).

.EXAMPLE
  .\Test-QuietHours.ps1 -Mode "Alert Outside"
  Exit 1 when device is outside quiet hours (so NinjaOne can alert); exit 0 when within quiet hours or no prefs.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [string]$PreferencePath = "C:\RMM\NinjaOne-QuietHours\quiet_hours.json",
  [Parameter(Mandatory = $false)]
  [string]$LogPath = "",
  [Parameter(Mandatory = $false)]
  [switch]$Quiet,
  [Parameter(Mandatory = $false)]
  [string]$Mode = $(if ($env:quietHoursMode) { $env:quietHoursMode } else { '' })
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Helpers
function Get-QuietPrefs {
  param([string]$Path)
  if (Test-Path $Path) {
    try {
      return Get-Content -Raw -Path $Path | ConvertFrom-Json
    } catch {
      Write-Verbose "Could not load quiet hours prefs: $($_.Exception.Message)"
      return $null
    }
  }
  return $null
}

# If primary path has no file, try user-profile paths (backward compatibility with pre-RMM deployments).
function Get-QuietPrefsWithFallback {
  param([string]$PrimaryPath = $PreferencePath)
  $p = Get-QuietPrefs -Path $PrimaryPath
  if ($null -ne $p) { return $p }
  $userFiles = Get-ChildItem -Path "C:\Users\*\AppData\Local\NinjaOne-QuietHours\quiet_hours.json" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
  if (-not $userFiles -or $userFiles.Count -eq 0) { return $null }
  return Get-QuietPrefs -Path $userFiles[0].FullName
}

function Test-TimeInRange {
  param(
    [Parameter(Mandatory)] [datetime]$Now,
    [Parameter(Mandatory)] [string]$Start,
    [Parameter(Mandatory)] [string]$End
  )
  # Accepts "HH:mm" 24-hour format. Handles wrap over midnight.
  $today = $Now.Date
  $startTime = [datetime]::ParseExact($Start, 'HH:mm', $null)
  $endTime   = [datetime]::ParseExact($End  , 'HH:mm', $null)

  $startDt = $today.AddHours($startTime.Hour).AddMinutes($startTime.Minute)
  $endDt   = $today.AddHours($endTime.Hour  ).AddMinutes($endTime.Minute)

  if ($endDt -le $startDt) {
    # Quiet window crosses midnight, e.g., 21:00 -> 07:00
    return (($Now -ge $startDt) -or ($Now -lt $endDt.AddDays(1)))
  } else {
    return ($Now -ge $startDt -and $Now -lt $endDt)
  }
}

function Test-QuietHours {
  <#
    .SYNOPSIS
      Returns $true if the current local time falls within the user's Quiet Hours.
    .DESCRIPTION
      Accepts a Quiet Hours JSON object (as returned by Get-QuietPrefs) or reads from default path.
      Supports single range per "weekday" and per "weekend", plus optional per-day overrides.
  #>
  param(
    $Prefs = $(Get-QuietPrefsWithFallback),
    [datetime]$Now = (Get-Date)
  )
  if (-not $Prefs) { return $false }

  $dow = [int]$Now.DayOfWeek  # Sunday=0
  $isWeekend = ($dow -in 0,6)

  # Prefer per-day override if present (perDay is optional in saved JSON)
  $perDay = $null
  if ($Prefs.PSObject.Properties['perDay']) {
    $perDay = $Prefs.perDay | Where-Object { $_.day -eq $dow }
  }
  if ($perDay) {
    return (Test-TimeInRange -Now $Now -Start $perDay.start -End $perDay.end)
  }

  if ($isWeekend -and $Prefs.PSObject.Properties['weekend'] -and $Prefs.weekend) {
    return (Test-TimeInRange -Now $Now -Start $Prefs.weekend.start -End $Prefs.weekend.end)
  }
  if (-not $isWeekend -and $Prefs.PSObject.Properties['weekdays'] -and $Prefs.weekdays) {
    return (Test-TimeInRange -Now $Now -Start $Prefs.weekdays.start -End $Prefs.weekdays.end)
  }
  return $false
}
#endregion Helpers

# Run test and log only when script is executed directly (not dot-sourced)
if ($MyInvocation.InvocationName -ne '.') {
  $prefs = Get-QuietPrefsWithFallback
  $inQuiet = $false
  if ($prefs) {
    $inQuiet = Test-QuietHours -Prefs $prefs -Now (Get-Date)
  }
  $status = if ($null -eq $prefs) { "NoPrefs" } else { "InQuietHours=$inQuiet" }
  $line = "$(Get-Date -Format 'o') $status"

  if ($LogPath) {
    $dir = [System.IO.Path]::GetDirectoryName($LogPath)
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Add-Content -Path $LogPath -Value $line -Encoding UTF8
  }

  if (-not $Quiet) {
    if ($null -eq $prefs) {
      Write-Host "Quiet hours: no preferences file found at $PreferencePath"
    } else {
      Write-Host $line
    }
  }

  # Exit code for alerting when Mode is set (NinjaOne script variable $env:quietHoursMode can set Mode)
  $modeKey = $Mode.Trim()
  if ($modeKey) {
    if ($null -eq $prefs) {
      exit 0
    }
    switch ($modeKey) {
      'Alert Within'  { if ($inQuiet) { exit 1 } else { exit 0 } }
      'Alert Outside' { if (-not $inQuiet) { exit 1 } else { exit 0 } }
      default         { exit 0 }
    }
  }
}

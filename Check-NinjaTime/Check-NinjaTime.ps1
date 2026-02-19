#Requires -Version 5.1
<#
.SYNOPSIS
    Multi-mode scheduling script with recurring window mode, configurable grace period,
    and support for environment variable defaults.

.EXIT CODES
  1 = Condition met (within window or at target time; NinjaOne triggers linked automation)
  2 = Condition not met (outside window, too far from target, or validation/parse error)

.DESCRIPTION
    This script supports several scheduling modes:
      - Once: A one‑time event based on a full date/time.
      - Daily: Runs every day at the time specified in TargetTime (only the time-of-day is used).
      - Weekly: Runs on one or more specified days of the week at the time specified in TargetTime.
      - Monthly: Runs on a specific numbered day of the month at the time specified in TargetTime.
      - Monthly, Day of Week: Runs on a specific occurrence of a weekday (e.g. second Tuesday, last Saturday).
      - Window: Executes if the current time is within (or about to reach) a recurring window.
               For Window mode you can specify a recurrence pattern:
                   * Daily (default): The window recurs every day.
                   * Weekly: The window recurs on specified day(s) of the week.
                   * Monthly: The window recurs on a specific occurrence of a weekday (e.g. second Tuesday).
               In recurring mode the date portion is ignored and only the time-of-day is used (Daily/Weekly).
               
    A configurable grace period ($TimeWindowMinutes) determines whether the script will sleep until the target time
    (or window start) if it is within that many minutes; otherwise, the script exits, allowing recurring execution.

.PARAMETER Mode
    The scheduling mode. Allowed values: Once, Daily, Weekly, Monthly, MonthlyDayOfWeek, Window.
    (Defaults to environment variable 'mode'.)

.PARAMETER TargetTime
    (For non-Window modes) Full date/time or time-of-day.
    (Defaults to environment variable 'targetTime'.)

.PARAMETER DayOfWeek
    (For Weekly or Monthly, Day of Week mode) One or more days (Weekly) or a single day (MonthlyDayOfWeek).
    (Defaults to environment variable 'weeklyDayOfWeek'.)

.PARAMETER DayOfMonth
    (For Monthly mode in non-window modes) The numeric day of the month (1–31).
    (Defaults to environment variable 'monthlyDayOfMonth'.)

.PARAMETER MonthlyOccurrence
    (For Monthly, Day of Week mode) Occurrence of the weekday in the month: 1, 2, 3, 4, or Last.
    (Defaults to environment variable 'monthlyOccurrence'.)

.PARAMETER WindowStart
    (For Window mode) The window’s start time.
    (Defaults to environment variable 'windowStart'.)
    Only the time portion will be used (e.g. "13:00").

.PARAMETER WindowEnd
    (For Window mode) The window’s end time.
    (Defaults to environment variable 'windowEnd'.)
    Only the time portion will be used (e.g. "23:00").

.PARAMETER WindowRecurrence
    (For Window mode) Specifies whether the window recurs "Daily", "Weekly", or "Monthly".
    (Defaults to environment variable 'windowRecurrence'; if not provided, defaults to "Daily".)

.PARAMETER WindowMonthlyOccurrence
    (For Window mode with Monthly recurrence) Occurrence of the weekday in the month: 1, 2, 3, 4, or Last.
    (Defaults to environment variable 'windowMonthlyOccurrence'.)

.PARAMETER WindowDayOfWeek
    (For Window mode with Weekly recurrence) One or more days of the week on which the window recurs.
    (Defaults to environment variable 'windowDayOfWeek'.)

.PARAMETER TimeWindowMinutes
    Grace period in minutes. If the target time or window start is within this many minutes,
    the script will sleep until it and exit 1. Set to interval + 1 (e.g., 16 for 15-min interval).
    (Defaults to environment variable 'timeWindowMinutes' or 5.)

.NOTES
    NinjaOne will invoke the linked action when this script exits 1 (condition met).
#>

[CmdletBinding()]
param(
    [string]$Mode,
    [string]$TargetTime,
    [string[]]$DayOfWeek,
    [int]$DayOfMonth,
    [string]$MonthlyOccurrence,
    [string]$WindowStart,
    [string]$WindowEnd,
    [string]$WindowRecurrence,
    [string]$WindowMonthlyOccurrence,
    [string[]]$WindowDayOfWeek,
    [int]$TimeWindowMinutes = 5
)

$ErrorActionPreference = 'Stop'

# Assign from environment variables when param not passed; trim strings.
if ([string]::IsNullOrWhiteSpace($Mode) -and $null -ne $env:mode -and -not [string]::IsNullOrWhiteSpace($env:mode))           { $Mode = $env:mode.Trim() }
if ([string]::IsNullOrWhiteSpace($TargetTime) -and $null -ne $env:targetTime -and -not [string]::IsNullOrWhiteSpace($env:targetTime)) { $TargetTime = $env:targetTime.Trim() }
if ((-not $DayOfWeek -or $DayOfWeek.Count -eq 0) -and $null -ne $env:weeklyDayOfWeek -and -not [string]::IsNullOrWhiteSpace($env:weeklyDayOfWeek)) { $DayOfWeek = @($env:weeklyDayOfWeek) }
if (-not $DayOfMonth -and $null -ne $env:monthlyDayOfMonth -and -not [string]::IsNullOrWhiteSpace($env:monthlyDayOfMonth)) {
    $parsed = 0
    if ([int]::TryParse($env:monthlyDayOfMonth.Trim(), [ref]$parsed) -and $parsed -ge 1 -and $parsed -le 31) { $DayOfMonth = $parsed }
}
if ([string]::IsNullOrWhiteSpace($MonthlyOccurrence) -and $null -ne $env:monthlyOccurrence -and -not [string]::IsNullOrWhiteSpace($env:monthlyOccurrence)) { $MonthlyOccurrence = $env:monthlyOccurrence.Trim() }
if ([string]::IsNullOrWhiteSpace($WindowStart) -and $null -ne $env:windowStart -and -not [string]::IsNullOrWhiteSpace($env:windowStart))   { $WindowStart = $env:windowStart.Trim() }
if ([string]::IsNullOrWhiteSpace($WindowEnd) -and $null -ne $env:windowEnd -and -not [string]::IsNullOrWhiteSpace($env:windowEnd))       { $WindowEnd = $env:windowEnd.Trim() }
if ([string]::IsNullOrWhiteSpace($WindowRecurrence) -and $null -ne $env:windowRecurrence -and -not [string]::IsNullOrWhiteSpace($env:windowRecurrence)) { $WindowRecurrence = $env:windowRecurrence.Trim() }
if ([string]::IsNullOrWhiteSpace($WindowMonthlyOccurrence) -and $null -ne $env:windowMonthlyOccurrence -and -not [string]::IsNullOrWhiteSpace($env:windowMonthlyOccurrence)) { $WindowMonthlyOccurrence = $env:windowMonthlyOccurrence.Trim() }
if ((-not $WindowDayOfWeek -or $WindowDayOfWeek.Count -eq 0) -and $null -ne $env:windowDayOfWeek -and -not [string]::IsNullOrWhiteSpace($env:windowDayOfWeek)) { $WindowDayOfWeek = @($env:windowDayOfWeek) }
if ($null -ne $env:timeWindowMinutes -and -not [string]::IsNullOrWhiteSpace($env:timeWindowMinutes)) {
    $parsed = 0
    if ([int]::TryParse($env:timeWindowMinutes.Trim(), [ref]$parsed) -and $parsed -ge 0) { $TimeWindowMinutes = $parsed }
}

# Normalize Mode and WindowRecurrence for case-insensitive comparison.
$Mode = ($Mode -as [string]).Trim().ToLowerInvariant()
if ($Mode -eq 'monthly, day of week') { $Mode = 'monthlydayofweek' }
if ([string]::IsNullOrWhiteSpace($Mode)) { Write-Error "Mode parameter is required. Set via parameter or NinjaOne script variable 'mode'."; exit 2 }
if ($WindowRecurrence) { $WindowRecurrence = $WindowRecurrence.Trim().ToLowerInvariant() }

# Parse comma/semicolon-separated day-of-week strings into arrays.
function Split-DayOfWeek {
    param([object]$InputValue)
    if (-not $InputValue) { return @() }
    $str = if ($InputValue -is [string]) { $InputValue } else { ($InputValue -join ',') }
    return @($str -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}
$DayOfWeek = Split-DayOfWeek -InputValue $DayOfWeek
$WindowDayOfWeek = Split-DayOfWeek -InputValue $WindowDayOfWeek

# Parse day name to [System.DayOfWeek] (case-insensitive).
function Get-DayOfWeekFromName {
    param([string]$DayName)
    if ([string]::IsNullOrWhiteSpace($DayName)) { return $null }
    $parsed = $null
    if ([Enum]::TryParse([System.DayOfWeek], $DayName.Trim(), $true, [ref]$parsed)) { return $parsed }
    return $null
}

# Normalize occurrence string to canonical form: "1", "2", "3", "4", or "last". Returns $null if invalid.
function Get-NormalizedMonthlyOccurrence {
    param([string]$Occurrence)
    if ([string]::IsNullOrWhiteSpace($Occurrence)) { return $null }
    $s = $Occurrence.Trim().ToLowerInvariant()
    if ($s -eq 'last') { return 'last' }
    $n = 0
    if ([int]::TryParse($s, [ref]$n) -and $n -ge 1 -and $n -le 4) { return [string]$n }
    return $null
}

# Get the date (midnight) of the nth occurrence of a weekday in the given month.
# Occurrence: "1", "2", "3", "4" (1st/2nd/3rd/4th), or "last".
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
    # Days until first occurrence of that weekday (0-6)
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
    # If we overshoot the month (e.g. 5th Tuesday in a 4-Tuesday month), cap at last occurrence
    if ($result.Month -ne $Month) {
        $result = $result.AddDays(-7)
    }
    return $result.Date
}

$MonthlyOccurrenceNormalized = Get-NormalizedMonthlyOccurrence -Occurrence $MonthlyOccurrence
$WindowMonthlyOccurrenceNormalized = Get-NormalizedMonthlyOccurrence -Occurrence $WindowMonthlyOccurrence

if ($Mode -eq 'monthlydayofweek' -and -not [string]::IsNullOrWhiteSpace($MonthlyOccurrence) -and $null -eq $MonthlyOccurrenceNormalized) {
    Write-Error "Invalid MonthlyOccurrence: '$($MonthlyOccurrence.Trim())'. Use 1, 2, 3, 4, or Last."
    exit 2
}

Write-Verbose "Mode: '$Mode', TargetTime: '$TargetTime', DayOfWeek: '$($DayOfWeek -join ', ')', TimeWindowMinutes: $TimeWindowMinutes"

# Convert input strings to DateTime or TimeSpan objects as needed.
if ($Mode -ne "window") {
    if ([string]::IsNullOrEmpty($TargetTime)) {
        Write-Error "TargetTime parameter is required for mode $Mode."
        exit 2
    }
    try {
        $TargetTime = [datetime]::Parse($TargetTime)
    } catch {
        Write-Error "TargetTime '$TargetTime' could not be parsed as a valid DateTime."
        exit 2
    }
} else {
    if ([string]::IsNullOrEmpty($WindowStart) -or [string]::IsNullOrEmpty($WindowEnd)) {
        Write-Error "WindowStart and WindowEnd parameters are required for Window mode."
        exit 2
    }
    try {
        # Parse the ISO8601 string and extract the local time portion in "HH:mm" format.
        $wsString = ([datetimeoffset]::Parse($WindowStart.Trim())).LocalDateTime.ToString("HH:mm")
        $weString = ([datetimeoffset]::Parse($WindowEnd.Trim())).LocalDateTime.ToString("HH:mm")
        # Convert the time strings to TimeSpan objects.
        $WindowStartTS = [TimeSpan]::Parse($wsString)
        $WindowEndTS   = [TimeSpan]::Parse($weString)
        Write-Verbose "Parsed WindowStart TimeSpan: $WindowStartTS, WindowEnd TimeSpan: $WindowEndTS"
    } catch {
        Write-Error "WindowStart or WindowEnd could not be parsed as valid TimeSpan values. Error: $_"
        exit 2
    }
}

function Get-NextOccurrence {
    param(
        [string]$Mode,
        [datetime]$TargetTime,
        [string[]]$DayOfWeek,
        [int]$DayOfMonth,
        [string]$MonthlyOccurrence
    )
    $now = Get-Date
    switch ($Mode) {
        "once" {
            return $TargetTime
        }
        "daily" {
            $todayOccurrence = $now.Date + $TargetTime.TimeOfDay
            if ($todayOccurrence -gt $now) {
                return $todayOccurrence
            } else {
                return $todayOccurrence.AddDays(1)
            }
        }
        "weekly" {
            if (-not $DayOfWeek -or $DayOfWeek.Count -eq 0) {
                Write-Error "DayOfWeek parameter is required for Weekly mode."
                exit 2
            }
            $occurrences = foreach ($dow in $DayOfWeek) {
                $targetDayVal = Get-DayOfWeekFromName -DayName $dow
                if ($null -eq $targetDayVal) { Write-Error "Invalid day of week: '$dow'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."; exit 2 }
                $targetDay = [int]$targetDayVal
                $currentDay = [int]$now.DayOfWeek
                $daysToAdd = $targetDay - $currentDay
                if ($daysToAdd -lt 0 -or ($daysToAdd -eq 0 -and ($now.TimeOfDay -ge $TargetTime.TimeOfDay))) {
                    $daysToAdd += 7
                }
                $now.Date.AddDays($daysToAdd) + $TargetTime.TimeOfDay
            }
            return $occurrences | Sort-Object | Select-Object -First 1
        }
        "monthly" {
            if (-not $DayOfMonth) {
                Write-Error "DayOfMonth parameter is required for Monthly mode."
                exit 2
            }
            $year = $now.Year
            $month = $now.Month
            try {
                $occurrence = Get-Date -Year $year -Month $month -Day $DayOfMonth -Hour $TargetTime.Hour -Minute $TargetTime.Minute -Second $TargetTime.Second
            } catch {
                Write-Error "Invalid DayOfMonth for the current month."
                exit 2
            }
            if ($occurrence -gt $now) {
                return $occurrence
            } else {
                $nextMonth = $now.AddMonths(1)
                $year = $nextMonth.Year
                $month = $nextMonth.Month
                try {
                    $occurrence = Get-Date -Year $year -Month $month -Day $DayOfMonth -Hour $TargetTime.Hour -Minute $TargetTime.Minute -Second $TargetTime.Second
                } catch {
                    Write-Error "Invalid DayOfMonth for the next month."
                    exit 2
                }
                return $occurrence
            }
        }
        "monthlydayofweek" {
            if (-not $DayOfWeek -or $DayOfWeek.Count -eq 0) {
                Write-Error "DayOfWeek parameter is required for Monthly, Day of Week mode (use a single day, e.g. Tuesday)."
                exit 2
            }
            if (-not $MonthlyOccurrence) {
                Write-Error "MonthlyOccurrence parameter is required for Monthly, Day of Week mode. Use 1, 2, 3, 4, or Last."
                exit 2
            }
            $dowVal = Get-DayOfWeekFromName -DayName $DayOfWeek[0]
            if ($null -eq $dowVal) { Write-Error "Invalid day of week: '$($DayOfWeek[0])'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."; exit 2 }
            $year = $now.Year
            $month = $now.Month
            $occurrenceDate = Get-NthWeekdayInMonth -Year $year -Month $month -DayOfWeek $dowVal -Occurrence $MonthlyOccurrence
            $occurrenceDt = $occurrenceDate + $TargetTime.TimeOfDay
            if ($occurrenceDt -gt $now) {
                return $occurrenceDt
            }
            $nextMonth = $now.AddMonths(1)
            $occurrenceDate = Get-NthWeekdayInMonth -Year $nextMonth.Year -Month $nextMonth.Month -DayOfWeek $dowVal -Occurrence $MonthlyOccurrence
            return $occurrenceDate + $TargetTime.TimeOfDay
        }
        default {
            Write-Error "Unsupported mode in Get-NextOccurrence."
            exit 2
        }
    }
}

$now = Get-Date

switch ($Mode) {
    "window" {
        if (-not $WindowRecurrence) { $WindowRecurrence = "daily" }
        if ($WindowRecurrence -eq "daily") {
            # Use the TimeSpans we extracted.
            $todayWindowStart = $now.Date + $WindowStartTS
            $todayWindowEnd   = $now.Date + $WindowEndTS
            if ($WindowEndTS -lt $WindowStartTS) {
                $todayWindowEnd = $todayWindowEnd.AddDays(1)
            }
            
            if ($now -lt $todayWindowStart) {
                $timeToWait = $todayWindowStart - $now
                if ($timeToWait.TotalMinutes -gt $TimeWindowMinutes) {
                    Write-Output "Today's window starts in more than $TimeWindowMinutes minutes. Exiting."
                    exit 2
                }
                $sleepSec = [math]::Min([math]::Ceiling($timeToWait.TotalSeconds), 2147483647)
                Write-Output "Current time is before today's window. Sleeping for $sleepSec seconds until window starts at $todayWindowStart."
                Start-Sleep -Seconds $sleepSec
            } elseif ($now -gt $todayWindowEnd) {
                Write-Output "Today's window has passed. Exiting."
                exit 2
            }
            Write-Output "Current time is within today's window. Condition met; NinjaOne will invoke the linked action."
            exit 1
        }
        elseif ($WindowRecurrence -eq "weekly") {
            if (-not $WindowDayOfWeek -or $WindowDayOfWeek.Count -eq 0) {
                Write-Error "WindowDayOfWeek parameter is required for Weekly window recurrence."
                exit 2
            }
            $windowOccurrences = foreach ($dow in $WindowDayOfWeek) {
                $targetDayVal = Get-DayOfWeekFromName -DayName $dow
                if ($null -eq $targetDayVal) { Write-Error "Invalid window day of week: '$dow'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."; exit 2 }
                $targetDay = [int]$targetDayVal
                $currentDay = [int]$now.DayOfWeek
                $daysToAdd = $targetDay - $currentDay
                if ($daysToAdd -lt 0 -or ($daysToAdd -eq 0 -and ($now.TimeOfDay -ge $WindowStartTS))) {
                    $daysToAdd += 7
                }
                $occurrenceStart = $now.Date.AddDays($daysToAdd) + $WindowStartTS
                $occurrenceEnd   = $now.Date.AddDays($daysToAdd) + $WindowEndTS
                if ($WindowEndTS -lt $WindowStartTS) {
                    $occurrenceEnd = $occurrenceEnd.AddDays(1)
                }
                [PSCustomObject]@{
                    Start = $occurrenceStart
                    End   = $occurrenceEnd
                }
            }
            $nextWindow = $windowOccurrences | Sort-Object Start | Select-Object -First 1

            if ($now -lt $nextWindow.Start) {
                $timeToWait = $nextWindow.Start - $now
                if ($timeToWait.TotalMinutes -gt $TimeWindowMinutes) {
                    Write-Output "Next window start ($($nextWindow.Start)) is not within the next $TimeWindowMinutes minutes. Exiting."
                    exit 2
                }
                $sleepSec = [math]::Min([math]::Ceiling($timeToWait.TotalSeconds), 2147483647)
                Write-Output "Current time is before the next weekly window. Sleeping for $sleepSec seconds until window starts at $($nextWindow.Start)."
                Start-Sleep -Seconds $sleepSec
            } elseif ($now -gt $nextWindow.End) {
                Write-Output "Current time is past the next window ($($nextWindow.End)). Exiting."
                exit 2
            }
            Write-Output "Current time is within the weekly window. Condition met; NinjaOne will invoke the linked action."
            exit 1
        }
        elseif ($WindowRecurrence -eq "monthly") {
            if (-not $WindowDayOfWeek -or $WindowDayOfWeek.Count -eq 0) {
                Write-Error "WindowDayOfWeek parameter is required for Monthly window recurrence (use a single day, e.g. Tuesday)."
                exit 2
            }
            if (-not $WindowMonthlyOccurrenceNormalized) {
                if (-not [string]::IsNullOrWhiteSpace($WindowMonthlyOccurrence)) {
                    Write-Error "Invalid WindowMonthlyOccurrence: '$($WindowMonthlyOccurrence.Trim())'. Use 1, 2, 3, 4, or Last."
                } else {
                    Write-Error "WindowMonthlyOccurrence parameter is required for Monthly window recurrence. Use 1, 2, 3, 4, or Last."
                }
                exit 2
            }
            $dowVal = Get-DayOfWeekFromName -DayName $WindowDayOfWeek[0]
            if ($null -eq $dowVal) { Write-Error "Invalid window day of week: '$($WindowDayOfWeek[0])'. Use Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, or Saturday."; exit 2 }
            $twMin = [double]$TimeWindowMinutes

            $windowDate = Get-NthWeekdayInMonth -Year $now.Year -Month $now.Month -DayOfWeek $dowVal -Occurrence $WindowMonthlyOccurrenceNormalized
            $windowStart = $windowDate + $WindowStartTS
            $windowEnd = $windowDate + $WindowEndTS
            if ($WindowEndTS -lt $WindowStartTS) {
                $windowEnd = $windowEnd.AddDays(1)
            }

            if ($now -ge $windowStart -and $now -le $windowEnd) {
                Write-Output "Current time is within the monthly window. Condition met; NinjaOne will invoke the linked action."
                exit 1
            }
            if ($now -lt $windowStart) {
                $ts = $windowStart - $now
                if ($ts.TotalMinutes -gt $twMin) {
                    Write-Output "This month's window starts at $windowStart, which is not within the next $TimeWindowMinutes minutes. Exiting."
                    exit 2
                }
                $sleepSec = [math]::Min([math]::Ceiling($ts.TotalSeconds), 2147483647.0)
                Write-Output "Current time is before this month's window. Sleeping for $sleepSec seconds until window starts at $windowStart."
                Start-Sleep -Seconds $sleepSec
                exit 1
            }
            # This month's window has passed; try next month
            $nextMonth = $now.AddMonths(1)
            $windowDate = Get-NthWeekdayInMonth -Year $nextMonth.Year -Month $nextMonth.Month -DayOfWeek $dowVal -Occurrence $WindowMonthlyOccurrenceNormalized
            $windowStart = $windowDate + $WindowStartTS
            $windowEnd = $windowDate + $WindowEndTS
            if ($WindowEndTS -lt $WindowStartTS) {
                $windowEnd = $windowEnd.AddDays(1)
            }
            if ($now -lt $windowStart) {
                $ts = $windowStart - $now
                if ($ts.TotalMinutes -gt $twMin) {
                    Write-Output "Next month's window starts at $windowStart, which is not within the next $TimeWindowMinutes minutes. Exiting."
                    exit 2
                }
                $sleepSec = [math]::Min([math]::Ceiling($ts.TotalSeconds), 2147483647.0)
                Write-Output "Current time is before next month's window. Sleeping for $sleepSec seconds until window starts at $windowStart."
                Start-Sleep -Seconds $sleepSec
                exit 1
            }
            if ($now -ge $windowStart -and $now -le $windowEnd) {
                Write-Output "Current time is within the monthly window. Condition met; NinjaOne will invoke the linked action."
                exit 1
            }
            Write-Output "Current time is past the next monthly window ($windowEnd). Exiting."
            exit 2
        }
        else {
            Write-Error "Unsupported WindowRecurrence value. Use Daily, Weekly, or Monthly."
            exit 2
        }
    }
    default {
        $nextOccurrence = Get-NextOccurrence -Mode $Mode -TargetTime $TargetTime -DayOfWeek $DayOfWeek -DayOfMonth $DayOfMonth -MonthlyOccurrence $MonthlyOccurrenceNormalized
        $timeDifference = $nextOccurrence - $now

        if ($timeDifference.TotalMinutes -gt $TimeWindowMinutes) {
            Write-Output "Scheduled time ($nextOccurrence) is not within the next $TimeWindowMinutes minutes. Exiting."
            exit 2
        } elseif ($timeDifference.TotalSeconds -gt 0) {
            $sleepSec = [math]::Min([math]::Ceiling($timeDifference.TotalSeconds), 2147483647)
            Write-Output "Sleeping for $sleepSec seconds until scheduled time: $nextOccurrence"
            Start-Sleep -Seconds $sleepSec
        } else {
            Write-Output "Scheduled time has already passed. Exiting."
            exit 2
        }
        Write-Output "Scheduled time reached. Condition met; NinjaOne will invoke the linked action."
        exit 1
    }
}

#Requires -Version 5.1

# --------------------------------------------------
# Author: Gavin Stone (NinjaOne)
# Attribution: Luke Whitelock (NinjaOne) for the Authentication Functions, Kyle Bohlander (Ninja One) - Added PS 5.1 Support
# Date: 22nd January 2024
# Description: Utilized in NinjaOne to schedule recurring maintenance mode for devices
# Version: 1.0
# --------------------------------------------------

# User editable variables:
$NinjaOneInstance = '' # This varies depending on region or environment. For example, if you are in the US, this would be 'app.ninjarmm.com'
$NinjaOneClientId = ''
$NinjaOneClientSecret = ''
$SkipCustomFieldTest = $false


# PRE REQUISITES --------------------------------------------------
# For API access you will need to generate a Client ID and Client Secret in NinjaOne
# Go to Administration > Apps > API and the Client App IDs tab. 
# Click the 'Add' button in the top right
# For application platform, select API Services (machine-to-machine)
# Name your token something you will recognize (e.g. Recurring Maintenance Mode Script API Token)
# Set Redirect URI to http://localhost
# Set the scopes to monitoring and management
# For allowed grant types, select client credentials only
# Click save in the top right, enter 2FA prompt if required.
# You will be presented with the client secret credential only once. Use the copy icon to copy this into the clipboard and store it somewhere secure. Enter this into the $NinjaOneClientSecret variable above
# Close this window, which will take you back to the Client App IDs tab. Click copy on the Client ID and store this somewhere secure. Enter this into the $NinjaOneClientId variable above

# You will need to create the following custom fields in NinjaOne
# ALL FIELDS MUST BE SET TO DEFINITION SCOPE OF DEVICE AND LOCATION AND ORGANIZATION
# Label: Recurring Maintenance - Enable Recurring Schedule | Name: recurringMaintenanceEnableRecurringSchedule | Type: Checkbox | API Permissions: Read/Write
# Description: A checkbox that is required to be set for a device to have recurring maintenance mode applied programmatically.
# Tooltip: Check this box to enable recurring maintenance mode for this device.

# Label: Recurring Maintenance - Time to Start (24h Format) | Name: recurringMaintenanceTimeToStart24hFormat | Type: Time | API Permissions: Read/Write
# Description: A time field is required to be set for a device to have recurring maintenance mode applied programmatically.
# Tooltip: Set the time that you want the recurring maintenance mode to start.

# Label: Recurring Maintenance - Total Minutes for Maintenance Mode | Name: recurringMaintenanceTotalMinutesForMaintenanceMode | Type: Integer | API Permissions: Read/Write | Set a Numeric Range of 1 - 1440
# Description: The number of minutes that the device will stay in Maintenance Mode for in the recurring schedule.
# Tooltip: Set the number of minutes that the device will stay in Maintenance Mode for. The maximum is 1440 minutes (24 hours)

# Label: Recurring Maintenance - Select Day | Name: recurringMaintenanceSelectDay | Type: Multi-select | API Permissions: Read/Write | Options: Every Sunday, Every Monday, Every Tuesday, Every Wednesday, Every Thursday, Every Friday, Every Saturday
# Description: The days of the week that the recurring maintenance mode will be set for.
# Tooltip: Select the days of the week that you want the recurring maintenance mode to be set for.

# Label: Recurring Maintenance - Date to Stop Applying Recurring Schedule | Name: recurringMaintenanceDateToStopApplyingRecurringSchedule | Type: Date | API Permissions: Read/Write
# Description: The date that the recurring maintenance mode will stop being applied.
# Tooltip: Select the date that you want the recurring maintenance mode to stop being applied. Please note, if a date is set here recurring maintenance mode will stop being applied on and from this date.

# Label: Recurring Maintenance - Suppress Scripting and Tasks | Name: recurringMaintenanceSuppressScriptingAndTasks | Type: Checkbox | API Permissions: Read/Write
# Description: When set, this will suppress all scripting and tasks on the device during the recurring maintenance mod.
# Tooltip: Check this box to suppress all scripting and tasks on the device during the recurring maintenance mode.

# Label: Recurring Maintenance - Suppress Patching | Name: recurringMaintenanceSuppressPatching | Type: Checkbox | API Permissions: Read/Write
# Description: When set, this will suppress all patching on the device during the recurring maintenance mode.
# Tooltip: Check this box to suppress all patching on the device during the recurring maintenance mode.

# Label: Recurring Maintenance - Suppress AV Scans | Name: recurringMaintenanceSuppressAvScans | Type: Checkbox | API Permissions: Read/Write
# Description: When set, this will suppress all AV Scans on the device during the recurring maintenance mode.
# Tooltip: Check this box to suppress all AV Scans on the device during the recurring maintenance mode.

# Label: Recurring Maintenance - Suppress Condition Based Alerting | Name: recurringMaintenanceSuppressConditionBasedAlerting | Type: Checkbox | API Permissions: Read/Write
# Description: When set, this will suppress all Condition Based Alerting on the device during the recurring maintenance mode.
# Tooltip: Check this box to suppress all Condition Based Alerting on the device during the recurring maintenance mode.

# Label: Recurring Maintenance - Last Result | Name: recurringMaintenanceLastResult | Type: Text (Multiline) | API Permissions: Read/Write
# Description: This is set automatically by the script to store relevant information from the script about setting Maintenance Mode.
# Tooltip: You do not need to set this field; the script populates it automatically.
#
# --- NEW FIELDS (optional; enable Daily, Weekly, Monthly, Monthly Day-of-Week scheduling) ---
# If recurringMaintenanceScheduleType is not set, the script treats the schedule as Weekly and uses recurringMaintenanceSelectDay (legacy).
#
# Label: Recurring Maintenance - Schedule Type | Name: recurringMaintenanceScheduleType | Type: Single-select | API Permissions: Read/Write | Options: Daily, Weekly, Monthly, MonthlyDayOfWeek
# Label: Recurring Maintenance - Day of Week | Name: recurringMaintenanceDayOfWeek | Type: Multi-select | API Permissions: Read/Write | Options: Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday
# Label: Recurring Maintenance - Day of Month | Name: recurringMaintenanceDayOfMonth | Type: Integer | API Permissions: Read/Write | Numeric Range: 1 - 31
# Label: Recurring Maintenance - Monthly Day of Week | Name: recurringMaintenanceMonthlyDayOfWeek | Type: Single-select | API Permissions: Read/Write | Options: Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday
# Label: Recurring Maintenance - Monthly Occurrence | Name: recurringMaintenanceMonthlyOccurrence | Type: Single-select | API Permissions: Read/Write | Options: 1, 2, 3, 4, Last


# Functions for Authentication
function Get-NinjaOneToken {
    [CmdletBinding()]
    param()

    if ($Script:NinjaOneInstance -and $Script:NinjaOneClientID -and $Script:NinjaOneClientSecret ) {
        if ($Script:NinjaTokenExpiry -and (Get-Date) -lt $Script:NinjaTokenExpiry) {
            return $Script:NinjaToken
        }
        else {

            if ($Script:NinjaOneRefreshToken) {
                $Body = @{
                    'grant_type'    = 'refresh_token'
                    'client_id'     = $Script:NinjaOneClientID
                    'client_secret' = $Script:NinjaOneClientSecret
                    'refresh_token' = $Script:NinjaOneRefreshToken
                }
            }
            else {

                $body = @{
                    grant_type    = 'client_credentials'
                    client_id     = $Script:NinjaOneClientID
                    client_secret = $Script:NinjaOneClientSecret
                    scope         = 'monitoring management'
                }
            }

            $token = Invoke-RestMethod -Uri "https://$($Script:NinjaOneInstance -replace '/ws','')/ws/oauth/token" -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing
    
            $Script:NinjaTokenExpiry = (Get-Date).AddSeconds($Token.expires_in)
            $Script:NinjaToken = $token
            

            Write-Host 'Fetched New Token'
            return $token
        }
        else {
            Throw 'Please run Connect-NinjaOne first'
        }
    }

}

function Connect-NinjaOne {
    [CmdletBinding()]
    param (
        [Parameter(mandatory = $true)]
        $NinjaOneInstance,
        [Parameter(mandatory = $true)]
        $NinjaOneClientID,
        [Parameter(mandatory = $true)]
        $NinjaOneClientSecret,
        $NinjaOneRefreshToken
    )

    $Script:NinjaOneInstance = $NinjaOneInstance
    $Script:NinjaOneClientID = $NinjaOneClientID
    $Script:NinjaOneClientSecret = $NinjaOneClientSecret
    $Script:NinjaOneRefreshToken = $NinjaOneRefreshToken
    

    try {
        $Null = Get-NinjaOneToken -ea Stop
    }
    catch {
        Throw "Failed to Connect to NinjaOne: $_"
    }

}

function Invoke-NinjaOneRequest {
    param(
        $Method,
        $Body,
        $InputObject,
        $Path,
        $QueryParams,
        [Switch]$Paginate,
        [Switch]$AsArray
    )

    $Token = Get-NinjaOneToken

    if ($InputObject) {
        if ($AsArray) {
            $Body = $InputObject | ConvertTo-Json -Depth 100
            if (($InputObject | Measure-Object).count -eq 1 ) {
                $Body = '[' + $Body + ']'
            }
        }
        else {
            $Body = $InputObject | ConvertTo-Json -Depth 100
        }
    }

    try {
        if ($Method -in @('GET', 'DELETE')) {
            if ($Paginate) {
            
                $After = 0
                $PageSize = 1000
                $NinjaResult = do {
                    $Result = Invoke-WebRequest -Uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)?pageSize=$PageSize&after=$After$(if ($QueryParams){"&$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json' -UseBasicParsing
                    $Result
                    $ResultCount = ($Result.id | Measure-Object -Maximum)
                    $After = $ResultCount.maximum
    
                } while ($ResultCount.count -eq $PageSize)
            }
            else {
                $NinjaResult = Invoke-WebRequest -Uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json; charset=utf-8' -UseBasicParsing
            }

        }
        elseif ($Method -in @('PATCH', 'PUT', 'POST')) {
            $NinjaResult = Invoke-WebRequest -Uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -Body $Body -ContentType 'application/json; charset=utf-8' -UseBasicParsing
        }
        else {
            Throw 'Unknown Method'
        }
    }
    catch {
        Throw "Error Occured: $_"
    }

    try {
        return $NinjaResult.content | ConvertFrom-Json -ea stop
    }
    catch {
        return $NinjaResult.content
    }

}

# Parse day name to [System.DayOfWeek] (case-insensitive). Accepts "Monday" or "Every Monday" style.
function Get-DayOfWeekFromName {
    param([string]$DayName)
    if ([string]::IsNullOrWhiteSpace($DayName)) { return $null }
    $trimmed = $DayName.Trim()
    if ($trimmed.StartsWith('Every ')) { $trimmed = $trimmed.Substring(6).Trim() }
    foreach ($v in [Enum]::GetValues([System.DayOfWeek])) {
        if ([string]::Equals($v.ToString(), $trimmed, [StringComparison]::OrdinalIgnoreCase)) { return $v }
    }
    return $null
}

# Normalize occurrence string to "1", "2", "3", "4", or "last". Returns $null if invalid.
function Get-NormalizedMonthlyOccurrence {
    param([string]$Occurrence)
    if ([string]::IsNullOrWhiteSpace($Occurrence)) { return $null }
    $s = $Occurrence.Trim().ToLowerInvariant()
    if ($s -eq 'last') { return 'last' }
    $n = 0
    if ([int]::TryParse($s, [ref]$n) -and $n -ge 1 -and $n -le 4) { return [string]$n }
    return $null
}

# Get the date (midnight) of the nth occurrence of a weekday in the given month. Occurrence: "1", "2", "3", "4", or "last".
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
    if ($result.Month -ne $Month) { $result = $result.AddDays(-7) }
    return $result.Date
}

# Returns the next maintenance start [datetime] for the given schedule. TimeUnix is in milliseconds (Ninja time field).
# Backward compat: if ScheduleType is null/empty, treats as Weekly and uses RecurringMaintenanceSelectDay ("Every Monday" etc.).
function Get-NextMaintenanceOccurrence {
    param(
        [string]$ScheduleType,
        [string]$TimeUnix,
        [array]$DayOfWeek,
        [object]$DayOfMonth,
        [string]$MonthlyDayOfWeek,
        [string]$MonthlyOccurrence,
        [array]$RecurringMaintenanceSelectDay
    )

    $now = Get-Date
    $timeZone = [TimeZoneInfo]::Local
    $utcEpoch = Get-Date "1970-01-01 00:00:00"
    $utcTime = $utcEpoch.AddMilliseconds([double]$TimeUnix)
    $startTimeOfDay = [TimeZoneInfo]::ConvertTimeFromUtc($utcTime, $timeZone).TimeOfDay

    # Resolve effective schedule type and weekly days (backward compat)
    $effectiveMode = if ([string]::IsNullOrWhiteSpace($ScheduleType)) { 'Weekly' } else { $ScheduleType.Trim().ToLowerInvariant() }
    if ($effectiveMode -eq 'monthlydayofweek') { $effectiveMode = 'MonthlyDayOfWeek' }

    $weeklyDays = @()
    if ($DayOfWeek -and $DayOfWeek.Count -gt 0) {
        $weeklyDays = @($DayOfWeek | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    if ($weeklyDays.Count -eq 0 -and $RecurringMaintenanceSelectDay -and $RecurringMaintenanceSelectDay.Count -gt 0) {
        $dayLookup = @{
            'Every Sunday' = 'Sunday'; 'Every Monday' = 'Monday'; 'Every Tuesday' = 'Tuesday'
            'Every Wednesday' = 'Wednesday'; 'Every Thursday' = 'Thursday'; 'Every Friday' = 'Friday'; 'Every Saturday' = 'Saturday'
        }
        foreach ($d in $RecurringMaintenanceSelectDay) {
            $mapped = $dayLookup[$d]
            if ($mapped) { $weeklyDays += $mapped }
        }
    }

    switch ($effectiveMode) {
        'daily' {
            $todayOccurrence = $now.Date + $startTimeOfDay
            if ($todayOccurrence -gt $now) { return $todayOccurrence }
            return $todayOccurrence.AddDays(1)
        }
        'weekly' {
            if (-not $weeklyDays -or $weeklyDays.Count -eq 0) { return $null }
            $occurrences = foreach ($dowName in $weeklyDays) {
                $targetDayVal = Get-DayOfWeekFromName -DayName $dowName
                if ($null -eq $targetDayVal) { continue }
                $targetDay = [int]$targetDayVal
                $currentDay = [int]$now.DayOfWeek
                $daysToAdd = $targetDay - $currentDay
                if ($daysToAdd -lt 0 -or ($daysToAdd -eq 0 -and ($now.TimeOfDay -ge $startTimeOfDay))) { $daysToAdd += 7 }
                $now.Date.AddDays($daysToAdd) + $startTimeOfDay
            }
            $candidates = $occurrences | Where-Object { $_ -gt $now } | Sort-Object
            if ($candidates) { return $candidates[0] }
            $nextWeekBase = $now.Date.AddDays(7)
            $occurrencesNext = foreach ($dowName in $weeklyDays) {
                $targetDayVal = Get-DayOfWeekFromName -DayName $dowName
                if ($null -eq $targetDayVal) { continue }
                $targetDay = [int]$targetDayVal
                $currentDayNext = [int]$nextWeekBase.DayOfWeek
                $daysToAdd = ($targetDay - $currentDayNext + 7) % 7
                if ($daysToAdd -eq 0) { $daysToAdd = 7 }
                $nextWeekBase.AddDays($daysToAdd) + $startTimeOfDay
            }
            if ($occurrencesNext) { return ($occurrencesNext | Sort-Object)[0] }
            return $null
        }
        'monthly' {
            $dom = 0
            if ($null -eq $DayOfMonth -or -not [int]::TryParse([string]$DayOfMonth, [ref]$dom) -or $dom -lt 1 -or $dom -gt 31) { return $null }
            $year = $now.Year
            $month = $now.Month
            try {
                $occurrence = Get-Date -Year $year -Month $month -Day $dom -Hour $startTimeOfDay.Hours -Minute $startTimeOfDay.Minutes -Second $startTimeOfDay.Seconds
            } catch { $occurrence = $null }
            if ($occurrence -and $occurrence -gt $now) { return $occurrence }
            $nextMonth = $now.AddMonths(1)
            try {
                $occurrence = Get-Date -Year $nextMonth.Year -Month $nextMonth.Month -Day $dom -Hour $startTimeOfDay.Hours -Minute $startTimeOfDay.Minutes -Second $startTimeOfDay.Seconds
            } catch { return $null }
            return $occurrence
        }
        'monthlydayofweek' {
            if ([string]::IsNullOrWhiteSpace($MonthlyDayOfWeek) -or [string]::IsNullOrWhiteSpace($MonthlyOccurrence)) { return $null }
            $dowVal = Get-DayOfWeekFromName -DayName $MonthlyDayOfWeek
            if ($null -eq $dowVal) { return $null }
            $occNorm = Get-NormalizedMonthlyOccurrence -Occurrence $MonthlyOccurrence
            if (-not $occNorm) { return $null }
            $occurrenceDate = Get-NthWeekdayInMonth -Year $now.Year -Month $now.Month -DayOfWeek $dowVal -Occurrence $occNorm
            $occurrenceDt = $occurrenceDate + $startTimeOfDay
            if ($occurrenceDt -gt $now) { return $occurrenceDt }
            $nextMonth = $now.AddMonths(1)
            $occurrenceDate = Get-NthWeekdayInMonth -Year $nextMonth.Year -Month $nextMonth.Month -DayOfWeek $dowVal -Occurrence $occNorm
            return $occurrenceDate + $startTimeOfDay
        }
        default { return $null }
    }
}

function Test-CustomField {
    param (
        [string]$CustomFieldName,
        [string]$CustomFieldType,
        [array]$CustomFieldList
    )

    $field = $CustomFieldList | Where-Object { $_.name -eq $CustomFieldName }

    if ($null -eq $field -or $field.count -eq 0) {
        Write-Output "Custom Field $CustomFieldName does not exist in Ninja. Please create this custom field and set it to a $CustomFieldType type"
        exit 1
    }
    if ($field.type -ne $CustomFieldType) {
        Write-Output "Custom Field $CustomFieldName is not set to a $CustomFieldType type. Please set this to a $CustomFieldType type"
        exit 1
    }
    if ($field.apiPermission -ne 'READ_WRITE') {
        Write-Output "Custom Field $CustomFieldName is not set to READ_WRITE on the API Permissions. Please set this to READ_WRITE"
        exit 1
    }
}

# END FUNCTIONS --------------------------------------------------

# Error Tracking
$ErrorLog = New-Object System.Collections.ArrayList

# Connect to NinjaOne API
try {
    Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientId -NinjaOneClientSecret $NinjaOneClientSecret
}
catch {
    Write-Output "Failed to connect to NinjaOne API: $_"
    exit 1
}

# Lets get a list of Custom Fields - We can make sure we have the ones we need
# If skipcustomfieldtest is set to true, we will skip this test
if ($SkipCustomFieldTest -eq $false) {
    $CustomFieldList = invoke-ninjaonerequest -method GET -path "device-custom-fields" -Paginate
    Test-CustomField -CustomFieldName 'recurringMaintenanceEnableRecurringSchedule' -CustomFieldType 'CHECKBOX' -CustomFieldList $CustomFieldList
    Test-CustomField -CustomFieldName 'recurringMaintenanceTimeToStart24hFormat' -CustomFieldType 'TIME' -CustomFieldList $CustomFieldList
    Test-CustomField -CustomFieldName 'recurringMaintenanceTotalMinutesForMaintenanceMode' -CustomFieldType 'NUMERIC' -CustomFieldList $CustomFieldList
    Test-CustomField -CustomFieldName 'recurringMaintenanceSelectDay' -CustomFieldType 'MULTI_SELECT' -CustomFieldList $CustomFieldList
    Test-CustomField -CustomFieldName 'recurringMaintenanceDateToStopApplyingRecurringSchedule' -CustomFieldType 'DATE' -CustomFieldList $CustomFieldList
    Test-CustomField -CustomFieldName 'recurringMaintenanceSuppressScriptingAndTasks' -CustomFieldType 'CHECKBOX' -CustomFieldList $CustomFieldList
    Test-CustomField -CustomFieldName 'recurringMaintenanceSuppressPatching' -CustomFieldType 'CHECKBOX' -CustomFieldList $CustomFieldList
    Test-CustomField -CustomFieldName 'recurringMaintenanceSuppressAvScans' -CustomFieldType 'CHECKBOX' -CustomFieldList $CustomFieldList
    Test-CustomField -CustomFieldName 'recurringMaintenanceSuppressConditionBasedAlerting' -CustomFieldType 'CHECKBOX' -CustomFieldList $CustomFieldList
    Test-CustomField -CustomFieldName 'recurringMaintenanceLastResult' -CustomFieldType 'TEXT_MULTILINE' -CustomFieldList $CustomFieldList
}

# Utilize the scoped custom fields API to get the custom fields we need
try {
    $MaintenanceModeCustomFields = Invoke-NinjaOneRequest -Method GET -Path 'queries/scoped-custom-fields?fields=recurringMaintenanceDateToStopApplyingRecurringSchedule,recurringMaintenanceEnableRecurringSchedule,recurringMaintenanceLastResult,recurringMaintenanceSelectDay,recurringMaintenanceSuppressAvScans,recurringMaintenanceSuppressConditionBasedAlerting,recurringMaintenanceSuppressPatching,recurringMaintenanceSuppressScriptingAndTasks,recurringMaintenanceTotalMinutesForMaintenanceMode,recurringMaintenanceTimeToStart24hFormat,recurringMaintenanceScheduleType,recurringMaintenanceDayOfWeek,recurringMaintenanceDayOfMonth,recurringMaintenanceMonthlyDayOfWeek,recurringMaintenanceMonthlyOccurrence,' -Paginate | Select-Object -ExpandProperty results    
}
catch {
    Write-Output "Failed to get scoped custom fields from NinjaOne API: $_"
    exit 1
}

# If the result of this has a count of 0, then we need to exit as there are no custom fields that match the criteria
if ($MaintenanceModeCustomFields.count -eq 0) {
    Write-Output "No custom fields found set that match the criteria for a recurring Maintenance Mode. Exiting..."
    exit 1
}

# Lets get a list of all the organizations that have recurringMaintenanceEnableRecurringSchedule set to true that are under the scope of ORGANIZATION
$OrganizationsWithRecurringMaintenanceEnabled = $MaintenanceModeCustomFields | Where-Object { ($_.fields.recurringMaintenanceEnableRecurringSchedule -eq 'True') -and $_.scope -eq 'ORGANIZATION' }

# Lets get a list of all the locations that have recurringMaintenanceEnableRecurringSchedule set to true that are under the scope of LOCATION
$LocationsWithRecurringMaintenanceEnabled = $MaintenanceModeCustomFields | Where-Object { ($_.fields.recurringMaintenanceEnableRecurringSchedule -eq 'True') -and $_.scope -eq 'LOCATION' }

# Lets get a list of all the devices that have recurringMaintenanceEnableRecurringSchedule set to true that are under the scope of NODE
$DevicesWithRecurringMaintenanceEnabled = $MaintenanceModeCustomFields | Where-Object { ($_.fields.recurringMaintenanceEnableRecurringSchedule -eq 'True') -and $_.scope -eq 'NODE' }

# Build up an array so we can use to store the data we need to use later
$FinalArrayActions = New-Object System.Collections.ArrayList
$OrganizationDevices = New-Object System.Collections.ArrayList
$LocationDevices = New-Object System.Collections.ArrayList

# So we're going to have to loop through every device manually because the scoped report doesn't give us the information we need
# Lets get a list of every node from the org, location and node level.
if (($OrganizationsWithRecurringMaintenanceEnabled | Measure-Object).Count -gt 0) {
    # Loop through each of these organizations
    foreach ($organization in $OrganizationsWithRecurringMaintenanceEnabled) {
        # Get all devices in this organization
        $OrgCall = Invoke-NinjaOneRequest -Method GET -Path "devices" -QueryParams "df=org = $($organization.entityId)" -Paginate
        $null = $OrganizationDevices.Add($OrgCall)
    }
}

# Locations Next
if (($LocationsWithRecurringMaintenanceEnabled | Measure-Object).Count -gt 0) {
    # Loop through each of these locations
    foreach ($location in $LocationsWithRecurringMaintenanceEnabled) {
        # Get all devices in this location
        $LocCall = Invoke-NinjaOneRequest -Method GET -Path "devices" -QueryParams "df=loc = $($location.entityId)" -Paginate
        $null = $LocationDevices.Add($LocCall)
    }
}

$UniqueDeviceIDs = New-Object System.Collections.Generic.List[Object]
$DevicesWithRecurringMaintenanceEnabled.entityId | Get-Unique | ForEach-Object { $UniqueDeviceIDs.Add($_) }
$LocationDevices.id | Get-Unique | ForEach-Object { $UniqueDeviceIDs.Add($_) }
$OrganizationDevices.id | Get-Unique | ForEach-Object { $UniqueDeviceIDs.Add($_) }
$UniqueDeviceIDs = $UniqueDeviceIDs | Sort-Object -Unique
foreach ($node in $UniqueDeviceIDs) {
    # Get device information
    $NodeResult = Invoke-NinjaOneRequest -Method GET -Path "device/$($node)"

    # Get the custom field information
    $NodeCustomField = Invoke-NinjaOneRequest -Method GET -Path "device/$($node)/custom-fields" -Paginate -QueryParams "withInheritance=true"

    $DeviceObject = [PSCustomObject]@{
        NodeID                                                  = $($Node)
        NodeOrganizationId                                      = $($NodeResult.organizationId)
        NodeLocationId                                          = $($NodeResult.locationId)
        NodeClass                                               = $($NodeResult.nodeClass)
        CurrentMaintenanceModeStatus                            = $($NodeResult.maintenance)
        SystemName                                              = $($NodeResult.systemName)
        Offline                                                 = $($NodeResult.offline)
        recurringMaintenanceEnableRecurringSchedule             = $($NodeCustomField.recurringMaintenanceEnableRecurringSchedule)
        recurringMaintenanceTimeToStart24hFormat                = $($NodeCustomField.recurringMaintenanceTimeToStart24hFormat)
        recurringMaintenanceTotalMinutesForMaintenanceMode       = $($NodeCustomField.recurringMaintenanceTotalMinutesForMaintenanceMode)
        recurringMaintenanceSelectDay                           = $($NodeCustomField.recurringMaintenanceSelectDay)
        recurringMaintenanceDateToStopApplyingRecurringSchedule = $($NodeCustomField.recurringMaintenanceDateToStopApplyingRecurringSchedule)
        recurringMaintenanceSuppressScriptingAndTasks           = $($NodeCustomField.recurringMaintenanceSuppressScriptingAndTasks)
        recurringMaintenanceSuppressPatching                    = $($NodeCustomField.recurringMaintenanceSuppressPatching)
        recurringMaintenanceSuppressAvScans                     = $($NodeCustomField.recurringMaintenanceSuppressAvScans)
        recurringMaintenanceSuppressConditionBasedAlerting      = $($NodeCustomField.recurringMaintenanceSuppressConditionBasedAlerting)
        recurringMaintenanceScheduleType                        = $($NodeCustomField.recurringMaintenanceScheduleType)
        recurringMaintenanceDayOfWeek                           = $($NodeCustomField.recurringMaintenanceDayOfWeek)
        recurringMaintenanceDayOfMonth                          = $($NodeCustomField.recurringMaintenanceDayOfMonth)
        recurringMaintenanceMonthlyDayOfWeek                    = $($NodeCustomField.recurringMaintenanceMonthlyDayOfWeek)
        recurringMaintenanceMonthlyOccurrence                   = $($NodeCustomField.recurringMaintenanceMonthlyOccurrence)
    }
    $null = $FinalArrayActions.Add($DeviceObject)
}

# We now have a list of all the devices, including taking into account inheritance, that have recurringMaintenanceEnableRecurringSchedule set to true
# It's time to do some sanity checks to make sure that all data is actionable.

$CountDeviceMarkedForMaintenance = 0
$CountDeviceInMaintenanceMode = 0
$CountDeviceAlreadyScheduled = 0
$CountDeviceWhereMaintenanceGotSet = 0
$CountDevicePastApplyDate = 0

# We will log each node as they will need removing from the recurring schedule
$NaughtyNodes = New-Object System.Collections.ArrayList

# This validates some of the node information
foreach ($node in $FinalArrayActions) {
    # Make sure recurringMaintenanceEnableRecurringSchedule is set to true
    if ($node.recurringMaintenanceEnableRecurringSchedule -ne $true) {
        $null = $NaughtyNodes.Add($node.NodeID)
        continue
    }
    else {
        $CountDeviceMarkedForMaintenance++
    }

    # Resolve effective schedule type (backward compat: null = Weekly using recurringMaintenanceSelectDay)
    $effectiveScheduleType = if ([string]::IsNullOrWhiteSpace($node.recurringMaintenanceScheduleType)) { 'Weekly' } else { $node.recurringMaintenanceScheduleType.Trim() }

    # Validate required fields per schedule type
    $scheduleError = $null
    switch ($effectiveScheduleType) {
        'Daily' { }
        'Weekly' {
            $hasWeeklyDays = ($node.recurringMaintenanceDayOfWeek -and @($node.recurringMaintenanceDayOfWeek).Count -gt 0) -or ($node.recurringMaintenanceSelectDay -and @($node.recurringMaintenanceSelectDay).Count -gt 0)
            if (-not $hasWeeklyDays) {
                $scheduleError = "Weekly schedule requires recurringMaintenanceDayOfWeek or recurringMaintenanceSelectDay (e.g. Every Monday)."
            }
        }
        'Monthly' {
            $dom = 0
            if ($null -eq $node.recurringMaintenanceDayOfMonth -or -not [int]::TryParse([string]$node.recurringMaintenanceDayOfMonth, [ref]$dom) -or $dom -lt 1 -or $dom -gt 31) {
                $scheduleError = "Monthly schedule requires recurringMaintenanceDayOfMonth (1-31)."
            }
        }
        'MonthlyDayOfWeek' {
            if ([string]::IsNullOrWhiteSpace($node.recurringMaintenanceMonthlyDayOfWeek) -or [string]::IsNullOrWhiteSpace($node.recurringMaintenanceMonthlyOccurrence)) {
                $scheduleError = "MonthlyDayOfWeek schedule requires recurringMaintenanceMonthlyDayOfWeek and recurringMaintenanceMonthlyOccurrence (1, 2, 3, 4, or Last)."
            } elseif ($null -eq (Get-NormalizedMonthlyOccurrence -Occurrence $node.recurringMaintenanceMonthlyOccurrence)) {
                $scheduleError = "recurringMaintenanceMonthlyOccurrence must be 1, 2, 3, 4, or Last."
            }
        }
        default {
            $scheduleError = "recurringMaintenanceScheduleType must be Daily, Weekly, Monthly, or MonthlyDayOfWeek."
        }
    }
    if ($scheduleError) {
        $null = $ErrorLog.Add("Node: $($node.NodeID). $scheduleError")
        $null = $NaughtyNodes.Add($node.NodeID)
        $UpdateBody = @{ "recurringMaintenanceLastResult" = "[Error] $(Get-Date) - $scheduleError" } | ConvertTo-Json
        Invoke-NinjaOneRequest -Method PATCH -Path "device/$($node.NodeID)/custom-fields" -Body $UpdateBody
        continue
    }

    # Check the window length
    if ($node.recurringMaintenanceTotalMinutesForMaintenanceMode -gt 1440 -or $node.recurringMaintenanceTotalMinutesForMaintenanceMode -lt 1) {
        $null = $ErrorLog.Add("Node: $($node.NodeID). recurringMaintenanceTotalMinutesForMaintenanceMode value is incorrect. Needs to be a value between 1 and 1440.")
        $null = $NaughtyNodes.Add($node.NodeID)
        $UpdateBody = @{
            "recurringMaintenanceLastResult" = "[Error] $(Get-Date) - recurringMaintenanceTotalMinutesForMaintenanceMode value is incorrect. Needs to be a value between 1 and 1440."
        } | ConvertTo-Json 
        Invoke-NinjaOneRequest -Method PATCH -Path "device/$($node.NodeID)/custom-fields" -Body $UpdateBody
        continue
    }
}


$CleanedObjects = $FinalArrayActions | Where-Object { $_.NodeID -notin $NaughtyNodes }

# Now lets loop through the final devices and set some Maintenance MODE!
foreach ($node in $CleanedObjects) {
    Write-Host $node.CurrentMaintenanceModeStatus.Status
    # We don't want to touch any device that is already in maintenance mode
    if ($node.CurrentMaintenanceModeStatus.Status -eq 'IN_MAINTENANCE') {
        $CountDeviceInMaintenanceMode++
        # Lets set the custom field to indicate we are going to not touch this device
        $UpdateBody = @{
            "recurringMaintenanceLastResult" = "[Info] $(Get-Date) - This device was due to be processed with a recurring scheduled maintenance mode, but was already in maintenance mode. No action taken."
        } | ConvertTo-Json 
        Invoke-NinjaOneRequest -Method PATCH -Path "device/$($node.NodeID)/custom-fields" -Body $UpdateBody
        continue
    }
    
    # Calculate the next start date from schedule type and options (backward compat: null ScheduleType = Weekly + recurringMaintenanceSelectDay)
    $NextRecurringScheduledStartDate = Get-NextMaintenanceOccurrence -ScheduleType $node.recurringMaintenanceScheduleType -TimeUnix $node.recurringMaintenanceTimeToStart24hFormat -DayOfWeek $node.recurringMaintenanceDayOfWeek -DayOfMonth $node.recurringMaintenanceDayOfMonth -MonthlyDayOfWeek $node.recurringMaintenanceMonthlyDayOfWeek -MonthlyOccurrence $node.recurringMaintenanceMonthlyOccurrence -RecurringMaintenanceSelectDay $node.recurringMaintenanceSelectDay

    if (-not $NextRecurringScheduledStartDate) {
        $UpdateBody = @{ "recurringMaintenanceLastResult" = "[Error] $(Get-Date) - Could not calculate next maintenance occurrence; check schedule type and options." } | ConvertTo-Json
        Invoke-NinjaOneRequest -Method PATCH -Path "device/$($node.NodeID)/custom-fields" -Body $UpdateBody
        continue
    }

    # Now we have the start date, we need to calculate the end date based on adding recurringMaintenanceTotalMinutesForMaintenanceMode to $NextRecurringScheduledStartDate
    $NextRecurringScheduledEndDate = $NextRecurringScheduledStartDate.AddMinutes($node.recurringMaintenanceTotalMinutesForMaintenanceMode)

    # We have a start and end date - lets make them Unix time (seconds)
    $UniversalStartDate = $NextRecurringScheduledStartDate.ToUniversalTime()
    $UniversalEndDate = $NextRecurringScheduledEndDate.ToUniversalTime()

    $StartTimeSpan = New-TimeSpan (Get-Date "1970-01-01 00:00:00") $UniversalStartDate
    $EndTimeSpan = New-TimeSpan (Get-Date "1970-01-01 00:00:00") $UniversalEndDate

    $UnixStartTime = $StartTimeSpan.TotalSeconds
    $UnixEndTime = $EndTimeSpan.TotalSeconds

    $stopDateLocal = $null
    if ($node.recurringMaintenanceDateToStopApplyingRecurringSchedule) {
        $UTC = (Get-Date "1970-01-01 00:00:00").AddMilliSeconds($node.recurringMaintenanceDateToStopApplyingRecurringSchedule)
        $TimeZone = [TimeZoneInfo]::Local
        $stopDateLocal = [TimeZoneInfo]::ConvertTimeFromUtc($UTC, $TimeZone)
    }

    # We don't want to touch any device if recurringMaintenanceDateToStopApplyingRecurringSchedule is set and is in the past
    if ($stopDateLocal -and ($stopDateLocal -lt $NextRecurringScheduledEndDate)) {
        $CountDevicePastApplyDate++
        # Lets set the custom field to indicate we are going to not touch this device
        $UpdateBody = @{
            "recurringMaintenanceLastResult" = "[Alert] $(Get-Date) - This device was due to be processed with a recurring scheduled maintenance mode, but recurringMaintenanceDateToStopApplyingRecurringSchedule is set and occurs before the end time of the recurring maintenance window. No action taken."
        } | ConvertTo-Json 
        Invoke-NinjaOneRequest -Method PATCH -Path "device/$($node.NodeID)/custom-fields" -Body $UpdateBody
        continue
    }

    # Now we need to determine if there is a schedule already set for this device - if there we want to check it matches the maintenance mode we are going to set
    if ((![string]::IsNullOrEmpty($node.CurrentMaintenanceModeStatus.start)) -or (![string]::IsNullOrEmpty($node.CurrentMaintenanceModeStatus.end))) {
        # We should only ever get here if there is a future schedule set, but the device is not in maintenance mode
        $CountDeviceAlreadyScheduled++

        $UTCstart = (Get-Date "1970-01-01 00:00:00").AddSeconds($node.CurrentMaintenanceModeStatus.start)
        $UTCend = (Get-Date "1970-01-01 00:00:00").AddSeconds($node.CurrentMaintenanceModeStatus.end)
        
        $TimeZone = [TimeZoneInfo]::Local

        $CurrentMaintenanceStart = [TimeZoneInfo]::ConvertTimeFromUtc($UTCstart, $TimeZone)
        $CurrentMaintenanceEnd = [TimeZoneInfo]::ConvertTimeFromUtc($UTCend, $TimeZone)

        # Check if the schedules match, if they do we can skip this node
        if ($CurrentMaintenanceStart -eq $NextRecurringScheduledStartDate -and $CurrentMaintenanceEnd -eq $NextRecurringScheduledEndDate) {
            continue
        }
    }

    # Lets figure out the modes we need to disable
    $DisabledModes = New-Object System.Collections.ArrayList
    if ($node.recurringMaintenanceSuppressScriptingAndTasks -eq 'True') {
        $null = $DisabledModes.Add('TASKS')
    }
    if ($node.recurringMaintenanceSuppressPatching -eq 'True') {
        $null = $DisabledModes.Add('PATCHING')
    }
    if ($node.recurringMaintenanceSuppressAvScans -eq 'True') {
        $null = $DisabledModes.Add('AVSCANS')
    }
    if ($node.recurringMaintenanceSuppressConditionBasedAlerting -eq 'True') {
        $null = $DisabledModes.Add('ALERTS')
    }

    # Build the body for the Maintenance Mode Set
    $mmBody = @{
        disabledFeatures = $DisabledModes
        start            = $UnixStartTime
        end              = $UnixEndTime
    } | ConvertTo-Json
    
    $CountDeviceWhereMaintenanceGotSet++
    Invoke-NinjaOneRequest -Method PUT -Path "device/$($node.NodeID)/maintenance" -Body $mmBody

    # Use the Custom Field to say what we've done
    $cfUpdateBody = @{
        "recurringMaintenanceLastResult" = "[Info] $(Get-Date) - This device attempted to schedule Maintenance mode on $($NextRecurringScheduledStartDate) for $($node.recurringMaintenanceTotalMinutesForMaintenanceMode) minutes. The Maintenance Mode was set to $($DisabledModes -join ',')"
    } | ConvertTo-Json 
    Invoke-NinjaOneRequest -Method PATCH -Path "device/$($node.NodeID)/custom-fields" -Body $cfUpdateBody

}

# Output some statistics
Write-Output "Any Warning Messages: $($ErrorLog -join "`r`n")"
Write-Output "Total Devices Marked For Recurring Maintenance Mode: $CountDeviceMarkedForMaintenance"
Write-Output "Total Devices Already In Maintenance Mode: $CountDeviceInMaintenanceMode"
Write-Output "Total Devices Already Scheduled For Maintenance Mode: $CountDeviceAlreadyScheduled"
Write-Output "Total Devices where Maintenance Mode was actually configured and set in this run: $CountDeviceWhereMaintenanceGotSet"
Write-Output "Total Devices Skipped as they had a date set in recurringMaintenanceDateToStopApplyingRecurringSchedule that overlapped with the maintenance window: $CountDevicePastApplyDate"


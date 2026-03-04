<#
.SYNOPSIS
    Sets weekly maintenance windows (and other custom fields) in NinjaOne from a CSV.

.DESCRIPTION
    Reads a CSV and updates NinjaOne custom fields for organizations, locations, or devices.
    Primary use: set weekly maintenance window (day, start time, end time) per org, location, or device.
    CSV must have "level" (organization | location | device) and "name" to identify the target;
    all other columns are custom field name = value (e.g. maintenanceDay, maintenanceStart, maintenanceEnd).
    For location level, "name" must be "organizationname,locationname" (comma-separated).
    Maintenance start/end can be sent as HH:MM for TEXT custom fields (local time per device) or as Unix ms for TIME custom fields.

.PARAMETER CsvPath
    Full path to the CSV file.

.PARAMETER NinjaOneInstance
    NinjaOne instance host (e.g. app.ninjarmm.com).

.PARAMETER NinjaOneClientId
    OAuth client ID.

.PARAMETER NinjaOneClientSecret
    OAuth client secret.

.PARAMETER OverwriteEmptyValues
    If set, empty CSV values are sent as null/empty to clear existing custom field data.
    If not set, empty values are omitted from the update payload.

.PARAMETER MaintenanceWindowAsLocalTime
    When set (default), maintenanceStart and maintenanceEnd are sent as HH:MM strings (no conversion to Unix ms).
    Use with TEXT custom fields in NinjaOne so each device uses the same local time (e.g. 11:00-16:00 in each device's timezone).
    When not set, HH:MM is converted to Unix ms for TIME custom fields (same UTC moment globally).

.EXAMPLE
    .\Set-CustomFieldMaintenanceWindow.ps1 -CsvPath "C:\data\maintenance-windows.csv" -NinjaOneInstance "app.ninjarmm.com" -NinjaOneClientId "..." -NinjaOneClientSecret "..."
.EXAMPLE
    .\Set-CustomFieldMaintenanceWindow.ps1 -CsvPath ".\maintenance.csv" -MaintenanceWindowAsLocalTime
    Sends maintenanceStart and maintenanceEnd as HH:MM to TEXT custom fields for local-time-per-device behavior.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$NinjaOneInstance = $env:NinjaOneInstance,

    [Parameter(Mandatory = $false)]
    [string]$NinjaOneClientId = $env:NinjaOneClientId,

    [Parameter(Mandatory = $false)]
    [string]$NinjaOneClientSecret = $env:NinjaOneClientSecret,

    [Parameter(Mandatory = $false)]
    [bool]$OverwriteEmptyValues = $false,

    [Parameter(Mandatory = $false)]
    [bool]$MaintenanceWindowAsLocalTime = $true
)

$ErrorActionPreference = 'Stop'
$script:UpdatedCount = 0
$script:SkippedCount = 0
$script:FailedCount = 0

#region Validation

if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    Write-Error "CsvPath is required. Pass -CsvPath or set environment variables."
    exit 1
}
if (-not (Test-Path -LiteralPath $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

try {
    $rows = Import-Csv -Path $CsvPath
} catch {
    Write-Error "Failed to import CSV from $CsvPath. $_"
    exit 1
}

if ($null -eq $rows -or $rows.Count -eq 0) {
    Write-Error "CSV file is empty or has no data rows."
    exit 1
}

$requiredColumns = @('level', 'name')
$columnNames = $rows[0].PSObject.Properties.Name
foreach ($col in $requiredColumns) {
    if ($columnNames -notcontains $col) {
        Write-Error "CSV is missing required column '$col'. Required columns: level, name."
        exit 1
    }
}

if (-not $NinjaOneInstance -or -not $NinjaOneClientId -or -not $NinjaOneClientSecret) {
    Write-Error "NinjaOne credentials required. Set -NinjaOneInstance, -NinjaOneClientId, -NinjaOneClientSecret or env vars NinjaOneInstance, NinjaOneClientId, NinjaOneClientSecret."
    exit 1
}

Write-Host "CSV import: $($rows.Count) rows. OverwriteEmptyValues: $OverwriteEmptyValues"

#endregion

#region Authentication

$body = @{
    grant_type    = 'client_credentials'
    client_id     = $NinjaOneClientId
    client_secret = $NinjaOneClientSecret
    scope         = 'monitoring management'
}

$API_AuthHeaders = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
$API_AuthHeaders.Add('accept', 'application/json')
$API_AuthHeaders.Add('Content-Type', 'application/x-www-form-urlencoded')

try {
    $authToken = Invoke-RestMethod -Uri "https://$NinjaOneInstance/oauth/token" -Method POST -Headers $API_AuthHeaders -Body $body
    $accessToken = $authToken.access_token
} catch {
    Write-Error "Failed to obtain NinjaOne access token. $_"
    exit 1
}

$headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
$headers.Add('accept', 'application/json')
$headers.Add('Authorization', "Bearer $accessToken")

#endregion

#region API Helper

function Invoke-NinjaAPIRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$Method = 'GET',
        [Parameter(Mandatory = $true)]$Headers,
        [string]$Body = $null
    )
    $maxRetries = 3
    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        try {
            $params = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $Headers
                ContentType = 'application/json'
            }
            if ($Body) { $params.Body = $Body }
            return Invoke-RestMethod @params
        } catch {
            Write-Warning "API request to $Uri failed on attempt $($retryCount + 1): $_"
            $retryCount++
            Start-Sleep -Seconds 2
        }
    }
    throw "API request failed after $maxRetries attempts: $Uri"
}

#endregion

#region Paginated GET helpers

function Get-AllOrganizations {
    param([hashtable]$Headers, [string]$BaseUrl)
    $all = [System.Collections.ArrayList]@()
    $after = 0
    $pageSize = 1000
    do {
        $uri = "$BaseUrl/api/v2/organizations?pageSize=$pageSize&after=$after"
        $page = Invoke-NinjaAPIRequest -Uri $uri -Method GET -Headers $Headers
        $pageList = @($page)
        if ($pageList.Count -eq 0) { break }
        foreach ($p in $pageList) { [void]$all.Add($p) }
        $lastId = $pageList[-1].id
        if ($null -ne $lastId) { $after = $lastId }
        if ($pageList.Count -lt $pageSize) { break }
    } while ($true)
    return $all
}

function Get-AllLocations {
    param([hashtable]$Headers, [string]$BaseUrl)
    $all = [System.Collections.ArrayList]@()
    $after = 0
    $pageSize = 1000
    do {
        $uri = "$BaseUrl/api/v2/locations?pageSize=$pageSize&after=$after"
        $page = Invoke-NinjaAPIRequest -Uri $uri -Method GET -Headers $Headers
        $pageList = @($page)
        if ($pageList.Count -eq 0) { break }
        foreach ($p in $pageList) { [void]$all.Add($p) }
        $lastId = $pageList[-1].id
        if ($null -ne $lastId) { $after = $lastId }
        if ($pageList.Count -lt $pageSize) { break }
    } while ($true)
    return $all
}

function Get-AllDevices {
    param([hashtable]$Headers, [string]$BaseUrl)
    $uri = "$BaseUrl/api/v2/devices-detailed"
    $devices = Invoke-NinjaAPIRequest -Uri $uri -Method GET -Headers $Headers
    return @($devices)
}

#endregion

#region Fetch entities

$baseUrl = "https://$NinjaOneInstance"
Write-Host "Fetching organizations..."
$organizations = Get-AllOrganizations -Headers $headers -BaseUrl $baseUrl
Write-Host "  Found $($organizations.Count) organizations."
Write-Host "Fetching locations..."
$locations = Get-AllLocations -Headers $headers -BaseUrl $baseUrl
Write-Host "  Found $($locations.Count) locations."
Write-Host "Fetching devices..."
$devices = Get-AllDevices -Headers $headers -BaseUrl $baseUrl
Write-Host "  Found $($devices.Count) devices."

#endregion

#region Build custom field payload from row (exclude level and name)

function Convert-TimeStringToUnixMs {
    <#
    .SYNOPSIS
        Converts HH:MM (or H:MM) to Unix time in milliseconds for NinjaOne TIME custom fields.
        Uses 1970-01-01 at the given time in local timezone so only the time-of-day is encoded.
    #>
    param([string]$TimeString)
    if ([string]::IsNullOrWhiteSpace($TimeString)) { return $null }
    $s = $TimeString.Trim()
    if ($s -match '^\d+$') {
        return [long]$s
    }
    if ($s -notmatch '^(\d{1,2}):(\d{2})$') { return $null }
    $hour = [int]$Matches[1]
    $minute = [int]$Matches[2]
    if ($hour -lt 0 -or $hour -gt 23 -or $minute -lt 0 -or $minute -gt 59) { return $null }
    $dt = [DateTime]::new(1970, 1, 1, $hour, $minute, 0)
    return [long]([DateTimeOffset]::new($dt).ToUnixTimeMilliseconds())
}

function Get-NormalizedMonthlyOccurrence {
    param([string]$Occurrence)
    if ([string]::IsNullOrWhiteSpace($Occurrence)) { return $null }
    $s = $Occurrence.Trim().ToLowerInvariant()
    if ($s -eq 'last') { return 'Last' }
    $n = 0
    if ([int]::TryParse($s, [ref]$n) -and $n -ge 1 -and $n -le 4) { return [string]$n }
    return $null
}

function Get-CustomFieldsFromRow {
    param([PSCustomObject]$Row, [bool]$OverwriteEmpty, [bool]$MaintenanceWindowAsLocalTime)
    $customFields = @{}
    foreach ($prop in $Row.PSObject.Properties) {
        if ($prop.Name -eq 'level' -or $prop.Name -eq 'name') { continue }
        $val = $prop.Value
        if ([string]::IsNullOrEmpty($val)) {
            if ($OverwriteEmpty) { $customFields[$prop.Name] = $null }
        } else {
            $customFields[$prop.Name] = $val
        }
    }
    # Recurrence: Daily | Weekly | Monthly (default Weekly when missing)
    $maintenanceRecurrence = ($Row.PSObject.Properties | Where-Object { $_.Name -eq 'maintenanceRecurrence' } | Select-Object -ExpandProperty Value) -as [string]
    if ([string]::IsNullOrWhiteSpace($maintenanceRecurrence)) { $maintenanceRecurrence = 'Weekly' }
    else {
        $maintenanceRecurrence = $maintenanceRecurrence.Trim().ToLowerInvariant()
        if ($maintenanceRecurrence -eq 'daily') { $maintenanceRecurrence = 'Daily' }
        elseif ($maintenanceRecurrence -eq 'weekly') { $maintenanceRecurrence = 'Weekly' }
        elseif ($maintenanceRecurrence -eq 'monthly') { $maintenanceRecurrence = 'Monthly' }
        else { $maintenanceRecurrence = 'Weekly' }
    }
    $customFields['maintenanceRecurrence'] = $maintenanceRecurrence

    # Convert TIME fields from HH:MM to Unix ms (NinjaOne API expects Unix ms for TIME type), unless local time (TEXT field)
    if (-not $MaintenanceWindowAsLocalTime) {
        if ($customFields.ContainsKey('maintenanceStart') -and $null -ne $customFields['maintenanceStart'] -and [string]$customFields['maintenanceStart'] -ne '') {
            $ms = Convert-TimeStringToUnixMs -TimeString ($customFields['maintenanceStart'] -as [string])
            if ($null -ne $ms) { $customFields['maintenanceStart'] = $ms }
        }
        if ($customFields.ContainsKey('maintenanceEnd') -and $null -ne $customFields['maintenanceEnd'] -and [string]$customFields['maintenanceEnd'] -ne '') {
            $ms = Convert-TimeStringToUnixMs -TimeString ($customFields['maintenanceEnd'] -as [string])
            if ($null -ne $ms) { $customFields['maintenanceEnd'] = $ms }
        }
    }
    return $customFields
}

#endregion

#region Resolve target ID by level and name

function Get-OrganizationIdByName {
    param([string]$Name, [array]$Orgs)
    $org = $Orgs | Where-Object { $_.name -eq $Name } | Select-Object -First 1
    if ($org) { return $org.id }; return $null
}

function Get-LocationIdByOrgAndLocationName {
    param([string]$OrgName, [string]$LocationName, [array]$Orgs, [array]$Locs)
    $org = $Orgs | Where-Object { $_.name -eq $OrgName } | Select-Object -First 1
    if (-not $org) { return $null }
    $loc = $Locs | Where-Object { $_.organizationId -eq $org.id -and $_.name -eq $LocationName } | Select-Object -First 1
    if ($loc) { return $loc.id }; return $null
}

function Get-DeviceIdByName {
    param([string]$Name, [array]$Devices)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    if ($Name -match '^\d+$') {
        $dev = $Devices | Where-Object { $_.id -eq [int]$Name } | Select-Object -First 1
        if ($dev) { return $dev.id }; return $null
    }
    $dev = $Devices | Where-Object { $_.systemName -eq $Name } | Select-Object -First 1
    if ($dev) { return $dev.id }; return $null
}

#endregion

# Process each row
foreach ($row in $rows) {
    $level = ($row.level -as [string]).Trim().ToLowerInvariant()
    $name = ($row.name -as [string]).Trim()
    $customFields = Get-CustomFieldsFromRow -Row $row -OverwriteEmpty $OverwriteEmptyValues -MaintenanceWindowAsLocalTime $MaintenanceWindowAsLocalTime

    if ($customFields.Count -eq 0) {
        Write-Warning "Skipping row (level=$level, name=$name): no custom field columns."
        $script:SkippedCount++
        continue
    }

    # Validate Monthly: require valid maintenanceOccurrence (1, 2, 3, 4, or Last)
    $recurrence = ($customFields['maintenanceRecurrence'] -as [string]).Trim().ToLowerInvariant()
    if ($recurrence -eq 'monthly') {
        $occRaw = ($customFields['maintenanceOccurrence'] -as [string])
        $occNormalized = Get-NormalizedMonthlyOccurrence -Occurrence $occRaw
        if ($null -eq $occNormalized) {
            Write-Warning "Skipping row (level=$level, name=$name): maintenanceRecurrence is Monthly but maintenanceOccurrence is missing or invalid. Use 1, 2, 3, 4, or Last."
            $script:SkippedCount++
            continue
        }
        $customFields['maintenanceOccurrence'] = $occNormalized
    }

    if ($level -notin 'organization', 'location', 'device') {
        Write-Warning "Skipping row: invalid level '$level'. Use organization, location, or device."
        $script:SkippedCount++
        continue
    }

    $targetId = $null
    $endpoint = $null
    $locationOrgId = $null

    switch ($level) {
        'organization' {
            $targetId = Get-OrganizationIdByName -Name $name -Orgs $organizations
            $endpoint = "organization"
        }
        'location' {
            $parts = $name -split ',', 2
            if ($parts.Count -lt 2 -or [string]::IsNullOrWhiteSpace($parts[1])) {
                Write-Warning "Skipping row (location): name must be 'organizationname,locationname'. Got: $name"
                $script:SkippedCount++
                continue
            }
            $orgName = $parts[0].Trim()
            $locName = $parts[1].Trim()
            $targetId = Get-LocationIdByOrgAndLocationName -OrgName $orgName -LocationName $locName -Orgs $organizations -Locs $locations
            if ($null -ne $targetId) {
                $locationOrgId = ($locations | Where-Object { $_.id -eq $targetId } | Select-Object -First 1).organizationId
            }
            $endpoint = "location"
        }
        'device' {
            $targetId = Get-DeviceIdByName -Name $name -Devices $devices
            $endpoint = "device"
        }
    }

    if ($null -eq $targetId) {
        Write-Warning "Skipping row: could not resolve $level for name '$name'."
        $script:SkippedCount++
        continue
    }

    if ($level -eq 'location') {
        if ($null -eq $locationOrgId) {
            Write-Warning "Skipping row (location '$name'): could not determine organization ID for location id=$targetId."
            $script:SkippedCount++
            continue
        }
        $uri = "$baseUrl/api/v2/organization/$locationOrgId/location/$targetId/custom-fields"
    } else {
        $uri = "$baseUrl/api/v2/$endpoint/$targetId/custom-fields"
    }
    $json = $customFields | ConvertTo-Json -Depth 3

    try {
        Invoke-NinjaAPIRequest -Uri $uri -Method PATCH -Headers $headers -Body $json | Out-Null
        Write-Host "Updated $level '$name' (id=$targetId)."
        $script:UpdatedCount++
    } catch {
        Write-Warning "Failed to update $level '$name': $_"
        $script:FailedCount++
    }

    Start-Sleep -Milliseconds 300
}

# Summary
Write-Host ""
Write-Host "Summary: Updated=$($script:UpdatedCount), Skipped=$($script:SkippedCount), Failed=$($script:FailedCount)"

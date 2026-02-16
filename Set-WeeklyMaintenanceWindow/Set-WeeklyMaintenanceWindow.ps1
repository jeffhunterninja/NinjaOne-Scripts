<#
.SYNOPSIS
    Sets weekly maintenance windows (and other custom fields) in NinjaOne from a CSV.

.DESCRIPTION
    Reads a CSV and updates NinjaOne custom fields for organizations, locations, or devices.
    Primary use: set weekly maintenance window (day, start time, end time) per org, location, or device.
    CSV must have "level" (organization | location | device) and "name" to identify the target;
    all other columns are custom field name = value (e.g. maintenanceDay, maintenanceStart, maintenanceEnd).
    For location level, "name" must be "organizationname,locationname" (comma-separated).

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

.EXAMPLE
    .\Set-WeeklyMaintenanceWindow.ps1 -CsvPath "C:\data\maintenance-windows.csv" -NinjaOneInstance "app.ninjarmm.com" -NinjaOneClientId "..." -NinjaOneClientSecret "..."
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
    [bool]$OverwriteEmptyValues = $false
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

function Get-NextOccurrenceUnixMs {
    <#
    .SYNOPSIS
        Returns Unix time in milliseconds for the next occurrence of the given day of week at the given time (local time).
    #>
    param([string]$DayName, [string]$TimeString)
    if ([string]::IsNullOrWhiteSpace($DayName) -or [string]::IsNullOrWhiteSpace($TimeString)) { return $null }
    $dayOfWeek = [System.DayOfWeek]::Sunday
    switch ($DayName.Trim()) {
        'Sunday'    { $dayOfWeek = [System.DayOfWeek]::Sunday; break }
        'Monday'    { $dayOfWeek = [System.DayOfWeek]::Monday; break }
        'Tuesday'   { $dayOfWeek = [System.DayOfWeek]::Tuesday; break }
        'Wednesday' { $dayOfWeek = [System.DayOfWeek]::Wednesday; break }
        'Thursday'  { $dayOfWeek = [System.DayOfWeek]::Thursday; break }
        'Friday'    { $dayOfWeek = [System.DayOfWeek]::Friday; break }
        'Saturday'  { $dayOfWeek = [System.DayOfWeek]::Saturday; break }
        default     { return $null }
    }
    if ($TimeString -notmatch '^(\d{1,2}):(\d{2})$') { return $null }
    $hour = [int]$Matches[1]
    $minute = [int]$Matches[2]
    if ($hour -lt 0 -or $hour -gt 23 -or $minute -lt 0 -or $minute -gt 59) { return $null }
    $today = Get-Date
    $daysToAdd = ($dayOfWeek - $today.DayOfWeek + 7) % 7
    if ($daysToAdd -eq 0) { $daysToAdd = 7 }
    $target = $today.Date.AddDays($daysToAdd).AddHours($hour).AddMinutes($minute).AddSeconds(0)
    return [long]([DateTimeOffset]::new($target).ToUnixTimeMilliseconds())
}

function Get-CustomFieldsFromRow {
    param([PSCustomObject]$Row, [bool]$OverwriteEmpty)
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
    # Convert maintenanceStart and maintenanceEnd to Unix time (ms) using maintenanceDay from the same row
    $maintenanceDay = ($Row.PSObject.Properties | Where-Object { $_.Name -eq 'maintenanceDay' } | Select-Object -ExpandProperty Value) -as [string]
    if (-not [string]::IsNullOrWhiteSpace($maintenanceDay)) {
        if ($customFields.ContainsKey('maintenanceStart') -and $null -ne $customFields['maintenanceStart'] -and $customFields['maintenanceStart'] -notmatch '^\d+$') {
            $ms = Get-NextOccurrenceUnixMs -DayName $maintenanceDay -TimeString ($customFields['maintenanceStart'] -as [string])
            if ($null -ne $ms) { $customFields['maintenanceStart'] = $ms }
        }
        if ($customFields.ContainsKey('maintenanceEnd') -and $null -ne $customFields['maintenanceEnd'] -and $customFields['maintenanceEnd'] -notmatch '^\d+$') {
            $ms = Get-NextOccurrenceUnixMs -DayName $maintenanceDay -TimeString ($customFields['maintenanceEnd'] -as [string])
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
    $customFields = Get-CustomFieldsFromRow -Row $row -OverwriteEmpty $OverwriteEmptyValues

    if ($customFields.Count -eq 0) {
        Write-Warning "Skipping row (level=$level, name=$name): no custom field columns."
        $script:SkippedCount++
        continue
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

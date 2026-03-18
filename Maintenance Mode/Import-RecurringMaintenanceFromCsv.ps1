#Requires -Version 5.1
<#
.SYNOPSIS
    Imports recurring maintenance custom fields from a CSV and updates NinjaOne devices, organizations, or locations.

.DESCRIPTION
    Each CSV column header is the custom field API name (e.g. recurringMaintenanceEnableRecurringSchedule,
    recurringMaintenanceScheduleType). One column must identify the target entity: id or deviceId for devices;
    optionally use scope (Device | Organization | Location) with organizationId or locationId.
    You can supply names instead of IDs: systemName (device), organizationName (organization), or locationName (location).
    Names are resolved to IDs via API lookup (case-insensitive). Duplicate names cause an error for that row.
    For large tenants, ID-based CSV rows are faster than name-based device rows (device lookup paginates all devices).
    Multi-select fields: use comma- or semicolon-separated values in a single cell; they are sent as arrays.
    Time field (recurringMaintenanceTimeToStart24hFormat): use the same format NinjaOne stores (typically Unix milliseconds).

.PARAMETER CsvPath
    Path to the CSV file. Required.

.PARAMETER OverwriteEmptyValues
    If set, empty CSV cells are sent as $null and will clear existing values. If not set, empty cells are omitted from the PATCH.

.PARAMETER WhatIf
    If set, only reports what would be updated without calling the API.

.EXAMPLE
    .\Import-RecurringMaintenanceFromCsv.ps1 -CsvPath .\maintenance-import.csv
.EXAMPLE
    .\Import-RecurringMaintenanceFromCsv.ps1 -CsvPath .\maintenance-import.csv -OverwriteEmptyValues -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,
    [switch]$OverwriteEmptyValues,
    [switch]$WhatIf
)

# User-editable: same as Recurring Maintenance Mode.ps1 for consistency
$NinjaOneInstance = ''
$NinjaOneClientId = ''
$NinjaOneClientSecret = ''

# Reserved column names: used to identify target entity, not sent as custom fields
$ReservedColumns = @('id', 'deviceId', 'organizationId', 'locationId', 'scope', 'systemName', 'organizationName', 'locationName')

# --- Auth (same pattern as Recurring Maintenance Mode.ps1) ---
function Get-NinjaOneToken {
    if (-not ($Script:NinjaOneInstance -and $Script:NinjaOneClientID -and $Script:NinjaOneClientSecret)) {
        throw 'Please set $NinjaOneInstance, $NinjaOneClientId, and $NinjaOneClientSecret.'
    }
    if ($Script:NinjaTokenExpiry -and (Get-Date) -lt $Script:NinjaTokenExpiry) {
        return $Script:NinjaToken
    }
    $body = @{
        grant_type    = 'client_credentials'
        client_id     = $Script:NinjaOneClientID
        client_secret = $Script:NinjaOneClientSecret
        scope         = 'monitoring management'
    }
    $token = Invoke-RestMethod -Uri "https://$($Script:NinjaOneInstance -replace '/ws','')/ws/oauth/token" -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing
    $Script:NinjaTokenExpiry = (Get-Date).AddSeconds($token.expires_in)
    $Script:NinjaToken = $token
    return $token
}

function Connect-NinjaOne {
    param(
        [Parameter(Mandatory = $true)] $NinjaOneInstance,
        [Parameter(Mandatory = $true)] $NinjaOneClientID,
        [Parameter(Mandatory = $true)] $NinjaOneClientSecret
    )
    $Script:NinjaOneInstance = $NinjaOneInstance
    $Script:NinjaOneClientID = $NinjaOneClientID
    $Script:NinjaOneClientSecret = $NinjaOneClientSecret
    $null = Get-NinjaOneToken
}

function Invoke-NinjaOnePatch {
    param([string]$Path, [hashtable]$Body)
    $token = Get-NinjaOneToken
    $json = $Body | ConvertTo-Json -Depth 10
    Invoke-WebRequest -Uri "https://$($Script:NinjaOneInstance)/api/v2/$Path" -Method PATCH -Headers @{ Authorization = "Bearer $($token.access_token)" } -Body $json -ContentType 'application/json; charset=utf-8' -UseBasicParsing | Out-Null
}

function Invoke-NinjaOneGet {
    param([string]$Path, [string]$QueryParams = '')
    $token = Get-NinjaOneToken
    $uri = "https://$($Script:NinjaOneInstance)/api/v2/$Path"
    if (-not [string]::IsNullOrWhiteSpace($QueryParams)) { $uri += "?$QueryParams" }
    $response = Invoke-WebRequest -Uri $uri -Method GET -Headers @{ Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json' -UseBasicParsing
    return $response.Content | ConvertFrom-Json
}

function Invoke-NinjaOneGetPaginated {
    param([string]$Path, [string]$QueryParams = '', [int]$PageSize = 1000)
    $after = 0
    $all = [System.Collections.ArrayList]::new()
    do {
        $q = "pageSize=$PageSize&after=$after"
        if (-not [string]::IsNullOrWhiteSpace($QueryParams)) { $q += "&$QueryParams" }
        $response = Invoke-NinjaOneGet -Path $Path -QueryParams $q
        if ($response -is [Array]) {
            $page = $response
        } elseif ($null -ne $response -and $response.PSObject.Properties['results']) {
            $page = @($response.results)
        } elseif ($null -ne $response -and $response.PSObject.Properties['data']) {
            $page = @($response.data)
        } else {
            $page = @($response)
        }
        foreach ($item in $page) { $null = $all.Add($item) }
        if ($page.Count -lt $PageSize) { break }
        if ($page.Count -eq 0) { break }
        $idVal = $page[0].PSObject.Properties['id']
        if (-not $idVal) { break }
        $after = ($page | ForEach-Object { $_.id } | Measure-Object -Maximum).Maximum
    } while ($true)
    return $all
}

# --- Name-to-ID lookups (cached) ---
$Script:DeviceByNameCache = $null
$Script:OrganizationByNameCache = $null
$Script:LocationByNameCache = $null

function Get-DeviceIdBySystemName {
    param([string]$SystemName)
    $key = [string]$SystemName
    if ([string]::IsNullOrWhiteSpace($key)) { return $null }
    if ($null -eq $Script:DeviceByNameCache) {
        $Script:DeviceByNameCache = @{}
        $devices = Invoke-NinjaOneGetPaginated -Path 'devices'
        foreach ($d in $devices) {
            $name = [string]$d.systemName
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $nameKey = $name.Trim().ToLowerInvariant()
            if (-not $Script:DeviceByNameCache.ContainsKey($nameKey)) {
                $Script:DeviceByNameCache[$nameKey] = [System.Collections.ArrayList]::new()
            }
            $null = $Script:DeviceByNameCache[$nameKey].Add([string]$d.id)
        }
    }
    $nameKey = $key.Trim().ToLowerInvariant()
    if (-not $Script:DeviceByNameCache.ContainsKey($nameKey)) { return $null }
    $ids = $Script:DeviceByNameCache[$nameKey]
    if ($ids.Count -eq 0) { return $null }
    if ($ids.Count -gt 1) { return 'AMBIGUOUS' }
    return $ids[0]
}

function Get-OrganizationIdByName {
    param([string]$OrganizationName)
    $key = [string]$OrganizationName
    if ([string]::IsNullOrWhiteSpace($key)) { return $null }
    if ($null -eq $Script:OrganizationByNameCache) {
        $Script:OrganizationByNameCache = @{}
        try {
            $orgs = Invoke-NinjaOneGetPaginated -Path 'organizations'
        } catch {
            $orgs = @()
        }
        foreach ($o in $orgs) {
            $name = [string]$o.name
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $nameKey = $name.Trim().ToLowerInvariant()
            if (-not $Script:OrganizationByNameCache.ContainsKey($nameKey)) {
                $Script:OrganizationByNameCache[$nameKey] = [System.Collections.ArrayList]::new()
            }
            $null = $Script:OrganizationByNameCache[$nameKey].Add([string]$o.id)
        }
    }
    $nameKey = $key.Trim().ToLowerInvariant()
    if (-not $Script:OrganizationByNameCache.ContainsKey($nameKey)) { return $null }
    $ids = $Script:OrganizationByNameCache[$nameKey]
    if ($ids.Count -eq 0) { return $null }
    if ($ids.Count -gt 1) { return 'AMBIGUOUS' }
    return $ids[0]
}

function Get-LocationIdByName {
    param([string]$LocationName)
    $key = [string]$LocationName
    if ([string]::IsNullOrWhiteSpace($key)) { return $null }
    if ($null -eq $Script:LocationByNameCache) {
        $Script:LocationByNameCache = @{}
        try {
            $locs = Invoke-NinjaOneGetPaginated -Path 'locations'
        } catch {
            $locs = @()
        }
        foreach ($l in $locs) {
            $name = [string]$l.name
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $nameKey = $name.Trim().ToLowerInvariant()
            if (-not $Script:LocationByNameCache.ContainsKey($nameKey)) {
                $Script:LocationByNameCache[$nameKey] = [System.Collections.ArrayList]::new()
            }
            $null = $Script:LocationByNameCache[$nameKey].Add([string]$l.id)
        }
    }
    $nameKey = $key.Trim().ToLowerInvariant()
    if (-not $Script:LocationByNameCache.ContainsKey($nameKey)) { return $null }
    $ids = $Script:LocationByNameCache[$nameKey]
    if ($ids.Count -eq 0) { return $null }
    if ($ids.Count -gt 1) { return 'AMBIGUOUS' }
    return $ids[0]
}

# --- Load CSV ---
if (-not (Test-Path -LiteralPath $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}
try {
    $rows = Import-Csv -Path $CsvPath -Encoding UTF8
} catch {
    Write-Error "Failed to import CSV: $_"
    exit 1
}
if (-not $rows -or $rows.Count -eq 0) {
    Write-Warning "CSV has no data rows."
    exit 0
}

# Determine target column for device-level: id or deviceId (optional if systemName is used)
$headers = $rows[0].PSObject.Properties.Name
$idColumn = if ($headers -contains 'id') { 'id' } elseif ($headers -contains 'deviceId') { 'deviceId' } else { $null }
$hasDeviceIdentifier = $idColumn -or ($headers -contains 'systemName')
$hasOrgIdentifier = ($headers -contains 'organizationId') -or ($headers -contains 'organizationName')
$hasLocIdentifier = ($headers -contains 'locationId') -or ($headers -contains 'locationName')
if (-not $hasDeviceIdentifier -and -not ($headers -contains 'scope')) {
    Write-Error "CSV must contain 'id' or 'deviceId' or 'systemName' for device updates, and/or 'scope' with 'organizationId'/'organizationName' or 'locationId'/'locationName' for org/location updates."
    exit 1
}
if (($headers -contains 'scope') -and -not $hasDeviceIdentifier -and -not $hasOrgIdentifier -and -not $hasLocIdentifier) {
    Write-Error "CSV with 'scope' must contain at least one of: deviceId/id/systemName (Device), organizationId/organizationName (Organization), locationId/locationName (Location)."
    exit 1
}

# Connect
try {
    Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientId -NinjaOneClientSecret $NinjaOneClientSecret
} catch {
    Write-Error "Failed to connect to NinjaOne: $_"
    exit 1
}

$updated = 0
$errors = [System.Collections.ArrayList]::new()

foreach ($row in $rows) {
    $scope = if ($headers -contains 'scope') { ([string]$row.scope).Trim().ToLowerInvariant() } else { 'device' }
    if ([string]::IsNullOrWhiteSpace($scope)) { $scope = 'device' }

    $entityId = $null
    $path = $null
    switch ($scope) {
        'device' {
            $entityId = if ($idColumn) { [string]$row.$idColumn } else { '' }
            $entityId = ([string]$entityId).Trim()
            if ([string]::IsNullOrWhiteSpace($entityId) -and ($headers -contains 'systemName')) {
                $systemName = ([string]$row.systemName).Trim()
                if (-not [string]::IsNullOrWhiteSpace($systemName)) {
                    $resolved = Get-DeviceIdBySystemName -SystemName $systemName
                    if ($resolved -eq 'AMBIGUOUS') {
                        $null = $errors.Add("Multiple devices found for systemName '$systemName'. Skipping row.")
                        continue
                    }
                    if ([string]::IsNullOrWhiteSpace($resolved)) {
                        $null = $errors.Add("No device found for systemName '$systemName'. Skipping row.")
                        continue
                    }
                    $entityId = $resolved
                }
            }
            if ([string]::IsNullOrWhiteSpace($entityId)) {
                $null = $errors.Add("Row has scope Device but missing deviceId/id or systemName. Supply deviceId, id, or systemName. Skipping.")
                continue
            }
            $path = "device/$entityId/custom-fields"
        }
        'organization' {
            $entityId = if ($headers -contains 'organizationId') { [string]$row.organizationId } else { '' }
            $entityId = ([string]$entityId).Trim()
            if ([string]::IsNullOrWhiteSpace($entityId) -and ($headers -contains 'organizationName')) {
                $orgName = ([string]$row.organizationName).Trim()
                if (-not [string]::IsNullOrWhiteSpace($orgName)) {
                    $resolved = Get-OrganizationIdByName -OrganizationName $orgName
                    if ($resolved -eq 'AMBIGUOUS') {
                        $null = $errors.Add("Multiple organizations found for organizationName '$orgName'. Skipping row.")
                        continue
                    }
                    if ([string]::IsNullOrWhiteSpace($resolved)) {
                        $null = $errors.Add("No organization found for organizationName '$orgName'. Skipping row.")
                        continue
                    }
                    $entityId = $resolved
                }
            }
            if ([string]::IsNullOrWhiteSpace($entityId)) {
                $null = $errors.Add("Row with scope Organization missing organizationId or organizationName. Skipping.")
                continue
            }
            $path = "organization/$entityId/custom-fields"
        }
        'location' {
            $entityId = if ($headers -contains 'locationId') { [string]$row.locationId } else { '' }
            $entityId = ([string]$entityId).Trim()
            if ([string]::IsNullOrWhiteSpace($entityId) -and ($headers -contains 'locationName')) {
                $locName = ([string]$row.locationName).Trim()
                if (-not [string]::IsNullOrWhiteSpace($locName)) {
                    $resolved = Get-LocationIdByName -LocationName $locName
                    if ($resolved -eq 'AMBIGUOUS') {
                        $null = $errors.Add("Multiple locations found for locationName '$locName'. Skipping row.")
                        continue
                    }
                    if ([string]::IsNullOrWhiteSpace($resolved)) {
                        $null = $errors.Add("No location found for locationName '$locName'. Skipping row.")
                        continue
                    }
                    $entityId = $resolved
                }
            }
            if ([string]::IsNullOrWhiteSpace($entityId)) {
                $null = $errors.Add("Row with scope Location missing locationId or locationName. Skipping.")
                continue
            }
            $path = "location/$entityId/custom-fields"
        }
        default {
            $null = $errors.Add("Unknown scope '$scope'; use Device, Organization, or Location. Skipping row.")
            continue
        }
    }

    $body = @{}
    foreach ($prop in $row.PSObject.Properties) {
        $name = $prop.Name
        if ($ReservedColumns -contains $name) { continue }
            $val = $prop.Value
        $isEmpty = [string]::IsNullOrWhiteSpace([string]$val)
        if ($isEmpty) {
            if ($OverwriteEmptyValues) { $body[$name] = $null }
            continue
        }
        $strVal = [string]$val
        # Multi-select: comma or semicolon separated -> array
        if ($strVal -match '[,;]') {
            $body[$name] = @($strVal -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        } else {
            $body[$name] = $strVal.Trim()
        }
    }

    if ($body.Count -eq 0) {
        $null = $errors.Add("Row for $scope $entityId has no custom field values to update. Skipping.")
        continue
    }

    if ($WhatIf) {
        Write-Host "WhatIf: Would PATCH $path with: $($body | ConvertTo-Json -Compress)"
        $updated++
        continue
    }

    try {
        Invoke-NinjaOnePatch -Path $path -Body $body
        Write-Host "Updated $scope $entityId"
        $updated++
    } catch {
        $null = $errors.Add("Failed to update $scope $entityId : $_")
    }
}

if ($errors.Count -gt 0) {
    Write-Warning "Errors:"
    $errors | ForEach-Object { Write-Warning $_ }
}
Write-Output "Done. Updated $updated of $($rows.Count) row(s)."

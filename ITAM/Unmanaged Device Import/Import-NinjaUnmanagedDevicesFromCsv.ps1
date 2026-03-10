<#
.SYNOPSIS
    Creates NinjaOne ITAM unmanaged devices from a CSV file.

.DESCRIPTION
    Reads a CSV of equipment (e.g. mouse, keyboard, headset, dock, monitors), resolves
    organizations, locations, and unmanaged device roles by name or ID, and creates each
    device via the NinjaOne API. Use after receiving and reconciling equipment to bulk-import
    into NinjaOne ITAM. All logic is standalone (no dot-sourcing).

.PARAMETER CsvPath
    Path to the CSV file. Required columns: Name, RoleName, and either (OrganizationName, LocationName)
    or (OrganizationId, LocationId). Optional: SerialNumber, WarrantyStartDate, WarrantyEndDate,
    Make, Model, PurchaseDate, PurchaseAmount for custom fields.

.PARAMETER BaseUrl
    NinjaOne instance base URL (e.g. ca.ninjarmm.com or https://ca.ninjarmm.com). Default: app.ninjarmm.com.

.PARAMETER ClientId
    NinjaOne API application Client ID. Can use $env:NinjaOneClientId if not provided.

.PARAMETER ClientSecret
    NinjaOne API application Client Secret. Can use $env:NinjaOneClientSecret if not provided.

.PARAMETER WhatIf
    Validate CSV and lookups only; do not create any devices.

.PARAMETER SkipErrors
    Continue on per-row failure and report all errors at the end.

.EXAMPLE
    .\Import-NinjaUnmanagedDevicesFromCsv.ps1 -CsvPath ".\Import-UnmanagedDevices-Example.csv" -BaseUrl ca.ninjarmm.com
.EXAMPLE
    .\Import-NinjaUnmanagedDevicesFromCsv.ps1 -CsvPath ".\equipment.csv" -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter()]
    [string]$BaseUrl = 'app.ninjarmm.com',

    [Parameter()]
    [string]$ClientId,

    [Parameter()]
    [string]$ClientSecret,

    [Parameter()]
    [switch]$WhatIf,

    [Parameter()]
    [switch]$SkipErrors
)

$ErrorActionPreference = 'Stop'

#region Inline OAuth and API (standalone, no dot-sourcing)

function Get-NinjaOAuthTokenInline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$ClientId,
        [Parameter(Mandatory)] [string]$ClientSecret,
        [Parameter(Mandatory)] [string]$BaseUrl,
        [int]$TimeoutSec = 30
    )
    $uri = "$($BaseUrl.TrimEnd('/'))/ws/oauth/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = 'client_credentials'
        scope         = 'monitoring management'
    }
    $headers = @{
        'Accept'       = 'application/json'
        'Content-Type' = 'application/x-www-form-urlencoded'
    }
    $resp = Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $body -TimeoutSec $TimeoutSec -ErrorAction Stop
    return $resp
}

function Invoke-NinjaOneApiInline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')] [string]$Method,
        [Parameter(Mandatory)] [string]$Endpoint,
        [string]$Query,
        $Body,
        [int]$TimeoutSec = 60,
        [Parameter(Mandatory)] [PSCustomObject]$Session
    )
    if (-not $Session.ExpiresAt -or (Get-Date) -ge $Session.ExpiresAt) {
        throw 'No active session or token expired. Re-authenticate.'
    }
    $uri = if ($Query -and $Query.Length -gt 0) {
        "$($Session.BaseUrl)/$($Endpoint.TrimStart('/'))?df=$Query"
    } else {
        "$($Session.BaseUrl)/$($Endpoint.TrimStart('/'))"
    }
    $headers = @{ Authorization = $Session.AuthHeader; 'Accept' = 'application/json' }
    $bodyParam = $null
    if ($Body) {
        $bodyParam = $Body | ConvertTo-Json -Depth 100
        $headers['Content-Type'] = 'application/json'
    }
    return Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers -Body $bodyParam -TimeoutSec $TimeoutSec -ErrorAction Stop
}

function Get-DateToUnixSeconds {
    param([datetime]$Date)
    return [int]($Date.ToUniversalTime() - [datetime]'1970-01-01').TotalSeconds
}

#endregion

# Resolve credentials
if (-not $ClientId) { $ClientId = $env:NinjaOneClientId }
if (-not $ClientSecret) { $ClientSecret = $env:NinjaOneClientSecret }
if ([string]::IsNullOrWhiteSpace($ClientId) -or [string]::IsNullOrWhiteSpace($ClientSecret)) {
    if ($PSVersionTable.PSVersion -lt [version]'7.0') {
        $ClientId     = Read-Host -Prompt 'ClientId'
        $ClientSecret = Read-Host -Prompt 'ClientSecret'
    } else {
        $ClientId     = Read-Host -MaskInput 'ClientId'
        $ClientSecret = Read-Host -MaskInput 'ClientSecret'
    }
}

# Normalize BaseUrl to https, no trailing slash
$base = $BaseUrl.Trim()
if (-not $base.StartsWith('http://', [StringComparison]::OrdinalIgnoreCase) -and -not $base.StartsWith('https://', [StringComparison]::OrdinalIgnoreCase)) {
    $base = "https://$base"
}
$base = $base.TrimEnd('/')

# OAuth and session
$tokenResp = Get-NinjaOAuthTokenInline -ClientId $ClientId -ClientSecret $ClientSecret -BaseUrl $base
$Session = [PSCustomObject]@{
    BaseUrl    = $base
    AuthHeader = "Bearer $($tokenResp.access_token)"
    ExpiresAt  = if ($tokenResp.expires_in) { (Get-Date).AddSeconds([int]$tokenResp.expires_in) } else { (Get-Date).AddHours(1) }
}

# Validate CSV exists and load
if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}
$rows = Import-Csv -LiteralPath $CsvPath -Encoding UTF8
if (-not $rows -or $rows.Count -eq 0) {
    Write-Warning 'CSV is empty or has no data rows.'
    exit 0
}

# Required column names (case-insensitive match when reading)
$requiredByName = @('Name', 'RoleName', 'OrganizationName', 'LocationName')
$requiredById   = @('Name', 'RoleName', 'OrganizationId', 'LocationId')
$allColumns     = $rows[0].PSObject.Properties.Name

function Test-RequiredColumns {
    param([string[]]$Required, [string[]]$Columns)
    foreach ($r in $Required) {
        $match = $Columns | Where-Object { $_ -ieq $r }
        if (-not $match) { return $false }
    }
    return $true
}

function Get-RowValue {
    param([PSCustomObject]$Row, [string]$ColumnName)
    $prop = $Row.PSObject.Properties | Where-Object { $_.Name -ieq $ColumnName } | Select-Object -First 1
    if (-not $prop) { return '' }
    $v = $prop.Value -as [string]
    return if ($null -eq $v) { '' } else { $v.Trim() }
}

$useById = $false
if ((Test-RequiredColumns -Required $requiredById -Columns $allColumns)) {
    $useById = $true
} elseif (-not (Test-RequiredColumns -Required $requiredByName -Columns $allColumns)) {
    Write-Error "CSV must have either (Name, RoleName, OrganizationName, LocationName) or (Name, RoleName, OrganizationId, LocationId). Found columns: $($allColumns -join ', ')"
    exit 1
}

# Cache: organizations, locations, unmanaged device roles
$organizations = Invoke-NinjaOneApiInline -Method GET -Endpoint 'v2/organizations' -Session $Session
$locations     = Invoke-NinjaOneApiInline -Method GET -Endpoint 'v2/locations' -Session $Session
$rolesRaw      = Invoke-NinjaOneApiInline -Method GET -Endpoint 'v2/noderole/list' -Session $Session
$roles         = $rolesRaw | Where-Object { $_.nodeClass -eq 'UNMANAGED_DEVICE' }

$created = 0
$failed  = 0
$errors  = [System.Collections.Generic.List[string]]::new()
$rowNum  = 0

foreach ($row in $rows) {
    $rowNum++
    $name = Get-RowValue -Row $row -ColumnName 'Name'
    $roleName = Get-RowValue -Row $row -ColumnName 'RoleName'
    if ([string]::IsNullOrWhiteSpace($roleName)) {
        $errors.Add("Row $rowNum`: RoleName is required.")
        if (-not $SkipErrors) { throw "Row $rowNum`: RoleName is required." }
        $failed++; continue
    }

    $orgId  = $null
    $locId  = $null
    if ($useById) {
        $orgId = Get-RowValue -Row $row -ColumnName 'OrganizationId'
        $locId = Get-RowValue -Row $row -ColumnName 'LocationId'
        if ([string]::IsNullOrWhiteSpace($orgId) -or [string]::IsNullOrWhiteSpace($locId)) {
            $errors.Add("Row $rowNum`: OrganizationId and LocationId are required when using ID-based columns.")
            if (-not $SkipErrors) { throw "Row $rowNum`: OrganizationId and LocationId are required." }
            $failed++; continue
        }
    } else {
        $orgName  = Get-RowValue -Row $row -ColumnName 'OrganizationName'
        $locName  = Get-RowValue -Row $row -ColumnName 'LocationName'
        $orgMatch = $organizations | Where-Object { $_.name -eq $orgName } | Select-Object -First 1
        if (-not $orgMatch) {
            $errors.Add("Row $rowNum`: Organization not found: '$orgName'.")
            if (-not $SkipErrors) { throw "Row $rowNum`: Organization not found: '$orgName'." }
            $failed++; continue
        }
        $orgId = $orgMatch.id
        $locMatch = $locations | Where-Object {
            $locName -and ($_.name -eq $locName) -and (($_.organizationID -eq $orgId) -or ($_.organizationId -eq $orgId))
        } | Select-Object -First 1
        if (-not $locMatch) {
            $errors.Add("Row $rowNum`: Location not found: '$locName' in organization '$orgName'.")
            if (-not $SkipErrors) { throw "Row $rowNum`: Location not found: '$locName'." }
            $failed++; continue
        }
        $locId = $locMatch.id
    }

    $roleMatch = $roles | Where-Object { $_.name -eq $roleName } | Select-Object -First 1
    if (-not $roleMatch) {
        $errors.Add("Row $rowNum`: Unmanaged device role not found: '$roleName'. Ensure the role exists in NinjaOne (e.g. Mouse, Keyboard, Displays, Headset, Dock).")
        if (-not $SkipErrors) { throw "Row $rowNum`: Role not found: '$roleName'." }
        $failed++; continue
    }
    $roleId = $roleMatch.id

    $displayName = $name
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $make  = Get-RowValue -Row $row -ColumnName 'Make'
        $model = Get-RowValue -Row $row -ColumnName 'Model'
        $displayName = if ($make -and $model) { "$make $model" } else { "Unmanaged $roleName $rowNum" }
    }

    $serial = Get-RowValue -Row $row -ColumnName 'SerialNumber'
    if ([string]::IsNullOrWhiteSpace($serial)) { $serial = $null }

    $warrantyStart = Get-Date
    $warrantyEnd   = (Get-Date).AddYears(3)
    $ws = Get-RowValue -Row $row -ColumnName 'WarrantyStartDate'
    $we = Get-RowValue -Row $row -ColumnName 'WarrantyEndDate'
    if ($ws -and ([datetime]::TryParse($ws, [ref]$null))) { $warrantyStart = [datetime]::Parse($ws) }
    if ($we -and ([datetime]::TryParse($we, [ref]$null))) { $warrantyEnd   = [datetime]::Parse($we) }

    $warrantyStartUnix = Get-DateToUnixSeconds -Date $warrantyStart
    $warrantyEndUnix   = Get-DateToUnixSeconds -Date $warrantyEnd

    $body = @{
        name              = $displayName
        roleId             = $roleId
        orgId              = $orgId
        locationId         = $locId
        warrantyStartDate  = $warrantyStartUnix
        warrantyEndDate    = $warrantyEndUnix
        serialNumber       = $serial
    }

    if ($WhatIf) {
        Write-Verbose "WhatIf: Would create unmanaged device '$displayName' (Role: $roleName, OrgId: $orgId, LocId: $locId)."
        $created++
        continue
    }

    try {
        $result = Invoke-NinjaOneApiInline -Method POST -Endpoint 'v2/itam/unmanaged-device' -Body $body -Session $Session
        $nodeId = $result.nodeId

        # Optional: PATCH custom fields if CSV has Make, Model, PurchaseDate, PurchaseAmount
        $make   = Get-RowValue -Row $row -ColumnName 'Make'
        $model  = Get-RowValue -Row $row -ColumnName 'Model'
        $purch  = Get-RowValue -Row $row -ColumnName 'PurchaseDate'
        $amount = Get-RowValue -Row $row -ColumnName 'PurchaseAmount'
        if ($make -or $model -or $purch -or $amount -or $serial) {
            $customFields = @{}
            if ($make)   { $customFields['manufacturer'] = $make }
            if ($model)  { $customFields['model'] = $model }
            if ($serial) { $customFields['itamAssetSerialNumber'] = $serial }
            if ($purch -and [datetime]::TryParse($purch, [ref]$null)) {
                $purchDate = [datetime]::Parse($purch)
                $customFields['itamAssetPurchaseDate'] = [int]($purchDate.ToUniversalTime() - [datetime]'1970-01-01').TotalSeconds * 1000
            }
            if ($amount -and ($amount -match '^\d+(\.\d+)?$')) {
                $customFields['itamAssetPurchaseAmount'] = [int][double]$amount
            }
            if ($customFields.Count -gt 0) {
                try {
                    Invoke-NinjaOneApiInline -Method PATCH -Endpoint "v2/device/$nodeId/custom-fields" -Body $customFields -Session $Session | Out-Null
                } catch {
                    Write-Verbose "Row $rowNum`: Custom fields PATCH failed (device created): $($_.Exception.Message)"
                }
            }
        }

        $created++
        Write-Verbose "Created: $displayName (nodeId: $nodeId)."
    } catch {
        $errors.Add("Row $rowNum` ($displayName): $($_.Exception.Message)")
        if (-not $SkipErrors) { throw }
        $failed++
    }
}

# Summary
Write-Host "Import complete. Created: $created, Failed: $failed."
if ($errors.Count -gt 0) {
    foreach ($e in $errors) {
        Write-Warning $e
    }
}

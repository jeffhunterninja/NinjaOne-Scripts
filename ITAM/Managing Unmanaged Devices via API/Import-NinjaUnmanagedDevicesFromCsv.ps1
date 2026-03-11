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
    Make, Model, PurchaseDate, PurchaseAmount, AssetStatus, ExpectedLifetime, EndOfLifeDate for custom fields.
    ExpectedLifetime is typically e.g. "1 years", "2 years" (stored lowercased). EndOfLifeDate should be a date (e.g. YYYY-MM-DD).

.PARAMETER BaseUrl
    NinjaOne instance base URL (e.g. ca.ninjarmm.com or https://ca.ninjarmm.com). Default: app.ninjarmm.com.

.PARAMETER ClientId
    NinjaOne API application Client ID. Can use $env:NinjaOneClientId if not provided.

.PARAMETER ClientSecret
    NinjaOne API application Client Secret. Can use $env:NinjaOneClientSecret if not provided.

.PARAMETER SkipErrors
    Continue on per-row failure and report all errors at the end.

.EXAMPLE
    .\Import-NinjaUnmanagedDevicesFromCsv.ps1 -CsvPath ".\Import-UnmanagedDevices-Example.csv" -BaseUrl ca.ninjarmm.com
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$CsvPath = '.\Import-UnmanagedDevices-Example.csv',

    [Parameter()]
    [string]$BaseUrl = 'ca.ninjarmm.com',

    [Parameter()]
    [string]$ClientId = '',

    [Parameter()]
    [string]$ClientSecret = '',

    [Parameter()]
    [switch]$SkipErrors = $true
)
# Resolve CSV path when default was used (run from script directory)
if ($CsvPath -eq '.\Import-UnmanagedDevices-Example.csv' -and $PSScriptRoot) {
    $CsvPath = Join-Path $PSScriptRoot 'Import-UnmanagedDevices-Example.csv'
}

$ErrorActionPreference = 'Stop'

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
        [int]$MaxRetries = 4,
        [Parameter(Mandatory)] [PSCustomObject]$Session
    )
    function Get-HttpStatusCodeInline {
        param($ErrorRecord)
        try {
            if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
                return [int]$ErrorRecord.Exception.Response.StatusCode
            }
        } catch { }
        return $null
    }

    function Get-RetryAfterSecondsInline {
        param($ErrorRecord)
        try {
            $resp = $ErrorRecord.Exception.Response
            if ($resp -and $resp.Headers['Retry-After']) {
                $raw = $resp.Headers['Retry-After'] | Select-Object -First 1
                $sec = 0
                if ([int]::TryParse($raw, [ref]$sec) -and $sec -gt 0) { return $sec }
            }
        } catch { }
        return $null
    }

    function Refresh-NinjaSessionInline {
        param([PSCustomObject]$TargetSession)
        $token = Get-NinjaOAuthTokenInline -ClientId $TargetSession.ClientId -ClientSecret $TargetSession.ClientSecret -BaseUrl $TargetSession.BaseUrl
        $TargetSession.AuthHeader = "Bearer $($token.access_token)"
        $TargetSession.ExpiresAt = if ($token.expires_in) { (Get-Date).AddSeconds([int]$token.expires_in - 60) } else { (Get-Date).AddMinutes(55) }
    }

    if (-not $Session.ExpiresAt -or (Get-Date) -ge $Session.ExpiresAt) {
        Refresh-NinjaSessionInline -TargetSession $Session
    }

    $uri = if ($Query -and $Query.Length -gt 0) {
        "$($Session.BaseUrl)/$($Endpoint.TrimStart('/'))?$Query"
    } else {
        "$($Session.BaseUrl)/$($Endpoint.TrimStart('/'))"
    }

    $attempt = 0
    while ($true) {
        $headers = @{ Authorization = $Session.AuthHeader; 'Accept' = 'application/json' }
        $bodyParam = $null
        if ($Body) {
            $bodyParam = $Body | ConvertTo-Json -Depth 100
            $headers['Content-Type'] = 'application/json'
        }

        try {
            return Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers -Body $bodyParam -TimeoutSec $TimeoutSec -ErrorAction Stop
        } catch {
            $status = Get-HttpStatusCodeInline -ErrorRecord $_
            $attempt++

            if ($status -eq 401 -and $attempt -le $MaxRetries) {
                Refresh-NinjaSessionInline -TargetSession $Session
                continue
            }

            $isRetryable = ($status -in @(408, 429, 500, 502, 503, 504))
            if (-not $isRetryable -or $attempt -gt $MaxRetries) { throw }

            $retryAfter = Get-RetryAfterSecondsInline -ErrorRecord $_
            $sleepSec = if ($retryAfter -and $retryAfter -gt 0) {
                [Math]::Min($retryAfter, 60)
            } else {
                [Math]::Min([Math]::Pow(2, $attempt), 30)
            }
            Write-Warning "Retrying $Method $Endpoint after HTTP $status in $sleepSec second(s) (attempt $attempt/$MaxRetries)."
            Start-Sleep -Seconds $sleepSec
        }
    }
}

function ConvertTo-ApiItemArrayInline {
    param($Response)
    if ($null -eq $Response) { return @() }
    if ($Response -is [Array]) { return @($Response) }
    if ($Response.PSObject.Properties['data']) { return @($Response.data) }
    if ($Response.PSObject.Properties['items']) { return @($Response.items) }
    if ($Response.PSObject.Properties['results']) { return @($Response.results) }
    return @($Response)
}

function Get-NinjaOnePagedInline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$Endpoint,
        [Parameter(Mandatory)] [PSCustomObject]$Session,
        [int]$PageSize = 200,
        [int]$MaxPages = 500
    )
    $out = [System.Collections.Generic.List[object]]::new()
    $offset = 0
    for ($page = 0; $page -lt $MaxPages; $page++) {
        $resp = Invoke-NinjaOneApiInline -Method GET -Endpoint $Endpoint -Query "limit=$PageSize&offset=$offset" -Session $Session
        $items = ConvertTo-ApiItemArrayInline -Response $resp
        if ($items.Count -eq 0) { break }

        foreach ($i in $items) { [void]$out.Add($i) }
        if ($items.Count -lt $PageSize) { break }
        $offset += $items.Count
    }
    return @($out)
}

function Get-DateToUnixSeconds {
    param([datetime]$Date)
    return [int]($Date.ToUniversalTime() - [datetime]'1970-01-01').TotalSeconds
}

function ConvertTo-UnixMilliseconds {
    param([datetime]$Date)
    return [int64](($Date.ToUniversalTime() - [datetime]'1970-01-01').TotalSeconds * 1000)
}

function ConvertTo-OptionalDateParseResult {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return [PSCustomObject]@{ Success = $true; Date = $null; Message = $null }
    }
    try {
        return [PSCustomObject]@{ Success = $true; Date = [datetime]$Value; Message = $null }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Date    = $null
            Message = "Invalid date '$Value'. Expected a valid date such as YYYY-MM-DD."
        }
    }
}

function ConvertTo-OptionalIntAmountParseResult {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return [PSCustomObject]@{ Success = $true; Amount = $null; Message = $null }
    }
    if ($Value -match '^\d+(\.\d+)?$') {
        return [PSCustomObject]@{ Success = $true; Amount = [int][double]$Value; Message = $null }
    }
    return [PSCustomObject]@{
        Success = $false
        Amount  = $null
        Message = "Invalid amount '$Value'. Use numeric values only."
    }
}

# Resolve credentials
$ClientId = if (-not $ClientId) { $env:NinjaOneClientId } else { $ClientId }
$ClientSecret = if (-not $ClientSecret) { $env:NinjaOneClientSecret } else { $ClientSecret }
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
    ExpiresAt  = if ($tokenResp.expires_in) { (Get-Date).AddSeconds([int]$tokenResp.expires_in - 60) } else { (Get-Date).AddMinutes(55) }
    ClientId   = $ClientId
    ClientSecret = $ClientSecret
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
    if ($null -eq $v) {
        return ''
    } else {
        return $v.Trim()
    }
}

$useById = $false
if ((Test-RequiredColumns -Required $requiredById -Columns $allColumns)) {
    $useById = $true
} elseif (-not (Test-RequiredColumns -Required $requiredByName -Columns $allColumns)) {
    Write-Error "CSV must have either (Name, RoleName, OrganizationName, LocationName) or (Name, RoleName, OrganizationId, LocationId). Found columns: $($allColumns -join ', ')"
    exit 1
}

# Cache: organizations, locations, unmanaged device roles
$organizations = Get-NinjaOnePagedInline -Endpoint 'v2/organizations' -Session $Session
$locations     = Get-NinjaOnePagedInline -Endpoint 'v2/locations' -Session $Session
$rolesRaw      = Get-NinjaOnePagedInline -Endpoint 'v2/noderole/list' -Session $Session
$roles         = $rolesRaw | Where-Object { $_.nodeClass -eq 'UNMANAGED_DEVICE' }

$created  = 0
$failed   = 0
$errors   = [System.Collections.Generic.List[string]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()
$rowNum   = 0

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
    [datetime]$parsedWarrantyStart = [datetime]::MinValue
    [datetime]$parsedWarrantyEnd   = [datetime]::MinValue
    if ($ws -and [datetime]::TryParse($ws, [ref]$parsedWarrantyStart)) { $warrantyStart = $parsedWarrantyStart }
    if ($we -and [datetime]::TryParse($we, [ref]$parsedWarrantyEnd)) { $warrantyEnd = $parsedWarrantyEnd }

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

    try {
        $result = Invoke-NinjaOneApiInline -Method POST -Endpoint 'v2/itam/unmanaged-device' -Body $body -Session $Session
        $nodeId = $result.nodeId

        # Optional: PATCH custom fields if CSV has Make, Model, PurchaseDate, PurchaseAmount, AssetStatus, ExpectedLifetime, EndOfLifeDate
        $make             = Get-RowValue -Row $row -ColumnName 'Make'
        $model            = Get-RowValue -Row $row -ColumnName 'Model'
        $purch            = Get-RowValue -Row $row -ColumnName 'PurchaseDate'
        $amount           = Get-RowValue -Row $row -ColumnName 'PurchaseAmount'
        $assetStatus      = Get-RowValue -Row $row -ColumnName 'AssetStatus'
        $expectedLifetime = Get-RowValue -Row $row -ColumnName 'ExpectedLifetime'
        $eolStr           = Get-RowValue -Row $row -ColumnName 'EndOfLifeDate'

        $parsedPurch = ConvertTo-OptionalDateParseResult -Value $purch
        if (-not $parsedPurch.Success -and -not [string]::IsNullOrWhiteSpace($purch)) {
            $warnings.Add("Row ${rowNum}: PurchaseDate - $($parsedPurch.Message) Value skipped.")
        }
        $parsedAmount = ConvertTo-OptionalIntAmountParseResult -Value $amount
        if (-not $parsedAmount.Success -and -not [string]::IsNullOrWhiteSpace($amount)) {
            $warnings.Add("Row ${rowNum}: PurchaseAmount - $($parsedAmount.Message) Value skipped.")
        }
        $parsedEol = ConvertTo-OptionalDateParseResult -Value $eolStr
        if (-not $parsedEol.Success -and -not [string]::IsNullOrWhiteSpace($eolStr)) {
            $warnings.Add("Row ${rowNum}: EndOfLifeDate - $($parsedEol.Message) Value skipped.")
        }

        if ($make -or $model -or $purch -or $amount -or $serial -or $assetStatus -or $expectedLifetime -or $eolStr) {
            $customFields = @{}
            if ($make)   { $customFields['manufacturer'] = $make }
            if ($model)  { $customFields['model'] = $model }
            if ($serial) { $customFields['itamAssetSerialNumber'] = $serial }
            if ($parsedPurch.Success -and $parsedPurch.Date) {
                $customFields['itamAssetPurchaseDate'] = (ConvertTo-UnixMilliseconds -Date $parsedPurch.Date)
            }
            if ($parsedAmount.Success -and $null -ne $parsedAmount.Amount) {
                $customFields['itamAssetPurchaseAmount'] = $parsedAmount.Amount
            }
            if ($assetStatus) {
                $customFields['itamAssetStatus'] = $assetStatus
            }
            if ($expectedLifetime) {
                $customFields['itamAssetExpectedLifetime'] = $expectedLifetime.Trim().ToLower()
            }
            if ($parsedEol.Success -and $parsedEol.Date) {
                $customFields['itamAssetEndOfLifeDate'] = (ConvertTo-UnixMilliseconds -Date $parsedEol.Date)
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
if ($warnings.Count -gt 0) {
    foreach ($w in $warnings) {
        Write-Warning $w
    }
}
if ($errors.Count -gt 0) {
    foreach ($e in $errors) {
        Write-Warning $e
    }
}

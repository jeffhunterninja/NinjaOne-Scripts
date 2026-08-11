<#
.SYNOPSIS
    Updates NinjaOne ITAM core asset fields on existing unmanaged devices from a CSV file.

.DESCRIPTION
    Update-only: does not create devices. Matches each CSV row by device Id (Ninja node id),
    verifies the device is an unmanaged device (nodeClass UNMANAGED_DEVICE), then PATCHes
    core ITAM asset fields via /v2/device/{id}/custom-fields. All logic is standalone (no
    dot-sourcing).

.PARAMETER CsvPath
    Path to the CSV file. Required column: Id. Optional: Name (logging only), AssetId,
    AssetStatus, SerialNumber, PurchaseDate, PurchaseAmount, ExpectedLifetime, EndOfLifeDate.
    Internal API names (itamAssetId, itamAssetStatus, etc.) are also accepted as headers.

.PARAMETER BaseUrl
    NinjaOne instance base URL (e.g. ca.ninjarmm.com or https://ca.ninjarmm.com). Default: app.ninjarmm.com.

.PARAMETER ClientId
    NinjaOne API application Client ID. Can use $env:NinjaOneClientId if not provided.

.PARAMETER ClientSecret
    NinjaOne API application Client Secret. Can use $env:NinjaOneClientSecret if not provided.

.PARAMETER OverwriteEmptyValues
    When set, empty CSV cells are sent as null to clear existing field values. Default: skip empty cells.

.PARAMETER SkipErrors
    Continue on per-row failure and report all errors at the end. Default: on.

.PARAMETER WhatIf
    Resolve devices and print PATCH payloads without calling the API.

.EXAMPLE
    .\Import-NinjaUnmanagedAssetFieldsFromCsv.ps1 -CsvPath ".\Import-UnmanagedAssetFields-Example.csv" -BaseUrl ca.ninjarmm.com

.EXAMPLE
    .\Import-NinjaUnmanagedAssetFieldsFromCsv.ps1 -CsvPath ".\asset-updates.csv" -BaseUrl app.ninjarmm.com -WhatIf
#>
[CmdletBinding()]
param(
    [Parameter()]
    [string]$CsvPath = '.\Import-UnmanagedAssetFields-Example.csv',

    [Parameter()]
    [string]$BaseUrl = 'app.ninjarmm.com',

    [Parameter()]
    [string]$ClientId = '',

    [Parameter()]
    [string]$ClientSecret = '',

    [Parameter()]
    [switch]$OverwriteEmptyValues,

    [Parameter()]
    [switch]$SkipErrors = $true,

    [Parameter()]
    [switch]$WhatIf
)

if ($CsvPath -eq '.\Import-UnmanagedAssetFields-Example.csv' -and $PSScriptRoot) {
    $CsvPath = Join-Path $PSScriptRoot 'Import-UnmanagedAssetFields-Example.csv'
}

$ErrorActionPreference = 'Stop'

#region API helpers (standalone)

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
        if ($null -ne $Body) {
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

function Get-NinjaOneDevicesPagedInline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [PSCustomObject]$Session,
        [int]$PageSize = 500
    )
    $out = [System.Collections.Generic.List[object]]::new()
    $seen = @{}

    foreach ($baseEp in @('v2/devices-detailed', 'v2/devices')) {
        $after = $null
        $gotAny = $false
        do {
            $queryParts = @("pageSize=$PageSize")
            if (-not [string]::IsNullOrWhiteSpace($after)) {
                $queryParts += "after=$([uri]::EscapeDataString($after))"
            }
            $query = $queryParts -join '&'
            $resp = $null
            try {
                $resp = Invoke-NinjaOneApiInline -Method GET -Endpoint $baseEp -Query $query -Session $Session
            } catch {
                Write-Verbose "Device query '$baseEp?$query' failed: $($_.Exception.Message)"
                break
            }

            $items = ConvertTo-ApiItemArrayInline -Response $resp
            foreach ($d in $items) {
                if (-not $d -or -not $d.PSObject.Properties['id']) { continue }
                $devId = 0
                if (-not [int]::TryParse([string]$d.id, [ref]$devId) -or $devId -le 0) { continue }
                if ($seen.ContainsKey($devId)) { continue }
                $seen[$devId] = $true
                [void]$out.Add($d)
                $gotAny = $true
            }

            $after = $null
            if ($resp -and $resp.PSObject.Properties['pageInfo'] -and $resp.pageInfo) {
                $hasMore = $false
                if ($resp.pageInfo.PSObject.Properties['hasMore']) {
                    $hasMore = [bool]$resp.pageInfo.hasMore
                }
                $nextCursor = $null
                if ($resp.pageInfo.PSObject.Properties['nextCursor'] -and $null -ne $resp.pageInfo.nextCursor) {
                    $nextCursor = [string]$resp.pageInfo.nextCursor
                }
                if ($hasMore -and -not [string]::IsNullOrWhiteSpace($nextCursor)) {
                    $after = $nextCursor
                }
            }
        } while ($after)

        if ($gotAny -or $out.Count -gt 0) { break }

        # Fallback: limit/offset paging if cursor paging returned nothing
        try {
            $fallback = Get-NinjaOnePagedInline -Endpoint $baseEp -Session $Session -PageSize $PageSize
            foreach ($d in $fallback) {
                if (-not $d -or -not $d.PSObject.Properties['id']) { continue }
                $devId = 0
                if (-not [int]::TryParse([string]$d.id, [ref]$devId) -or $devId -le 0) { continue }
                if ($seen.ContainsKey($devId)) { continue }
                $seen[$devId] = $true
                [void]$out.Add($d)
            }
            if ($out.Count -gt 0) { break }
        } catch {
            Write-Verbose "Limit/offset fallback for '$baseEp' failed: $($_.Exception.Message)"
        }
    }

    return @($out)
}

#endregion

#region Value helpers

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

function Get-RowValue {
    param(
        [PSCustomObject]$Row,
        [string[]]$ColumnNames
    )
    foreach ($columnName in $ColumnNames) {
        $prop = $Row.PSObject.Properties | Where-Object { $_.Name -ieq $columnName } | Select-Object -First 1
        if (-not $prop) { continue }
        $v = $prop.Value -as [string]
        if ($null -eq $v) { return '' }
        return $v.Trim()
    }
    return $null  # column(s) not present in CSV
}

function Test-ColumnPresent {
    param(
        [string[]]$AllColumns,
        [string[]]$ColumnNames
    )
    foreach ($columnName in $ColumnNames) {
        if ($AllColumns | Where-Object { $_ -ieq $columnName }) { return $true }
    }
    return $false
}

function Get-DeviceDisplayName {
    param($Device)
    if (-not $Device) { return $null }
    foreach ($propName in @('displayName', 'systemName', 'name')) {
        $p = $Device.PSObject.Properties[$propName]
        if ($p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) {
            return ([string]$p.Value).Trim()
        }
    }
    return $null
}

#endregion

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
    BaseUrl      = $base
    AuthHeader   = "Bearer $($tokenResp.access_token)"
    ExpiresAt    = if ($tokenResp.expires_in) { (Get-Date).AddSeconds([int]$tokenResp.expires_in - 60) } else { (Get-Date).AddMinutes(55) }
    ClientId     = $ClientId
    ClientSecret = $ClientSecret
}

# Validate CSV exists and load
if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}
$rows = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8)
if (-not $rows -or $rows.Count -eq 0) {
    Write-Warning 'CSV is empty or has no data rows.'
    exit 0
}

$allColumns = @($rows[0].PSObject.Properties.Name)
if (-not (Test-ColumnPresent -AllColumns $allColumns -ColumnNames @('Id'))) {
    Write-Error "CSV must have required column 'Id'. Found columns: $($allColumns -join ', ')"
    exit 1
}

Write-Host "CSV loaded: $($rows.Count) row(s). Loading devices from NinjaOne..."
$devices = Get-NinjaOneDevicesPagedInline -Session $Session
$deviceById = @{}
foreach ($d in $devices) {
    $devId = 0
    if ([int]::TryParse([string]$d.id, [ref]$devId) -and $devId -gt 0) {
        $deviceById[$devId] = $d
    }
}
Write-Host "Indexed $($deviceById.Count) device(s)."

$updated  = 0
$skipped  = 0
$failed   = 0
$errors   = [System.Collections.Generic.List[string]]::new()
$warnings = [System.Collections.Generic.List[string]]::new()
$rowNum   = 0

# Friendly + internal header aliases for core asset fields
$fieldSpecs = @(
    @{ ApiKey = 'itamAssetId';             Aliases = @('AssetId', 'itamAssetId');             Kind = 'text' }
    @{ ApiKey = 'itamAssetStatus';         Aliases = @('AssetStatus', 'itamAssetStatus');     Kind = 'text' }
    @{ ApiKey = 'itamAssetSerialNumber';   Aliases = @('SerialNumber', 'itamAssetSerialNumber'); Kind = 'text' }
    @{ ApiKey = 'itamAssetPurchaseDate';   Aliases = @('PurchaseDate', 'itamAssetPurchaseDate'); Kind = 'date' }
    @{ ApiKey = 'itamAssetPurchaseAmount'; Aliases = @('PurchaseAmount', 'itamAssetPurchaseAmount'); Kind = 'amount' }
    @{ ApiKey = 'itamAssetExpectedLifetime'; Aliases = @('ExpectedLifetime', 'itamAssetExpectedLifetime'); Kind = 'lifetime' }
    @{ ApiKey = 'itamAssetEndOfLifeDate';  Aliases = @('EndOfLifeDate', 'itamAssetEndOfLifeDate'); Kind = 'date' }
)

foreach ($row in $rows) {
    $rowNum++
    $idRaw = Get-RowValue -Row $row -ColumnNames @('Id')
    $nameForLog = Get-RowValue -Row $row -ColumnNames @('Name')
    if ([string]::IsNullOrWhiteSpace($nameForLog)) { $nameForLog = $null }

    $deviceId = 0
    if ([string]::IsNullOrWhiteSpace($idRaw) -or -not [int]::TryParse($idRaw, [ref]$deviceId) -or $deviceId -le 0) {
        $msg = "Row ${rowNum}: Id is required and must be a positive integer."
        $errors.Add($msg)
        if (-not $SkipErrors) { throw $msg }
        $failed++
        continue
    }

    $label = if ($nameForLog) { "$nameForLog (ID $deviceId)" } else { "ID $deviceId" }

    if (-not $deviceById.ContainsKey($deviceId)) {
        $msg = "Row ${rowNum}: Device not found for $label."
        $errors.Add($msg)
        if (-not $SkipErrors) { throw $msg }
        $failed++
        continue
    }

    $device = $deviceById[$deviceId]
    $nodeClass = $null
    if ($device.PSObject.Properties['nodeClass']) {
        $nodeClass = [string]$device.nodeClass
    }
    if (-not [string]::Equals($nodeClass, 'UNMANAGED_DEVICE', [StringComparison]::OrdinalIgnoreCase)) {
        $ncLabel = if ([string]::IsNullOrWhiteSpace($nodeClass)) { '(missing)' } else { $nodeClass }
        $msg = "Row ${rowNum}: Skipping $label — nodeClass is '$ncLabel' (expected UNMANAGED_DEVICE)."
        $warnings.Add($msg)
        Write-Warning $msg
        $skipped++
        continue
    }

    $displayName = Get-DeviceDisplayName -Device $device
    if (-not $displayName) { $displayName = $label }

    $patchBody = @{}

    foreach ($spec in $fieldSpecs) {
        if (-not (Test-ColumnPresent -AllColumns $allColumns -ColumnNames $spec.Aliases)) {
            continue
        }

        $raw = Get-RowValue -Row $row -ColumnNames $spec.Aliases
        # Column present but empty
        if ([string]::IsNullOrWhiteSpace($raw)) {
            if ($OverwriteEmptyValues) {
                $patchBody[$spec.ApiKey] = $null
            }
            continue
        }

        switch ($spec.Kind) {
            'text' {
                $patchBody[$spec.ApiKey] = $raw
            }
            'lifetime' {
                $patchBody[$spec.ApiKey] = $raw.Trim().ToLower()
            }
            'date' {
                $parsed = ConvertTo-OptionalDateParseResult -Value $raw
                if (-not $parsed.Success) {
                    $warnings.Add("Row ${rowNum}: $($spec.Aliases[0]) - $($parsed.Message) Value skipped.")
                    continue
                }
                if ($parsed.Date) {
                    $patchBody[$spec.ApiKey] = (ConvertTo-UnixMilliseconds -Date $parsed.Date)
                }
            }
            'amount' {
                $parsed = ConvertTo-OptionalIntAmountParseResult -Value $raw
                if (-not $parsed.Success) {
                    $warnings.Add("Row ${rowNum}: $($spec.Aliases[0]) - $($parsed.Message) Value skipped.")
                    continue
                }
                if ($null -ne $parsed.Amount) {
                    $patchBody[$spec.ApiKey] = $parsed.Amount
                }
            }
        }
    }

    if ($patchBody.Count -eq 0) {
        $msg = "Row ${rowNum}: No asset field values to update for $label."
        $warnings.Add($msg)
        Write-Warning $msg
        $skipped++
        continue
    }

    $json = $patchBody | ConvertTo-Json -Depth 5 -Compress
    if ($WhatIf) {
        Write-Host "[WhatIf] Would PATCH device/$deviceId ($displayName): $json"
        $updated++
        continue
    }

    try {
        Invoke-NinjaOneApiInline -Method PATCH -Endpoint "v2/device/$deviceId/custom-fields" -Body $patchBody -Session $Session | Out-Null
        Write-Verbose "Updated: $displayName (ID $deviceId) — $json"
        $updated++
    } catch {
        $msg = "Row ${rowNum} ($label): $($_.Exception.Message)"
        $errors.Add($msg)
        if (-not $SkipErrors) { throw }
        $failed++
    }
}

# Summary
$mode = if ($WhatIf) { 'WhatIf complete' } else { 'Import complete' }
Write-Host "$mode. Updated: $updated, Skipped: $skipped, Failed: $failed."
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

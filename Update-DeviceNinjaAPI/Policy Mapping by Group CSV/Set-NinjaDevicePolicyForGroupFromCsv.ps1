#Requires -Version 5.1
<#
.SYNOPSIS
  Sets device policy for all devices in one or more NinjaOne groups from a CSV file.

.DESCRIPTION
  Reads a CSV with policyName, groupName, and groupId columns. Resolves names to IDs via
  GET /v2/policies and GET /v2/groups, then PATCHes each device in each group with the
  resolved policyId. When both groupId and groupName are present, groupId takes precedence.

  Evaluate and test in a controlled setting before use in production.

.EXIT CODES
  0 = Success (all device updates succeeded, or groups had no devices)
  1 = Auth failure, API failure, or one or more device updates failed
  2 = Validation error (missing/invalid parameters, credentials, or CSV resolve errors)

.PARAMETER CsvPath
  Path to the CSV file. Optional on the command line; if omitted, the script uses $LocalCsvPath
  (edit that variable near the top of the script body). Required columns: policyName, groupName, groupId.
  Per row: policyName required; at least one of groupName or groupId required.
  groupId is used when non-empty; otherwise groupName is resolved via GET /v2/groups.

.PARAMETER NinjaOneInstance
  NinjaOne instance hostname or base URL. Optional; defaults to env:NINJA_BASE_URL or https://app.ninjarmm.com.

.PARAMETER NinjaOneClientId
  OAuth client ID. Optional; defaults to env:NINJA_CLIENT_ID.

.PARAMETER NinjaOneClientSecret
  OAuth client secret. Optional; defaults to env:NINJA_CLIENT_SECRET.

.PARAMETER WhatIf
  Preview which devices would be updated without applying changes. Provided by SupportsShouldProcess.

.PARAMETER UseWsPaths
  Use /ws/oauth/token instead of /oauth/token (required for some instances). Default: $false.

.PARAMETER ThrottleMs
  Delay in milliseconds between each device PATCH. Default: 0.

.EXAMPLE
  .\Set-NinjaDevicePolicyForGroupFromCsv.ps1 -CsvPath .\Set-NinjaDevicePolicyForGroup-Example.csv

.EXAMPLE
  .\Set-NinjaDevicePolicyForGroupFromCsv.ps1 -CsvPath .\mapping.csv -WhatIf -ThrottleMs 300
#>

$ErrorActionPreference = 'Stop'

# Edit $LocalCsvPath when you run the script without -CsvPath (F5, dot-source, etc.).
$LocalCsvPath = ''

if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $CsvPath = $LocalCsvPath
}

# Defaults live here (not in param) so dot-sourcing (. script.ps1) does not pass
# default values as positional args and break parameter binding.
if ([string]::IsNullOrWhiteSpace($NinjaOneInstance)) { $NinjaOneInstance = $env:NINJA_BASE_URL }
if ([string]::IsNullOrWhiteSpace($NinjaOneClientId)) { $NinjaOneClientId = $env:NINJA_CLIENT_ID }
if ([string]::IsNullOrWhiteSpace($NinjaOneClientSecret)) { $NinjaOneClientSecret = $env:NINJA_CLIENT_SECRET }
if (-not $PSBoundParameters.ContainsKey('ThrottleMs')) { $ThrottleMs = 0 }
if (-not (Get-Variable -Name WhatIf -ErrorAction SilentlyContinue)) { $WhatIf = $false }

if ([string]::IsNullOrWhiteSpace($NinjaOneClientId) -or [string]::IsNullOrWhiteSpace($NinjaOneClientSecret)) {
    Write-Error "NinjaOneClientId and NinjaOneClientSecret are required. Pass -NinjaOneClientId and -NinjaOneClientSecret, or set NINJA_CLIENT_ID and NINJA_CLIENT_SECRET environment variables."
    exit 2
}

if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
    Write-Error "CsvPath not found or not a file: $CsvPath"
    exit 2
}

$rawRows = Import-Csv -LiteralPath $CsvPath -Encoding UTF8
if (-not $rawRows -or $rawRows.Count -eq 0) {
    Write-Error "CSV is empty or has no data rows: $CsvPath"
    exit 2
}

$headers = $rawRows[0].PSObject.Properties.Name
$policyCol = $headers | Where-Object { [string]::Equals($_, 'policyName', [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
$groupNameCol = $headers | Where-Object { [string]::Equals($_, 'groupName', [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
$groupIdCol = $headers | Where-Object { [string]::Equals($_, 'groupId', [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
if (-not $policyCol -or -not $groupNameCol -or -not $groupIdCol) {
    Write-Error "CSV must have columns: policyName, groupName, groupId. Found: $($headers -join ', ')"
    exit 2
}

$csvRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$rowNum = 1
foreach ($r in $rawRows) {
    $rowNum++
    $policyName = ($r.PSObject.Properties[$policyCol].Value -as [string]).Trim()
    $groupName = ($r.PSObject.Properties[$groupNameCol].Value -as [string]).Trim()
    $groupIdRaw = ($r.PSObject.Properties[$groupIdCol].Value -as [string]).Trim()

    if ([string]::IsNullOrWhiteSpace($policyName) -and [string]::IsNullOrWhiteSpace($groupName) -and [string]::IsNullOrWhiteSpace($groupIdRaw)) {
        continue
    }

    if ([string]::IsNullOrWhiteSpace($policyName)) {
        Write-Error "Row $rowNum : policyName is required and cannot be empty (groupName='$groupName', groupId='$groupIdRaw')."
        exit 2
    }
    if ([string]::IsNullOrWhiteSpace($groupName) -and [string]::IsNullOrWhiteSpace($groupIdRaw)) {
        Write-Error "Row $rowNum : at least one of groupName or groupId is required (policyName='$policyName')."
        exit 2
    }

    $csvRows.Add([PSCustomObject]@{
            PolicyName = $policyName
            GroupName  = $groupName
            GroupIdRaw = $groupIdRaw
            RowNumber  = $rowNum
        })
}

if ($csvRows.Count -eq 0) {
    Write-Error "CSV has no valid data rows with policyName and groupName or groupId."
    exit 2
}

$NinjaBaseUrl = $NinjaOneInstance
if ([string]::IsNullOrWhiteSpace($NinjaBaseUrl)) { $NinjaBaseUrl = 'https://app.ninjarmm.com' }
$NinjaBaseUrl = $NinjaBaseUrl.Trim()
if ($NinjaBaseUrl -notmatch '^https?://') { $NinjaBaseUrl = "https://$NinjaBaseUrl" }

$baseUrl = $NinjaBaseUrl
$oauthPath = if ($UseWsPaths) { 'ws/oauth/token' } else { 'oauth/token' }
$tokenUri = "$baseUrl/$oauthPath"
$body = @{
    grant_type    = 'client_credentials'
    client_id     = $NinjaOneClientId.Trim()
    client_secret = $NinjaOneClientSecret.Trim()
    scope         = 'monitoring management'
}
$authHeaders = @{
    'accept'       = 'application/json'
    'Content-Type' = 'application/x-www-form-urlencoded'
}
try {
    $authResp = Invoke-RestMethod -Uri $tokenUri -Method POST -Headers $authHeaders -Body $body
    $accessToken = $authResp | Select-Object -ExpandProperty 'access_token' -ErrorAction SilentlyContinue
    if (-not $accessToken) { throw "Token response did not include access_token." }
    Write-Verbose "Obtained NinjaOne access token."
} catch {
    Write-Error "Failed to obtain NinjaOne access token. $($_.Exception.Message)"
    exit 1
}
$headers = @{
    'accept'        = 'application/json'
    'Authorization' = "Bearer $accessToken"
}

function ConvertTo-NinjaApiList {
    param($Response)
    if ($null -eq $Response) { return @() }
    if ($Response -is [System.Array]) { return @($Response) }
    foreach ($prop in @('results', 'data', 'items', 'groups')) {
        if ($Response.PSObject.Properties[$prop]) {
            $inner = $Response.$prop
            if ($inner -is [System.Array]) { return @($inner) }
            if ($null -ne $inner) { return @($inner) }
        }
    }
    return @($Response)
}

function Get-NinjaScalarInt {
    param($Value)
    if ($null -eq $Value) { return $null }
    $current = $Value
    if ($current -is [System.Array]) {
        $current = $current | Where-Object { $null -ne $_ } | Select-Object -First 1
    }
    if ($null -eq $current) { return $null }
    try { return [int]$current } catch { return $null }
}

function Get-NinjaApiEntityList {
    param($Response)
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($item in (ConvertTo-NinjaApiList -Response $Response)) {
        if ($item -is [System.Array]) {
            foreach ($sub in $item) {
                if ($null -ne $sub) { [void]$list.Add($sub) }
            }
        } else {
            [void]$list.Add($item)
        }
    }
    return $list
}

try {
    $groups = Get-NinjaApiEntityList -Response (Invoke-RestMethod -Uri "$baseUrl/v2/groups" -Method GET -Headers $headers)
} catch {
    Write-Error "Failed to get groups (GET /v2/groups). $($_.Exception.Message)"
    exit 1
}

try {
    $policies = Get-NinjaApiEntityList -Response (Invoke-RestMethod -Uri "$baseUrl/v2/policies" -Method GET -Headers $headers)
} catch {
    Write-Error "Failed to get policies (GET /v2/policies). $($_.Exception.Message)"
    exit 1
}

$groupIdByNumericId = @{}
foreach ($g in $groups) {
    $gid = Get-NinjaScalarInt -Value $g.id
    if ($null -ne $gid -and $gid -gt 0) { $groupIdByNumericId[$gid] = $g }
}

$errors = [System.Collections.Generic.List[string]]::new()
$groupPolicyMap = [System.Collections.Generic.Dictionary[int, [PSCustomObject]]]::new()

foreach ($row in $csvRows) {
    $resolvedGroupId = $null
    $groupIdRaw = $row.GroupIdRaw

    if (-not [string]::IsNullOrWhiteSpace($groupIdRaw)) {
        if ($groupIdRaw -notmatch '^\d+$' -or [int]$groupIdRaw -le 0) {
            $errors.Add("Row $($row.RowNumber): groupId must be a positive integer (value='$groupIdRaw', policyName='$($row.PolicyName)').")
            continue
        }
        $resolvedGroupId = [int]$groupIdRaw
        if (-not $groupIdByNumericId.ContainsKey($resolvedGroupId)) {
            $errors.Add("Row $($row.RowNumber): groupId $resolvedGroupId not found in GET /v2/groups (policyName='$($row.PolicyName)').")
            continue
        }
    } else {
        $groupName = $row.GroupName
        $group = $groups | Where-Object { $_.name -and ($_.name -eq $groupName) } | Select-Object -First 1
        if (-not $group) {
            $group = $groups | Where-Object { $_.name -and [string]::Equals($_.name, $groupName, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
        }
        if (-not $group) {
            $errors.Add("Row $($row.RowNumber): groupName not found in GET /v2/groups: '$groupName' (policyName='$($row.PolicyName)').")
            continue
        }
        $resolvedGroupId = Get-NinjaScalarInt -Value $group.id
        if ($null -eq $resolvedGroupId -or $resolvedGroupId -le 0) {
            $errors.Add("Row $($row.RowNumber): groupName '$groupName' resolved to an invalid group id.")
            continue
        }
    }

    $policyName = $row.PolicyName
    $policy = $policies | Where-Object { $_.name -and ($_.name -eq $policyName) } | Select-Object -First 1
    if (-not $policy) {
        $policy = $policies | Where-Object { $_.name -and [string]::Equals($_.name, $policyName, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
    }
    if (-not $policy) {
        $errors.Add("Row $($row.RowNumber): policyName not found in GET /v2/policies: '$policyName' (groupId=$resolvedGroupId).")
        continue
    }

    $policyId = Get-NinjaScalarInt -Value $policy.id
    if ($null -eq $policyId -or $policyId -le 0) {
        $errors.Add("Row $($row.RowNumber): policyName '$policyName' resolved to an invalid policy id (groupId=$resolvedGroupId).")
        continue
    }

    $groupPolicyMap[$resolvedGroupId] = [PSCustomObject]@{
        GroupId    = $resolvedGroupId
        PolicyId   = $policyId
        PolicyName = $policyName
    }
}

if ($errors.Count -gt 0) {
    foreach ($e in $errors) { Write-Error $e }
    Write-Error "Resolve errors: $($errors.Count). Fix CSV or NinjaOne data and re-run."
    exit 2
}

$allResults = [System.Collections.Generic.List[object]]::new()
$anyDeviceFailed = $false

foreach ($entry in $groupPolicyMap.Values) {
    $groupId = $entry.GroupId
    $policyId = $entry.PolicyId
    $policyName = $entry.PolicyName

    $groupDeviceIdsUrl = "$baseUrl/v2/group/$groupId/device-ids"
    try {
        $groupDevices = Invoke-RestMethod -Uri $groupDeviceIdsUrl -Method GET -Headers $headers
    } catch {
        Write-Error "Failed to fetch device IDs for group $groupId (policy '$policyName'). $_"
        exit 1
    }

    $deviceIds = @($groupDevices)
    if (-not $deviceIds -or $deviceIds.Count -eq 0) {
        Write-Host "No devices found in group $groupId (policy '$policyName')."
        $allResults.Add([PSCustomObject]@{
                GroupId         = $groupId
                PolicyId        = $policyId
                PolicyName      = $policyName
                TotalDevices    = 0
                UpdatedCount    = 0
                FailedCount     = 0
                FailedDeviceIds = @()
            })
        continue
    }

    Write-Verbose "Found $($deviceIds.Count) device(s) in group $groupId (policy '$policyName')."
    $updated = 0
    $failed = 0
    $failedDeviceIds = [System.Collections.Generic.List[object]]::new()
    $total = $deviceIds.Count
    $current = 0

    foreach ($deviceId in $deviceIds) {
        $current++
        $deviceUpdateUrl = "$baseUrl/api/v2/device/$deviceId"
        $requestBody = @{ policyId = $policyId }
        $json = $requestBody | ConvertTo-Json

        if ($WhatIf) {
            Write-Host "[WhatIf] Group $groupId : would assign policy $policyId to $deviceId"
        } else {
            try {
                Write-Progress -Activity "Updating devices in group $groupId" -Status "Assigning policy to device $current of $total" -PercentComplete ([int](100 * $current / $total))
                Invoke-RestMethod -Method Patch -Uri $deviceUpdateUrl -Headers $headers -Body $json -ContentType 'application/json'
                Write-Host "Group $groupId : assigning policy $policyId to $deviceId"
                $updated++
            } catch {
                Write-Error "Failed to update device $deviceId in group $groupId. $_"
                $failed++
                $failedDeviceIds.Add($deviceId)
            }
        }
        if ($ThrottleMs -gt 0 -and $current -lt $total) {
            Start-Sleep -Milliseconds $ThrottleMs
        }
    }

    Write-Progress -Activity "Updating devices in group $groupId" -Completed

    $groupResult = [PSCustomObject]@{
        GroupId         = $groupId
        PolicyId        = $policyId
        PolicyName      = $policyName
        TotalDevices    = $total
        UpdatedCount    = $updated
        FailedCount     = $failed
        FailedDeviceIds = @($failedDeviceIds)
    }
    $allResults.Add($groupResult)
    Write-Output $groupResult

    if ($failed -gt 0) {
        $anyDeviceFailed = $true
    }
}

if ($anyDeviceFailed) {
    $failedSummary = ($allResults | Where-Object { $_.FailedCount -gt 0 } | ForEach-Object {
            "group $($_.GroupId): $($_.FailedDeviceIds -join ', ')"
        }) -join '; '
    Write-Error "One or more device updates failed. $failedSummary"
    exit 1
}

exit 0

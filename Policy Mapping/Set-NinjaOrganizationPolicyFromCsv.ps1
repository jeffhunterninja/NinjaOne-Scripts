#Requires -Version 5.1
<#
.SYNOPSIS
  Updates NinjaOne organization policy assignments from a CSV file.

.DESCRIPTION
  Reads a CSV with columns OrganizationName, PolicyName, and DeviceRole. Resolves
  names to IDs via NinjaOne APIs (organizations, policies, and GET /api/v2/roles).
  For each organization in the CSV, builds the list of nodeRoleId-to-policyId
  assignments and PUTs them to api/v2/organization/{id}/policies. Uses client
  credentials (machine-to-machine); no browser.

.PARAMETER CsvPath
  Path to the CSV file. Required columns: OrganizationName, PolicyName, DeviceRole.
  DeviceRole must be non-empty; resolved via GET /api/v2/roles (name or ID).

.PARAMETER NinjaOneInstance
  NinjaOne base URL (e.g. https://app.ninjarmm.com). Optional; defaults to env NINJA_BASE_URL or https://app.ninjarmm.com.

.PARAMETER NinjaOneClientId
  OAuth client ID. Optional; defaults to env NINJA_CLIENT_ID.

.PARAMETER NinjaOneClientSecret
  OAuth client secret. Optional; defaults to env NINJA_CLIENT_SECRET.

.PARAMETER WhatIf
  List organizations and intended (role -> policy) changes without calling PUT.

.EXIT CODES
  0 = Success
  1 = Auth or API error
  2 = Validation error (missing/invalid parameters, credentials, or CSV data)
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$NinjaOneInstance,

    [Parameter(Mandatory = $false)]
    [string]$NinjaOneClientId,

    [Parameter(Mandatory = $false)]
    [string]$NinjaOneClientSecret,

    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# Resolve credentials and base URL
$clientId = if (-not [string]::IsNullOrWhiteSpace($NinjaOneClientId)) { $NinjaOneClientId } else { $env:NINJA_CLIENT_ID }
$clientSecret = if (-not [string]::IsNullOrWhiteSpace($NinjaOneClientSecret)) { $NinjaOneClientSecret } else { $env:NINJA_CLIENT_SECRET }
$NinjaBaseUrl = if (-not [string]::IsNullOrWhiteSpace($NinjaOneInstance)) { $NinjaOneInstance } else { if (-not [string]::IsNullOrWhiteSpace($env:NINJA_BASE_URL)) { $env:NINJA_BASE_URL } else { 'https://app.ninjarmm.com' } }

if ([string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($clientSecret)) {
    Write-Error "Missing required API credentials. Set -NinjaOneClientId and -NinjaOneClientSecret, or env NINJA_CLIENT_ID and NINJA_CLIENT_SECRET."
    exit 2
}

$NinjaBaseUrl = $NinjaBaseUrl.Trim()
if ($NinjaBaseUrl -notmatch '^https?://') { $NinjaBaseUrl = "https://$NinjaBaseUrl" }

# Validate CSV path
if (-not (Test-Path -LiteralPath $CsvPath -PathType Leaf)) {
    Write-Error "CsvPath not found or not a file: $CsvPath"
    exit 2
}

# Load CSV and normalize column names (case-insensitive)
$rawRows = Import-Csv -LiteralPath $CsvPath -Encoding UTF8
if (-not $rawRows -or $rawRows.Count -eq 0) {
    Write-Error "CSV is empty or has no data rows: $CsvPath"
    exit 2
}

$headers = $rawRows[0].PSObject.Properties.Name
$orgCol = $headers | Where-Object { $_ -eq 'OrganizationName' }
$policyCol = $headers | Where-Object { $_ -eq 'PolicyName' }
$roleCol = $headers | Where-Object { $_ -eq 'DeviceRole' }
if (-not $orgCol -or -not $policyCol -or -not $roleCol) {
    Write-Error "CSV must have columns: OrganizationName, PolicyName, DeviceRole. Found: $($headers -join ', ')"
    exit 2
}

# Build normalized rows and validate DeviceRole required
$csvRows = [System.Collections.Generic.List[PSCustomObject]]::new()
$rowNum = 1
foreach ($r in $rawRows) {
    $rowNum++
    $orgName = ($r.PSObject.Properties[$orgCol].Value -as [string]).Trim()
    $policyName = ($r.PSObject.Properties[$policyCol].Value -as [string]).Trim()
    $deviceRole = ($r.PSObject.Properties[$roleCol].Value -as [string]).Trim()
    if ([string]::IsNullOrWhiteSpace($deviceRole)) {
        Write-Error "Row $rowNum : DeviceRole is required and cannot be empty (OrganizationName='$orgName', PolicyName='$policyName')."
        exit 2
    }
    $csvRows.Add([PSCustomObject]@{
            OrganizationName = $orgName
            PolicyName      = $policyName
            DeviceRole      = $deviceRole
        })
}

# Remove rows with empty org or policy so we can skip blank lines; then require at least one
$csvRows = @($csvRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.OrganizationName) -and -not [string]::IsNullOrWhiteSpace($_.PolicyName) })
if ($csvRows.Count -eq 0) {
    Write-Error "CSV has no valid rows with non-empty OrganizationName and PolicyName."
    exit 2
}

# Authenticate (client credentials)
$tokenUri = "$NinjaBaseUrl/ws/oauth/token"
$authBody = @{
    grant_type    = 'client_credentials'
    client_id     = $clientId.Trim()
    client_secret = $clientSecret.Trim()
    scope         = 'monitoring management'
}
$authHeaders = @{
    'accept'       = 'application/json'
    'Content-Type' = 'application/x-www-form-urlencoded'
}
try {
    $authResp = Invoke-RestMethod -Uri $tokenUri -Method POST -Headers $authHeaders -Body $authBody
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

# GET /v2/organizations
try {
    $organizations = Invoke-RestMethod -Uri "$NinjaBaseUrl/v2/organizations" -Method GET -Headers $headers
} catch {
    Write-Error "Failed to get organizations. $($_.Exception.Message)"
    exit 1
}

# GET /v2/policies
try {
    $policies = Invoke-RestMethod -Uri "$NinjaBaseUrl/v2/policies" -Method GET -Headers $headers
} catch {
    Write-Error "Failed to get policies. $($_.Exception.Message)"
    exit 1
}

# GET /api/v2/roles
try {
    $roles = Invoke-RestMethod -Uri "$NinjaBaseUrl/api/v2/roles" -Method GET -Headers $headers
} catch {
    Write-Error "Failed to get roles (GET /api/v2/roles). $($_.Exception.Message)"
    exit 1
}

# Resolve role by name (case-insensitive) or by numeric ID
function Resolve-RoleId {
    param([string]$DeviceRole, [array]$RolesList)
    $trimmed = $DeviceRole.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return $null }
    if ($trimmed -match '^\d+$') {
        $byId = $RolesList | Where-Object { $_.id -eq [int]$trimmed } | Select-Object -First 1
        if ($byId) { return $byId.id }
    }
    $byName = $RolesList | Where-Object { $_.name -and ($_.name -eq $trimmed) } | Select-Object -First 1
    if ($byName) { return $byName.id }
    $byNameCi = $RolesList | Where-Object { $_.name -and [string]::Equals($_.name, $trimmed, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
    if ($byNameCi) { return $byNameCi.id }
    return $null
}

# Build assignments per organization: orgName -> [ { nodeRoleId, policyId } ]
# Dedupe by nodeRoleId: last row wins for a given role
$orgAssignments = [System.Collections.Generic.Dictionary[string, [System.Collections.Generic.Dictionary[int, int]]]]::new()
$errors = [System.Collections.Generic.List[string]]::new()

foreach ($row in $csvRows) {
    $orgName = $row.OrganizationName
    $policyName = $row.PolicyName
    $deviceRole = $row.DeviceRole

    $org = $organizations | Where-Object { $_.name -eq $orgName } | Select-Object -First 1
    if (-not $org) {
        $errors.Add("Organization not found: '$orgName' (PolicyName='$policyName', DeviceRole='$deviceRole')")
        continue
    }

    $policy = $policies | Where-Object { $_.name -eq $policyName } | Select-Object -First 1
    if (-not $policy) {
        $errors.Add("Policy not found: '$policyName' (OrganizationName='$orgName', DeviceRole='$deviceRole')")
        continue
    }

    $roleId = Resolve-RoleId -DeviceRole $deviceRole -RolesList $roles
    if (-not $roleId) {
        $errors.Add("DeviceRole not found in /api/v2/roles: '$deviceRole' (OrganizationName='$orgName', PolicyName='$policyName')")
        continue
    }

    if (-not $orgAssignments.ContainsKey($orgName)) {
        $orgAssignments[$orgName] = [System.Collections.Generic.Dictionary[int, int]]::new()
    }
    $orgAssignments[$orgName][$roleId] = $policy.id
}

if ($errors.Count -gt 0) {
    foreach ($e in $errors) { Write-Error $e }
    Write-Error "Resolve errors: $($errors.Count). Fix CSV or NinjaOne data and re-run."
    exit 2
}

# Apply: for each org, PUT policy assignments
$orgIds = @{}
foreach ($o in $organizations) { $orgIds[$o.name] = $o.id }

foreach ($orgName in $orgAssignments.Keys) {
    $roleToPolicy = $orgAssignments[$orgName]
    $orgId = $orgIds[$orgName]
    $assignments = @()
    foreach ($rid in $roleToPolicy.Keys) {
        $assignments += @{ nodeRoleId = $rid; policyId = $roleToPolicy[$rid] }
    }
    $bodyJson = $assignments | ConvertTo-Json -Compress

    if ($PSCmdlet.ShouldProcess($orgName, "PUT organization policy assignments ($($assignments.Count) role(s))")) {
        $policiesUrl = "$NinjaBaseUrl/api/v2/organization/$orgId/policies"
        try {
            Invoke-RestMethod -Method PUT -Uri $policiesUrl -Headers $headers -Body $bodyJson -ContentType 'application/json' | Out-Null
            Write-Host "Updated policy assignment for organization: $orgName ($($assignments.Count) role(s))."
        } catch {
            Write-Error "Failed to assign policies for organization '$orgName'. $($_.Exception.Message)"
            exit 1
        }
    } else {
        $whatIfDetail = ($assignments | ForEach-Object { "role $($_.nodeRoleId)->policy $($_.policyId)" }) -join ', '
        Write-Host "[WhatIf] Would update organization '$orgName' with $($assignments.Count) assignment(s): $whatIfDetail"
    }
}

exit 0

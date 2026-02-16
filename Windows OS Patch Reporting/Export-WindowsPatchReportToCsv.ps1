<#
This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
This script exports Windows device patch installations (one line per KB per device) to CSV for a given month.

Uses the same data and logic as Publish-WindowsPatchReport.ps1 but outputs CSV instead of NinjaOne Documents/KB.

The NinjaOne module can be forked here: https://github.com/lwhitelock/NinjaOneDocs
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$ReportMonth = $env:reportMonth,  # Optional (e.g., "December 2024"); defaults to current month
    [Parameter()]
    [string]$OutputPath = $env:outputPath,   # Optional; default: .\WindowsPatchReport_<YYYYMM>.csv
    [Parameter()]
    [Switch]$PerOrganization = [System.Convert]::ToBoolean($env:perOrganization)  # If set, write one CSV per organization
)

# Check for required PowerShell version (7+)
if (!($PSVersionTable.PSVersion.Major -ge 7)) {
    try {
        if (!(Test-Path "$env:SystemDrive\Program Files\PowerShell\7")) {
            Write-Output 'Does not appear Powershell 7 is installed'
            exit 1
        }

        # Refresh PATH
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path', 'User')

        # Restart script in PowerShell 7
        pwsh -File "`"$PSCommandPath`"" @PSBoundParameters

    }
    catch {
        Write-Output 'PowerShell 7 was not installed. Update PowerShell and try again.'
        throw $Error
    }
    finally { exit $LASTEXITCODE }
}

# Install or update the NinjaOneDocs module
try {
    $moduleName = "NinjaOneDocs"
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Install-Module -Name $moduleName -Force -AllowClobber
    }
    Import-Module $moduleName
}
catch {
    Write-Error "Failed to import NinjaOneDocs module. Error: $_"
    exit
}

# NinjaOne credentials - store in secure NinjaOne custom fields or set locally
$NinjaOneInstance = Ninja-Property-Get ninjaoneInstance
$NinjaOneClientId = Ninja-Property-Get ninjaoneClientId
$NinjaOneClientSecret = Ninja-Property-Get ninjaoneClientSecret

if (!$NinjaOneInstance -or !$NinjaOneClientId -or !$NinjaOneClientSecret) {
    Write-Output "Missing required API credentials"
    exit 1
}

try {
    Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientId -NinjaOneClientSecret $NinjaOneClientSecret
}
catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit
}

function Convert-ActivityTime {
    param([Parameter(Mandatory)]$TimeValue)

    if ($TimeValue -is [datetime]) {
        return $TimeValue
    } elseif ($TimeValue -is [int64]) {
        return [System.DateTimeOffset]::FromUnixTimeSeconds($TimeValue).DateTime
    } elseif ($TimeValue -is [double]) {
        $roundedUnixTime = [long][math]::Floor($TimeValue)
        return [System.DateTimeOffset]::FromUnixTimeSeconds($roundedUnixTime).DateTime
    } else {
        return $TimeValue
    }
}

function ConvertTo-QueryParamString {
    param([Parameter(Mandatory)][hashtable]$QueryParams)
    ($QueryParams.GetEnumerator() | ForEach-Object {
        "$($_.Key)=$([System.Uri]::EscapeDataString([string]$_.Value))"
    }) -join '&'
}

function Get-ReportDateRange {
    param([string]$ReportMonth)
    if ($ReportMonth) {
        try {
            $ParsedDate = [datetime]::ParseExact($ReportMonth, "MMMM yyyy", [cultureinfo]::InvariantCulture)
            $FirstDayOfMonth = Get-Date -Year $ParsedDate.Year -Month $ParsedDate.Month -Day 1
        }
        catch {
            Write-Error "Invalid ReportMonth format. Use 'MMMM yyyy' (e.g., 'December 2024')."
            throw
        }
    } else {
        $FirstDayOfMonth = Get-Date -Day 1
    }
    $LastDayOfMonth = $FirstDayOfMonth.AddMonths(1).AddDays(-1)
    [PSCustomObject]@{
        FirstDayOfMonth = $FirstDayOfMonth
        LastDayOfMonth  = $LastDayOfMonth
        currentMonth    = $FirstDayOfMonth.ToString("MMMM")
        currentYear     = $FirstDayOfMonth.ToString("yyyy")
        FirstDayString  = $FirstDayOfMonth.ToString('yyyyMMdd')
        LastDayString   = $LastDayOfMonth.ToString('yyyyMMdd')
    }
}

try {
    $dateRange = Get-ReportDateRange -ReportMonth $ReportMonth
}
catch {
    exit 1
}

$FirstDayOfMonth = $dateRange.FirstDayOfMonth
$LastDayOfMonth = $dateRange.LastDayOfMonth
$currentMonth = $dateRange.currentMonth
$currentYear = $dateRange.currentYear
$FirstDayString = $dateRange.FirstDayString
$LastDayString = $dateRange.LastDayString
$yyyyMM = $FirstDayOfMonth.ToString('yyyyMM')

Write-Output "Generating report for: $($FirstDayOfMonth.ToString('MMMM yyyy'))"
Write-Output "Report Date Range: $($FirstDayOfMonth.ToShortDateString()) - $($LastDayOfMonth.ToShortDateString())"

# Fetch devices and organizations
try {
    $devicesQueryParams = @{ df = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)' }
    $devices = Invoke-NinjaOneRequest -Method GET -Path 'devices-detailed' -QueryParams (ConvertTo-QueryParamString -QueryParams $devicesQueryParams)
    $organizations = Invoke-NinjaOneRequest -Method GET -Path 'organizations'
}
catch {
    Write-Error "Failed to retrieve devices or organizations. Error: $_"
    exit
}

# Fetch patch installations for the date range
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    status          = 'Installed'
    installedBefore = $LastDayString
    installedAfter  = $FirstDayString
}
$QueryParamString = ConvertTo-QueryParamString -QueryParams $queryParams
$patchinstallsResponse = Invoke-NinjaOneRequest -Method GET -Path 'queries/os-patch-installs' -QueryParams $QueryParamString -Paginate
$patchinstalls = if ($patchinstallsResponse.results) { $patchinstallsResponse.results } else { @($patchinstallsResponse | Select-Object -ExpandProperty 'results') }

# Index devices and organizations for lookup
$deviceIndex = @{}
foreach ($device in $devices) {
    $deviceIndex[$device.id] = $device
}
$organizationIndex = @{}
foreach ($organization in $organizations) {
    $organizationIndex[$organization.id] = $organization
}

# Build table: one row per KB per device (exclude Defender Security Intelligence updates)
$table = [System.Collections.ArrayList]::new()
foreach ($patchinstall in $patchinstalls) {
    if ($patchinstall.name -like "*Security Intelligence Update for Microsoft Defender Antivirus*") {
        continue
    }

    $device = $deviceIndex[$patchinstall.deviceId]
    if (-not $device) { continue }

    $organization = $organizationIndex[$device.organizationId]
    if (-not $organization) { continue }

    $installedAt = Convert-ActivityTime $patchinstall.installedAt
    $timestamp = if ($patchinstall.timestamp) { Convert-ActivityTime $patchinstall.timestamp } else { $null }

    $row = [PSCustomObject]@{
        OrganizationName = $organization.name
        DeviceName      = $device.systemName
        PatchName       = $patchinstall.name
        KBNumber        = $patchinstall.kbNumber
        InstalledAt     = $installedAt
        Timestamp       = $timestamp
        DeviceId        = $patchinstall.deviceId
    }
    [void]$table.Add($row)
}

$exportParams = @{
    NoTypeInformation = $true
    Encoding          = [System.Text.Encoding]::UTF8
}

if ($PerOrganization) {
    $table | Group-Object -Property OrganizationName | ForEach-Object {
        $safeName = ($_.Name -replace '[^\w\s\-]', '' -replace '\s+', '_').Trim()
        $path = Join-Path (Get-Location) "WindowsPatchReport_${safeName}_${yyyyMM}.csv"
        $_.Group | Export-Csv -Path $path @exportParams
        Write-Output ('Exported {0} rows to {1}' -f $_.Group.Count, $path)
    }
} else {
    if (-not $OutputPath) {
        $OutputPath = Join-Path (Get-Location) "WindowsPatchReport_${yyyyMM}.csv"
    }
    $table | Export-Csv -Path $OutputPath @exportParams
    Write-Output ('Exported {0} rows to {1}' -f $table.Count, $OutputPath)
}

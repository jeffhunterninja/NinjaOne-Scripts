<#
This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
This script exports Windows device patch installations for a given month to local HTML.

Uses the same data and logic as Publish-WindowsPatchReport.ps1 but outputs HTML locally instead of NinjaOne Documents/KB.
Optional: use the PSWriteHTML module for styled reports (install if missing).

The NinjaOne module can be forked here: https://github.com/lwhitelock/NinjaOneDocs
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$ReportMonth = $env:reportMonth,
    [Parameter()]
    [string]$OutputPath = $env:outputPath,
    [Parameter()]
    [Switch]$PerOrganization = [System.Convert]::ToBoolean($env:perOrganization),
    [Parameter()]
    [Switch]$UsePSWriteHTML = [System.Convert]::ToBoolean($env:usePSWriteHTML)
)

# Check for required PowerShell version (7+)
# When NinjaOne runs with Windows PowerShell 5 and we respawn to pwsh, the child process
# may not have NinjaOne RMM context, so Ninja-Property-Get can return empty → invalid_client.
# If we must respawn, fetch credentials in this process and pass via env so the child can connect.
if (!($PSVersionTable.PSVersion.Major -ge 7)) {
    try {
        if (!(Test-Path "$env:SystemDrive\Program Files\PowerShell\7")) {
            Write-Output 'Does not appear Powershell 7 is installed'
            exit 1
        }
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path', 'User')
        # Load module and get credentials in this process (has NinjaOne context), then pass to pwsh
        $moduleName = "NinjaOneDocs"
        if (-not (Get-Module -ListAvailable -Name $moduleName)) { Install-Module -Name $moduleName -Force -AllowClobber }
        Import-Module $moduleName -ErrorAction Stop
        $env:_NinjaOneInstance = Ninja-Property-Get ninjaoneInstance
        $env:_NinjaOneClientId = Ninja-Property-Get ninjaoneClientId
        $env:_NinjaOneClientSecret = Ninja-Property-Get ninjaoneClientSecret
        pwsh -File "`"$PSCommandPath`"" @PSBoundParameters
    }
    catch {
        Write-Output 'PowerShell 7 was not installed. Update PowerShell and try again.'
        throw $Error
    }
    finally { exit $LASTEXITCODE }
}

# NinjaOneDocs module
try {
    $moduleName = "NinjaOneDocs"
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Install-Module -Name $moduleName -Force -AllowClobber
    }
    Import-Module $moduleName
}
catch {
    Write-Error "Failed to import NinjaOneDocs module. Error: $_"
    exit 1
}

# Use credentials passed from parent (respawn) if present; otherwise get from NinjaOne context
if ($env:_NinjaOneInstance -and $env:_NinjaOneClientId -and $env:_NinjaOneClientSecret) {
    $NinjaOneInstance = $env:_NinjaOneInstance
    $NinjaOneClientId = $env:_NinjaOneClientId
    $NinjaOneClientSecret = $env:_NinjaOneClientSecret
    Remove-Item Env:_NinjaOneInstance, Env:_NinjaOneClientId, Env:_NinjaOneClientSecret -ErrorAction SilentlyContinue
} else {
    $NinjaOneInstance = Ninja-Property-Get ninjaoneInstance
    $NinjaOneClientId = Ninja-Property-Get ninjaoneClientId
    $NinjaOneClientSecret = Ninja-Property-Get ninjaoneClientSecret
}

if (!$NinjaOneInstance -or !$NinjaOneClientId -or !$NinjaOneClientSecret) {
    Write-Output "Missing required API credentials"
    exit 1
}

try {
    Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientId -NinjaOneClientSecret $NinjaOneClientSecret
}
catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit 1
}

function Convert-ActivityTime {
    param([Parameter(Mandatory)]$TimeValue)
    if ($TimeValue -is [datetime]) { return $TimeValue }
    if ($TimeValue -is [int64]) { return [System.DateTimeOffset]::FromUnixTimeSeconds($TimeValue).DateTime }
    if ($TimeValue -is [double]) {
        $roundedUnixTime = [long][math]::Floor($TimeValue)
        return [System.DateTimeOffset]::FromUnixTimeSeconds($roundedUnixTime).DateTime
    }
    return $TimeValue
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

function Get-CategorizedPatchActivities {
    param([Parameter(Mandatory = $false)][array]$Activities = @())
    if ($null -eq $Activities) { $Activities = @() }
    $patchScans = [System.Collections.ArrayList]::new()
    $patchScanFailures = [System.Collections.ArrayList]::new()
    $patchApplicationCycles = [System.Collections.ArrayList]::new()
    $patchApplicationFailures = [System.Collections.ArrayList]::new()
    foreach ($activity in $Activities) {
        if ($activity.activityResult -match "SUCCESS") {
            if ($activity.statusCode -match "PATCH_MANAGEMENT_SCAN_COMPLETED") { [void]$patchScans.Add($activity) }
            elseif ($activity.statusCode -match "PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED") { [void]$patchApplicationCycles.Add($activity) }
        } elseif ($activity.activityResult -match "FAILURE") {
            if ($activity.statusCode -match "PATCH_MANAGEMENT_SCAN_COMPLETED") { [void]$patchScanFailures.Add($activity) }
            elseif ($activity.statusCode -match "PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED") { [void]$patchApplicationFailures.Add($activity) }
        }
    }
    [PSCustomObject]@{
        PatchScans              = @($patchScans)
        PatchScanFailures       = @($patchScanFailures)
        PatchApplicationCycles  = @($patchApplicationCycles)
        PatchApplicationFailures = @($patchApplicationFailures)
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
    exit 1
}

# Patch installations for date range
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    status          = 'Installed'
    installedBefore = $LastDayString
    installedAfter  = $FirstDayString
}
$QueryParamString = ConvertTo-QueryParamString -QueryParams $queryParams
$patchinstallsResponse = Invoke-NinjaOneRequest -Method GET -Path 'queries/os-patch-installs' -QueryParams $QueryParamString -Paginate
$patchinstalls = if ($patchinstallsResponse.results) { $patchinstallsResponse.results } else { @($patchinstallsResponse | Select-Object -ExpandProperty 'results') }

# Patch failures
$queryParams = @{ df = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'; status = 'Failed' }
$QueryParamString = ConvertTo-QueryParamString -QueryParams $queryParams
$patchfailuresResponse = Invoke-NinjaOneRequest -Method GET -Path 'queries/os-patch-installs' -QueryParams $QueryParamString -Paginate
$patchfailures = if ($patchfailuresResponse.results) { $patchfailuresResponse.results } else { @($patchfailuresResponse | Select-Object -ExpandProperty 'results') }

# Activities for scan/apply stats
try {
    $queryParams2 = @{
        df       = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
        class    = 'DEVICE'
        type     = 'PATCH_MANAGEMENT'
        status   = 'in (PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED, PATCH_MANAGEMENT_SCAN_COMPLETED, PATCH_MANAGEMENT_FAILURE)'
        after    = $FirstDayString
        before   = $LastDayString
        pageSize = 1000
    }
    $QueryParamString2 = ConvertTo-QueryParamString -QueryParams $queryParams2
    $allActivities = Invoke-NinjaOneRequest -Method GET -Path 'activities' -QueryParams $QueryParamString2 -Paginate
    # Normalize: API or -Paginate may return null .activities when there are no results
    $activitiesList = if ($null -ne $allActivities -and $null -ne $allActivities.activities) { @($allActivities.activities) } else { @() }
    $firstDayUnix = [System.DateTimeOffset]::new($FirstDayOfMonth).ToUnixTimeSeconds()
    $lastDayUnix = [System.DateTimeOffset]::new($LastDayOfMonth.AddHours(23).AddMinutes(59).AddSeconds(59)).ToUnixTimeSeconds()
    $filteredActivities = $activitiesList | Where-Object { $_.activityTime -ge $firstDayUnix -and $_.activityTime -le $lastDayUnix }
    $userActivities = @($filteredActivities | ForEach-Object { $_.activityTime = Convert-ActivityTime $_.activityTime; $_ })
    $categorized = Get-CategorizedPatchActivities -Activities $userActivities
    $patchScans = $categorized.PatchScans
    $patchApplicationCycles = $categorized.PatchApplicationCycles
}
catch {
    Write-Error "Failed to retrieve activities. Error: $_"
    exit 1
}

# Index devices and organizations
$deviceIndex = @{}
foreach ($device in $devices) { $deviceIndex[$device.id] = $device }
$organizationIndex = @{}
foreach ($organization in $organizations) { $organizationIndex[$organization.id] = $organization }

# Initialize organization tracking
foreach ($organization in $organizations) {
    Add-Member -InputObject $organization -NotePropertyName "PatchScans" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "PatchFailures" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "PatchInstalls" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "PatchApplications" -NotePropertyValue @() -Force
}

foreach ($patchScan in $patchScans) {
    $device = $deviceIndex[$patchScan.deviceId]
    if (-not $device) { continue }
    $organization = $organizationIndex[$device.organizationId]
    if (-not $organization) { continue }
    $organization.PatchScans += $patchScan
}
foreach ($patchApplicationCycle in $patchApplicationCycles) {
    $device = $deviceIndex[$patchApplicationCycle.deviceId]
    if (-not $device) { continue }
    $organization = $organizationIndex[$device.organizationId]
    if (-not $organization) { continue }
    $organization.PatchApplications += $patchApplicationCycle
}
foreach ($patchinstall in $patchinstalls) {
    $device = $deviceIndex[$patchinstall.deviceId]
    if (-not $device) { continue }
    $organization = $organizationIndex[$device.organizationId]
    if (-not $organization) { continue }
    $patchinstall | Add-Member -NotePropertyName "DeviceName" -NotePropertyValue $device.systemName -Force
    $organization.PatchInstalls += $patchinstall
}
foreach ($patchfailure in $patchfailures) {
    $device = $deviceIndex[$patchfailure.deviceId]
    if (-not $device) { continue }
    $organization = $organizationIndex[$device.organizationId]
    if (-not $organization) { continue }
    $organization.PatchFailures += $patchfailure
}

# Build per-org report data: tracked updates (exclude Defender Security Intelligence) with converted dates
$orgReportData = [System.Collections.ArrayList]::new()
foreach ($organization in $organizations) {
    $currentDeviceIds = ($devices | Where-Object { $_.organizationId -eq $organization.id }).id
    $currentPatchInstalls = $patchinstalls | Where-Object { $_.deviceId -in $currentDeviceIds }
    $trackedUpdates = $currentPatchInstalls | Where-Object { $_.name -notlike "*Security Intelligence Update for Microsoft Defender Antivirus*" }
    $tableRows = [System.Collections.ArrayList]::new()
    foreach ($install in $trackedUpdates) {
        $install.installedAt = Convert-ActivityTime $install.installedAt
        if ($install.timestamp) { $install.timestamp = Convert-ActivityTime $install.timestamp }
        $row = [PSCustomObject]@{
            OrganizationName = $organization.name
            DeviceName      = $install.DeviceName
            PatchName       = $install.name
            KBNumber        = $install.kbNumber
            InstalledAt     = $install.installedAt
            DeviceId        = $install.deviceId
        }
        [void]$tableRows.Add($row)
    }
    $stats = [PSCustomObject]@{
        PatchScanCycles    = ($organization.PatchScans).Count
        PatchApplyCycles   = ($organization.PatchApplications).Count
        PatchInstallations = $trackedUpdates.Count
        FailedPatches      = ($organization.PatchFailures).Count
    }
    [void]$orgReportData.Add([PSCustomObject]@{
        Organization = $organization
        TableRows    = @($tableRows)
        Stats        = $stats
    })
}

# ----- Simple HTML (default) -----
function Get-SimpleHtmlTable {
    param([array]$Rows, [string]$NinjaOneInstanceForLinks)
    if ($null -eq $Rows -or $Rows.Count -eq 0) {
        return "<table class='report'><tbody><tr><td>No data</td></tr></tbody></table>"
    }
    $props = @('OrganizationName','DeviceName','PatchName','KBNumber','InstalledAt')
    $html = "<table class='report'><thead><tr>"
    foreach ($p in $props) { $html += "<th>$p</th>" }
    $html += "</tr></thead><tbody>"
    foreach ($obj in $Rows) {
        $html += "<tr>"
        foreach ($p in $props) {
            $val = $obj.$p
            if ($p -eq 'DeviceName' -and $NinjaOneInstanceForLinks -and $obj.DeviceId) {
                $url = "https://$NinjaOneInstanceForLinks/#/deviceDashboard/$($obj.DeviceId)/overview"
                $html += "<td><a href='$url' target='_blank'>$val</a></td>"
            } else {
                $html += "<td>$val</td>"
            }
        }
        $html += "</tr>"
    }
    $html += "</tbody></table>"
    return $html
}

function Write-SimpleHtmlReport {
    param(
        [System.Collections.ArrayList]$OrgReportData,
        [string]$ReportTitle,
        [string]$DateRangeText,
        [string]$OutPath,
        [switch]$PerOrg,
        [string]$NinjaOneInstanceForLinks
    )
    $style = @"
<style>
body { font-family: Segoe UI, sans-serif; margin: 20px; }
h1 { color: #333; }
.report { border-collapse: collapse; width: 100%; margin-top: 12px; }
.report th, .report td { border: 1px solid #ddd; padding: 8px; text-align: left; }
.report th { background: #f0f0f0; }
.report tr:nth-child(even) { background: #f9f9f9; }
.stats { margin: 12px 0; }
</style>
"@
    if (-not $PerOrg) {
        $fullTableRows = [System.Collections.ArrayList]::new()
        foreach ($orgBlock in $OrgReportData) {
            foreach ($r in $orgBlock.TableRows) { [void]$fullTableRows.Add($r) }
        }
        $tableHtml = Get-SimpleHtmlTable -Rows $fullTableRows -NinjaOneInstanceForLinks $NinjaOneInstanceForLinks
        $body = "<h1>$ReportTitle</h1><p class='stats'>$DateRangeText</p>$tableHtml"
        $html = "<!DOCTYPE html><html><head><title>$ReportTitle</title>$style</head><body>$body</body></html>"
        $html | Set-Content -Path $OutPath -Encoding UTF8
        $totalRows = ($fullTableRows | Measure-Object).Count
        Write-Output "Exported $totalRows rows to $OutPath"
        return
    }
    $baseDir = [System.IO.Path]::GetDirectoryName($OutPath)
    if (-not $baseDir) { $baseDir = (Get-Location).Path }
    $count = 0
    foreach ($orgBlock in $OrgReportData) {
        $safeName = ($orgBlock.Organization.name -replace '[^\w\s\-]', '' -replace '\s+', '_').Trim()
        $path = Join-Path $baseDir "WindowsPatchReport_${safeName}_${yyyyMM}.html"
        $tableHtml = Get-SimpleHtmlTable -Rows $orgBlock.TableRows -NinjaOneInstanceForLinks $NinjaOneInstanceForLinks
        $statsHtml = "Patch Scan Cycles: $($orgBlock.Stats.PatchScanCycles) | Patch Apply Cycles: $($orgBlock.Stats.PatchApplyCycles) | Patch Installations: $($orgBlock.Stats.PatchInstallations) | Failed Patches: $($orgBlock.Stats.FailedPatches)"
        $body = "<h1>$($orgBlock.Organization.name) - $ReportTitle</h1><p class='stats'>$DateRangeText</p><p class='stats'>$statsHtml</p>$tableHtml"
        $fullHtml = "<!DOCTYPE html><html><head><title>$($orgBlock.Organization.name) - $ReportTitle</title>$style</head><body>$body</body></html>"
        $fullHtml | Set-Content -Path $path -Encoding UTF8
        $count += $orgBlock.TableRows.Count
        Write-Output "Exported $($orgBlock.TableRows.Count) rows to $path"
    }
    Write-Output "Total rows exported: $count"
}

# ----- PSWriteHTML report -----
function Write-PSWriteHtmlReport {
    param(
        [System.Collections.ArrayList]$OrgReportData,
        [string]$ReportTitle,
        [string]$DateRangeText,
        [string]$OutPath,
        [switch]$PerOrg,
        [string]$yyyyMM
    )
    if (-not (Get-Module -ListAvailable -Name PSWriteHTML)) {
        Install-Module -Name PSWriteHTML -AllowClobber -Force
    }
    Import-Module -Name PSWriteHTML
    $baseDir = [System.IO.Path]::GetDirectoryName($OutPath)
    if (-not $baseDir) { $baseDir = (Get-Location).Path }
    if (-not $PerOrg) {
        $allRows = [System.Collections.ArrayList]::new()
        foreach ($orgBlock in $OrgReportData) { foreach ($r in $orgBlock.TableRows) { [void]$allRows.Add($r) } }
        $allRows = @($allRows)
        New-HTML -Title $ReportTitle -FilePath $OutPath {
            New-HTMLSection -HeaderText $ReportTitle {
                New-HTMLText -Text $DateRangeText
            }
            New-HTMLSection -HeaderText 'Patch Installations' {
                New-HTMLTable -DataTable $allRows
            }
        }
        Write-Output "Exported $($allRows.Count) rows to $OutPath"
        return
    }
    $count = 0
    foreach ($orgBlock in $OrgReportData) {
        $safeName = ($orgBlock.Organization.name -replace '[^\w\s\-]', '' -replace '\s+', '_').Trim()
        $path = Join-Path $baseDir "WindowsPatchReport_${safeName}_${yyyyMM}.html"
        $rows = @($orgBlock.TableRows)
        $title = "$($orgBlock.Organization.name) - $ReportTitle"
        New-HTML -Title $title -FilePath $path {
            New-HTMLSection -HeaderText $title {
                New-HTMLText -Text $DateRangeText
            }
            New-HTMLSection -HeaderText 'Patch Statistics' {
                New-HTMLTable -DataTable $orgBlock.Stats
            }
            New-HTMLSection -HeaderText 'Patch Installations' {
                New-HTMLTable -DataTable $rows
            }
        }
        $count += $rows.Count
        Write-Output "Exported $($rows.Count) rows to $path"
    }
    Write-Output "Total rows exported: $count"
}

# ----- Main output -----
$reportTitle = "Windows Patch Report - $currentMonth $currentYear"
$dateRangeText = "Report period: $($FirstDayOfMonth.ToShortDateString()) - $($LastDayOfMonth.ToShortDateString())"
if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location) "WindowsPatchReport_${yyyyMM}.html"
}

if ($UsePSWriteHTML) {
    Write-PSWriteHtmlReport -OrgReportData $orgReportData -ReportTitle $reportTitle -DateRangeText $dateRangeText -OutPath $OutputPath -PerOrg:$PerOrganization -yyyyMM $yyyyMM
} else {
    Write-SimpleHtmlReport -OrgReportData $orgReportData -ReportTitle $reportTitle -DateRangeText $dateRangeText -OutPath $OutputPath -PerOrg:$PerOrganization -NinjaOneInstanceForLinks $NinjaOneInstance
}

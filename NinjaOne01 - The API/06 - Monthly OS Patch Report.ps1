<#

This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

#>

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"

# Get the current date
$today = Get-Date -Format "HHmmss"
# define base folder
$basefolder = "C:\Users\JeffHunter\NinjaReports\"
# define file paths
$patchinginforeport = $basefolder + "monthlypatchinginfo" + $today + "report.csv"

# API authentication details
$body = @{
    grant_type = "client_credentials"
    client_id = $NinjaOneClientId
    client_secret = $NinjaOneClientSecret
    scope = "monitoring"
}

$API_AuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$API_AuthHeaders.Add("accept", 'application/json')
$API_AuthHeaders.Add("Content-Type", 'application/x-www-form-urlencoded')

try {
    $auth_token = Invoke-RestMethod -Uri https://$NinjaOneInstance/oauth/token -Method POST -Headers $API_AuthHeaders -Body $body
    $access_token = $auth_token | Select-Object -ExpandProperty 'access_token' -EA 0
}
catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit
}
# Check if we successfully obtained an access token
if (-not $access_token) {
    Write-Host "Failed to obtain access token. Please check your client ID and client secret."
    exit
}

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("accept", 'application/json')
$headers.Add("Authorization", "Bearer $access_token")

# Initialize the date (ensure $date is defined)
$date = Get-Date

# Calculate the first day of the previous month using explicit checks
if ($date.Month -eq 1) {
    $prevMonth = 12
    $year = $date.Year - 1
} else {
    $prevMonth = $date.Month - 1
    $year = $date.Year
}

# Create a DateTime object for the first day of the previous month
$firstDayOfPreviousMonth = Get-Date -Year $year -Month $prevMonth -Day 1

# Calculate the last day of the previous month
$lastDayOfPreviousMonth = $firstDayOfPreviousMonth.AddMonths(1).AddDays(-1)

# Format dates as strings in the format 'yyyyMMdd'
$firstDayString = $firstDayOfPreviousMonth.ToString('yyyyMMdd')
$lastDayString = $lastDayOfPreviousMonth.ToString('yyyyMMdd')

# Output the results
Write-Host "First day of previous month: $firstDayString"
Write-Host "Last day of previous month: $lastDayString"

# define ninja urls
$devices_url = "https://$NinjaOneInstance/v2/devices?df=class%20in%20(WINDOWS_WORKSTATION,%20WINDOWS_SERVER)"
$organizations_url = "https://$NinjaOneInstance/v2/organizations"
$activities_url = "https://$NinjaOneInstance/api/v2/activities?df=class%20in%20(WINDOWS_WORKSTATION,%20WINDOWS_SERVER)&class=DEVICE&type=PATCH_MANAGEMENT&status=in%20(PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED,%20PATCH_MANAGEMENT_SCAN_COMPLETED,%20PATCH_MANAGEMENT_FAILURE)&after=" + $firstDayString + "&before=" + $lastDayString + "&pageSize=1000"
$patchreport_url = "https://$NinjaOneInstance/api/v2/queries/os-patch-installs?df=class%20in%20(WINDOWS_WORKSTATION,%20WINDOWS_SERVER)&status=Installed&installedAfter=" + $firstDayString + "&installedBefore=" + $lastDayString
$failedpatch_url = "https://$NinjaOneInstance/api/v2/queries/os-patch-installs?df=class%20in%20(WINDOWS_WORKSTATION,%20WINDOWS_SERVER)&status=Failed"

# call ninja urls
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
    $request = Invoke-RestMethod -Uri $activities_url -Method GET -Headers $headers -Verbose
    $patchinstalls = Invoke-RestMethod -Uri $patchreport_url -Method GET -Headers $headers | Select-Object -ExpandProperty 'results'
    $patchfailures = Invoke-RestMethod -Uri $failedpatch_url -Method GET -Headers $headers | Select-Object -ExpandProperty 'results'
    $organizations = Invoke-RestMethod -Uri $organizations_url -Method GET -Headers $headers
}
catch {
    Write-Error "Failed to retrieve required data from NinjaOne API. Error: $_"
    exit}

$userActivities = $request.activities
$activitiesRemaining = $true
$olderThan = $userActivities[-1].id

# Loop while there are still activities available from the API response.
while($activitiesRemaining -eq $true) {
    $activities_url = "https://$NinjaOneInstance/api/v2/activities?df=class%20in%20(WINDOWS_WORKSTATION,%20WINDOWS_SERVER)&class=DEVICE&type=PATCH_MANAGEMENT&status=in%20(PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED,%20PATCH_MANAGEMENT_SCAN_COMPLETED,%20PATCH_MANAGEMENT_FAILURE)&after=" + $firstDayString + "&before=" + $lastDayString + "&olderThan=" + $olderThan + "&pageSize=1000"
    $response = Invoke-RestMethod -Uri $activities_url -Method GET -Headers $headers

    if ($response.activities.count -eq 0) {
        $activitiesRemaining = $false
    } else {
        $userActivities += $response.activities
        $olderThan = $response.activities[-1].id
    }
}


# Filter user activities
$patchScans = @()
$patchScanFailures = @()
$patchApplicationCycles = @()
$patchApplicationFailures = @()

foreach ($activity in $userActivities) {
    if ($activity.activityResult -match "SUCCESS") {
        if ($activity.statusCode -match "PATCH_MANAGEMENT_SCAN_COMPLETED") {
            $patchScans += $activity
        } elseif ($activity.statusCode -match "PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED") {
            $patchApplicationCycles += $activity
        }
    } elseif ($activity.activityResult -match "FAILURE") {
        if ($activity.statusCode -match "PATCH_MANAGEMENT_SCAN_COMPLETED") {
            $patchScanFailures += $activity
        } elseif ($activity.statusCode -match "PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED") {
            $patchApplicationFailures += $activity
        }
    }
}

# Index devices by ID for faster lookup
$deviceIndex = @{}
foreach ($device in $devices) {
    $deviceIndex[$device.id] = $device
}

# Initialize organization objects with PatchFailures property
foreach ($organization in $organizations) {
    Add-Member -InputObject $organization -NotePropertyName "PatchScans" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "PatchFailures" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "PatchInstalls" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "Workstations" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "Servers" -NotePropertyValue @() -Force

}

# Assign devices to organizations
foreach ($device in $devices) {
    $currentOrg = $organizations | Where-Object { $_.id -eq $device.organizationId }
    if ($device.nodeClass.EndsWith("_SERVER")) {
        $currentOrg.Servers += $device.systemName
    } elseif ($device.nodeClass.EndsWith("_WORKSTATION") -or $device.nodeClass -eq "MAC") {
        $currentOrg.Workstations += $device.systemName
    }
}

# Process patch scans
foreach ($patchScan in $patchScans) {
    $device = $deviceIndex[$patchScan.deviceId]
    $patchScan | Add-Member -NotePropertyName "DeviceName" -NotePropertyValue $device.systemName -Force
    $patchScan | Add-Member -NotePropertyName "OrgID" -NotePropertyValue $device.organizationId -Force
    $organization = $organizations | Where-Object { $_.id -eq $device.organizationId }
    $organization.PatchScans += $patchScan
}

# Process patch installations
foreach ($patchinstall in $patchinstalls) {
    $device = $deviceIndex[$patchinstall.deviceId]
    $patchinstall | Add-Member -NotePropertyName "DeviceName" -NotePropertyValue $device.systemName -Force
    $patchinstall | Add-Member -NotePropertyName "OrgID" -NotePropertyValue $device.organizationId -Force
    $organization = $organizations | Where-Object { $_.id -eq $device.organizationId }
    $organization.PatchInstalls += $patchinstall
}

# Process patch installation failures
foreach ($patchfailure in $patchfailures) {
    $device = $deviceIndex[$patchfailure.deviceId]
    $patchfailure | Add-Member -NotePropertyName "DeviceName" -NotePropertyValue $device.systemName -Force
    $patchfailure | Add-Member -NotePropertyName "OrgID" -NotePropertyValue $device.organizationId -Force
    $organization = $organizations | Where-Object { $_.id -eq $device.organizationId }
    $organization.PatchFailures += $patchfailure
}


# Generate results
$results = foreach ($organization in $organizations) {
    [PSCustomObject]@{
        Name = $organization.Name
        Workstations = ($organization.Workstations).Count
        Servers = ($organization.Servers).Count
        Total = ($organization.Workstations).Count + ($organization.Servers).Count
        PatchScans = ($organization.PatchScans).Count
        PatchInstalls = ($organization.PatchInstalls).Count
        PatchFailures = ($organization.PatchFailures).Count
    }
}

# Export results
Write-Output $results | Format-Table
$results | Export-CSV -NoTypeInformation -Path $patchinginforeport

Write-Host "CSV file has been created successfully!"

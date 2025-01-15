<#

PowerShell script to generate a report of Windows device patch installations for the previous month in a specific organization using NinjaOne API
This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

#>

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"

# API authentication details
$body = @{
    grant_type = "client_credentials"
    client_id = $NinjaOneClientId
    client_secret = $NinjaOneClientSecret
    scope = "monitoring"
}

# Headers for the authentication request
$API_AuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$API_AuthHeaders.Add("accept", 'application/json')
$API_AuthHeaders.Add("Content-Type", 'application/x-www-form-urlencoded')

# Authenticate and retrieve the access token
try {
    $auth_token = Invoke-RestMethod -Uri https://$NinjaOneInstance/oauth/token -Method POST -Headers $API_AuthHeaders -Body $body
    $access_token = $auth_token.access_token
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

# Headers for subsequent API requests, including the obtained access token
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
$FirstDayString = $firstDayOfPreviousMonth.ToString('yyyyMMdd')
$LastDayString = $lastDayOfPreviousMonth.ToString('yyyyMMdd')

# Output the results
Write-Host "First day of previous month: $firstDayString"
Write-Host "Last day of previous month: $lastDayString"

# Defining the file path for the output CSV report
$today = Get-Date -format "yyyyMMdd"
$patchinfo_report = "C:\Users\JeffHunter\NinjaReports\${today}_Patch_Report.csv"

# Defining API endpoints for device and patch information
$devices_url = "https://$NinjaOneInstance/v2/devices"
$patchreport_url = "https://$NinjaOneInstance/api/v2/queries/os-patch-installs?df=class%20in%20(WINDOWS_WORKSTATION,%20WINDOWS_SERVER)&status=Installed&installedBefore=$LastDayString&installedAfter=$FirstDayString"

# Fetching devices and patch installation details
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
    $patchinstalls = Invoke-RestMethod -Uri $patchreport_url -Method GET -Headers $headers | Select-Object -ExpandProperty 'results'
}
catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit
}

# Processing each patch installation to enrich with device and organization information
foreach ($patchinstall in $patchinstalls) {
    $currentDevice = $devices | Where-Object {$_.id -eq $patchinstall.deviceId} | Select-Object -First 1
    # Adding device name to each patch installation record
    Add-Member -InputObject $patchinstall -NotePropertyName "DeviceName" -NotePropertyValue $currentDevice.systemName
    # Converting timestamps from Unix time to DateTime format
    $patchinstall.installedAt = ([DateTimeOffset]::FromUnixTimeSeconds($patchinstall.installedAt).DateTime).ToString()
    $patchinstall.timestamp = ([DateTimeOffset]::FromUnixTimeSeconds($patchinstall.timestamp).DateTime).ToString()
}

# Displaying the patch installations in a formatted table
$patchinstalls | Select-Object name, status, installedAt, kbNumber, DeviceName | Format-Table

# Exporting the patch installation details to a CSV file
$patchinstalls | Select-Object name, status, installedAt, kbNumber, DeviceName | Export-CSV -NoTypeInformation -Path $patchinfo_report

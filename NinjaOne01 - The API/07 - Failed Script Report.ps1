<#

This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

#>

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"

# Script parameters
$days = 3 # Number of days in the past to get data for - increasing this number will correspondingly increase the script runtime
$scriptName = "Process Log"

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

# Obtain the authentication token
try {
    $authResponse = Invoke-RestMethod -Uri https://$NinjaOneInstance/oauth/token -Method POST -Headers $API_AuthHeaders -Body $body
    $access_token = $authResponse.access_token
} catch {
    Write-Error "Failed to authenticate. Error: $_"
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

# Date calculations
$date = (Get-Date).AddDays(-$days).ToString("yyyyMMdd")
$today = Get-Date -Format "yyyyMMdd"

# File path for the report
$failedScriptsReport = "C:\Users\JeffHunter\NinjaReports\${today}_Script_Report.csv"

# Define API endpoints
$devices_url = "https://$NinjaOneInstance/v2/devices"
$activities_url = "https://$NinjaOneInstance/api/v2/activities?class=DEVICE&type=ACTION&status=COMPLETED&after=${date}&pageSize=1000"

# Fetch devices and initial activities
# The /activities/ endpoint is limited to 1000 entries at once, so pagination must be utilized for larger data sets
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
    $activitiesResponse = Invoke-RestMethod -Uri $activities_url -Method GET -Headers $headers
} catch {
    Write-Error "Failed to retrieve data. Error: $_"
    exit
}

$userActivities = $activitiesResponse.activities
$activitiesRemaining = $true
$olderThan = $userActivities[-1].id

# Paginate through remaining activities if applicable
while($activitiesRemaining) {
    $activities_url = "https://$NinjaOneInstance/api/v2/activities?type=ACTION&status=COMPLETED&after=${date}&olderThan=${olderThan}&pageSize=1000"
    $response = Invoke-RestMethod -Uri $activities_url -Method GET -Headers $headers

    if ($response.activities.count -eq 0) {
        $activitiesRemaining = $false
    } else {
        $userActivities += $response.activities
        $olderThan = $response.activities[-1].id
    }
}

# Convert Unix timestamp to readable date time format - time will be in UTC
foreach ($activity in $userActivities) {
    $activity.activityTime = ([System.DateTimeOffset]::FromUnixTimeSeconds($activity.activityTime)).DateTime.ToString()
}

# Filter through all script completed activities looking explicitly for failed script activities that match the script name
$failedScripts = $userActivities | Where-Object { $_.activityResult -match "FAILURE" -and $_.sourceName.substring(4) -like "*$scriptName*"} | Select-Object deviceId,activityResult,activityTime,activityType,subject,message

# Map device names to failed script activities
foreach ($failedScript in $failedScripts) {
    $failedScript | Add-Member -NotePropertyName "DeviceName" -NotePropertyValue ""
    $device = $devices | Where-Object {$_.id -eq $failedScript.deviceId}
    $failedScript.DeviceName = $device.systemName
}

# Prepare final data for reporting
$failedScripts = $failedScripts | Select-Object deviceName,activityResult,activityTime,activityType,message

if ($failedScripts.Count -eq 0) {
    Write-Host "No failed script executions have been found for the script and time period specified."
} else {
    Write-Host ($failedScripts | Format-Table | Out-String)
    # Uncomment the line below to enable CSV export
    $failedScripts | Export-Csv -NoTypeInformation -Path $failedScriptsReport
    Write-Host "CSV file containing failed scripts has been created successfully!"
    Write-Host "Go to $failedScriptsReport to find your Failed Scripts Report"
}

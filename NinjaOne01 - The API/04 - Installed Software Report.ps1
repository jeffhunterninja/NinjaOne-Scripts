<#

This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

Script Notes:
# This utilizes a third-party module to generate an HTML report called PSWriteHTML
# Script can be modified to produce a CSV if preferred

#>

# Install and import the required module
# Check if PSWriteHTML module is installed
$module = Get-Module -ListAvailable -Name PSWriteHTML
if (-not $module) {
    # If the module is not installed, install it
    Install-Module -Name PSWriteHTML -AllowClobber -Force
}
# Import the PSWriteHTML module
Import-Module -Name PSWriteHTML

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"

# Body for authentication
$body = @{
    grant_type = "client_credentials"
    client_id = $NinjaOneClientId
    client_secret = $NinjaOneClientSecret
    scope = "monitoring"
}

# Headers for authentication
$API_AuthHeaders = @{
    'accept' = 'application/json'
    'Content-Type' = 'application/x-www-form-urlencoded'
}

# Authenticate and get access token
try {
    $auth_token = Invoke-RestMethod -Uri "https://$NinjaOneInstance/oauth/token" -Method POST -Headers $API_AuthHeaders -Body $body
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
# Headers for subsequent API calls
$headers = @{
    'accept' = 'application/json'
    'Authorization' = "Bearer $access_token"
}

# Define Ninja URLs
$devices_url = "https://$NinjaOneInstance/v2/devices-detailed"
$organizations_url = "https://$NinjaOneInstance/v2/organizations-detailed"
$locations_url = "https://$NinjaOneInstance/v2/locations"
$softwares_url = "https://$NinjaOneInstance/v2/queries/software"

# Call Ninja URLs to get data
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
    $organizations = Invoke-RestMethod -Uri $organizations_url -Method GET -Headers $headers
    $locations = Invoke-RestMethod -Uri $locations_url -Method GET -Headers $headers
    $softwares = Invoke-RestMethod -Uri $softwares_url -Method GET -Headers $headers
}
catch {
    Write-Error "Failed to retrieve data from NinjaOne API. Error: $_"
    exit
}

# Define application names for filtering
$appNames = @("Chrome", "Firefox", "Edge")

# Filter software results
$filteredObjCreated = $softwares.results | Where-Object { $_.name -ne $null } | Select-Object name, version, deviceId, publisher
$filteredObj = $appNames | ForEach-Object {
    $appName = $_
    $filteredObjCreated | Where-Object { $_.name -like "*$appName*" }
} | Sort-Object deviceId -Unique

# Add device name, organization name, and location name to make a complete report
foreach ($device in $devices){
    $currentDev = $filteredObj | Where-Object {$_.deviceId -eq $device.id}
    $currentDev | Add-Member -MemberType NoteProperty -Name 'DeviceName' -Value $device.systemname -Force
    $currentDev | Add-Member -MemberType NoteProperty -Name 'OrgID' -Value $device.organizationId -Force
    $currentDev | Add-Member -MemberType NoteProperty -Name 'LocID' -Value $device.locationId -Force       
}
foreach ($organization in $organizations){
    $currentOrg = $filteredObj | Where-Object {$_.OrgID -eq $organization.id}
    $currentOrg | Add-Member -MemberType NoteProperty -Name 'OrgName' -Value $organization.name -Force  
}
foreach ($location in $locations){
    $currentLoc = $filteredObj | Where-Object {$_.LocID -eq $location.id}
    $currentLoc | Add-Member -MemberType NoteProperty -Name 'LocName' -Value $location.name -Force 
}

# Output the HTML view
$filteredObj | Out-HtmlView

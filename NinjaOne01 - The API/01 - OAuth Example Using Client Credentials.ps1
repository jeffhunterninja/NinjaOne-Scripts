<#

This is provided as an educational example of how to interact with the NinjaAPI using the client credentials grant type.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

#>

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"

# Body for authentication
$body = @{
    grant_type = "client_credentials"
    client_id = $NinjaOneClientId
    client_secret = $NinjaOneClientSecret
    scope = "monitoring management"
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

# Call Ninja URLs to get data
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
    $organizations = Invoke-RestMethod -Uri $organizations_url -Method GET -Headers $headers
}
catch {
    Write-Error "Failed to retrieve organizations and devices from NinjaOne API. Error: $_"
    exit
}

# Extend organizations objects with additional properties to classify devices
Foreach ($organization in $organizations) {
    Add-Member -InputObject $organization -NotePropertyName "Workstations" -NotePropertyValue @()
    Add-Member -InputObject $organization -NotePropertyName "Servers" -NotePropertyValue @()
}

# Loop through all devices and copy each device to corresponding organization, with separate properties for storing servers and workstations
Foreach ($device in $devices) {
    $currentOrg = $organizations | Where-Object {$_.id -eq $device.organizationId}
    if ($device.nodeClass.EndsWith("_SERVER")) {
        $currentOrg.servers += $device.systemName
    } elseif ($device.nodeClass.EndsWith("_WORKSTATION") -or $device.nodeClass -eq "MAC") {
        $currentOrg.workstations += $device.systemName
    }
}

# Create and display a summary report of organizations and their device counts broken down by servers and workstations, plus total devices
$reportSummary = Foreach ($organization in $organizations) {
    [PSCustomObject]@{
        Name = $organization.Name
        Workstations = $organization.workstations.length
        Servers = $organization.servers.length
        TotalDevices = ($organization.workstations.length + $organization.servers.length)
    }
}

# Display the summary report in a table format
$reportSummary | Format-Table | Out-String

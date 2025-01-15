<#

This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

Script Notes:
# Further evaluation and testing with VMware and Hyper-V hosts is recommended to ensure accurate tabulation
# More information here: https://ninjarmm.zendesk.com/hc/en-us/community/posts/4424760908813/comments/4445857839757

Attributions:
# Juan Miguel, Steve Mohring, Alexander Wissfeld

#>

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"

# Initialize the body parameters for OAuth 2.0 authentication
$body = @{
    grant_type = "client_credentials" # Defines the type of grant being requested
    client_id = $NinjaOneClientId # Your NinjaRMM client application ID
    client_secret = $NinjaOneClientSecret # Your NinjaRMM client application secret
    scope = "monitoring" # The scope of access requested
}

# Create a dictionary to hold headers for the authentication request
$API_AuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$API_AuthHeaders.Add("accept", 'application/json')
$API_AuthHeaders.Add("Content-Type", 'application/x-www-form-urlencoded')

# Authenticate with NinjaRMM and retrieve the access token
$auth_token = Invoke-RestMethod -Uri https://$NinjaOneInstance/oauth/token -Method POST -Headers $API_AuthHeaders -Body $body
$access_token = $auth_token | Select-Object -ExpandProperty 'access_token' -EA 0

# Prepare headers for subsequent API requests using the obtained access token
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("accept", 'application/json')
$headers.Add("Authorization", "Bearer $access_token")

# Get the current date in yyyyMMdd format
$today = Get-Date -format "yyyyMMdd"

# Define file path for the licenses report
$licenses_report = "C:\Users\JeffHunter\NinjaReports\" + $today + "_Ninja_Licenses_Report.csv"

# Define API endpoints for devices and organizations
$devices_url = "https://$NinjaOneInstance/v2/devices"
$organizations_url = "https://$NinjaOneInstance/v2/organizations"
$remotes_url = "https://$NinjaOneInstance/v2/group/16/device-ids"
$bitdefenders_url = "https://$NinjaOneInstance/v2/group/8/device-ids"
$webroots_url = "https://$NinjaOneInstance/v2/group/7/device-ids"

# Retrieve data from NinjaRMM API
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
    $organizations = Invoke-RestMethod -Uri $organizations_url -Method GET -Headers $headers
    $remotes = Invoke-RestMethod -Uri $remotes_url -Method GET -Headers $headers
    $bitdefenders = Invoke-RestMethod -Uri $bitdefenders_url -Method GET -Headers $headers
    $webroots = Invoke-RestMethod -Uri $webroots_url -Method GET -Headers $headers
}
catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit
}


# Extend organizations objects with additional properties to classify devices
Foreach ($organization in $organizations) {
    Add-Member -InputObject $organization -NotePropertyName "Workstations" -NotePropertyValue @()
    Add-Member -InputObject $organization -NotePropertyName "Servers" -NotePropertyValue @()
    Add-Member -InputObject $organization -NotePropertyName "Networks" -NotePropertyValue @()
    Add-Member -InputObject $organization -NotePropertyName "Remotes" -NotePropertyValue @()
    Add-Member -InputObject $organization -NotePropertyName "Bitdefenders" -NotePropertyValue @()
    Add-Member -InputObject $organization -NotePropertyName "Webroots" -NotePropertyValue @()
}

# Enumerate devices and assign them to their respective organization and category
Write-Host 'Enumerating everything ...'
Foreach ($device in $devices) {
    $currentOrg = $organizations | Where-Object {$_.id -eq $device.organizationId}
    if ($device.nodeClass.EndsWith("_SERVER")) {
        $currentOrg.servers += $device.systemName
    } elseif ($device.nodeClass.EndsWith("_WORKSTATION") -or $device.nodeClass -eq "MAC") {
        $currentOrg.workstations += $device.systemName
    } elseif ($device.nodeClass.StartsWith("NMS")) {
        $currentOrg.networks += $device.id
    }
    if ($remotes.Contains($device.id)) { $currentOrg.remotes += $device.systemName }
    if ($bitdefenders.Contains($device.id)) { $currentOrg.bitdefenders += $device.systemName }
    if ($webroots.Contains($device.id)) { $currentOrg.webroots += $device.systemName }
}
Write-Host 'Done ✅'

# Create and display a summary report of organizations and their device counts
$reportSummary = Foreach ($organization in $organizations) {
    [PSCustomObject]@{
        Name = $organization.Name
        Workstations = $organization.workstations.length
        Servers = $organization.servers.length
        TotalDevices = ($organization.workstations.length + $organization.servers.length)
        NetworkDevices = $organization.networks.length
        RemoteAccessEnabled = $organization.remotes.length
        BitdefenderEnabled = $organization.bitdefenders.length
        WebrootEnabled = $organization.webroots.length
    }
}

# Display the summary report in a table format
$reportSummary | Format-Table | Out-String

# Export the report to a CSV file
$reportSummary | Export-CSV -NoTypeInformation -Path $licenses_report

# Confirm completion and report location
Write-Host "CSV files have been created with success!"
Write-Host "Go to $licenses_report to find your Licenses Report"

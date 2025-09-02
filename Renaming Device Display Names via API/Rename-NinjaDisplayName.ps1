<#
This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancements may be necessary to handle larger datasets.
#>

# Your NinjaRMM credentials
$NinjaOneInstance = 'ca.ninjarmm.com' # Adjust if necessary based on your region
$NinjaOneClientId = ''
$NinjaOneClientSecret = ''

# Prepare the body for authentication
$body = @{
    grant_type = "client_credentials"
    client_id = $NinjaOneClientId
    client_secret = $NinjaOneClientSecret
    scope = "monitoring management"
}

# Prepare headers for authentication request
$API_AuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$API_AuthHeaders.Add("accept", 'application/json')
$API_AuthHeaders.Add("Content-Type", 'application/x-www-form-urlencoded')

# Obtain the authentication token
try {
    $auth_token = Invoke-RestMethod -Uri https://$NinjaOneInstance/oauth/token -Method POST -Headers $API_AuthHeaders -Body $body
    $access_token = $auth_token | Select-Object -ExpandProperty 'access_token' -EA 0
} catch {
    Write-Error "Failed to obtain authentication token. $_"
    exit 1
}

# Prepare headers for subsequent API requests
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("accept", 'application/json')
$headers.Add("Authorization", "Bearer $access_token")

# Import device data from a CSV file
$deviceimports = Import-CSV -Path "C:\Users\JeffHunter\Documents\NinjaReports\csvexample2.csv"

# Process each device import entry
$assets = foreach ($deviceimport in $deviceimports) {
    [PSCustomObject]@{
        Name = $deviceimport.systemName
        DisplayName = $deviceimport.displayName
        ID = $null
    }
}

# Fetch the detailed list of devices from NinjaOne
$devices_url = "https://$NinjaOneInstance/v2/devices"
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
} catch {
    Write-Error "Failed to fetch devices. $_"
    exit 1
}


# Match devices and add their IDs to the assets
foreach ($device in $devices) {
    $currentDev = $assets | Where-Object { $_.Name -eq $device.systemName }
    if ($null -ne $currentDev) {
        $currentDev.ID = $device.id
    }
}

# Update the display names for each asset
foreach ($asset in $assets) {
    if ($null -ne $asset.ID) {
        # Define NinjaOne API endpoint for updating display name
        $displayname_url = "https://$NinjaOneInstance/api/v2/device/" + $asset.ID

        # Extract display name and prepare the request body
        $displayname = $asset.DisplayName
        $request_body = @{
            displayName = "$displayname"
        }

        # Convert the request body to JSON
        $json = $request_body | ConvertTo-Json

        Write-Host "Changing display name for:" $asset.Name "to" $asset.DisplayName

        # Update the display name via the API
        try {
            Invoke-RestMethod -Method 'Patch' -Uri $displayname_url -Headers $headers -Body $json -ContentType "application/json" -Verbose
        } catch {
            Write-Error "Failed to update display name for $($asset.Name). $_"
        }
    } else {
        Write-Warning "Skipping update for $($asset.Name) as ID is null."
    }
}

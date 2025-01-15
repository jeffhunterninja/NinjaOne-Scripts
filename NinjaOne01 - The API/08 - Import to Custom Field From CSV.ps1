<#

This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

#>

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"

# Define authentication details
$body = @{
    grant_type = "client_credentials"
    client_id = $NinjaOneClientId # Replace with your actual client ID
    client_secret = $NinjaOneClientSecret # Replace with your actual client secret
    scope = "monitoring management"
}

# Set headers for the authentication request
$API_AuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$API_AuthHeaders.Add("accept", 'application/json')
$API_AuthHeaders.Add("Content-Type", 'application/x-www-form-urlencoded')

# Obtain an access token from the NinjaRMM OAuth token endpoint
try {
    $auth_token = Invoke-RestMethod -Uri https://$NinjaOneInstance/oauth/token -Method POST -Headers $API_AuthHeaders -Body $body
    $access_token = $auth_token.access_token
}
catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit}

# Check if we successfully obtained an access token
if (-not $access_token) {
    Write-Host "Failed to obtain access token. Please check your client ID and client secret."
    exit
}

# Set headers for subsequent API requests using the obtained access token
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("accept", 'application/json')
$headers.Add("Authorization", "Bearer $access_token")

# Import device data from a CSV file
$deviceimports = Import-CSV -Path "C:\Users\JeffHunter\OneDrive - NinjaOne\Scripting\NinjaOne01 - The API\NinjaOne01 - The API\Resources\Import to Device CF.csv"

# Prepare the devices list from the imported CSV data
$assets = Foreach ($deviceimport in $deviceimports) {
    [PSCustomObject]@{
        Name = $deviceimport.name
        AssetOwner = $deviceimport.assetOwner
        ID = $deviceimport.Id
    }
}

# Update custom fields for each device
foreach ($asset in $assets) {
    # Construct the URL for the device's custom fields endpoint
    $customfields_url = "https://$NinjaOneInstance/api/v2/device/" + $asset.ID + "/custom-fields"

    # Prepare the request body with the custom field data
    $request_body = @{
        assetOwner = $asset.AssetOwner
    }

    # Convert the request body to JSON format
    $json = $request_body | ConvertTo-Json

    # Display the current operation
    Write-Host "Patching custom fields for: " $asset.Name

    # Send a PATCH request to update the custom fields for the device
    Invoke-RestMethod -Method 'Patch' -Uri $customfields_url -Headers $headers -Body $json -ContentType "application/json" -Verbose
}

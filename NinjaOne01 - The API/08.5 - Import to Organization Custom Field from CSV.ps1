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
    client_id = $NinjaOneClientId  # Replace with your actual client ID
    client_secret = $NinjaOneClientSecret  # Replace with your actual client secret
    scope = "monitoring management"
}

# Load the CSV file containing the organization data
$deviceimports = Import-CSV -Path "C:\Users\JeffHunter\OneDrive - NinjaRMM\Scripting\Final Versions\CSVs\Import to Org CF.csv"

# Set up headers for the authentication request
$API_AuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$API_AuthHeaders.Add("accept", 'application/json')
$API_AuthHeaders.Add("Content-Type", 'application/x-www-form-urlencoded')

# Authenticate and retrieve access token
$auth_uri = "https://$NinjaOneInstance/oauth/token"
$auth_token = Invoke-RestMethod -Uri $auth_uri -Method POST -Headers $API_AuthHeaders -Body $body
$access_token = $auth_token.access_token

# Prepare headers for subsequent API requests
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("accept", 'application/json')
$headers.Add("Authorization", "Bearer $access_token")

# Fetch organizations
$organizations_url = "https://$NinjaOneInstance/v2/organizations"
try {
    $organizations = Invoke-RestMethod -Uri $organizations_url -Method GET -Headers $headers
}
catch {
    Write-Error "Failed to retrieve organizations from NinjaOne API. Error: $_"
    exit
}

# Process each entry in the CSV
$assets = Foreach ($deviceimport in $deviceimports) {
    [PSCustomObject]@{
        ID = 0
        DisplayName = $deviceimport.'Organization Name'
        OrgCustomField = $deviceimport.'Custom Field'
        OrgVariable = $deviceimport.'Organization Variable'
    }
}

# Update the ID for each asset based on matching organization name
foreach ($organization in $organizations) {
    foreach ($asset in $assets) {
        if ($asset.DisplayName -like $organization.name) {
            $asset.ID = $organization.id
        }
    }
}

# Patch custom fields for each organization
foreach ($asset in $assets) {
    $customfields_url = "https://$NinjaOneInstance/api/v2/organization/$($asset.ID)/custom-fields"
  
    # Dynamically construct request body using the custom field name and value
    $request_body = @{
        $asset.OrgCustomField = $asset.OrgVariable
    }

    $json = $request_body | ConvertTo-Json

    Write-Host "Patching custom fields for: $($asset.DisplayName) with an organization ID of: $($asset.ID)"
    Write-Host "Writing into URL: $customfields_url"

    # Perform the PATCH request
    try {
        Invoke-RestMethod -Method 'Patch' -Uri $customfields_url -Headers $headers -Body $json -ContentType "application/json" -Verbose
    }
    catch {
        Write-Error "Failed to connect to NinjaOne API. Error: $_"
        exit
    }
}

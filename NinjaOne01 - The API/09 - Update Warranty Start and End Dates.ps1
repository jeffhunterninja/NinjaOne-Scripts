<#

This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

#>

$NinjaOneInstance     = "app.ninjarmm.com"
$NinjaOneClientId     = "-"
$NinjaOneClientSecret = "-"

# Import device data from a CSV file
$warrantyimports = Import-CSV -Path "C:\Users\JeffHunter\OneDrive - NinjaOne\Scripting\NinjaOne01 - The API\NinjaOne01 - The API\Resources\warranty_data.csv"

function Convert-ToUnixTime {
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [DateTime]$DateTime
    )
    try {
        # Ensure $DateTime is a valid DateTime object
        if (-not $DateTime -or -not ($DateTime -is [DateTime])) {
            Write-Output "Invalid DateTime input: '$DateTime'"
            return 0
        }

        # Convert to UTC and calculate Unix timestamp
        $unixTime = [Math]::Floor((($DateTime.ToUniversalTime()) - [datetime]'1970-01-01T00:00:00Z').TotalSeconds)
        return $unixTime
    }
    catch {
        Write-Error "Failed to convert to Unix time: $_"
    }
}

# Example usage:
$exampleDate = "2025-12-12"
try {
    $unixTimestamp = Convert-ToUnixTime -DateTime ([DateTime]::Parse($exampleDate))
    Write-Output "Unix Timestamp: $unixTimestamp"
}
catch {
    Write-Error "Error processing input: $_"
}


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
    $auth_token = Invoke-RestMethod -Uri https://$NinjaOneInstance/ws/oauth/token -Method POST -Headers $API_AuthHeaders -Body $body
    $access_token = $auth_token | Select-Object -ExpandProperty 'access_token' -EA 0
} catch {
    Write-Error "Failed to obtain authentication token. $_"
    exit 1
}

# Prepare headers for subsequent API requests
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("accept", 'application/json')
$headers.Add("Authorization", "Bearer $access_token")

# Fetch the detailed list of devices from NinjaOne
$devices_url = "https://$NinjaOneInstance/ws/api/v2/devices-detailed"
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
} catch {
    Write-Error "Failed to fetch devices. $_"
    exit 1
}

# Process each device import entry and add systemName if matched
$assets = foreach ($warrantyimport in $warrantyimports) {
    # Find the matching device by ID
    $device = $devices | Where-Object { $_.id -eq $warrantyimport.Id }

    # Create the asset object with the systemName if the device is found
    if ($device) {
        [PSCustomObject]@{
            Name = $warrantyimport.name
            StartDate = $warrantyimport.WarrantyStart
            EndDate = $warrantyimport.WarrantyEnd
            FullfillDate = $warrantyimport.MftrFullfill
            ID = $warrantyimport.Id
        }
    }

}

# Update the display names for each asset
foreach ($asset in $assets) {
    if ($null -ne $asset.ID) {
        # Define NinjaOne API endpoint for updating warranty information
        $warranty_url = "https://$NinjaOneInstance/api/v2/device/" + $asset.ID
        
        $WarrantyFields = @{
            'startDate' = Convert-ToUnixTime -DateTime $asset.StartDate
            'endDate'   = Convert-ToUnixTime -DateTime $asset.EndDate
            'manufacturerFulfillmentDate' = Convert-ToUnixTime -DateTime $asset.FullfillDate
            }   

        $request_body = @{
            warranty = $WarrantyFields
        }

        # Convert the request body to JSON
        $json = $request_body | ConvertTo-Json

        Write-Host "Uploading warranty data for:" $asset.ID

        # Update the warranty info via the API
        try {
            Invoke-RestMethod -Method 'Patch' -Uri $warranty_url -Headers $headers -Body $json -ContentType "application/json" -Verbose
        } catch {
            Write-Error "Failed to update set warranty info for $($asset.ID). $_"
        }
    } else {
        Write-Warning "Skipping warranty update for $($asset.ID) as ID is null."
    }
}

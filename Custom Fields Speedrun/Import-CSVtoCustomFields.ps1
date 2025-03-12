<#
This script demonstrates how to interact with the NinjaOne API, import device data from a CSV,
match it with devices from NinjaOne, and update custom fields dynamically based on CSV headers.
It includes enhanced error handling, logging, and dynamic request body generation.

Parameters:
- $OverwriteEmptyValues: Determines if empty CSV values are included as $null (overwriting existing data)
  or excluded from the update payload. (Default: $false)

Before running this script:
- Ensure the CSV file (csvexample.csv) has at least the following columns: Id, name.
- Additional CSV columns (e.g. assetOwner, location, etc.) will be used as custom fields.
- Replace $NinjaOneClientId and $NinjaOneClientSecret with your credentials.
#>

param(
    [bool]$OverwriteEmptyValues = $false
)

# Your NinjaRMM credentials
$NinjaOneInstance = ''  # Varies by region/environment (e.g. 'app.ninjarmm.com' for US)
$NinjaOneClientId = ''                  # Enter your client id here
$NinjaOneClientSecret = ''              # Enter your client secret here

# Import device data from a CSV file
$csvPath = "C:\Users\JeffHunter\OneDrive - NinjaOne\Custom Fields Speedrun\datatoimport.csv"

try {
    $deviceimports = Import-Csv -Path $csvPath
} catch {
    Write-Error "Failed to import CSV file from $csvPath. $_"
    exit 1
}

# Validate CSV has required columns (Id and name)
$requiredColumns = @("Id", "name")
foreach ($col in $requiredColumns) {
    if (-not ($deviceimports[0].PSObject.Properties.Name -contains $col)) {
        Write-Error "CSV file is missing required column '$col'. Please verify the CSV structure."
        exit 1
    }
}

Write-Host "CSV Import successful. Processing $($deviceimports.Count) entries..."

# Prepare the body for authentication
$body = @{
    grant_type    = "client_credentials"
    client_id     = $NinjaOneClientId
    client_secret = $NinjaOneClientSecret
    scope         = "monitoring management"
}

# Prepare headers for authentication request
$API_AuthHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$API_AuthHeaders.Add("accept", "application/json")
$API_AuthHeaders.Add("Content-Type", "application/x-www-form-urlencoded")

# Obtain the authentication token
try {
    $auth_token = Invoke-RestMethod -Uri "https://$NinjaOneInstance/oauth/token" -Method POST -Headers $API_AuthHeaders -Body $body
    $access_token = $auth_token.access_token
} catch {
    Write-Error "Failed to obtain authentication token. $_"
    exit 1
}

# Prepare headers for subsequent API requests
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("accept", "application/json")
$headers.Add("Authorization", "Bearer $access_token")

# Fetch the detailed list of devices from NinjaOne
$devices_url = "https://$NinjaOneInstance/v2/devices-detailed"
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
} catch {
    Write-Error "Failed to fetch devices. $_"
    exit 1
}

# Function: Invoke-NinjaAPIRequest with a retry mechanism
function Invoke-NinjaAPIRequest {
    param (
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$Method = 'GET',
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [string]$Body = $null
    )

    $maxRetries = 3
    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        try {
            return Invoke-RestMethod -Uri $Uri -Method $Method -Headers $Headers -Body $Body -ContentType "application/json"
        } catch {
            Write-Error "API request to $Uri failed on attempt $($retryCount + 1): $_"
            $retryCount++
            Start-Sleep -Seconds 2
        }
    }
    Write-Error "API request to $Uri failed after $maxRetries attempts."
    return $null
}

# Process each device import entry and prepare asset objects with dynamic custom fields.
$assets = foreach ($deviceimport in $deviceimports) {
    # Find the matching device by ID
    $device = $devices | Where-Object { $_.id -eq $deviceimport.Id }

    # Build a dynamic dictionary of custom fields from CSV.
    # Exclude known fields ('Id' and 'name'); include all other columns.
    $customFields = @{}
    foreach ($property in $deviceimport.PSObject.Properties) {
        if ($property.Name -notin @("Id", "name")) {
            # Check for empty value.
            if ([string]::IsNullOrEmpty($property.Value)) {
                if ($OverwriteEmptyValues) {
                    # Include the property with a $null value to overwrite existing data.
                    $customFields[$property.Name] = $null
                } else {
                    # Skip the property to leave current data intact.
                    continue
                }
            } else {
                $customFields[$property.Name] = $property.Value
            }
        }
    }

    if ($device) {
        [PSCustomObject]@{
            Name         = $deviceimport.name
            ID           = $deviceimport.Id
            SystemName   = $device.systemName
            CustomFields = $customFields
        }
    } else {
        Write-Warning "Device ID $($deviceimport.Id) not found in the devices list."
        [PSCustomObject]@{
            Name         = $deviceimport.name
            ID           = $deviceimport.Id
            SystemName   = $null
            CustomFields = $customFields
        }
    }
}

# Debug: Print out the assets imported.
Write-Host "Imported Assets:"
$assets | ForEach-Object { Write-Host "ID: $($_.ID) - Name: $($_.Name) - SystemName: $($_.SystemName)" }

# Update custom fields for each asset (only if SystemName is not null and there are custom fields to update)
foreach ($asset in $assets) {
    if (($null -ne $asset.SystemName) -and $asset.CustomFields.Count -gt 0) {
        # Define NinjaOne API endpoint for updating custom fields.
        $customfields_url = "https://$NinjaOneInstance/api/v2/device/$($asset.ID)/custom-fields"

        # Convert the dynamic custom fields dictionary to JSON.
        $json = $asset.CustomFields | ConvertTo-Json -Depth 3

        Write-Host "Patching custom fields for: $($asset.SystemName) with data:"
        Write-Host $json

        # Update the custom fields via the API using our helper function.
        $result = Invoke-NinjaAPIRequest -Uri $customfields_url -Method 'Patch' -Headers $headers -Body $json
        if ($result -eq $null) {
            Write-Error "Failed to update custom fields for $($asset.Name)."
        }
        
        # Optional: Delay to help with API rate limits.
        Start-Sleep -Seconds 1
    } else {
        Write-Warning "Skipping update for asset with ID $($asset.ID) as SystemName is null or no custom fields provided."
    }
}

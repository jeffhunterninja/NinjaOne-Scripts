<#

This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.

## Additional notes: 
## This works with the 6.0 version of NinjaOne
## Currently only contacts can be assigned to devices via the API
## Technicians and end users will need to be manually assigned within the NinjaOne webapp for the time being

===
CSV is formatted like this:
name,device,email 
John Doe,Laptop123,john.doe@example.com
Jane Smith,Desktop456,jane.smith@example.com
Bob Johnson,Tablet789,bob.johnson@example.com

Only device name and email are truly required
#>

# Your NinjaRMM credentials
$NinjaOneInstance = 'ca.ninjarmm.com' # This varies depending on region or environment. For example, if you are in the US, this would be 'app.ninjarmm.com'
$NinjaOneClientId = ''
$NinjaOneClientSecret = ''
try {
    # Import device data from a CSV file
    $userimports = Import-CSV -Path "C:\Users\jeffh\Downloads\sample_devices.csv"
}
catch {
    Write-Output "There was an issue importing the CSV, please confirm file path."
    exit
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
$contacts_url = "https://$NinjaOneInstance/ws/api/v2/contacts"
$users_url = "https://$NinjaOneInstance/ws/api/v2/users"
try {
    $devices = Invoke-RestMethod -Uri $devices_url -Method GET -Headers $headers
    $contacts = Invoke-RestMethod -Uri $contacts_url -Method GET -Headers $headers
    $users = Invoke-RestMethod -Uri $users_url -Method GET -Headers $headers
} catch {
    Write-Error "Failed to fetch from API. $_"
    exit 1
}

$assets = foreach ($userimport in $userimports) {
    $importEmail = ($userimport.email -as [string]).Trim().ToLower()
    if ([string]::IsNullOrWhiteSpace($importEmail)) {
        Write-Warning "Skipping CSV row with missing email: $($userimport | ConvertTo-Json -Depth 1)"
        continue
    }

    # Try to find a matching user account first (preferred)
    $currentUser = $null
    if ($users) {
        $currentUser = $users | Where-Object { ($_.'email' -as [string]) -and ($_.email.Trim().ToLower() -eq $importEmail) } | Select-Object -First 1
    }

    # If not found in users, fall back to contacts
    $matchedSource = 'users'
    if (-not $currentUser -and $contacts) {
        $currentUser = $contacts | Where-Object {
            # contacts may use different property names; try common ones
            ( ($_.email -as [string]) -and ($_.email.Trim().ToLower() -eq $importEmail) ) -or
            ( ($_.Email -as [string]) -and ($_.Email.Trim().ToLower() -eq $importEmail) )
        } | Select-Object -First 1

        if ($currentUser) { $matchedSource = 'contacts' }
    }

    if ($currentUser) {
        # Normalize the fields we expect (email and uid). Contacts may have slightly different shapes.
        $resolvedEmail = ($currentUser.email -as [string])
        if (-not $resolvedEmail) { $resolvedEmail = ($currentUser.Email -as [string]) }

        # OwnerUid field may be 'uid' or 'id' depending on endpoint; try common names
        $ownerUid = $null
        if ($currentUser.uid) { $ownerUid = $currentUser.uid }
        elseif ($currentUser.id) { $ownerUid = $currentUser.id }
        elseif ($currentUser.Id) { $ownerUid = $currentUser.Id }

        [PSCustomObject]@{
            Name      = $userimport.name
            Device    = $userimport.device
            AssetOwner = $importEmail
            Email     = $resolvedEmail
            OwnerUid  = $ownerUid
            MatchedFrom = $matchedSource
        }
    } else {
        Write-Warning "User/Contact $($userimport.email) not found in users or contacts; no user assignment will be processed for device '$($userimport.device)'."
    }
}

# Create hash tables for displayname and systemname with case-insensitive keys
$displayNameMap = @{}
$systemNameMap  = @{}

# This correlates the deviceId in NinjaOne by matching the device name provided by the CSV.
# The display name is used for matching purposes. If there is no display name, the systemName is used.
foreach ($device in $devices) {
    if ($device.displayname) {
        $key = $device.displayname.ToLower()
        if (-not $displayNameMap.ContainsKey($key)) {
            $displayNameMap[$key] = $device
        }
        else {
            # Handle duplicate displaynames
            Write-Warning "Duplicate displayname found: $($device.displayname)"
        }
    }
    elseif ($device.systemname) {
        # Only checks for systemname if displayname is not present
        $key = $device.systemname.ToLower()
        if (-not $systemNameMap.ContainsKey($key)) {
            $systemNameMap[$key] = $device
        }
        else {
            # Handle duplicate systemnames
            Write-Warning "Duplicate systemname found: $($device.systemname)"
        }
    }
}

# Loop through each asset and assign IDs based on the hash tables
foreach ($asset in $assets) {
    $assetNameLower = $asset.Device.ToLower()
    Add-Member -InputObject $asset -NotePropertyName "ID" -NotePropertyValue '0' -Force
    if ($displayNameMap.ContainsKey($assetNameLower)) {
        $asset.ID = $displayNameMap[$assetNameLower].id
    }
    elseif ($systemNameMap.ContainsKey($assetNameLower)) {
        $asset.ID = $systemNameMap[$assetNameLower].id
    }
    else {
        # Optional: Log or handle assets with no matching device
        Write-Warning "No matching device found for asset: $($asset.Device)"
    }
}

# Debugging: Print out the assets imported
Write-Host "Imported Assets:"
$assets | ForEach-Object { Write-Host "Device: $($_.Device) - Device ID:$($_.ID) - Email: $($_.Email) - Name: $($_.Name) - Uid: $($_.OwnerUid)" }

# Assign the user for each asset
foreach ($asset in $assets) {
    if ($null -ne $asset.ID -or $asset.OwnerUid) {
        # Define NinjaOne API endpoint for updating the assigned user
        $assigneduser_url = "https://$NinjaOneInstance/ws/api/v2/device/" + $asset.ID + "/owner/" + $asset.ownerUid

        Write-Host "Updating assigned user for:" $asset.Device "with data:" $assetowner

        # Assign the user via the API
        try {
            Invoke-RestMethod -Method 'Post' -Uri $assigneduser_url -Headers $headers -Body $json -ContentType "application/json" -Verbose
        } catch {
            Write-Error "Failed to assign user for $($asset.Name). $_"
        }
    } else {
        Write-Warning "Skipping updating user $($asset.Name) as deviceId or ownerUid is null."
    }
}

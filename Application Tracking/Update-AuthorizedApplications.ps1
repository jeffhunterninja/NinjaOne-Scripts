param (
    [string]$Option1,  # comma-separated string of organizations or "ALL" that should automatically authorize all software currently installed
    [bool]$Option2,  # True/False parameter for execution execution of targets (specified in Option1)
    [string]$Option3,  # comma-separated string of organizations or "ALL"
    [string]$Option4,  # comma-separated string of software to add per org (specified in Option2)
    [string]$Option5,  # comma-separated string of organizations or "All"
    [string]$Option6,  # comma-separated string of software to remove per org (specified in Option5)
    [string]$Option7,  # comma-separated string of devices
    [string]$Option8,  # comma-separated string of software to add per device (specified in Option7)
    [string]$Option9,  # comma-separated string of devices
    [string]$Option10  # comma-separated string of software to remove per device (specified in Option9)
)



if ($env:commaSeparatedListOfOrganizationsToUpdate -and $env:commaSeparatedListOfOrganizationsToUpdate -notlike "null") { 
    $Option1 = $env:commaSeparatedListOfOrganizationsToUpdate 
}

if ($env:updateOrganizationsBasedOnCurrentSoftwareInventory -and $env:updateOrganizationsBasedOnCurrentSoftwareInventory -notlike "null") { 
    $Option2 = [System.Convert]::ToBoolean($env:updateOrganizationsBasedOnCurrentSoftwareInventory) 
}

if ($env:appendToOrganizations -and $env:appendToOrganizations -notlike "null") { 
    $Option3 = $env:appendToOrganizations 
}

if ($env:softwareToAppend -and $env:softwareToAppend -notlike "null") { 
    $Option4 = $env:softwareToAppend 
}

if ($env:removeFromOrganizations -and $env:removeFromOrganizations -notlike "null") { 
    $Option5 = $env:removeFromOrganizations 
}

if ($env:softwareToRemove -and $env:softwareToRemove -notlike "null") { 
    $Option6 = $env:softwareToRemove 
}

if ($env:appendToDevices -and $env:appendToDevices -notlike "null") { 
    $Option7 = $env:appendToDevices 
}

if ($env:deviceSoftwareToAppend -and $env:deviceSoftwareToAppend -notlike "null") { 
    $Option8 = $env:deviceSoftwareToAppend 
}

if ($env:removeFromDevices -and $env:removeFromDevices -notlike "null") { 
    $Option9 = $env:removeFromDevices 
}

if ($env:deviceSoftwareToRemove -and $env:deviceSoftwareToRemove -notlike "null") { 
    $Option10 = $env:deviceSoftwareToRemove 
}

# Configuration
$NinjaOneInstance = Ninja-Property-Get ninjaoneInstance
$NinjaOneClientId = Ninja-Property-Get ninjaoneClientId
$NinjaOneClientSecret = Ninja-Property-Get ninjaoneClientSecret

# Authentication
$authBody = @{
    grant_type    = "client_credentials"
    client_id     = $NinjaOneClientId
    client_secret = $NinjaOneClientSecret
    scope         = "monitoring management"
}
$authHeaders = @{
    accept        = 'application/json'
    "Content-Type" = 'application/x-www-form-urlencoded'
}

try {
    $authResponse = Invoke-RestMethod -Uri "https://$NinjaOneInstance/oauth/token" -Method POST -Headers $authHeaders -Body $authBody
    $accessToken = $authResponse.access_token
} catch {
    Write-Error "Failed to authenticate with NinjaOne API: $_"
    exit 1
}

# Headers for API requests
$headers = @{
    accept        = 'application/json'
    Authorization = "Bearer $accessToken"
}

# Fetch organizations from NinjaOne
$organizationsUrl = "https://$NinjaOneInstance/v2/organizations"
try {
    $organizations = Invoke-RestMethod -Uri $organizationsUrl -Method GET -Headers $headers
} catch {
    Write-Error "Failed to fetch organizations: $_"
    exit 1
}

# Fetch organizations from NinjaOne
$devicesUrl = "https://$NinjaOneInstance/v2/devices?df=class%20in%20(WINDOWS_WORKSTATION,%20WINDOWS_SERVER)"
try {
    $devices = Invoke-RestMethod -Uri $devicesUrl -Method GET -Headers $headers
} catch {
    Write-Error "Failed to fetch organizations: $_"
    exit 1
}

# Validate parameters and enforce conditions
function Test-Parameters {
    if ($Option1 -and $Option2) {
        Write-Host "Updating organizations to authorize all current installed applications. Based on this selection, any other sections of this script will not run." -ForegroundColor Green
    }

    if ($Option2 -and -not $Option1) {
        Write-Error "Attempted to set authorized applications to all current installed software, but there was no target for updating authorized software" -ErrorAction Stop
    }

    if ($Option4 -and -not $Option3) {
        Write-Error "You provided software to add, but no target organizations. Specify organizations or use 'ALL'." -ErrorAction Stop
    }
    
    if ($Option3 -and -not $Option4) {
        Write-Error "You provided target organizations, but no software to add. Specify a list of software to add." -ErrorAction Stop
    }
    
    if ($Option6 -and -not $Option5) {
        Write-Error "You provided software to remove, but did not provide target organizations." -ErrorAction Stop
    }
    
    if ($Option5 -and -not $Option6) {
        Write-Error "You provided target organizations, but no software to remove." -ErrorAction Stop
    }
    
    if ($Option7 -and -not $Option8) {
        Write-Error "You provided devices to update, but no software to add." -ErrorAction Stop
    }
    
    if ($Option8 -and -not $Option7) {
        Write-Error "You provided software to add, but no devices to target." -ErrorAction Stop
    }
    
    if ($Option10 -and -not $Option9) {
        Write-Error "You provided software to remove, but no devices to target." -ErrorAction Stop
    }
    
    if ($Option9 -and -not $Option10) {
        Write-Error "You provided devices to update, but no software to remove." -ErrorAction Stop
    }    

    # List all options to check for length
    $ParametersToCheck = @($Option1, $Option2, $Option3, $Option4, $Option5, $Option6, $Option7, $Option8, $Option9, $Option10)

    # Check each parameter's length
    foreach ($param in $ParametersToCheck) {
        if ($param -and $param.Length -gt 2048) {
            Write-Error "Your entry exceeds the character limit of 2048." -ErrorAction Stop
        }
    }
    return 0
}

# Call validation function
if (Test-Parameters -eq 1) {
    # Exit after Option1 execution
    Write-Host "Exiting script after executing overwrite with current software inventory." -ForegroundColor Yellow
    return
}

# Create a HashSet of device IDs for quick lookup
$deviceIds = @($devices.id) | ForEach-Object { $_ } | Sort-Object -Unique

# Section 1 - Updating organizations with current software inventory as authorized applications
if ($Option2) {
    if ($Option1 -ieq "ALL") {
        Write-Host "Updating all organizations with newly authorized software..." -ForegroundColor Cyan
        # Process each organization
        foreach ($organization in $organizations) {
            if ($null -ne $organization) {
                $orgId = $organization.id
                $softwareInventoryUrl = "https://$NinjaOneInstance/v2/queries/software?df=org%3D$orgId"

                # Retrieve software inventory for the organization
                try {
                    $softwareInventory = Invoke-RestMethod -Uri $softwareInventoryUrl -Method GET -Headers $headers
                     # Filter software inventory objects where the deviceId exists in the deviceIds
                     $filteredSoftwareInventory = $softwareInventory.results | Where-Object {
                        $deviceIds -contains $_.deviceId
                    }
                    $softwareNames = $filteredSoftwareInventory | ForEach-Object { $_.name } | Sort-Object | Get-Unique
                   
                } catch {
                    Write-Error "Failed to retrieve software inventory for $($OrgName): $_"
                    continue
                }

                # Prepare the custom field value
                $softwareList = $softwareNames -join ","
                Write-Host "Software inventory for '$OrgName': $softwareList"

                # Prepare request body to update custom field
                $customFieldsUrl = "https://$NinjaOneInstance/api/v2/organization/$orgId/custom-fields"
                $requestBody = @{
                    softwareList = @{html = $softwareList}
                } | ConvertTo-Json -Depth 10

                # Update custom field
                try {
                    Start-Sleep -Milliseconds 200 # Rate limiting
                    Invoke-RestMethod -Method PATCH -Uri $customFieldsUrl -Headers $headers -Body $requestBody -ContentType "application/json"
                    Write-Host "Successfully updated custom field for '$OrgName'."
                } catch {
                    Write-Error "Failed to update custom field for $($OrgName): $_"
                }
            } else {
                Write-Warning "Organization '$OrgName' not found in NinjaOne."
            }
        }
    } else {
        if ($Option1) {
            Write-Host "Updating select organizations with newly authorized software..." -ForegroundColor Cyan
            # Process each organization
            if ($Option1 -like "*,*") {
                $Option1 = $Option1 -split ',' | ForEach-Object { "$_".Trim() }
            }
            foreach ($OrgName in $Option1) {
                $organization = $organizations | Where-Object { $_.name -eq $OrgName }

                if ($null -ne $organization) {
                    $orgId = $organization.id
                    $softwareInventoryUrl = "https://$NinjaOneInstance/v2/queries/software?df=org%3D$orgId"

                    # Retrieve software inventory for the organization
                    try {
                        $softwareInventory = Invoke-RestMethod -Uri $softwareInventoryUrl -Method GET -Headers $headers
                         # Filter software inventory objects where the deviceId exists in the deviceIds
                         $filteredSoftwareInventory = $softwareInventory.results | Where-Object {
                            $deviceIds -contains $_.deviceId
                        }
                        $softwareNames = $filteredSoftwareInventory | ForEach-Object { $_.name } | Sort-Object | Get-Unique
                       
                    } catch {
                        Write-Error "Failed to retrieve software inventory for $($OrgName): $_"
                        continue
                    }

                    # Prepare the custom field value
                    $softwareList = $softwareNames -join ","
                    Write-Host "Software inventory for '$OrgName': $softwareList"

                    # Prepare request body to update custom field
                    $customFieldsUrl = "https://$NinjaOneInstance/api/v2/organization/$orgId/custom-fields"
                    $requestBody = @{
                        softwareList = @{html = $softwareList}
                    } | ConvertTo-Json -Depth 10

                    # Update custom field
                    try {
                        Start-Sleep -Milliseconds 200 # Rate limiting
                        Invoke-RestMethod -Method PATCH -Uri $customFieldsUrl -Headers $headers -Body $requestBody -ContentType "application/json"
                        Write-Host "Successfully updated custom field for '$OrgName'."
                    } catch {
                        Write-Error "Failed to update custom field for $($OrgName): $_"
                    }
                } else {
                    Write-Warning "Organization '$OrgName' not found in NinjaOne."
                }
            }
        }
    }
    Write-Host "Executing update of organizations..." -ForegroundColor Cyan
    return
}

# Section 2 - Add software to authorized list for all or select organizations
if ($Option3) {
    if ($Option3 -ieq "ALL") {
        Write-Host "Updating all organizations with newly authorized software..." -ForegroundColor Cyan
        $NewCustomFieldValues = $Option4 -split ","
            foreach ($organization in $organizations) {            
                if ($null -ne $organization) {
                    $orgId = $organization.id
                    $customFieldsUrl = "https://$NinjaOneInstance/api/v2/organization/$orgId/custom-fields"
            
                    # Retrieve existing custom field values
                    try {
                        $customFields = Invoke-RestMethod -Uri $customFieldsUrl -Method GET -Headers $headers
                        $existingValue = $customFields.softwareList.text -as [string]
                    } catch {
                        Write-Error "Failed to retrieve custom fields for $($OrgName): $_"
                        continue
                    }
            
                    # Combine existing and new values
                    if ($existingValue) {
                        $combinedValues = ($existingValue -split ",") + $NewCustomFieldValues
                        $updatedValue = ($combinedValues | Sort-Object | Get-Unique) -join ","
                    } else {
                        $updatedValue = $NewCustomFieldValues -join ","
                    }
            
                    # Prepare request body
                    $requestBody = @{
                        softwareList = @{html = $updatedValue}
                    } | ConvertTo-Json -Depth 10
            
                    # Update custom field
                    Write-Host "Updating custom field for organization '$OrgName' with ID: $orgId"
                    Write-Host "New value: $updatedValue"
            
                    try {
                        Start-Sleep -Milliseconds 200 # Rate limiting
                        Invoke-RestMethod -Method PATCH -Uri $customFieldsUrl -Headers $headers -Body $requestBody -ContentType "application/json"
                        Write-Host "Successfully updated custom field for '$OrgName'."
                    } catch {
                        Write-Error "Failed to update custom field for $($OrgName): $_"
                    }
                } else {
                    Write-Warning "Organization '$OrgName' not found in NinjaOne."
                }
            }
    } else {
        if ($Option4) {
            Write-Host "Updating select organizations with newly authorized software..." -ForegroundColor Cyan
            # Input: CSV-formatted strings
            # Convert input CSV strings to arrays
            $OrgNames = $Option3 -split ","
            $NewCustomFieldValues = $Option4 -split ","
            foreach ($OrgName in $OrgNames) {
                $organization = $organizations | Where-Object { $_.name -eq $OrgName }
            
                if ($null -ne $organization) {
                    $orgId = $organization.id
                    $customFieldsUrl = "https://$NinjaOneInstance/api/v2/organization/$orgId/custom-fields"
            
                    # Retrieve existing custom field values
                    try {
                        $customFields = Invoke-RestMethod -Uri $customFieldsUrl -Method GET -Headers $headers
                        $existingValue = $customFields.softwareList.text -as [string]
                    } catch {
                        Write-Error "Failed to retrieve custom fields for $($OrgName): $_"
                        continue
                    }
            
                    # Combine existing and new values
                    if ($existingValue) {
                        $combinedValues = ($existingValue -split ",") + $NewCustomFieldValues
                        $updatedValue = ($combinedValues | Sort-Object | Get-Unique) -join ","
                    } else {
                        $updatedValue = $NewCustomFieldValues -join ","
                    }
            
                    # Prepare request body
                    $requestBody = @{
                        softwareList = @{html = $updatedValue}
                    } | ConvertTo-Json -Depth 10
            
                    # Update custom field
                    Write-Host "Updating custom field for organization '$OrgName' with ID: $orgId"
                    Write-Host "New value: $updatedValue"
            
                    try {
                        Start-Sleep -Milliseconds 200 # Rate limiting
                        Invoke-RestMethod -Method PATCH -Uri $customFieldsUrl -Headers $headers -Body $requestBody -ContentType "application/json"
                        Write-Host "Successfully updated custom field for '$OrgName'."
                    } catch {
                        Write-Error "Failed to update custom field for $($OrgName): $_"
                    }
                } else {
                    Write-Warning "Organization '$OrgName' not found in NinjaOne."
                }
            }
            }
    }
}

# Section 3 - Remove software from all or select organizations
if ($Option5) {
    if ($Option5 -ieq "ALL") {
        Write-Host "Updating all organizations to remove software from authorized list..." -ForegroundColor Gray
        $ValuesToRemove = $Option6 -split ","

            # Process each organization
            foreach ($organization in $organizations) {
                if ($null -ne $organization) {
                    $orgId = $organization.id
                    $customFieldsUrl = "https://$NinjaOneInstance/api/v2/organization/$orgId/custom-fields"

                    # Retrieve existing custom field values
                    try {
                        $customFields = Invoke-RestMethod -Uri $customFieldsUrl -Method GET -Headers $headers
                        $existingValue = $customFields.softwareList.text -as [string]
                    } catch {
                        Write-Error "Failed to retrieve custom fields for $($OrgName): $_"
                        continue
                    }

                    # Remove specified values from existing custom field value
                    if ($existingValue) {
                        $filteredValues = ($existingValue -split ",") | Where-Object { $_ -notin $ValuesToRemove }
                        $updatedValue = $filteredValues -join ","
                    } else {
                        Write-Warning "No existing value for '$OrgName', skipping update."
                        continue
                    }


                    # Prepare request body
                    $requestBody = @{
                        softwareList = @{html = $updatedValue}
                    } | ConvertTo-Json -Depth 10

                    # Update custom field
                    Write-Host "Updating custom field for organization '$OrgName' with ID: $orgId"
                    Write-Host "Updated value: $updatedValue"

                    try {
                        Start-Sleep -Milliseconds 200 # Rate limiting
                        Invoke-RestMethod -Method PATCH -Uri $customFieldsUrl -Headers $headers -Body $requestBody -ContentType "application/json"
                        Write-Host "Successfully updated custom field for '$OrgName'."
                    } catch {
                        Write-Error "Failed to update custom field for $($OrgName): $_"
                    }
                } else {
                    Write-Warning "Organization '$OrgName' not found in NinjaOne."
                }
            }
    } else {
        if ($Option6) {
            Write-Host "Updating select organizations to remove software from authorized list..." -ForegroundColor Gray
            # Convert input CSV strings to arrays
            $OrgNames = $Option5 -split ","
            $ValuesToRemove = $Option6 -split ","

            # Process each organization
            foreach ($OrgName in $OrgNames) {
                $organization = $organizations | Where-Object { $_.name -eq $OrgName }

                if ($null -ne $organization) {
                    $orgId = $organization.id
                    $customFieldsUrl = "https://$NinjaOneInstance/api/v2/organization/$orgId/custom-fields"

                    # Retrieve existing custom field values
                    try {
                        $customFields = Invoke-RestMethod -Uri $customFieldsUrl -Method GET -Headers $headers
                        $existingValue = $customFields.softwareList.text -as [string]
                    } catch {
                        Write-Error "Failed to retrieve custom fields for $($OrgName): $_"
                        continue
                    }

                    # Remove specified values from existing custom field value
                    if ($existingValue) {
                        $filteredValues = ($existingValue -split ",") | Where-Object { $_ -notin $ValuesToRemove }
                        $updatedValue = $filteredValues -join ","
                    } else {
                        Write-Warning "No existing value for '$OrgName', skipping update."
                        continue
                    }


                    # Prepare request body
                    $requestBody = @{
                        softwareList = @{html = $updatedValue}
                    } | ConvertTo-Json -Depth 10

                    # Update custom field
                    Write-Host "Updating custom field for organization '$OrgName' with ID: $orgId"
                    Write-Host "Updated value: $updatedValue"

                    try {
                        Start-Sleep -Milliseconds 200 # Rate limiting
                        Invoke-RestMethod -Method PATCH -Uri $customFieldsUrl -Headers $headers -Body $requestBody -ContentType "application/json"
                        Write-Host "Successfully updated custom field for '$OrgName'."
                    } catch {
                        Write-Error "Failed to update custom field for $($OrgName): $_"
                    }
                } else {
                    Write-Warning "Organization '$OrgName' not found in NinjaOne."
                }
            }
        }
    }
}

# Section 4 - Adding authorized software to select devices
if ($Option7) {
    if ($Option7 -ieq "ALL") {
        Write-Error "Not possible to update software to all devices, provide a CSV target..." -ForegroundColor Red
    } else {
        if ($Option8) {
            Write-Host "Updating select devices with newly authorized software..." -ForegroundColor Cyan
            # Convert CSV strings to arrays
            $DeviceNames = $Option7 -split ","
            $NewCustomFieldValues = $Option8 -split ","

            # Process each device in the list
            foreach ($DeviceName in $DeviceNames) {
                $device = $devices | Where-Object { $_.systemName -eq $DeviceName }

                if ($null -ne $device) {
                    # Define the URL for fetching/updating custom fields
                    $customFieldsUrl = "https://$NinjaOneInstance/api/v2/device/$($device.id)/custom-fields"

                    # Retrieve existing custom field values
                    try {
                        $customFields = Invoke-RestMethod -Uri $customFieldsUrl -Method GET -Headers $headers
                        $existingValue = $customFields.deviceSoftwareList.text -as [string]
                    } catch {
                        Write-Error "Failed to retrieve custom fields for $($DeviceName): $_"
                        continue
                    }

                    # Combine existing and new values
                    if ($existingValue) {
                        $combinedValues = ($existingValue -split ",") + $NewCustomFieldValues
                        $updatedValue = ($combinedValues | Sort-Object | Get-Unique) -join ","
                    } else {
                        $updatedValue = $NewCustomFieldValues -join ","
                    }

                    # Prepare the updated custom field body
                    $requestBody = @{
                        deviceSoftwareList = @{html=$updatedValue}
                    } | ConvertTo-Json -Depth 10

                    # Update the custom field
                    try {
                        Start-Sleep -Milliseconds 200 # Rate limiting
                        Invoke-RestMethod -Uri $customFieldsUrl -Method PATCH -Headers $headers -Body $requestBody -ContentType "application/json"
                        Write-Host "Updated custom field for $DeviceName with value: $updatedValue"
                    } catch {
                        Write-Error "Failed to update custom field for $($DeviceName): $_"
                    }
                } else {
                    Write-Warning "Device $DeviceName not found in NinjaOne."
                }
            }

        }
    }
}

# Section 5 - Removing authorized software from select devices
if ($Option9) {
    if ($Option9 -ieq "ALL") {
        Write-Error "Not possible to remove software across all devices, provide a CSV target..." -ForegroundColor Red
        # Add your "ALL" logic here
    } else {
        if ($Option10) {
            Write-Host "Updating select devices to remove software from authorized list..." -ForegroundColor Cyan
            # Convert input CSV strings to arrays
            $DeviceNames = $Option9 -split ","
            $ValuesToRemove = $Option10 -split ","

            # Process each device
            foreach ($DeviceName in $DeviceNames) {
                $device = $devices | Where-Object { $_.systemName -eq $DeviceName }

                if ($null -ne $device) {
                    $deviceId = $device.id
                    $customFieldsUrl = "https://$NinjaOneInstance/api/v2/device/$deviceId/custom-fields"

                    # Retrieve existing custom field values
                    try {
                        $customFields = Invoke-RestMethod -Uri $customFieldsUrl -Method GET -Headers $headers
                        $existingValue = $customFields.deviceSoftwareList.text -as [string]
                    } catch {
                        Write-Error "Failed to retrieve custom fields for $($DeviceName): $_"
                        continue
                    }

                    # Remove specified values from existing custom field value
                    if ($existingValue) {
                        $filteredValues = ($existingValue -split ",") | Where-Object { $_ -notin $ValuesToRemove }
                        $updatedValue = $filteredValues -join ","
                    } else {
                        Write-Warning "No existing value for '$DeviceName', skipping update."
                        continue
                    }
                    if (!$updatedValue) {
                        $updatedValue = $null
                    }
                    # Prepare request body
                    $requestBody = @{
                        deviceSoftwareList = $updatedValue
                    } | ConvertTo-Json -Depth 10

                    # Update custom field
                    Write-Host "Updating custom field for device '$DeviceName' with ID: $deviceId"
                    Write-Host "Updated value: $updatedValue"

                    try {
                        Start-Sleep -Milliseconds 200 # Rate limiting
                        Invoke-RestMethod -Method PATCH -Uri $customFieldsUrl -Headers $headers -Body $requestBody -ContentType "application/json"
                        Write-Host "Successfully updated custom field for '$DeviceName'."
                    } catch {
                        Write-Error "Failed to update custom field for $($DeviceName): $_"
                    }
                } else {
                    Write-Warning "Device '$DeviceName' not found in NinjaOne."
                }
            }
        }
    }
}

Write-Host "Script execution completed." -ForegroundColor Green

#Requires -Version 5.1
<#
.SYNOPSIS
  Syncs NinjaOne patch status (pending, approved, failed) from the API into device custom fields.

.DESCRIPTION
  Retrieves patch data from NinjaOne API queries, compares with device custom fields
  pendingPatches, approvedPatches, and failedPatches, and updates or clears them. Clears
  stale values when a device no longer has patches in a category. Intended for scheduled
  runs (e.g. hourly) from an API server or automation host. Requires NinjaOneDocs module
  and credentials from NinjaOne custom properties.

.EXIT CODES
  0 = Success
  1 = Missing credentials or PowerShell 7 / module failure
  2 = API connect or request error
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# Check for required PowerShell version (7+)
if (!($PSVersionTable.PSVersion.Major -ge 7)) {
    try {
        if (!(Test-Path "$env:SystemDrive\Program Files\PowerShell\7")) {
            Write-Output 'Does not appear Powershell 7 is installed'
            exit 1
        }

        # Refresh PATH
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path', 'User')
        
        # Restart script in PowerShell 7
        pwsh -File "`"$PSCommandPath`"" @PSBoundParameters
        
    }
    catch {
        Write-Output 'PowerShell 7 was not installed. Update PowerShell and try again.'
        throw $Error
    }
    finally { exit $LASTEXITCODE }
}

# Install or update the NinjaOneDocs module or create your own fork here https://github.com/lwhitelock/NinjaOneDocs
try {
    $moduleName = "NinjaOneDocs"
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Install-Module -Name $moduleName -Force -AllowClobber
    }
    Import-Module $moduleName
}
catch {
    Write-Error "Failed to import NinjaOneDocs module. Error: $_"
    exit 2
}

# Your NinjaRMM credentials - these should be stored in secure NinjaOne custom fields
$NinjaOneInstance = Ninja-Property-Get ninjaoneInstance
$NinjaOneClientId = Ninja-Property-Get ninjaoneClientId
$NinjaOneClientSecret = Ninja-Property-Get ninjaoneClientSecret

if ([string]::IsNullOrWhiteSpace($NinjaOneInstance) -or [string]::IsNullOrWhiteSpace($NinjaOneClientId) -or [string]::IsNullOrWhiteSpace($NinjaOneClientSecret)) {
    Write-Output "Missing required API credentials (ninjaoneInstance, ninjaoneClientId, ninjaoneClientSecret). Set all three in NinjaOne custom properties."
    exit 1
}

# Connect to NinjaOne using the Connect-NinjaOne function
try {
    Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientId -NinjaOneClientSecret $NinjaOneClientSecret
}
catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit 2
}

function Compare-And-UpdateCustomFields {
    param (
        [string]$deviceId,
        [string]$fieldName,
        [string]$newValue
    )
    try {
        $currentFields = Invoke-NinjaOneRequest -Method GET -Path "device/$deviceId/custom-fields"
        $currentValue = $currentFields."$fieldName"
        Write-Verbose "Retrieved value of $currentValue"
    } catch {
        Write-Warning "Failed to retrieve custom fields for device ID $deviceId. Error: $_"
        return
    }

    # Compare current value with new value
    if ($currentValue -ne $newValue) {
        Write-Host "$(Get-Date) - Updating custom field '$fieldName' for device ID $deviceId"
        $request_body = @{
            $fieldName = $newValue
        } | ConvertTo-Json

        # Perform the update
        try {
            Invoke-NinjaOneRequest -Method PATCH -Path "device/$deviceId/custom-fields" -Body $request_body
            Write-Host "Successfully updated '$fieldName' for device ID $deviceId"
        } catch {
            Write-Warning "Failed to update custom fields for device ID $deviceId. Error: $_"
        }
    } else {
        Write-Host "$(Get-Date) - No update needed for '$fieldName' on device ID $deviceId"
    }
}

$pendingCF = "pendingPatches"
$approvedCF = "approvedPatches"
$failedCF = "failedPatches"

# Define query parameters for patch installations
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    status          = 'FAILED'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Call Invoke-NinjaOneRequest using splatting
$patchfailures = Invoke-NinjaOneRequest -Method GET -Path 'queries/os-patch-installs' -QueryParams $QueryParamString | Select-Object -ExpandProperty 'results'

# Define query parameters for patch installations
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    status          = 'MANUAL'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Call Invoke-NinjaOneRequest using splatting
$pendingpatches = Invoke-NinjaOneRequest -Method GET -Path 'queries/os-patches' -QueryParams $QueryParamString | Select-Object -ExpandProperty 'results'

# Define query parameters for patch installations
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    status          = 'APPROVED'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Call Invoke-NinjaOneRequest using splatting
$approvedpatches = Invoke-NinjaOneRequest -Method GET -Path 'queries/os-patches' -QueryParams $QueryParamString | Select-Object -ExpandProperty 'results'

# Define query parameters for patch installations
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    fields = 'pendingPatches'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Call Invoke-NinjaOneRequest using splatting
$pendingcustomfields = Invoke-NinjaOneRequest -Method GET -Path 'queries/custom-fields-detailed' -QueryParams $QueryParamString -Paginate | Select-Object -ExpandProperty 'results'

# Define query parameters for patch installations
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    fields = 'failedPatches'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Call Invoke-NinjaOneRequest using splatting
$failedcustomfields = Invoke-NinjaOneRequest -Method GET -Path 'queries/custom-fields-detailed' -QueryParams $QueryParamString -Paginate | Select-Object -ExpandProperty 'results'

# Define query parameters for patch installations
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    fields = 'approvedPatches'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Call Invoke-NinjaOneRequest using splatting
$approvedcustomfields = Invoke-NinjaOneRequest -Method GET -Path 'queries/custom-fields-detailed' -QueryParams $QueryParamString -Paginate | Select-Object -ExpandProperty 'results'

# Process pending patches
$groupedpending = $pendingpatches | Group-Object -Property deviceId
# Process pending patches
$groupedfailed = $patchfailures | Group-Object -Property deviceId
# Process pending patches
$groupedapproved = $approvedpatches | Group-Object -Property deviceId


foreach ($group in $groupedpending) {
    $deviceId = $group.Name
    $updatesForDevice = $group.Group

    # Convert updates to JSON string for comparison
    $newValue = ($updatesForDevice | ForEach-Object { $_.name }) -join ","
    Compare-And-UpdateCustomFields -deviceId $deviceId -fieldName "pendingPatches" -newValue $newValue
}

foreach ($group in $groupedfailed) {
    $deviceId = $group.Name
    $updatesForDevice = $group.Group

    # Convert updates to JSON string for comparison
    $newValue = ($updatesForDevice | ForEach-Object { $_.name }) -join ","
    Compare-And-UpdateCustomFields -deviceId $deviceId -fieldName "failedPatches" -newValue $newValue
}

foreach ($group in $groupedapproved) {
    $deviceId = $group.Name
    $updatesForDevice = $group.Group

    # Convert updates to JSON string for comparison
    $newValue = ($updatesForDevice | ForEach-Object { $_.name }) -join ","
    Compare-And-UpdateCustomFields -deviceId $deviceId -fieldName "approvedPatches" -newValue $newValue
}


# Create hashtables for quick membership checks
$PendingDeviceIds   = @{}
$FailedDeviceIds    = @{}
$ApprovedDeviceIds  = @{}

$groupedpending | ForEach-Object   { $PendingDeviceIds[[string]$_.Name]   = $true }
$groupedfailed | ForEach-Object    { $FailedDeviceIds[[string]$_.Name]    = $true }
$groupedapproved | ForEach-Object  { $ApprovedDeviceIds[[string]$_.Name]  = $true }

# Check for stale pendingPatches
foreach ($cf in $pendingcustomfields) {
    # Convert deviceId to string to match the keys in the hashtable
    $deviceId = [string]$cf.deviceId
    $currentPending = $cf.pendingPatches

    # If there's data in pendingPatches but the device isn't in the current $groupedpending list, it's stale
    if ([string]::IsNullOrWhiteSpace($currentPending) -eq $false -and -not $PendingDeviceIds.ContainsKey($deviceId)) {
        Write-Host "$(Get-Date) - Stale pendingPatches found for device $deviceId. Clearing field."
        Compare-And-UpdateCustomFields -deviceId $deviceId -fieldName "pendingPatches" -newValue ""
    }
}

# Check for stale failedPatches
foreach ($cf in $failedcustomfields) {
    # Convert deviceId to string to match the keys in the hashtable
    $deviceId = [string]$cf.deviceId
    $currentFailed = $cf.failedPatches

    # If there's data in failedPatches but the device isn't in the current $groupedfailed list, it's stale
    if ([string]::IsNullOrWhiteSpace($currentFailed) -eq $false -and -not $FailedDeviceIds.ContainsKey($deviceId)) {
        Write-Host "$(Get-Date) - Stale failedPatches found for device $deviceId. Clearing field."
        Compare-And-UpdateCustomFields -deviceId $deviceId -fieldName "failedPatches" -newValue ""
    }
}

# Check for stale approvedPatches
foreach ($cf in $approvedcustomfields) {
    $deviceId = [string]$cf.deviceId
    $currentApproved = $cf.approvedPatches

    # If there's data in approvedPatches but the device isn't in the current $groupedapproved list, it's stale
    if ([string]::IsNullOrWhiteSpace($currentApproved) -eq $false -and -not $ApprovedDeviceIds.ContainsKey($deviceId)) {
        Write-Host "$(Get-Date) - Stale approvedPatches found for device $deviceId. Clearing field."
        Compare-And-UpdateCustomFields -deviceId $deviceId -fieldName "approvedPatches" -newValue ""
    }
}

exit 0

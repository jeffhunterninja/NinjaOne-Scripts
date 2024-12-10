<#
.SYNOPSIS
    Generates a NinjaOne status report, including unauthorized applications, device details, and location mapping.

.DESCRIPTION
    This script:
    - Connects to the NinjaOne API using credentials.
    - Retrieves devices, organizations, and custom fields.
    - Matches devices with their corresponding custom field values (e.g., unauthorized applications).
    - Produces a status report in CSV format containing devices, locations, organizations, and custom fields.
    - Outputs a master list of all unauthorized applications across all devices.

.PARAMETER Action
    The script checks for PowerShell 7+ support and automatically restarts in PowerShell 7 if required.

.PARAMETER NinjaOneInstance
    [string] Your NinjaOne instance URL. Retrieved from a secure custom field.

.PARAMETER NinjaOneClientId
    [string] Your NinjaOne API client ID. Retrieved from a secure custom field.

.PARAMETER NinjaOneClientSecret
    [string] Your NinjaOne API client secret. Retrieved from a secure custom field.

.INPUTS
    - NinjaOne credentials (securely stored custom fields).
    - PowerShell 7 must be installed and available.
    - The script uses the `NinjaOneDocs` module (GitHub: https://github.com/lwhitelock/NinjaOneDocs).

.OUTPUTS
    - A CSV report saved in the `C:\temp\` directory.
    - A master list of unauthorized applications displayed in the console.

.EXAMPLE
    # Run the script to generate the NinjaOne Status Report
    .\ScriptName.ps1

    Output:
        - Report saved to C:\temp\yyyyMMdd_Ninja_Status_Report.csv
        - A list of unauthorized applications is displayed in the console.

.EXAMPLE
    # Restart in PowerShell 7 if necessary and generate the report
    pwsh -File .\ScriptName.ps1

.NOTES
    - **Dependencies**: PowerShell 7+, `NinjaOneDocs` module.
    - The script ensures PowerShell 7 is installed and uses it for execution.
    - Devices and organizations are enriched with location and custom field details.
    - NinjaOne API credentials are retrieved securely using the NinjaOne property functions.

.LINK
    NinjaOneDocs GitHub Module: https://github.com/lwhitelock/NinjaOneDocs
#>


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
    exit
}


# Your NinjaRMM credentials - these should be stored in secure NinjaOne custom fields
$NinjaOneInstance = Ninja-Property-Get ninjaoneInstance
$NinjaOneClientId = Ninja-Property-Get ninjaoneClientId
$NinjaOneClientSecret = Ninja-Property-Get ninjaoneClientSecret

if (!$ninjaoneInstance -and !$NinjaOneClientId -and !$NinjaOneClientSecret) {
    Write-Output "Missing required API credentials"
    exit 1
}

# Connect to NinjaOne using the Connect-NinjaOne function
try {
    Connect-NinjaOne -NinjaOneInstance $NinjaOneInstance -NinjaOneClientID $NinjaOneClientId -NinjaOneClientSecret $NinjaOneClientSecret
}
catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit
}
    
# Get date of today
$today = Get-Date -format "yyyyMMdd"

# file paths
$status_report = "C:\temp\" + $today + "_Ninja_Status_Report.csv"

# Fetch devices and organizations using module functions
try {
    $devices = Invoke-NinjaOneRequest -Method GET -Path 'devices' -QueryParams "df=class%20in%20(WINDOWS_WORKSTATION,%20WINDOWS_SERVER)"
    $organizations = Invoke-NinjaOneRequest -Method GET -Path 'organizations'
    $locations = Invoke-NinjaOneRequest -Method GET -Path 'locations'
}
catch {
    Write-Error "Failed to retrieve devices, location, or organizations. Error: $_"
    exit
}

# Define query parameters for patch installations
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    fields          = 'unauthorizedApplications'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Call Invoke-NinjaOneRequest using splatting
$customfields = Invoke-NinjaOneRequest -Method GET -Path 'queries/custom-fields-detailed' -QueryParams $QueryParamString -Paginate | Select-Object -ExpandProperty 'results'

$customFieldIDs = $customfields | Select-Object -ExpandProperty deviceId
$matchingdevices = $devices | Where-Object { $customFieldIDs -contains $_.id }

$assets = Foreach ($device in $matchingdevices) {
    [PSCustomObject]@{
        DeviceName = $device.systemName
        DeviceID = $device.id
        LocationName = 0
        LocationID = $device.locationId
        OrganizationName = 0
        OrganizationID = $device.organizationId
        CustomField = 0
    }
}
foreach ($location in $locations) {
        $currentDev = $assets | Where-Object {$_.LocationID -eq $location.id}
    $currentDev | Add-Member -MemberType NoteProperty -Name 'LocationName' -Value $location.name -Force
    }

foreach ($organization in $organizations) {
        $currentDev = $assets | Where-Object {$_.OrganizationID -eq $organization.id}
    $currentDev | Add-Member -MemberType NoteProperty -Name 'OrganizationName' -Value $organization.name -Force
    }

foreach ($customfield in $customfields) {
    $currentDev = $assets | Where-Object {$_.DeviceID -eq $customfield.deviceId}
$currentDev | Add-Member -MemberType NoteProperty -Name 'CustomField' -Value $customfield.fields.value -Force
}

    #remove the IDs that aren't necessary for the report
$assets | Select-Object devicename, customfield, locationname, OrganizationName | Format-Table | Out-String

Write-Host 'Creating the final report'

$assets | Select-Object devicename, customfield, locationname, OrganizationName | Export-CSV -NoTypeInformation -Path $status_report  

Write-Host "csv files have been created with success!"
Write-Host "Go to " $status_report " to find your Status Report"

# Extract, split, and clean the CustomField values
$uniqueValues = $assets.CustomField `
    -split ',\s*' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique

# Create a master comma-separated list of the unique values
$masterList = $uniqueValues -join ', '

# Output the master list
Write-Host "All unauthorized applications across all organizations: $masterList"

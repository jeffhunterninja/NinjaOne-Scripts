<#
.SYNOPSIS
    Backs up or restores authorized software lists for organizations and devices in NinjaOne.

.DESCRIPTION
    This script provides two primary actions: 
    1. **Backup**: Retrieves and saves the authorized software lists for organizations and devices to a JSON file.
    2. **Restore**: Restores authorized software lists for organizations or devices from a backup JSON file.

    The script supports filtering targets for restoration (organizations or devices) and selecting the most recent backup file automatically.

.PARAMETER Action
    [string] Specifies the action to perform.
    Valid values: "Backup" or "Restore".

.PARAMETER BackupFile
    [string] The path to the backup JSON file to be used for restoration. 
    Required when performing a Restore action.

.PARAMETER BackupDirectory
    [string] The directory containing backup files. 
    If no BackupFile is specified, the script will automatically select the most recent backup file.

.PARAMETER TargetType
    [string] The type of targets to restore.
    Valid values: "All", "Organizations", or "Devices".

.PARAMETER RestoreTargets
    [string] A comma-separated list of specific targets to restore (organization names or device names). 
    If TargetType is "All", this parameter is ignored.

.EXAMPLE
    # Perform a backup of all authorized software lists
    .\ScriptName.ps1 -Action Backup

.EXAMPLE
    # Restore authorized software lists for all organizations from the most recent backup
    .\ScriptName.ps1 -Action Restore -BackupDirectory "C:\Backups" -TargetType Organizations

.EXAMPLE
    # Restore authorized software for specific devices from a specific backup file
    .\ScriptName.ps1 -Action Restore -BackupFile "C:\Backups\Backup_20240923_120000.json" `
                     -TargetType Devices -RestoreTargets "Device1,Device2"

.EXAMPLE
    # Automatically find the most recent backup and restore software lists for all devices
    .\ScriptName.ps1 -Action Restore -BackupDirectory "C:\Backups" -TargetType Devices

.INPUTS
    Environment variables:
        - $env:action          : Specifies the action ("Backup" or "Restore").
        - $env:backupFile      : Path to the backup file for restoration.

.OUTPUTS
    - For **Backup**: A JSON file containing authorized software lists is saved to the specified or default location.
    - For **Restore**: Updates authorized software lists in NinjaOne and outputs success/failure messages.

.NOTES
    - PowerShell Version: 5.1 or later.
    - The script requires API credentials for NinjaOne.

#>

param(
    # Action Parameter to Determine Backup or Restore
    [Parameter(Mandatory = $false, HelpMessage = "Specify the action to perform: Backup or Restore.")]
    [string]$Action,

    # Restore-Specific Parameters
    [Parameter(Mandatory = $false, HelpMessage = "Path to the backup JSON file for restoration.")]
    [string]$BackupFile,

    [Parameter(Mandatory = $false, HelpMessage = "Directory containing backup files for restoration.")]
    [string]$BackupDirectory,

    [Parameter(Mandatory = $false, HelpMessage = "Type of target to restore: All, Organizations, or Devices.")]
    [string]$TargetType,

    [Parameter(Mandatory = $false, HelpMessage = "Comma-separated list of specific targets to restore.")]
    [string]$RestoreTargets
)
if ($env:action -and $env:action -notlike "null") { $Action = $env:action }
if ($env:backupFile -and $env:backupFile -notlike "null") { $BackupFile = $env:backupFile }
if ($env:backupDirectory -and $env:backupDirectory -notlike "null") { $BackupDirectory = $env:backupDirectory }
if ($env:targetType -and $env:targetType -notlike "null") { $TargetType = $env:targetType }
if ($env:restoreTargets -and $env:restoreTargets -notlike "null") { $RestoreTargets = $env:restoreTargets }

try {
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
}
catch {
    Write-Error "Failed to authenticate with NinjaOne API: $_"
    exit 1
}
try {
    $authResponse = Invoke-RestMethod -Uri "https://$NinjaOneInstance/oauth/token" -Method POST -Headers $authHeaders -Body $authBody
    $accessToken = $authResponse.access_token
    # Headers for API requests
    $headers = @{
        accept        = 'application/json'
        Authorization = "Bearer $accessToken"
    }
} catch {
    Write-Error "Failed to authenticate with NinjaOne API: $_"
    exit 1
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

function Backup-AuthorizedSoftware {
    param(
        [string]$OutputFile = "Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    )

    Write-Host "Starting backup of authorized software lists..."

    $backupData = [ordered]@{
        Organizations = @()
        Devices       = @()
    }

    # Backup Organizations
    foreach ($org in $organizations) {
        $orgId = $org.id
        $orgName = $org.name
        $customFieldsUrl = "https://$NinjaOneInstance/api/v2/organization/$orgId/custom-fields"
        try {
            $customFields = Invoke-RestMethod -Uri $customFieldsUrl -Method GET -Headers $headers
            $softwareList = $customFields.softwareList.text -as [string]
        } catch {
            Write-Warning "Failed to retrieve custom fields for organization '$orgName': $_"
            $softwareList = $null
        }

        $backupData.Organizations += [ordered]@{
            Name         = $orgName
            Id           = $orgId
            SoftwareList = $softwareList
        }
    }

    # Backup Devices
    foreach ($dev in $devices) {
        $deviceId = $dev.id
        $deviceName = $dev.systemName
        $customFieldsUrl = "https://$NinjaOneInstance/api/v2/device/$deviceId/custom-fields"
        try {
            $customFields = Invoke-RestMethod -Uri $customFieldsUrl -Method GET -Headers $headers
            $deviceSoftwareList = $customFields.deviceSoftwareList.text -as [string]
        } catch {
            Write-Warning "Failed to retrieve custom fields for device '$deviceName': $_"
            $deviceSoftwareList = $null
        }

        # Only add the device to backup if a value is present
        if ([string]::IsNullOrWhiteSpace($deviceSoftwareList)) {
            Write-Host "Skipping backup for device '$deviceName' as no software list is set." -ForegroundColor Yellow
            continue
        }

        $backupData.Devices += [ordered]@{
            Name              = $deviceName
            Id                = $deviceId
            DeviceSoftwareList = $deviceSoftwareList
        }
    }


    # Save backup to file
    $backupJson = $backupData | ConvertTo-Json -Depth 10
    $backupJson | Out-File $OutputFile -Encoding UTF8

    Write-Host "Backup complete. File saved to: $OutputFile" -ForegroundColor Green
}

function Get-MostRecentBackupFile {
    param (
        [string]$Directory,
        [string[]]$SearchKeywords
    )

    # Construct the search pattern by combining keywords with wildcards
    $searchPattern = $SearchKeywords | ForEach-Object { "*$_*" }

    # Initialize an array to hold matching files
    $matchingFiles = @()

    foreach ($pattern in $searchPattern) {
        $files = Get-ChildItem -Path $Directory -Filter $pattern -File -ErrorAction SilentlyContinue
        if ($files) {
            $matchingFiles += $files
        }
    }

    if ($matchingFiles.Count -eq 0) {
        throw "No backup files found in directory '$Directory' matching the keywords: $($SearchKeywords -join ', ')"
    }

    # Select the most recent file based on LastWriteTime
    $mostRecentFile = $matchingFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    return $mostRecentFile.FullName
}
function Restore-AuthorizedSoftware {
    [CmdletBinding(DefaultParameterSetName = 'Directory')]
    param(
        # Parameter Set 1: Specify Backup File
        [Parameter(Mandatory = $false, Position = 0, ParameterSetName = 'File')]
        [string]$BackupFile,

        # Parameter Set 2: Specify Backup Directory and Keywords
        [Parameter(Mandatory = $false, ParameterSetName = 'Directory')]
        [string]$BackupDirectory,

        # Common Parameters
        [Parameter(Mandatory = $true)]
        [ValidateSet("All","Organizations","Devices")]
        [string]$TargetType,

        [Parameter(Mandatory = $false)]
        [string]$RestoreTargets
    )

    # Function to find the most recent backup file matching the keywords
    

    # Determine the backup file based on the parameter set
    switch ($PSCmdlet.ParameterSetName) {
        'File' {
            $BackupFilePath = $BackupFile
            Write-Host "Using specified backup file: $BackupFilePath" -ForegroundColor Cyan
        }
        'Directory' {
            try {
                $BackupFilePath = Get-MostRecentBackupFile -Directory $BackupDirectory -SearchKeywords "Backup"
                Write-Host "Using most recent backup file: $BackupFilePath" -ForegroundColor Cyan
            } catch {
                Write-Error $_
                return
            }
        }
    }

    # Ensure that Targets are not specified when TargetType is 'All'
    if ($TargetType -eq "All" -and $RestoreTargets) {
        Write-Error "Cannot specify targets when TargetType is 'All'. Remove the Targets parameter or choose a different TargetType."
        return
    }

    # Read and parse the backup file
    try {
        $backupData = Get-Content $BackupFilePath -Raw | ConvertFrom-Json
    } catch {
        Write-Error "Failed to parse JSON from '$BackupFilePath': $_"
        return
    }

    # Ensure that Organizations and Devices are at least empty arrays if not present
    if (-not $backupData.Organizations) {
        $backupData.Organizations = @()
    }
    if (-not $backupData.Devices) {
        $backupData.Devices = @()
    }

    Write-Host "Starting restore from backup: $BackupFilePath" -ForegroundColor Cyan

    # Split targets if provided
    $TargetList = $null
    if ($RestoreTargets) {
        $TargetList = $RestoreTargets -split "," | ForEach-Object { $_.Trim() }
    }

    # Function to restore a single organization's authorized software
    function Restore-Organization($OrgData) {
        $orgId = $OrgData.Id
        $orgName = $OrgData.Name
        $updatedValue = $OrgData.SoftwareList

        if (-not $updatedValue) {
            Write-Host "Organization '$orgName' has no software list to restore." -ForegroundColor Yellow
            return
        }

        $customFieldsUrl = "https://$NinjaOneInstance/api/v2/organization/$orgId/custom-fields"
        $requestBody = @{
            softwareList = @{ html = $updatedValue }
        } | ConvertTo-Json -Depth 10

        try {
            Invoke-RestMethod -Method PATCH -Uri $customFieldsUrl -Headers $Headers -Body $requestBody -ContentType "application/json"
            Write-Host "Successfully restored authorized software for organization '$orgName'." -ForegroundColor Green
        } catch {
            Write-Error "Failed to restore authorized software for '$orgName': $_"
        }
    }

    # Function to restore a single device's authorized software
    function Restore-Device($DevData) {
        $deviceId = $DevData.Id
        $deviceName = $DevData.Name
        $updatedValue = $DevData.DeviceSoftwareList

        $customFieldsUrl = "https://$NinjaOneInstance/api/v2/device/$deviceId/custom-fields"
        $requestBody = if ($updatedValue) {
            @{ deviceSoftwareList = @{ html = $updatedValue } }
        } else {
            @{ deviceSoftwareList = $null }
        }
        $requestBody = $requestBody | ConvertTo-Json -Depth 10

        try {
            Invoke-RestMethod -Method PATCH -Uri $customFieldsUrl -Headers $Headers -Body $requestBody -ContentType "application/json"
            Write-Host "Successfully restored authorized software for device '$deviceName'." -ForegroundColor Green
        } catch {
            Write-Error "Failed to restore authorized software for '$($deviceName)': $_"
        }
    }

    switch ($TargetType) {
        "All" {
            # Restore all organizations
            foreach ($org in $backupData.Organizations) {
                Restore-Organization $org
            }
            # Restore all devices
            foreach ($dev in $backupData.Devices) {
                Restore-Device $dev
            }
        }

        "Organizations" {
            if (-not $TargetList) {
                # Restore all organizations
                foreach ($org in $backupData.Organizations) {
                    Restore-Organization $org
                }
            } else {
                # Prepare a lookup table only if we have organizations
                $OrgByName = @{}
                if ($backupData.Organizations.Count -gt 0) {
                    $OrgByName = $backupData.Organizations | Group-Object -Property Name -AsHashTable -AsString
                }
                foreach ($targetName in $TargetList) {
                    if ($OrgByName -and $OrgByName.ContainsKey($targetName)) {
                        Restore-Organization $OrgByName[$targetName]
                    } else {
                        Write-Warning "No matching organization found for '$targetName'. Skipping."
                    }
                }
            }
        }

        "Devices" {
            if (-not $TargetList) {
                # Restore all devices
                foreach ($dev in $backupData.Devices) {
                    Restore-Device $dev
                }
            } else {
                # Prepare a lookup table only if we have devices
                $DevByName = @{}
                if ($backupData.Devices.Count -gt 0) {
                    $DevByName = $backupData.Devices | Group-Object -Property Name -AsHashTable -AsString
                }

                foreach ($targetName in $TargetList) {
                    if ($DevByName -and $DevByName.ContainsKey($targetName)) {
                        Restore-Device $DevByName[$targetName]
                    } else {
                        Write-Warning "No matching device found for '$targetName'. Skipping."
                    }
                }
            }
        }
    }

    Write-Host "Restore process completed." -ForegroundColor Green
}

# Conditional Logic Based on $Action
switch ($Action) {
    "Backup" {
        # Backup Logic
        $OutputFile = "C:\Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
        Backup-AuthorizedSoftware -OutputFile $OutputFile
    }

    "Restore" {
        # Validation for Restore Parameters
        if (-not $BackupFile -and -not $BackupDirectory) {
            Write-Error "Error: You must provide either a BackupFile or BackupDirectory for restoration."
            exit
        }
        
        if (-not $BackupFile) {
            # If no BackupFile is provided, find the most recent backup file in the BackupDirectory
            try {
                $BackupFile = Get-MostRecentBackupFile -Directory $BackupDirectory -SearchKeywords "Backup"
                Write-Host "Using most recent backup file: $BackupFilePath" -ForegroundColor Cyan
            } catch {
                Write-Error $_
                exit
            }
        }    

        Restore-AuthorizedSoftware -BackupFile $BackupFile `
                                   -TargetType $TargetType `
                                   -RestoreTargets $RestoreTargets
    }

    default {
        Write-Error "Invalid Action. Use 'Backup' or 'Restore'."
    }
}



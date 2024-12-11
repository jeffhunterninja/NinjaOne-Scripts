<#
This is provided as an educational example of how to interact with the NinjaAPI.
Any scripts should be evaluated and tested in a controlled setting before being utilized in production.
As this script is an educational example, further improvements and enhancement may be necessary to handle larger datasets.
This is a PowerShell script to generate a report of Windows device patch installations for the current month using NinjaOne API

Thank you to Luke Whitelock (https://mspp.io) and @skolte (https://github.com/freezscholte/) on Discord!

The NinjaOne module can be forked here: https://github.com/lwhitelock/NinjaOneDocs

#>

[CmdletBinding()]
param (
    [Parameter()]
    [Switch]$CreateKB = [System.Convert]::ToBoolean($env:sendToKnowledgeBase),
    [Parameter()]
    [Switch]$CreateDocument = [System.Convert]::ToBoolean($env:sendToDocumentation),
    [Parameter()]
    [Switch]$CreateGlobalKB = [System.Convert]::ToBoolean($env:globalOverview),
    [Parameter()]
    [string]$ReportMonth = [System.Convert]::ToString($env:reportMonth) # Optional parameter (e.g., "December 2024")
)

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

# Initialize start time
$Start = Get-Date

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
if ($CreateKB -or $CreateDocument -or $CreateGlobalKB) {
# Define the month and year for the report
if ($ReportMonth) {
    try {
        # Parse input as "MMMM yyyy" (e.g., "December 2024")
        $ParsedDate = [datetime]::ParseExact($ReportMonth, "MMMM yyyy", [cultureinfo]::InvariantCulture)

        # Set the first and last day of the specified month
        $FirstDayOfMonth = Get-Date -Year $ParsedDate.Year -Month $ParsedDate.Month -Day 1
        $LastDayOfMonth = $FirstDayOfMonth.AddMonths(1).AddDays(-1)
        $currentMonth = $FirstDayOfMonth.ToString("MMMM")  # Get full month name (e.g., December)
        $currentYear = $FirstDayOfMonth.ToString("yyyy")   # Get year (e.g., 2024)
    }
    catch {
        Write-Error "Invalid ReportMonth format. Use 'MMMM yyyy' (e.g., 'December 2024')."
        exit 1
    }
}
else {
    # Default to the current month and year
    $FirstDayOfMonth = Get-Date -Day 1
    $LastDayOfMonth = (Get-Date -Day 1).AddMonths(1).AddDays(-1)
    # Define the current month and year
    $currentMonth = (Get-Date).ToString("MMMM")
    $currentYear = (Get-Date).Year.ToString()
}

# Formatting for API query parameters
$FirstDayString = $FirstDayOfMonth.ToString('yyyyMMdd')
$LastDayString = $LastDayOfMonth.ToString('yyyyMMdd')

# Display the date range being used
Write-Output "Generating report for: $($FirstDayOfMonth.ToString('MMMM yyyy'))"
Write-Output "Report Date Range: $($FirstDayOfMonth.ToShortDateString()) - $($LastDayOfMonth.ToShortDateString())"


# Collect user activities
$patchScans = @()
$patchScanFailures = @()
$patchApplicationCycles = @()
$patchApplicationFailures = @()

# Fetch devices and organizations using module functions
try {
    $devices = Invoke-NinjaOneRequest -Method GET -Path 'devices-detailed' -QueryParams "df=class%20in%20(WINDOWS_WORKSTATION,%20WINDOWS_SERVER)"
    $organizations = Invoke-NinjaOneRequest -Method GET -Path 'organizations'

}
catch {
    Write-Error "Failed to retrieve devices or organizations. Error: $_"
    exit
}

# Define query parameters for patch installations
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    status          = 'Installed'
    installedBefore = $LastDayString
    installedAfter  = $FirstDayString
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Call Invoke-NinjaOneRequest using splatting
$patchinstalls = Invoke-NinjaOneRequest -Method GET -Path 'queries/os-patch-installs' -QueryParams $QueryParamString | Select-Object -ExpandProperty 'results'

# Define query parameters for patch failures
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    status          = 'Failed'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Call Invoke-NinjaOneRequest using splatting
$patchfailures = Invoke-NinjaOneRequest -Method GET -Path 'queries/os-patch-installs' -QueryParams $QueryParamString | Select-Object -ExpandProperty 'results'

# Define query parameters for pending patches
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    status          = 'Manual'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Define query parameters for pending patches
$queryParams = @{
    df              = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
    status          = 'Approved'
}

# Format the query parameters into a string (URL encoding)
$QueryParamString = ($queryParams.GetEnumerator() | ForEach-Object { 
    "$($_.Key)=$($_.Value -replace ' ', '%20')"
}) -join '&'

# Fetch activities with built-in pagination
try {
    $queryParams2 = @{
        df     = 'class in (WINDOWS_WORKSTATION, WINDOWS_SERVER)'
        class  = 'DEVICE'
        type   = 'PATCH_MANAGEMENT'
        status = 'in (PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED, PATCH_MANAGEMENT_SCAN_COMPLETED, PATCH_MANAGEMENT_FAILURE)'
        after  = $FirstDayString
        before = $LastDayString
        pageSize = 1000
    }
    # Format the query parameters into a string
    # Manually replace spaces with %20 for proper URL formatting
    $QueryParamString2 = ($queryParams2.GetEnumerator() | ForEach-Object { 
        "$($_.Key)=$($_.Value -replace ' ', '%20')"
    }) -join '&'
    $allActivities = Invoke-NinjaOneRequest -Method GET -Path 'activities' -QueryParams $QueryParamString2 -Paginate

# Step 1: Get the first day of the current month
$firstDayOfCurrentMonth = Get-Date -Year (Get-Date).Year -Month (Get-Date).Month -Day 1

# Step 2: Convert the first day of the current month to Unix time (seconds since 1970)
$firstDayUnix = [System.DateTimeOffset]::new($firstDayOfCurrentMonth).ToUnixTimeSeconds()

# Step 3: Filter activities that occurred on or after the first day of the current month
$filteredActivities = $allActivities.activities | Where-Object {
    $_.activityTime -ge $firstDayUnix
}

# Now $filteredActivities contains only the activities that occurred on or after the first day of the current month

    # Convert the JSON response to activities array if needed
    if ($filteredActivities) {
        $userActivities = $filteredActivities
    }

    # Convert Unix timestamps (in seconds) or already formatted DateTime to readable DateTime
    $userActivities = $filteredActivities | ForEach-Object {
        $activity = $_
        # Safely handle the activityTime conversion
        $unixTime = $activity.activityTime

        if ($unixTime -is [int64]) {
            # Convert Unix time (seconds) to DateTime if it's an integer
            $activityTime = [System.DateTimeOffset]::FromUnixTimeSeconds($unixTime).DateTime
        }
        elseif ($unixTime -is [datetime]) {
            # If already a DateTime object, use it directly
            $activityTime = $unixTime
        }
        else {
            Write-Error "Invalid activity time format encountered: $unixTime"
            $activityTime = $null
        }

        # Update the activity's time field with the converted DateTime
        if ($activityTime) {
            $activity.activityTime = $activityTime
        }

        # Return the updated activity object
        $activity
    }

} catch {
    Write-Error "Failed to retrieve activities. Error: $_"
    exit
}

foreach ($activity in $userActivities) {
    if ($activity.activityResult -match "SUCCESS") {
        if ($activity.statusCode -match "PATCH_MANAGEMENT_SCAN_COMPLETED") {
            $patchScans += $activity
        } elseif ($activity.statusCode -match "PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED") {
            $patchApplicationCycles += $activity
        }
    } elseif ($activity.activityResult -match "FAILURE") {
        if ($activity.statusCode -match "PATCH_MANAGEMENT_SCAN_COMPLETED") {
            $patchScanFailures += $activity
        } elseif ($activity.statusCode -match "PATCH_MANAGEMENT_APPLY_PATCH_COMPLETED") {
            $patchApplicationFailures += $activity
        }
    }
}

# Index devices by ID for faster lookup
$deviceIndex = @{}
foreach ($device in $devices) {
    $deviceIndex[$device.id] = $device
}

# Initialize organization objects with tracked properties
foreach ($organization in $organizations) {
    Add-Member -InputObject $organization -NotePropertyName "PatchScans" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "PatchFailures" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "PatchInstalls" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "PatchApplications" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "Workstations" -NotePropertyValue @() -Force
    Add-Member -InputObject $organization -NotePropertyName "Servers" -NotePropertyValue @() -Force
}

# Assign devices to organizations
foreach ($device in $devices) {
    $currentOrg = $organizations | Where-Object { $_.id -eq $device.organizationId }
    if ($device.nodeClass.EndsWith("_SERVER")) {
        $currentOrg.Servers += $device.systemName
    } elseif ($device.nodeClass.EndsWith("_WORKSTATION") -or $device.nodeClass -eq "MAC") {
        $currentOrg.Workstations += $device.systemName
    }
}

# Process patch scans
foreach ($patchScan in $patchScans) {
    $device = $deviceIndex[$patchScan.deviceId]
    $patchScan | Add-Member -NotePropertyName "DeviceName" -NotePropertyValue $device.systemName -Force
    $patchScan | Add-Member -NotePropertyName "OrgID" -NotePropertyValue $device.organizationId -Force
    $organization = $organizations | Where-Object { $_.id -eq $device.organizationId }
    $organization.PatchScans += $patchScan
}

# Process patch application/update cycles
foreach ($patchApplicationCycle in $patchApplicationCycles) {
    $device = $deviceIndex[$patchApplicationCycle.deviceId]
    $patchApplicationCycle | Add-Member -NotePropertyName "DeviceName" -NotePropertyValue $device.systemName -Force
    $patchApplicationCycle | Add-Member -NotePropertyName "OrgID" -NotePropertyValue $device.organizationId -Force
    $organization = $organizations | Where-Object { $_.id -eq $device.organizationId }
    $organization.PatchApplications += $patchApplicationCycle
}

# Process patch installations
foreach ($patchinstall in $patchinstalls) {
    $device = $deviceIndex[$patchinstall.deviceId]
    $patchinstall | Add-Member -NotePropertyName "DeviceName" -NotePropertyValue $device.systemName -Force
    $patchinstall | Add-Member -NotePropertyName "OrgID" -NotePropertyValue $device.organizationId -Force
    $organization = $organizations | Where-Object { $_.id -eq $device.organizationId }
    $organization.PatchInstalls += $patchinstall
}

# Process patch installation failures
foreach ($patchfailure in $patchfailures) {
    $device = $deviceIndex[$patchfailure.deviceId]
    $patchfailure | Add-Member -NotePropertyName "DeviceName" -NotePropertyValue $device.systemName -Force
    $patchfailure | Add-Member -NotePropertyName "OrgID" -NotePropertyValue $device.organizationId -Force
    $organization = $organizations | Where-Object { $_.id -eq $device.organizationId }
    $organization.PatchFailures += $patchfailure
}

# Function to convert an array of objects to an HTML table
function ConvertTo-ObjectToHtmlTable {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Objects,

        [string]$NinjaOneInstance,

        # Add a parameter for properties you want to exclude
        [string[]]$ExcludedProperties = @('RowColour','deviceId')
    )

    $html = "<table class='table table-striped'>"
    $html += "<thead><tr>"
    
    # Exclude specified properties from the header
    foreach ($property in $Objects[0].PSObject.Properties.Name) {
        if (-not ($ExcludedProperties -contains $property)) {
            $html += "<th>$property</th>"
        }
    }
    $html += "</tr></thead><tbody>"

    foreach ($obj in $Objects) {
        # Apply row color if RowColour is present
        $rowColour = if ($obj.RowColour) { " class='$($obj.RowColour)'" } else { "" }
        $html += "<tr$rowColour>"
        
        # Iterate over object properties and exclude specified ones
        foreach ($property in $obj.PSObject.Properties.Name) {
            if (-not ($ExcludedProperties -contains $property)) {
                $value = $obj.$property

                # Use 'systemName' to identify the column for the link
                if ($property -eq 'DeviceName') {
                    $url = "https://$NinjaOneInstance/#/deviceDashboard/$($obj.deviceId)/overview"
                    $html += "<td><a href='$url' target='_blank'>$value</a></td>"
                } else {
                    $html += "<td>$value</td>"
                }
            }
        }
        $html += "</tr>"
    }

    $html += "</tbody></table>"
    return $html
}

# Define the Patch Report Template
$PatchReportTemplate = [PSCustomObject]@{
    name          = 'Patch Installation Reports'
    allowMultiple = $true
    fields        = @(
        [PSCustomObject]@{
            fieldLabel                = 'Patch Installation Details'
            fieldName                 = 'patchDetails'
            fieldType                 = 'WYSIWYG'
            fieldTechnicianPermission = 'READ_ONLY'
            fieldScriptPermission     = 'NONE'
            fieldApiPermission        = 'READ_WRITE'
            fieldContent              = @{
                required         = $False
                advancedSettings = @{
                    expandLargeValueOnRender = $True
                }
            }
        },
        [PSCustomObject]@{
            fieldLabel                = 'Patch Installations'
            fieldName                 = 'patchInstallations'
            fieldType                 = 'WYSIWYG'
            fieldTechnicianPermission = 'READ_ONLY'
            fieldScriptPermission     = 'NONE'
            fieldApiPermission        = 'READ_WRITE'
            fieldContent              = @{
                required         = $False
                advancedSettings = @{
                    expandLargeValueOnRender = $True
                }
            }
        }
    )
}
} else {
    Write-Output "Skipping organizational document and KB creation"
}


if ($CreateKB -or $CreateDocument) {
    # Create or retrieve the document template
    $DocTemplate = Invoke-NinjaOneDocumentTemplate $PatchReportTemplate

    # Fetch existing Patch Report Documents
    $PatchReportDocs = Invoke-NinjaOneRequest -Method GET -Path 'organization/documents' -QueryParams "templateIds=$($DocTemplate.id)"

    # Initialize lists for document updates and creations
    [System.Collections.Generic.List[PSCustomObject]]$NinjaDocUpdates = @()
    [System.Collections.Generic.List[PSCustomObject]]$NinjaDocCreation = @()
    [System.Collections.Generic.List[PSCustomObject]]$NinjaKBUpdates = @()
    [System.Collections.Generic.List[PSCustomObject]]$NinjaKBCreation = @()

    try {
        foreach ($organization in $organizations) {
            # Filter devices and patch installs for the current organization
            $currentDevices = $devices | Where-Object { $_.organizationId -eq $organization.id }
            
            # Ensure $currentDevices is an array to safely access .Count
            $currentDevices = @($currentDevices)
            
            # Check if there are no devices in the current organization
            if ($currentDevices.Count -le 0) {
                Write-Host "No devices found for organization '$($organization.name)'. Skipping to next organization."
                continue  # Skip to the next iteration of the loop
            }

            # Extract unique device IDs
            $currentDeviceIds = $currentDevices | Select-Object -ExpandProperty id -Unique

            # Filter patch installs for the current devices
            $currentPatchInstalls = $patchinstalls | Where-Object { $_.deviceId -in $currentDeviceIds }

            # Exclude specific security intelligence updates
            $trackedUpdates = $currentPatchInstalls | Where-Object {
                $_.name -notlike "*Security Intelligence Update for Microsoft Defender Antivirus*"
            }

            Write-Host "Processing organization: $($organization.name)"
            Write-Host "Number of tracked updates: $($trackedUpdates.Count)"

            # Check if there are no tracked updates
            if ($trackedUpdates.Count -le 0) {
                Write-Host "No tracked updates found for organization '$($organization.name)'. Skipping."
                continue  # Skip to the next organization
            }

            # Process each patch installation
            foreach ($install in $trackedUpdates) {
                # Retrieve the corresponding device
                $current_Device = $currentDevices | Where-Object { $_.id -eq $install.deviceId } | Select-Object -First 1

                # Add new properties to the patch installation object
                $install | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $current_Device.systemName -Force
                
                # Convert Unix timestamps to readable format
                $install.installedAt = ([DateTimeOffset]::FromUnixTimeSeconds($install.installedAt).DateTime).ToString()
                $install.timestamp = ([DateTimeOffset]::FromUnixTimeSeconds($install.timestamp).DateTime).ToString()
            }

            # Prepare the data for the table
            $tableDevices = $trackedUpdates | 
                Select-Object DeviceName,
                            name, 
                            installedAt,
                            kbNumber,
                            deviceId

            # Validate $tableDevices
            if ($null -eq $tableDevices -or $tableDevices.Count -eq 0) {
                Write-Host "No data available to generate HTML table for organization '$($organization.name)'. Skipping."
                continue  # Skip to the next organization
            }

            # Generate HTML table
            $htmltable = ConvertTo-ObjectToHtmlTable -Objects $tableDevices -NinjaOneInstance $NinjaOneInstance

                if ($CreateDocument) {
                    # Initialize statistics for the current organization
                    $patchStatistics = [PSCustomObject]@{
                        'Patch Scan Cycles' = ($organization.PatchScans).Count
                        'Patch Apply Cycles' = ($organization.PatchApplications).Count
                        'Patch Installations' = ($trackedUpdates).Count
                        'Failed Patches' = ($organization.PatchFailures).Count
                    }

                    $patchwidget = Get-NinjaOneInfoCard -Title "Patch Statistics" -Data $patchStatistics
                    # Filter the documents
                    $MatchedDoc = $PatchReportDocs | Where-Object { 
                        $_.organizationId -eq $organization.id -and
                        $_.documentName -match $currentMonth -and 
                        $_.documentName -match $currentYear -and 
                        $_.documentName -match "Patch Installation"
                    }
            
                    $MatchCount = ($MatchedDoc | Measure-Object).Count
            
                    if ($MatchCount -eq 0) {
                        Write-Host "No match found for $($organization.name)"
                    } elseif ($MatchCount -gt 1) {
                        Throw "Multiple NinjaOne Documents ($($MatchedDoc.documentId -join ',')) matched to $($organization.name)"
                        continue
                    } else {
                        Write-Host "Matched document ID: $($MatchedDoc.documentId) for $($organization.name)"
                    }
            
                    $DocFields = @{
                        'patchDetails'      = @{'html' = $patchwidget }
                        'patchInstallations'  = @{'html' = $htmltable }
                    }
            
                    if ($MatchedDoc) {
                        $UpdateObject = [PSCustomObject]@{
                            documentId   = $MatchedDoc.documentId
                            documentName = "$($organization.name) Patch Installation Report - $currentMonth $currentYear"
                            fields       = $DocFields
                        }
            
                        $NinjaDocUpdates.Add($UpdateObject)
            
                    } else {
                        $CreateObject = [PSCustomObject]@{
                            documentName       = "$($organization.name) Patch Installation Report - $currentMonth $currentYear"
                            documentTemplateId = $DocTemplate.id
                            organizationId     = [int]$organization.id
                            fields             = $DocFields
                        }
            
                        $NinjaDocCreation.Add($CreateObject)
                    }
                }
                    if ($CreateKB) {
                    $MatchingName = "$($organization.name) Patch Installation Report - $currentMonth $currentYear"

                    # Fetch existing Patch Report KB Articles
                    $PatchReportKBs = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/organization/articles' -QueryParams "articleName=$($MatchingName)"
                    
                    $KBMatchCount = ($PatchReportKBs | Measure-Object).Count
                    Write-Output "Found $($organization.name) had $($KBMatchCount) match"
                    $KBFields = @{
                        'html' = $htmltable
                        }   

                    if ($KBMatchCount -eq 0) {
                        Write-Host "No match found for $($MatchingName)"
                    } elseif ($KBMatchCount -gt 1) {
                        Throw "Multiple NinjaOne KBs with ($($PatchReportKBs.id -join ',')) matched to $($MatchingName)"
                        continue
                    } else {
                        Write-Host "Matched document ID: $($PatchReportKBs.id) for $($MatchingName)"
                    }

                    if ($PatchReportKBs) {
                        $UpdateObject = [PSCustomObject]@{
                            id   = $PatchReportKBs.Id
                            organizationId = $($organization.id)
                            destinationFolderPath = "$($organization.name) Monthly Patch Installation Reports"
                            name = "$($organization.name) Patch Installation Report - $currentMonth $currentYear"
                            content       = $KBFields
                        }

                        $NinjaKBUpdates.Add($UpdateObject)

                    } else {
                        $CreateObject = [PSCustomObject]@{
                            name       = "$($organization.name) Patch Installation Report - $currentMonth $currentYear"
                            organizationId = $($organization.id)
                            destinationFolderPath = "$($organization.name) Monthly Patch Installation Reports"
                            content             = $KBFields
                        }
                        $NinjaKBCreation.Add($CreateObject)
                    }
        }
    }


        # Perform the bulk updates of data
        if (($NinjaDocCreation | Measure-Object).count -ge 1) {
            Write-Host "Creating Documents"
            $CreatedDocs = Invoke-NinjaOneRequest -Path "organization/documents" -Method POST -InputObject $NinjaDocCreation -AsArray
            Write-Host "Created $(($CreatedDocs | Measure-Object).count) Documents"
        }

        if (($NinjaDocUpdates | Measure-Object).count -ge 1) {
            Write-Host "Updating Documents"
            $UpdatedDocs = Invoke-NinjaOneRequest -Path "organization/documents" -Method PATCH -InputObject $NinjaDocUpdates -AsArray
            Write-Host "Updated $(($UpdatedDocs | Measure-Object).count) Documents"
        }

        Write-Output "$(Get-Date): Complete Total Runtime: $((New-TimeSpan -Start $Start -End (Get-Date)).TotalSeconds) seconds"

        # Perform the bulk updates of data to the Knowledge Base
        if (($NinjaKBCreation | Measure-Object).count -ge 1) {
            Write-Host "Creating KB Articles"
            $CreatedKBs = Invoke-NinjaOneRequest -Path "knowledgebase/articles/" -Method POST -InputObject $NinjaKBCreation -AsArray
            Write-Host "Created $(($CreatedKBs | Measure-Object).count) KB Articles"
        }

        if (($NinjaKBUpdates | Measure-Object).count -ge 1) {
            Write-Host "Updating KB Articles"
            $UpdatedKBs = Invoke-NinjaOneRequest -Path "knowledgebase/articles" -Method PATCH -InputObject $NinjaKBUpdates -AsArray
            Write-Host "Updated $(($UpdatedKBs | Measure-Object).count) KB Articles"
        }

        Write-Output "$(Get-Date): Complete Total Runtime: $((New-TimeSpan -Start $Start -End (Get-Date)).TotalSeconds) seconds"

    }
    catch {
        Write-Output "Failed to Generate Documentation. Line number: $($_.InvocationInfo.ScriptLineNumber) Error: $($_.Exception.Message)"
        exit 1
    }
}
if ($CreateGlobalKB) {
    try {
        # Initialize an array to hold the aggregated data
    $AggregatedPatchInstalls = @()

    # Process each patch installation
    foreach ($patchinstall in $patchinstalls) {
        # Exclude specific security intelligence updates
        if ($patchinstall.name -like "*Security Intelligence Update for Microsoft Defender Antivirus*") {
            continue
        }

        # Get the device associated with the patchinstall
        $device = $deviceIndex[$patchinstall.deviceId]
        if (-not $device) {
            continue
        }

        # Get the organization associated with the device
        $organization = $organizations | Where-Object { $_.id -eq $device.organizationId }
        if (-not $organization) {
            continue
        }

        # Add properties to the patchinstall
        $patchinstall | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $device.systemName -Force
        $patchinstall | Add-Member -MemberType NoteProperty -Name "OrganizationName" -Value $organization.name -Force

        # Convert Unix timestamps to readable format
        $patchinstall.installedAt = ([DateTimeOffset]::FromUnixTimeSeconds($patchinstall.installedAt).DateTime).ToString()
        $patchinstall.timestamp = ([DateTimeOffset]::FromUnixTimeSeconds($patchinstall.timestamp).DateTime).ToString()

        # Collect the necessary properties into a custom object
        $AggregatedPatchInstall = [PSCustomObject]@{
            'OrganizationName' = $patchinstall.OrganizationName
            'DeviceName'       = $patchinstall.DeviceName
            'PatchName'        = $patchinstall.name
            'InstalledAt'      = $patchinstall.installedAt
            'KBNumber'         = $patchinstall.kbNumber
            'DeviceId'         = $patchinstall.deviceId
        }

        $AggregatedPatchInstalls += $AggregatedPatchInstall
    }
    # Generate HTML table
    $htmltableAggregated = ConvertTo-ObjectToHtmlTable -Objects $AggregatedPatchInstalls -NinjaOneInstance $NinjaOneInstance

    $KBFields = @{
    'html' = $htmltableAggregated
    }   

    $GlobalMatchingName = "Patch Installation Report - $currentMonth $currentYear"

    # Fetch existing Patch Report KB Articles
    $PatchReportKBs = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/global/articles' -QueryParams "articleName=$($GlobalMatchingName)"

    $MatchCount = ($PatchReportKBs | Measure-Object).Count

    if ($MatchCount -eq 0) {
        Write-Host "No match found for $($GlobalMatchingName)"
    } elseif ($MatchCount -gt 1) {
        Throw "Multiple NinjaOne KBs with ($($PatchReportKBs.documentId -join ',')) matched to $($GlobalMatchingName)"
        continue
    } else {
        Write-Host "Matched document ID: $($PatchReportKBs.Id) for $($GlobalMatchingName)"
    }

    if ($PatchReportKBs) {
        $UpdateObject = [PSCustomObject]@{
            id   = $PatchReportKBs.Id
            name = "Patch Installation Report - $currentMonth $currentYear"
            content       = $KBFields
        }
        $UpdatedGlobalKB = Invoke-NinjaOneRequest -Path "knowledgebase/articles" -Method PATCH -InputObject $UpdateObject -AsArray
    } else {
        $CreateObject = [PSCustomObject]@{
            name       = "Patch Installation Report - $currentMonth $currentYear"
            destinationFolderPath = "Monthly Patch Installation Reports"
            content             = $KBFields
        }
        $CreatedGlobalKB = Invoke-NinjaOneRequest -Path "knowledgebase/articles" -Method POST -InputObject $CreateObject -AsArray
    }
    }
    catch {
        Write-Output "Failed to Generate Global KB Article for All Windows Patches. Line number: $($_.InvocationInfo.ScriptLineNumber) Error: $($_.Exception.Message)"
        exit 1
    }
}

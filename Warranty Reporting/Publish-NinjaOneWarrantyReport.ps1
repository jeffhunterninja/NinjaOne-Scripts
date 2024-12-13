<#
.SYNOPSIS
    Publishes a comprehensive warranty report for NinjaOne-managed devices.

.DESCRIPTION
    The Publish-NinjaOneWarrantyReport script interacts with the NinjaAPI to generate detailed warranty reports for devices managed through NinjaOne. It automates data retrieval, processes warranty information, generates CSV and HTML reports, and updates or creates Knowledge Base (KB) articles and documentation within NinjaOne.

.PARAMETER CreateKB
    Switch to enable the creation or updating of global Knowledge Base (KB) articles. If not specified, the script checks the `sendToKnowledgeBase` environment variable.

.PARAMETER CreateDocument
    Switch to enable the creation or updating of organization-specific documentation. If not specified, the script checks the `sendToDocumentation` environment variable.

.PARAMETER CreateGlobalKB
    Switch to enable the creation or updating of a global warranty KB article. If not specified, the script checks the `globalOverview` environment variable.

.EXAMPLE
    .\Publish-NinjaOneWarrantyReport.ps1 -CreateKB -CreateDocument -CreateGlobalKB

    This command runs the script and enables the creation or updating of both global KB articles and organization-specific documentation.

#>

[CmdletBinding()]
param (
    [Parameter()]
    [Switch]$CreateKB = [System.Convert]::ToBoolean($env:sendToKnowledgeBase),
    [Parameter()]
    [Switch]$CreateDocument = [System.Convert]::ToBoolean($env:sendToDocumentation),
    [Parameter()]
    [Switch]$CreateGlobalKB = [System.Convert]::ToBoolean($env:globalOverview)
)

# File path to export
$today            = (Get-Date -Format "yyyy-MM-dd")
$warranty_report  = "C:\temp\$today`_NinjaOne_Warranty_Report.csv"

# Do you want a list of all devices or only devices with warranty information present?
# Toggle variable to control filtering
$filterWarranty = $false  # Set to $true to show only devices with warranty info, $false to show all
# Define the threshold for 'warning' (90 days)
$daysThreshold = 90
$CreateWarrantyGlobalKB = $true
$CreateKB = $true
$CreateDocument = $true

# Establish current month and year for report title
    $currentMonth = (Get-Date).ToString("MMMM")
    $currentYear = (Get-Date).Year.ToString()

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

function New-StatCard {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Description,

        [Parameter(Mandatory = $false)]
        [string]$Color = "# ",  # Default color: green

        [Parameter(Mandatory = $false)]
        [string]$IconClass = "fa-solid fa-arrow-down-up-across-line",  # Default icon

        [Parameter(Mandatory = $false)]
        [string]$Id = ""  # Optional ID attribute
    )

    # Prepare the ID attribute if provided
    if ($Id) {
        $IdAttr = "id='$Id'"
    } else {
        $IdAttr = ""
    }

    # Construct the HTML using a here-string with proper variable expansion
    $html = @"
<td style='border: 0px; white-space: nowrap'>
    <div class='stat-card' style='display: flex;'>
        <div class='stat-value' $IdAttr style='color: $Color;'>$Value</div>
        <div class='stat-desc'><i class='$IconClass'></i>&nbsp;&nbsp;$Description</div>
    </div>
</td>
"@

    return $html
} 
function ConvertTo-ObjectToHtmlTable {
    param (
        [Parameter(Mandatory = $true)]
        [array]$Objects,

        [string]$NinjaOneInstance,

        [string[]]$ExcludedProperties = @('RowColour','Device ID'),

        # Define a custom section order for RowColour
        [string[]]$SectionOrder = @("danger", "warning", "success", "unknown"),

        # Define custom headings corresponding to each RowColour
        [hashtable]$SectionHeadings = @{
            "danger"   = "Warranty Expired"
            "warning"    = "Warranty Expiring Soon"
            "success" = "Under Warranty"
            "unknown"  = "No Warranty Information Available"
        }
    )

    # Define a template map for different group outputs.
    # {HEADING} will be replaced by the group's heading text.
    $infoCardTemplates = @{
        "danger" = @"
<div class="info-card error">
  <i class="info-icon fa-solid fa-circle-exclamation"></i>
  <div class="info-text">
    <div class="info-title">{HEADING}</div>
    <div class="info-description">
      The warranty has expired for the following devices:
    </div>
  </div>
</div>
"@

        "warning" = @"
<div class="info-card warning">
  <i class="info-icon fa-solid fa-triangle-exclamation"></i>
  <div class="info-text">
    <div class="info-title">{HEADING}</div>
    <div class="info-description">
      The warranty expiration for these devices will occur within the next 90 days.
    </div>
  </div>
</div>
"@

        # If other groups (like yellow or green) do not have a special template,
        # you can either leave them off and handle defaults later, or define a default template:
        "success" = @"
<div class="info-card success">
<i class="info-icon fa-solid fa-circle-check"></i>
<div class="info-text">
    <div class="info-title">{HEADING}</div>
    <div class="info-description">
    These devices are under warranty
    </div>
</div>
</div>
"@

        "unknown" = @"
        <div class="info-card">
        <i class="info-icon fa-solid fa-circle-info"></i>
        <div class="info-text">
          <div class="info-title">{HEADING}</div>
          <div class="info-description">
            No warranty information is available for these devices.
          </div>
        </div>
      </div>
"@
    }

    # Group the objects by RowColour
    $groups = $Objects | Group-Object -Property RowColour

    # Order the groups based on the specified SectionOrder
    $orderedGroups = foreach ($color in $SectionOrder) {
        $groups | Where-Object { $_.Name -eq $color }
    }

    # Add any groups not in the SectionOrder to the end
    $remainingGroups = $groups | Where-Object { $SectionOrder -notcontains $_.Name }
    $orderedGroups += $remainingGroups

    # Filter to non-empty groups
    $nonEmptyGroups = $orderedGroups | Where-Object { $_.Group.Count -gt 0 }

    if ($nonEmptyGroups.Count -eq 0) {
        Write-Warning "No objects to display."
        return "<p>No objects.</p>"
    }

    # We'll build a master HTML string containing multiple sections
    $html = [System.Text.StringBuilder]::new()

    foreach ($group in $nonEmptyGroups) {
        $currentGroupName = $group.Name

        # Determine heading for this group
        $currentGroupHeading = $SectionHeadings[$currentGroupName]
        if (-not $currentGroupHeading) {
            $currentGroupHeading = $currentGroupName
        }

        # Get the template. If none exists for this group, use a default template.
        $template = $infoCardTemplates[$currentGroupName]
        if (-not $template) {
            # Default template if no match:
            $template = @"
<div class="info-card info">
  <i class="info-icon fa-solid fa-info-circle"></i>
  <div class="info-text">
    <div class="info-title">{HEADING}</div>
    <div class="info-description">
      This section corresponds to the $currentGroupName group.
    </div>
  </div>
</div>
"@
        }

        # Replace {HEADING} in the template with the actual heading
        $renderedTemplate = $template.Replace("{HEADING}", $currentGroupHeading)

        # Append the rendered template to the HTML
        [void]$html.AppendLine($renderedTemplate)

        # Determine the first object to generate headers
        $firstObject = $group.Group[0]

        # Start a new table for this group
        [void]$html.Append("<table class='table table-striped'>")

        # Create table headers
        [void]$html.Append("<thead><tr>")
        $visibleProperties = $firstObject.PSObject.Properties.Name | Where-Object { $ExcludedProperties -notcontains $_ }
        foreach ($property in $visibleProperties) {
            [void]$html.Append("<th>$property</th>")
        }
        [void]$html.Append("</tr></thead><tbody>")

        # Render each object's row
        foreach ($obj in $group.Group) {
            $rowColour = if ($obj.RowColour) { " class='$($obj.RowColour)'" } else { "" }
            [void]$html.Append("<tr$rowColour>")

            foreach ($property in $visibleProperties) {
                $value = $obj.$property
                if ($property -eq 'Device Name') {
                    $url = "https://$NinjaOneInstance/#/deviceDashboard/$($obj.'Device ID')/overview"
                    [void]$html.Append("<td><a href='$url' target='_blank'>$value</a></td>")
                } else {
                    [void]$html.Append("<td>$value</td>")
                }
            }

            [void]$html.Append("</tr>")
        }

        [void]$html.Append("</tbody></table>")
    }

    return $html.ToString()
}
function Convert-ActivityTime {
    param([Parameter(Mandatory)]$TimeValue)

    # Handle null or whitespace input
    if ([string]::IsNullOrWhiteSpace($TimeValue)) {
        return $null
    }

    if ($TimeValue -is [datetime]) {
        return $TimeValue.ToString("MM/dd/yyyy")
    } elseif ($TimeValue -is [int64]) {
        if ($TimeValue -gt 253402300799) {
            # Handle Unix timestamp in milliseconds
            return [System.DateTimeOffset]::FromUnixTimeMilliseconds($TimeValue).DateTime.ToString("MM/dd/yyyy")
        }
        # Handle Unix timestamp in seconds
        return [System.DateTimeOffset]::FromUnixTimeSeconds($TimeValue).DateTime.ToString("MM/dd/yyyy")
    } elseif ($TimeValue -is [double]) {
        # Round down double Unix timestamp
        $roundedUnixTime = [long][math]::Floor($TimeValue)
        return [System.DateTimeOffset]::FromUnixTimeSeconds($roundedUnixTime).DateTime.ToString("MM/dd/yyyy")
    } elseif ($TimeValue -is [string]) {
        # Attempt to parse string into a DateTime object
        try {
            return [datetime]::Parse($TimeValue).ToString("MM/dd/yyyy")
        } catch {
            return $TimeValue  # Return as-is if parsing fails
        }
    } else {
        # Return as-is for unsupported types
        return $TimeValue
    }
}

# Fetch devices and organizations using module functions
try {
    $devices = Invoke-NinjaOneRequest -Method GET -Path 'devices-detailed' -QueryParams "expand=warranty"
    $organizations = Invoke-NinjaOneRequest -Method GET -Path 'organizations'
    $locations = Invoke-NinjaOneRequest -Method GET -Path 'locations'
}
catch {
    Write-Error "Failed to retrieve devices, locations, or organizations. Error: $_"
    exit
}

# Build assets list
$assets = foreach ($device in $devices) {
    [PSCustomObject]@{
        "Device Name"         = $device.systemName
        "Device ID"           = $device.id
        "Location Name"       = 0
        "Location ID"         = $device.locationId
        "Organization Name"   = 0
        "Organization ID"     = $device.organizationId
        "Warranty Start"      = $device.references.warranty.startDate
        "Warranty End"        = $device.references.warranty.endDate
        "Mftr Fullfill"       = $device.references.warranty.manufacturerFulfillmentDate
    }
}


# Add location names
foreach ($location in $locations) {
    $currentDev = $assets | Where-Object { $_.'Location ID' -eq $location.id }
    $currentDev | Add-Member -MemberType NoteProperty -Name 'Location Name' -Value $location.name -Force
}

# Add organization names
foreach ($organization in $organizations) {
    $currentDev = $assets | Where-Object { $_.'Organization ID' -eq $organization.id }
    $currentDev | Add-Member -MemberType NoteProperty -Name 'Organization Name' -Value $organization.name -Force
}

# Convert Unix timestamps to DateTime, skipping missing values
$assets | ForEach-Object {
    if ($_.'Warranty Start') {
        $_.'Warranty Start' = Convert-ActivityTime $_.'Warranty Start'
    }
    if ($_.'Warranty End') {
        $_.'Warranty End' = Convert-ActivityTime $_.'Warranty End'
    }
    if ($_.'Mftr Fullfill') {
        $_.'Mftr Fullfill' = Convert-ActivityTime $_.'Mftr Fullfill'
    }
}

# Process the WarrantyEnd property
$assets | ForEach-Object {
    # Default RowColour to 'unknown'
    $rowColour = 'unknown'
    $WarrantyStatus = 'Unknown'

    # Check if WarrantyEnd is present and processable as a date
    try {
        if ($_.'Warranty End') {
            $warrantyEndDate = [datetime]$_.'Warranty End'
            $today = (Get-Date)
            $daysRemaining = ($warrantyEndDate - $today).Days

            # Determine RowColour based on the days remaining
            if ($daysRemaining -gt $daysThreshold) {
                $rowColour = 'success' # More than 90 days in the future
                $WarrantyStatus = 'Valid'
            }
            elseif ($daysRemaining -le $daysThreshold -and $daysRemaining -ge 0) {
                $rowColour = 'warning' # Within 90 days
                $WarrantyStatus = 'Support Ending'
            }
            elseif ($daysRemaining -lt 0) {
                $rowColour = 'danger'  # Warranty expired
                $WarrantyStatus = 'Expired'
            }
        }
    }
    catch {
        # Do nothing, keep RowColour as 'unknown'
    }

    # Add the RowColour property
    $_ | Add-Member -NotePropertyName "RowColour" -NotePropertyValue $rowColour -Force
    $_ | Add-Member -NotePropertyName "Warranty Status" -NotePropertyValue $WarrantyStatus -Force
}

# Apply filtering logic based on $filterWarranty variable
if ($filterWarranty) {
    # Filter objects with at least one valid property
    $assets = $assets | Where-Object { $_.'Warranty Start' -or $_.'Warranty End' }
}

# Remove the IDs that aren't necessary for the report
$finalReport = $assets | Select-Object `
    "Device Name", `
    "Warranty Status", `
    "Warranty Start", `
    "Warranty End", `
    RowColour, `
    "Organization Name", `
    "Location Name", `
    "Device ID"

# Display the report as a table
$finalReport | Format-Table | Out-String
# Create HTML table of warranty status by device
$tableHtml = ConvertTo-ObjectToHtmlTable -Objects $finalReport

# Export the final report to CSV
Write-Host 'Creating the final report...'
#$finalReport | Export-Csv -NoTypeInformation -Path $warranty_report

Write-Host "A CSV report has been generated for this information."
Write-Host "Go to $warranty_report to find your NinjaOne Warranty Report!"

$WarrantyCounts = @{
    "Expired" = ($assets | Where-Object { $_.RowColour -eq "danger" }).Count
    "Expiring" = ($assets | Where-Object { $_.RowColour -eq "warning" }).Count
    "Valid" = ($assets | Where-Object { $_.RowColour -eq "success" }).Count
    "Unknown" = ($assets | Where-Object { $_.RowColour -eq "unknown" }).Count
}

# Generate Stat Cards
$expiredWarranties = New-StatCard `
    -Value $WarrantyCounts."Expired" `
    -Description "Expired Warranties" `
    -Color "#FF4500" `
    -IconClass "fas fa-times-circle" `
    -Id "expiredWarranties"

$expiringWarranties = New-StatCard `
    -Value $WarrantyCounts."Expiring" `
    -Description "Expiring Warranties" `
    -Color "#FFD700" `
    -IconClass "fas fa-exclamation-triangle" `
    -Id "expiringWarranties"

$validWarranties = New-StatCard `
    -Value $WarrantyCounts."Valid" `
    -Description "Under Warranty" `
    -Color "#008001" `
    -IconClass "fas fa-check-circle" `
    -Id "validWarranties"

# Combine Stat Cards into a Table Row
$WarrantyTable = @"
<table style='border: 0px;'>
    <tbody>
        <tr>
            $expiredWarranties
            $expiringWarranties
            $validWarranties
        </tr>
    </tbody>
</table>
"@

if ($CreateWarrantyGlobalKB) {
    try {
    # Generate HTML table

    $combinedHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Warranty Report</title>
    <!-- Include any required CSS here -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        /* Example styles, customize as needed */
        .stat-card { /* Your styles here */ }
        .info-card { /* Your styles here */ }
        /* Add additional CSS to ensure proper styling */
    </style>
</head>
<body>
    <div class="container-fluid">
        $WarrantyTable
        <hr>
        $tableHtml
    </div>
</body>
</html>
"@
    $htmltableAggregated = $combinedHtml #ConvertTo-ObjectToHtmlTable -Objects $finalReport

    $KBFields = @{
    'html' = $htmltableAggregated
    }   

    $GlobalMatchingName = "Warranty Report - $currentMonth $currentYear"

    # Fetch existing Warranty Report KB Articles
    $WarrantyReportKBs = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/global/articles' -QueryParams "articleName=$($GlobalMatchingName)"

    $MatchCount = ($WarrantyReportKBs | Measure-Object).Count

    if ($MatchCount -eq 0) {
        Write-Host "No match found for $($GlobalMatchingName)"
    } elseif ($MatchCount -gt 1) {
        Throw "Multiple NinjaOne KBs with ($($WarrantyReportKBs.documentId -join ',')) matched to $($GlobalMatchingName)"
        continue
    } else {
        Write-Host "Matched document ID: $($WarrantyReportKBs.Id) for $($GlobalMatchingName)"
    }

    if ($WarrantyReportKBs) {
        $UpdateObject = [PSCustomObject]@{
            id   = $WarrantyReportKBs.Id
            name = "Warranty Report - $currentMonth $currentYear"
            content       = $KBFields
        }
        $UpdatedGlobalKB = Invoke-NinjaOneRequest -Path "knowledgebase/articles" -Method PATCH -InputObject $UpdateObject -AsArray
    } else {
        $CreateObject = [PSCustomObject]@{
            name       = "Warranty Report - $currentMonth $currentYear"
            destinationFolderPath = "Monthly Warranty Report"
            content             = $KBFields
        }
        $CreatedGlobalKB = Invoke-NinjaOneRequest -Path "knowledgebase/articles" -Method POST -InputObject $CreateObject -AsArray
    }
    }
    catch {
        Write-Output "Failed to Generate Global KB Article for All Warranties. Line number: $($_.InvocationInfo.ScriptLineNumber) Error: $($_.Exception.Message)"
        exit 1
    }
}

# Define the Patch Report Template
$WarrantyReportTemplate = [PSCustomObject]@{
    name          = 'Warranty Reports'
    allowMultiple = $true
    fields        = @(
        [PSCustomObject]@{
            fieldLabel                = 'Warranty Stats'
            fieldName                 = 'warrantyStats'
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
            fieldLabel                = 'Warranty Status By Device'
            fieldName                 = 'warrantyStatus'
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

if ($CreateKB -or $CreateDocument) {
    # Create or retrieve the document template
    $DocTemplate = Invoke-NinjaOneDocumentTemplate $WarrantyReportTemplate

    # Fetch existing Patch Report Documents
    $WarrantyReportDocs = Invoke-NinjaOneRequest -Method GET -Path 'organization/documents' -QueryParams "templateIds=$($DocTemplate.id)"

    # Initialize lists for document updates and creations
    [System.Collections.Generic.List[PSCustomObject]]$NinjaDocUpdates = @()
    [System.Collections.Generic.List[PSCustomObject]]$NinjaDocCreation = @()
    [System.Collections.Generic.List[PSCustomObject]]$NinjaKBUpdates = @()
    [System.Collections.Generic.List[PSCustomObject]]$NinjaKBCreation = @()

    try {
        foreach ($organization in $organizations) {
            # Filter devices and patch installs for the current organization
            $currentDevices = $assets | Where-Object { $_.'Organization ID' -eq $organization.id }
            $WarrantyCounts = @{
    "Expired" = ($currentDevices | Where-Object { $_.RowColour -eq "danger" }).Count
    "Expiring" = ($currentDevices | Where-Object { $_.RowColour -eq "warning" }).Count
    "Valid" = ($currentDevices | Where-Object { $_.RowColour -eq "success" }).Count
    "Unknown" = ($currentDevices | Where-Object { $_.RowColour -eq "unknown" }).Count
}

# Generate Stat Cards
$orgexpiredWarranties = New-StatCard `
    -Value $WarrantyCounts."Expired" `
    -Description "Expired Warranties" `
    -Color "#FF4500" `
    -IconClass "fas fa-times-circle" `
    -Id "expiredWarranties"

$orgexpiringWarranties = New-StatCard `
    -Value $WarrantyCounts."Expiring" `
    -Description "Expiring Warranties" `
    -Color "#FFD700" `
    -IconClass "fas fa-exclamation-triangle" `
    -Id "expiringWarranties"

$orgvalidWarranties = New-StatCard `
    -Value $WarrantyCounts."Valid" `
    -Description "Under Warranty" `
    -Color "#008001" `
    -IconClass "fas fa-check-circle" `
    -Id "validWarranties"

# Combine Stat Cards into a Table Row
$orgWarrantyTable = @"
<table style='border: 0px;'>
    <tbody>
        <tr>
            $orgexpiredWarranties
            $orgexpiringWarranties
            $orgvalidWarranties
        </tr>
    </tbody>
</table>
"@
            # Ensure $currentDevices is an array to safely access .Count
            $currentDevices = @($currentDevices)
            
            # Check if there are no devices in the current organization
            if ($currentDevices.Count -le 0) {
                Write-Host "No devices found for organization '$($organization.name)'. Skipping to next organization."
                continue  # Skip to the next iteration of the loop
            }

            # Extract unique device IDs
            $currentDeviceIds = $currentDevices | Select-Object -ExpandProperty 'Device ID' -Unique

            # Prepare the data for the table
            $tableDevices = $currentDevices | 
                Select-Object "Device Name",
                            "Warranty Status",
                            "Warranty Start", 
                            "Warranty End",
                            "Location Name",
                            "Device ID",
                            RowColour
                            

            # Validate $tableDevices
            if ($null -eq $tableDevices -or $tableDevices.Count -eq 0) {
                Write-Host "No data available to generate HTML table for organization '$($organization.name)'. Skipping."
                continue  # Skip to the next organization
            }

            # Generate HTML table
            $orghtmltable = ConvertTo-ObjectToHtmlTable -Objects $tableDevices -NinjaOneInstance $NinjaOneInstance

                if ($CreateDocument) {
                    # Filter the documents
                    $MatchedDoc = $WarrantyReportDocs | Where-Object { 
                        $_.organizationId -eq $organization.id -and
                        $_.documentName -match $currentMonth -and 
                        $_.documentName -match $currentYear -and 
                        $_.documentName -match "Warranty Report"
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
                        'warrantyStats'      = @{'html' =   $orgWarrantyTable }
                        'warrantyStatus'  = @{'html' = $orghtmltable }
                    }
            
                    if ($MatchedDoc) {
                        $UpdateObject = [PSCustomObject]@{
                            documentId   = $MatchedDoc.documentId
                            documentName = "$($organization.name) Warranty Report - $currentMonth $currentYear"
                            fields       = $DocFields
                        }
            
                        $NinjaDocUpdates.Add($UpdateObject)
            
                    } else {
                        $CreateObject = [PSCustomObject]@{
                            documentName       = "$($organization.name) Warranty Report - $currentMonth $currentYear"
                            documentTemplateId = $DocTemplate.id
                            organizationId     = [int]$organization.id
                            fields             = $DocFields
                        }
            
                        $NinjaDocCreation.Add($CreateObject)
                    }
                }
                    if ($CreateKB) {
                    $MatchingName = "$($organization.name) Warranty Report - $currentMonth $currentYear"
    $orgcombinedHtml = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>$($organization.name) Warranty Report</title>
    <!-- Include any required CSS here -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        /* Example styles, customize as needed */
        .stat-card { /* Your styles here */ }
        .info-card { /* Your styles here */ }
        /* Add additional CSS to ensure proper styling */
    </style>
</head>
<body>
    <div class="container-fluid">
        $orgWarrantyTable
        <hr>
        $orghtmltable
    </div>
</body>
</html>
"@

                    # Fetch existing Patch Report KB Articles
                    $WarrantyReportKBs = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/organization/articles' -QueryParams "articleName=$($MatchingName)"
                    
                    $KBMatchCount = ($WarrantyReportKBs | Measure-Object).Count
                    Write-Output "Found $($organization.name) had $($KBMatchCount) match"
                    $KBFields = @{
                        'html' = $orgcombinedHtml
                        }   

                    if ($KBMatchCount -eq 0) {
                        Write-Host "No match found for $($MatchingName)"
                    } elseif ($KBMatchCount -gt 1) {
                        Throw "Multiple NinjaOne KBs with ($($WarrantyReportKBs.id -join ',')) matched to $($MatchingName)"
                        continue
                    } else {
                        Write-Host "Matched document ID: $($WarrantyReportKBs.id) for $($MatchingName)"
                    }

                    if ($WarrantyReportKBs) {
                        $UpdateObject = [PSCustomObject]@{
                            id   = $WarrantyReportKBs.Id
                            organizationId = $($organization.id)
                            destinationFolderPath = "$($organization.name) Monthly Warranty Reports"
                            name = "$($organization.name) Warranty Report - $currentMonth $currentYear"
                            content       = $KBFields
                        }

                        $NinjaKBUpdates.Add($UpdateObject)

                    } else {
                        $CreateObject = [PSCustomObject]@{
                            name       = "$($organization.name) Warranty Report - $currentMonth $currentYear"
                            organizationId = $($organization.id)
                            destinationFolderPath = "$($organization.name) Monthly Warranty Reports"
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

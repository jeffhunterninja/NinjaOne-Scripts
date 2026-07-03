#Requires -Version 5.1

<#
.SYNOPSIS
    Retrieve the system's last boot time (Win32_OperatingSystem.LastBootUpTime) and optionally save it to a custom field.
.DESCRIPTION
    Retrieve the system's last boot time (Win32_OperatingSystem.LastBootUpTime) and optionally save it to a custom field.
.EXAMPLE
    (No Parameters)

    Retrieving the last boot time from Win32_OperatingSystem.

    ### Last Boot Time ###
    12/18/2024 11:20 AM

PARAMETER: -CustomField "ExampleInput"
    Optionally save the last boot time to a custom field of your choosing.

.NOTES
    Minimum OS Architecture Supported: Windows 10, Windows Server 2016
    Version: 1.0
    Release Notes: Initial Release
#>

[CmdletBinding()]
param (
    [Parameter()]
    [String]$CustomField
)

begin {
    # If script form variables are used, replace the command line parameters with their value.
    if ($env:lastBootTimeCustomField -and $env:lastBootTimeCustomField -notlike "null") { $CustomField = $env:lastBootTimeCustomField }

    # Check if a custom field value was provided.
    if ($CustomField) {
        # Remove any leading or trailing whitespace.
        $CustomField = $CustomField.Trim()

        # If, after trimming, the custom field is empty, print an error and exit.
        if (!$CustomField) {
            Write-Host -Object "[Error] Please enter a valid custom field."
            exit 1
        }
    }

    function Set-NinjaProperty {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $True)]
            [String]$Name,
            [Parameter()]
            [String]$Type,
            [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
            $Value,
            [Parameter()]
            [String]$DocumentName,
            [Parameter()]
            [Switch]$Piped
        )
        # Remove the non-breaking space character
        if ($Type -eq "WYSIWYG") {
            $Value = $Value -replace ' ', '&nbsp;'
        }
        
        # Measure the number of characters in the provided value
        $Characters = $Value | ConvertTo-Json | Measure-Object -Character | Select-Object -ExpandProperty Characters
    
        # Throw an error if the value exceeds the character limit of 200,000 characters
        if ($Piped -and $Characters -ge 200000) {
            throw [System.ArgumentOutOfRangeException]::New("Character limit exceeded: the value is greater than or equal to 200,000 characters.")
        }
    
        if (!$Piped -and $Characters -ge 45000) {
            throw [System.ArgumentOutOfRangeException]::New("Character limit exceeded: the value is greater than or equal to 45,000 characters.")
        }
        
        # Initialize a hashtable for additional documentation parameters
        $DocumentationParams = @{}
    
        # If a document name is provided, add it to the documentation parameters
        if ($DocumentName) { $DocumentationParams["DocumentName"] = $DocumentName }
        
        # Define a list of valid field types
        $ValidFields = "Attachment", "Checkbox", "Date", "Date or Date Time", "Decimal", "Dropdown", "Email", "Integer", "IP Address", "MultiLine", "MultiSelect", "Phone", "Secure", "Text", "Time", "URL", "WYSIWYG"
    
        # Warn the user if the provided type is not valid
        if ($Type -and $ValidFields -notcontains $Type) { Write-Warning "$Type is an invalid type. Please check here for valid types: https://ninjarmm.zendesk.com/hc/en-us/articles/16973443979789-Command-Line-Interface-CLI-Supported-Fields-and-Functionality" }
        
        # Define types that require options to be retrieved
        $NeedsOptions = "Dropdown"
    
        # If the property is being set in a document or field and the type needs options, retrieve them
        if ($DocumentName) {
            if ($NeedsOptions -contains $Type) {
                $NinjaPropertyOptions = Ninja-Property-Docs-Options -AttributeName $Name @DocumentationParams 2>&1
            }
        }
        else {
            if ($NeedsOptions -contains $Type) {
                $NinjaPropertyOptions = Ninja-Property-Options -Name $Name 2>&1
            }
        }
        
        # Throw an error if there was an issue retrieving the property options
        if ($NinjaPropertyOptions.Exception) { throw $NinjaPropertyOptions }
            
        # Process the property value based on its type
        switch ($Type) {
            "Checkbox" {
                # Convert the value to a boolean for Checkbox type
                $NinjaValue = [System.Convert]::ToBoolean($Value)
            }
            "Date or Date Time" {
                # Convert the value to a Unix timestamp for Date or Date Time type
                $Date = (Get-Date $Value).ToUniversalTime()
                $TimeSpan = New-TimeSpan (Get-Date "1970-01-01 00:00:00") $Date
                $NinjaValue = $TimeSpan.TotalSeconds
            }
            "Dropdown" {
                # Convert the dropdown value to its corresponding GUID
                $Options = $NinjaPropertyOptions -replace '=', ',' | ConvertFrom-Csv -Header "GUID", "Name"
                $Selection = $Options | Where-Object { $_.Name -eq $Value } | Select-Object -ExpandProperty GUID
            
                # Throw an error if the value is not present in the dropdown options
                if (!($Selection)) {
                    throw [System.ArgumentOutOfRangeException]::New("Value is not present in dropdown options.")
                }
            
                $NinjaValue = $Selection
            }
            default {
                # For other types, use the value as is
                $NinjaValue = $Value
            }
        }
            
        # Set the property value in the document if a document name is provided
        if ($DocumentName) {
            $CustomFieldResult = Ninja-Property-Docs-Set -AttributeName $Name -AttributeValue $NinjaValue @DocumentationParams 2>&1
        }
        else {
            try {
                # Otherwise, set the standard property value
                if ($Piped) {
                    $CustomFieldResult = $NinjaValue | Ninja-Property-Set-Piped -Name $Name 2>&1
                }
                else {
                    $CustomFieldResult = Ninja-Property-Set -Name $Name -Value $NinjaValue 2>&1
                }
            }
            catch {
                Write-Host -Object "[Error] Failed to set custom field."
                throw $_.Exception.Message
            }
        }
            
        # Throw an error if setting the property failed
        if ($CustomFieldResult.Exception) {
            throw $CustomFieldResult
        }
    }

    function Test-IsElevated {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $p = New-Object System.Security.Principal.WindowsPrincipal($id)
        $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    }

    if (!$ExitCode) {
        $ExitCode = 0
    }
}
process {
    # Check if the current user is elevated (running as Administrator).
    if (!(Test-IsElevated)) {
        Write-Host -Object "[Error] Access Denied. Please run with Administrator privileges."
        exit 1
    }

    # Inform the user that the script is retrieving the last boot time.
    Write-Host -Object "Retrieving the last boot time from Win32_OperatingSystem."

    try {
        # Query Win32_OperatingSystem for the LastBootUpTime property.
        $OperatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    }
    catch {
        Write-Host -Object "[Error] $($_.Exception.Message)"
        Write-Host -Object "[Error] Failed to retrieve operating system information."
        exit 1
    }

    # LastBootUpTime is already returned as a [DateTime] object by Get-CimInstance.
    $LastBootUpTime = $OperatingSystem.LastBootUpTime

    if (!$LastBootUpTime) {
        Write-Host -Object "[Error] Failed to retrieve the last boot time."
        exit 1
    }

    # Format the boot time for a friendly, readable output.
    $FormattedLastBootUpTime = "$($LastBootUpTime.ToShortDateString()) $($LastBootUpTime.ToShortTimeString())"

    # If a custom field is specified, set it to the formatted last boot time.
    if ($CustomField) {
        Write-Host -Object ""
        try {
            Write-Host "Attempting to set the Custom Field '$CustomField'."
            Set-NinjaProperty -Name $CustomField -Value $FormattedLastBootUpTime
            Write-Host "Successfully set the Custom Field '$CustomField'!"
        }
        catch {
            Write-Host "[Error] $($_.Exception.Message)"
            $ExitCode = 1
        }
    }

    # Print the result for reference.
    Write-Host -Object "`n### Last Boot Time ###"
    Write-Host -Object $FormattedLastBootUpTime

    # Exit with the previously set exit code (defaulting to 0 if not set).
    exit $ExitCode
}
end {
    
    
    
}

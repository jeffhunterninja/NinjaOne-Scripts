#Requires -Version 5.1
<#
.SYNOPSIS
    Retrieves the Windows device timezone and writes it to a NinjaOne device custom field.

.DESCRIPTION
    Gets the current system timezone using Get-TimeZone (PowerShell 5.1+) and writes the
    timezone ID (e.g. "Eastern Standard Time") to a NinjaOne device custom field via
    Ninja-Property-Set when the script runs under the NinjaOne agent. The timezone ID is
    consistent and machine-parseable. You must create the device custom field in NinjaOne
    (Devices -> Custom Fields) before running this script; a TEXT field is sufficient.

.PARAMETER CustomFieldName
    Name of the NinjaOne device custom field to write the timezone to. Default: timezone.
    Can be overridden by environment variable timezoneCustomField (e.g. from NinjaOne script variables).

.PARAMETER NoNinjaWrite
    If set, only outputs the timezone to the host/pipeline; does not call Ninja-Property-Set.
    Use for local testing when not running under NinjaOne.

.EXIT CODES
    0 = Success
    1 = Error (e.g. failure to get timezone or Ninja-Property-Set failed when not using -NoNinjaWrite)

.EXAMPLE
    .\Set-DeviceTimezoneCustomField.ps1
    Run on device via NinjaOne; writes timezone ID to custom field "timezone".

.EXAMPLE
    .\Set-DeviceTimezoneCustomField.ps1 -CustomFieldName deviceTimezone
    Writes to custom field "deviceTimezone".

.EXAMPLE
    .\Set-DeviceTimezoneCustomField.ps1 -NoNinjaWrite
    Detects and displays timezone only; does not write to NinjaOne.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CustomFieldName = 'timezone',

    [Parameter(Mandatory = $false)]
    [switch]$NoNinjaWrite
)

$ErrorActionPreference = 'Stop'

# Allow env override for custom field name (e.g. from NinjaOne script variables)
if (-not [string]::IsNullOrWhiteSpace($env:timezoneCustomField)) {
    $CustomFieldName = $env:timezoneCustomField.Trim()
}

try {
    $tz = Get-TimeZone
    $timezoneId = $tz.Id
    $displayName = $tz.DisplayName

    Write-Host "Timezone ID: $timezoneId"
    Write-Host "Display name: $displayName"

    if (-not $NoNinjaWrite) {
        if (Get-Command Ninja-Property-Set -ErrorAction SilentlyContinue) {
            try {
                Ninja-Property-Set $CustomFieldName $timezoneId
                Write-Host "Written to NinjaOne custom field '$CustomFieldName'."
            } catch {
                Write-Warning "Ninja-Property-Set failed: $($_.Exception.Message)"
                exit 1
            }
        } else {
            Write-Warning "Ninja-Property-Set is not available. Run this script from NinjaOne on a managed device, or use -NoNinjaWrite for local testing."
            exit 1
        }
    }

    if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('NoNinjaWrite')) {
        [PSCustomObject]@{
            TimezoneId   = $timezoneId
            DisplayName  = $displayName
        } | Write-Output
    }

    exit 0
} catch {
    Write-Host "An error occurred: $_"
    exit 1
}

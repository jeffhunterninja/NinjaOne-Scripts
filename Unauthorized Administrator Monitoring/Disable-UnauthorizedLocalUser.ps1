#Requires -Version 5.1
<#
.SYNOPSIS
    Disables local user accounts listed in the NinjaOne unauthorized admins custom field.

.DESCRIPTION
    Reads unauthorizedLocalAdmins (or unauthorizedAdmins) from Ninja-Property-Get and disables
    each listed account with Disable-LocalUser. Accepts newline- or comma-separated values.
    Use only after confirming the listed users are actually unauthorized. Test in a lab first.

.EXIT CODES
    0 = Completed (individual failures are warned, not fatal)
    1 = Could not read Ninja property or critical failure

.PARAMETER NinjaProperty
    Ninja custom field name to read. Default: unauthorizedLocalAdmins (fallback: unauthorizedAdmins).
#>
[CmdletBinding()]
param(
    [string]$NinjaProperty = 'unauthorizedLocalAdmins'
)

$ErrorActionPreference = 'Stop'

# Prefer newline-separated; fallback to unauthorizedAdmins (often comma-separated)
$raw = Ninja-Property-Get $NinjaProperty -ErrorAction SilentlyContinue
if ([string]::IsNullOrWhiteSpace($raw)) {
    $raw = Ninja-Property-Get 'unauthorizedAdmins' -ErrorAction SilentlyContinue
}
if ([string]::IsNullOrWhiteSpace($raw) -or $raw -match '^No unauthorized|^Error:') {
    Write-Host "No unauthorized admins list to process or list is empty/error."
    exit 0
}

$badusers = $raw -split "[\r\n,]+" | Where-Object { $_.Trim() -ne '' }
foreach ($baduser in $badusers) {
    $name = $baduser.Trim()
    try {
        Disable-LocalUser -Name $name -ErrorAction Stop
        Write-Host "Disabled local user: $name"
    } catch {
        Write-Warning "Failed to disable '$name': $_"
    }
}
exit 0

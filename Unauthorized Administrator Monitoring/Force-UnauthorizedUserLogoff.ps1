#Requires -Version 5.1
<#
.SYNOPSIS
    Forces logoff of user sessions for accounts listed in the NinjaOne unauthorized admins custom field.

.DESCRIPTION
    Reads unauthorizedLocalAdmins (or unauthorizedAdmins) from Ninja-Property-Get, finds each user's
    session with query user, and runs logoff for that session. Accepts newline- or comma-separated
    values. Use only after confirming the listed users are actually unauthorized. Test in a lab first.

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
        $userSession = query user | Where-Object { $_ -match [regex]::Escape($name) }
        if ($userSession) {
            $sessionId = $userSession.Trim().Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)[2]
            if ($sessionId -match '^\d+$') {
                logoff $sessionId
                Write-Host "Successfully logged off user: $name"
            } else {
                Write-Warning "Could not parse session ID for user: $name"
            }
        } else {
            Write-Warning "User $name is not logged in."
        }
    } catch {
        Write-Warning "Failed to log off user $name. Error: $_"
    }
}
exit 0

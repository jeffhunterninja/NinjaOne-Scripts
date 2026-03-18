#Requires -Version 5.1
<#
.SYNOPSIS
    Compares local administrators to an allowed list and writes results to NinjaOne custom fields.

.DESCRIPTION
    Enumerates local administrators via ADSI (Administrators group SID S-1-5-32-544). The allowed list
    is built from: built-in "Administrator" and "Domain Admins", Ninja org docs (User Management / Admins),
    org property authorizedAdminsOrg, and device custom field authorizedAdmins. Writes the current
    local admin list to LocalAdmins and any violations to unauthorizedAdmins / unauthorizedLocalAdmins.
    Comparison is case-insensitive; allowed entries may be "DOMAIN\user" or "user".

.EXIT CODES
    0 = No unauthorized admins found
    1 = One or more unauthorized admins found
    2 = Script error (e.g. ADSI failure, Ninja property error)

.PARAMETER NoNinjaWrite
    If set, only outputs results to the pipeline; does not call Ninja-Property-Set. Use for testing without Ninja.

.EXAMPLE
    .\Test-UnauthorizedAdministrators.ps1

.EXAMPLE
    .\Test-UnauthorizedAdministrators.ps1 -NoNinjaWrite
#>
[CmdletBinding()]
param(
    [switch]$NoNinjaWrite
)

$ErrorActionPreference = 'Stop'

try {
    $groupSID = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-544")
    $groupName = $groupSID.Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
    $localAdminsPaths = ([ADSI]"WinNT://$env:COMPUTERNAME/$groupName").PSBase.Invoke('Members') | ForEach-Object { $_.GetType().InvokeMember('AdsPath', 'GetProperty', $null, $_, $null) }

    $adminList = @()
    $allowedAdmins = @()

    # Well-known allowed entries (once, before loop)
    $allowedAdmins += "Administrator"
    $allowedAdmins += "Domain Admins"

    # Organization-level allowed admins (Ninja docs)
    try {
        $ninjaOrgAdminsRaw = Ninja-Property-Docs-Get 'User Management' 'Admins' authorizedOrgAdmins
        if (![string]::IsNullOrWhiteSpace($ninjaOrgAdminsRaw)) {
            $allowedAdmins += ($ninjaOrgAdminsRaw -split "`r`n") | Where-Object { $_.Trim() -ne '' }
        }
    } catch {
        # Org docs may not exist
    }

    # Global allowed list (Ninja org property)
    $globalAdminsRaw = Ninja-Property-Get authorizedAdminsOrg -ErrorAction SilentlyContinue
    if (![string]::IsNullOrWhiteSpace($globalAdminsRaw)) {
        $allowedAdmins += ($globalAdminsRaw -split ",\s*") | Where-Object { $_.Trim() -ne '' }
    }

    # Device-specific allowed admins
    $deviceAdminsRaw = Ninja-Property-Get authorizedAdmins -ErrorAction SilentlyContinue
    if (![string]::IsNullOrWhiteSpace($deviceAdminsRaw)) {
        $allowedAdmins += ($deviceAdminsRaw -split ",\s*|`r`n") | Where-Object { $_.Trim() -ne '' }
    }

    $allowedAdmins = $allowedAdmins | Sort-Object -Unique
    # Case-insensitive lookup: normalize to lowercase; match by full name or bare account name
    $allowedSet = @{}
    foreach ($a in $allowedAdmins) {
        $key = $a.Trim().ToLowerInvariant()
        if ($key -ne '') { $allowedSet[$key] = $true }
        $bare = $a.Trim().Split('\')[-1].ToLowerInvariant()
        if ($bare -ne '' -and -not $allowedSet.ContainsKey($bare)) { $allowedSet[$bare] = $true }
    }

    $unauthorizedList = @()

    foreach ($path in $localAdminsPaths) {
        $parts = $path.Split('/', [StringSplitOptions]::RemoveEmptyEntries)
        $name = $parts[-1]
        $domain = $parts[-2]
        if ($domain -eq "WinNT:") {
            $fullName = $name
        } else {
            $fullName = "$domain\$name"
        }
        $adminList += $fullName

        $fullLower = $fullName.ToLowerInvariant()
        $bareLower = $name.ToLowerInvariant()
        $isAllowed = $allowedSet.ContainsKey($fullLower) -or $allowedSet.ContainsKey($bareLower)
        if (-not $isAllowed) {
            $unauthorizedList += $fullName
        }
    }

    $adminList = $adminList | Sort-Object -Unique

    if (-not $NoNinjaWrite) {
        Ninja-Property-Set LocalAdmins ($adminList -join "`r`n")
    }

    if ($unauthorizedList.Count -gt 0) {
        Write-Host "Unauthorized admins found: $($unauthorizedList -join ', ')"
        if (-not $NoNinjaWrite) {
            Ninja-Property-Set unauthorizedAdmins ($unauthorizedList -join ", ")
            Ninja-Property-Set unauthorizedLocalAdmins ($unauthorizedList -join "`r`n")
        }
        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('NoNinjaWrite')) {
            [PSCustomObject]@{ LocalAdmins = $adminList; UnauthorizedAdmins = $unauthorizedList; ExitCode = 1 } | Write-Output
        }
        exit 1
    } else {
        Write-Host "No unauthorized admins found."
        if (-not $NoNinjaWrite) {
            Ninja-Property-Set unauthorizedAdmins "No unauthorized admins found!"
            Ninja-Property-Set unauthorizedLocalAdmins "No unusual local admins found!"
        }
        if ($PSCmdlet.MyInvocation.BoundParameters.ContainsKey('NoNinjaWrite')) {
            [PSCustomObject]@{ LocalAdmins = $adminList; UnauthorizedAdmins = @(); ExitCode = 0 } | Write-Output
        }
        exit 0
    }
}
catch {
    Write-Host "An error occurred: $_"
    if (-not $NoNinjaWrite) {
        try { Ninja-Property-Set unauthorizedAdmins "Error: $($_.Exception.Message)" } catch { }
    }
    exit 2
}

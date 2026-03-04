<#
.SYNOPSIS
    Removes Wi-Fi profiles by policy: blocklisted SSIDs (Blocklist mode) or the connected SSID when not in the allowlist (Allowlist mode).

.DESCRIPTION
    Reads the same NinjaOne custom fields as the Guest Wifi Check (wifiCheckMode, blocklistedWifiNetworks/guestWifiNetwork, allowedWifiNetworks).
    Blocklist mode: deletes each saved Wi-Fi profile whose SSID is in the blocklist.
    Allowlist mode: if the device is currently connected to an SSID not in the allowlist, deletes that profile only.
    All logic is in-line; no dot-sourcing. Deleting Wi-Fi profiles typically requires elevated rights (admin or SYSTEM).

.PARAMETER NoNinjaWrite
    Reserved for testing; this script does not write to NinjaOne. Mode and lists are still read from NinjaOne when available.

.PARAMETER Mode
    Override mode: Blocklist or Allowlist. If not set, script reads wifiCheckMode from NinjaOne (default Blocklist).

.PARAMETER WhatIf
    Report which profile(s) would be deleted without running netsh wlan delete profile.

.EXAMPLE
    .\Remove-WifiProfilesByPolicy.ps1
.EXAMPLE
    .\Remove-WifiProfilesByPolicy.ps1 -WhatIf -Mode Blocklist
.NOTES
    Exit codes: 0 = success (profiles removed or nothing to do); 2 = error (WLAN unavailable, netsh failure, or exception).
    Run with sufficient privileges to delete profiles (e.g. as SYSTEM or elevated admin when scheduled in NinjaOne).
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$NoNinjaWrite,
    [ValidateSet('Blocklist', 'Allowlist')]
    [string]$Mode = '',
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# Get value from Ninja custom field; return $null on missing/failure
function Get-NinjaProperty {
    param([string]$Name)
    try {
        if (Get-Command Ninja-Property-Get -ErrorAction SilentlyContinue) {
            $v = Ninja-Property-Get $Name
            if ([string]::IsNullOrWhiteSpace($v)) { return $null }
            return $v.Trim()
        }
    } catch { }
    return $null
}

# Normalize SSID list for comparison: trim and lowercase
function Get-NormalizedSsidList {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }
    return $Raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' } | ForEach-Object { $_.ToLowerInvariant() }
}

# Return list of saved profile names (exact display names from netsh wlan show profiles)
function Get-SavedProfileNames {
    param([string]$ProfilesRaw)
    $names = @()
    $matches = [regex]::Matches($ProfilesRaw, '(?im)All User Profile\s*:\s*(.+)$')
    foreach ($m in $matches) {
        $name = $m.Groups[1].Value.Trim()
        if ($name -and $name -ne 'SSID') { $names += $name }
    }
    return $names | Select-Object -Unique
}

# Return list of currently connected SSIDs (from netsh wlan show interfaces)
function Get-ConnectedSsidList {
    $ifRaw = netsh.exe wlan show interfaces 2>&1 | Out-String
    if ($ifRaw -match 'The Wireless AutoConfig Service \(wlansvc\) is not running\.') { return @() }
    if ($ifRaw -match 'There is no wireless interface on the system\.') { return @() }
    $connected = @()
    $blocks = $ifRaw -split '(?m)^\s*$'
    foreach ($block in $blocks) {
        if ($block -match '(?im)^\s*State\s*:\s*connected\b') {
            $m = [regex]::Match($block, '(?im)^\s*SSID\s*:\s*(.+)$')
            if ($m.Success) {
                $ssid = $m.Groups[1].Value.Trim()
                if ($ssid -and $ssid -ne 'SSID') { $connected += $ssid }
            }
        }
    }
    return $connected | Select-Object -Unique
}

# Delete a single Wi-Fi profile by name. Escapes double quotes in profile name for netsh.
function Remove-WifiProfileByName {
    param([string]$ProfileName, [bool]$DryRun)
    if ([string]::IsNullOrWhiteSpace($ProfileName)) { return $false }
    if ($DryRun) {
        Write-Host "WhatIf: would delete profile: $ProfileName"
        return $true
    }
    $escaped = $ProfileName -replace '"', '""'
    $nameArg = "name=`"$escaped`""
    $result = & netsh.exe wlan delete profile $nameArg 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Failed to delete profile '$ProfileName': $result"
        return $false
    }
    Write-Host "Removed profile: $ProfileName"
    return $true
}

try {
    $wifiRaw = netsh.exe wlan show profiles 2>&1 | Out-String

    # WLAN unavailable
    if ($wifiRaw -match 'The Wireless AutoConfig Service \(wlansvc\) is not running\.') {
        Write-Host 'The Wireless AutoConfig Service (wlansvc) is not running. No profiles removed.'
        exit 0
    }
    if ($wifiRaw -match 'There is no wireless interface on the system\.') {
        Write-Host 'There is no wireless interface on the system. No profiles removed.'
        exit 0
    }

    # Resolve mode: param override else Ninja else Blocklist
    $modeResolved = $Mode
    if ([string]::IsNullOrWhiteSpace($modeResolved)) {
        $modeResolved = Get-NinjaProperty 'wifiCheckMode'
        if ([string]::IsNullOrWhiteSpace($modeResolved)) { $modeResolved = 'Blocklist' }
        $modeResolved = $modeResolved.Trim()
        if ($modeResolved -notin 'Blocklist', 'Allowlist') { $modeResolved = 'Blocklist' }
    }

    $dryRun = $WhatIf

    if ($modeResolved -eq 'Blocklist') {
        $blocklistRaw = Get-NinjaProperty 'blocklistedWifiNetworks'
        if ([string]::IsNullOrWhiteSpace($blocklistRaw)) { $blocklistRaw = Get-NinjaProperty 'guestWifiNetwork' }
        $blocklist = Get-NormalizedSsidList $blocklistRaw
        if ($blocklist.Count -eq 0) {
            Write-Host 'Blocklist mode: no blocklisted SSIDs configured. Nothing to remove.'
            exit 0
        }

        $savedNames = Get-SavedProfileNames $wifiRaw
        $removed = 0
        foreach ($profileName in $savedNames) {
            $normalized = $profileName.ToLowerInvariant()
            if ($blocklist -contains $normalized) {
                if (Remove-WifiProfileByName -ProfileName $profileName -DryRun $dryRun) { $removed++ }
            }
        }
        if ($dryRun -and $removed -eq 0) {
            Write-Host 'WhatIf: no saved profiles match the blocklist.'
        }
        exit 0
    }

    # Allowlist mode: only consider currently connected SSID
    $allowlistRaw = Get-NinjaProperty 'allowedWifiNetworks'
    $allowlist = Get-NormalizedSsidList $allowlistRaw
    if ($allowlist.Count -eq 0) {
        Write-Host 'Allowlist mode: no allowed SSIDs configured. Nothing to remove.'
        exit 0
    }

    $connectedList = Get-ConnectedSsidList
    if ($connectedList.Count -eq 0) {
        Write-Host 'Allowlist mode: not connected to Wi-Fi. Nothing to remove.'
        exit 0
    }

    $removedAny = $false
    foreach ($connectedSsid in $connectedList) {
        $normalized = $connectedSsid.ToLowerInvariant()
        if ($allowlist -notcontains $normalized) {
            if (Remove-WifiProfileByName -ProfileName $connectedSsid -DryRun $dryRun) { $removedAny = $true }
        }
    }
    if (-not $removedAny -and -not $dryRun) {
        Write-Host 'Allowlist mode: current connection is in the allowlist. Nothing to remove.'
    }
    exit 0
} catch {
    Write-Warning $_.Exception.Message
    exit 2
}

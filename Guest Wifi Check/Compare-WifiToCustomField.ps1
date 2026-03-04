<#
.SYNOPSIS
    Checks Wi-Fi against a blocklist (alert when blocklisted networks appear) or allowlist (alert on deviation).

.DESCRIPTION
    Supports two modes via NinjaOne custom field wifiCheckMode: Blocklist or Allowlist.
    Blocklist: alerts when the device has saved or is connected to any network in the blocklist (e.g. guest Wi-Fi).
    Allowlist: alerts when any saved or connected SSID is not in the allowed list (deviation).
    Reads blocklist from guestWifiNetwork or blocklistedWifiNetworks; allowlist from allowedWifiNetworks.
    Scope (wifiCheckScope): SavedOnly, ConnectedOnly, or Both. Writes wifinetworks, currentWifiNetwork, and optional wifiCheckStatus.

.PARAMETER NoNinjaWrite
    When set, does not call Ninja-Property-Set. Use for local/testing without writing to NinjaOne.

.PARAMETER Mode
    Override mode: Blocklist or Allowlist. If not set, script reads wifiCheckMode from NinjaOne (default Blocklist).

.PARAMETER Scope
    Override scope: SavedOnly, ConnectedOnly, or Both. If not set, script reads wifiCheckScope from NinjaOne (default Both).

.EXAMPLE
    .\10 - Check for blacklisted wifi network.ps1
.EXAMPLE
    .\10 - Check for blacklisted wifi network.ps1 -NoNinjaWrite -Mode Blocklist
.NOTES
    Exit codes: 0 = OK; 1 = blocklisted found or deviation; 2 = error.
#>
[CmdletBinding()]
param(
    [switch]$NoNinjaWrite,
    [ValidateSet('Blocklist', 'Allowlist')]
    [string]$Mode = '',
    [ValidateSet('SavedOnly', 'ConnectedOnly', 'Both')]
    [string]$Scope = ''
)

$ErrorActionPreference = 'Stop'

# Normalize SSID for comparison: trim and lowercase
function Get-NormalizedSsidList {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }
    return $Raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' } | ForEach-Object { $_.ToLowerInvariant() }
}

# Safe write to Ninja custom field
function Set-NinjaProperty {
    param([string]$Name, [string]$Value)
    if ($NoNinjaWrite) { return }
    try {
        if (Get-Command Ninja-Property-Set -ErrorAction SilentlyContinue) {
            Ninja-Property-Set $Name $Value
        }
    } catch {
        Write-Warning "Ninja-Property-Set failed ($Name): $($_.Exception.Message)"
    }
}

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

# Write current connected SSID(s) to Ninja and return the list
function Update-CurrentWifiNetworkField {
    param([string]$CustomFieldName = 'currentWifiNetwork')
    $ifRaw = netsh.exe wlan show interfaces 2>&1 | Out-String
    if ($ifRaw -match 'The Wireless AutoConfig Service \(wlansvc\) is not running\.') {
        Set-NinjaProperty $CustomFieldName 'wlansvc is not running'
        return @()
    }
    if ($ifRaw -match 'There is no wireless interface on the system\.') {
        Set-NinjaProperty $CustomFieldName 'N/A'
        return @()
    }
    $connected = Get-ConnectedSsidList
    if ($connected.Count -gt 0) {
        Set-NinjaProperty $CustomFieldName ($connected -join ', ')
    } else {
        Set-NinjaProperty $CustomFieldName 'Disconnected'
    }
    return $connected
}

try {
    $wifiRaw = netsh.exe wlan show profiles 2>&1 | Out-String

    # Resolve mode: param override else Ninja else Blocklist
    $modeResolved = $Mode
    if ([string]::IsNullOrWhiteSpace($modeResolved)) {
        $modeResolved = Get-NinjaProperty 'wifiCheckMode'
        if ([string]::IsNullOrWhiteSpace($modeResolved)) { $modeResolved = 'Blocklist' }
        $modeResolved = $modeResolved.Trim()
        if ($modeResolved -notin 'Blocklist','Allowlist') { $modeResolved = 'Blocklist' }
    }

    # Resolve scope: param override else Ninja else Both
    $scopeResolved = $Scope
    if ([string]::IsNullOrWhiteSpace($scopeResolved)) {
        $scopeResolved = Get-NinjaProperty 'wifiCheckScope'
        if ([string]::IsNullOrWhiteSpace($scopeResolved)) { $scopeResolved = 'Both' }
        $scopeResolved = $scopeResolved.Trim()
        if ($scopeResolved -notin 'SavedOnly','ConnectedOnly','Both') { $scopeResolved = 'Both' }
    }

    # Disabled WLAN autoconfig or no WLAN interface
    if ($wifiRaw -match 'The Wireless AutoConfig Service \(wlansvc\) is not running\.') {
        Set-NinjaProperty 'wifinetworks' 'wlansvc is not running'
        $connectedList = Update-CurrentWifiNetworkField
        Set-NinjaProperty 'wifiCheckStatus' 'OK'
        exit 0
    }
    if ($wifiRaw -match 'There is no wireless interface on the system\.') {
        Set-NinjaProperty 'wifinetworks' 'N/A'
        $connectedList = Update-CurrentWifiNetworkField
        Set-NinjaProperty 'wifiCheckStatus' 'OK'
        exit 0
    }

    # Parse saved SSIDs
    $savedSsids = @()
    $matches = [regex]::Matches($wifiRaw, '(?im)All User Profile\s*:\s*(.+)$')
    foreach ($m in $matches) {
        $ssid = $m.Groups[1].Value.Trim()
        if ($ssid -and $ssid -ne 'SSID') { $savedSsids += $ssid }
    }
    $savedSsids = $savedSsids | Select-Object -Unique

    Set-NinjaProperty 'wifinetworks' ($savedSsids -join "`n")
    $connectedList = Update-CurrentWifiNetworkField

    # SSIDs to evaluate based on scope (normalized to lowercase for comparison)
    $toCheck = @()
    if ($scopeResolved -in 'SavedOnly','Both') {
        $toCheck += $savedSsids | ForEach-Object { $_.ToLowerInvariant() }
    }
    if ($scopeResolved -in 'ConnectedOnly','Both') {
        $toCheck += $connectedList | ForEach-Object { $_.ToLowerInvariant() }
    }
    $toCheck = $toCheck | Select-Object -Unique

    if ($modeResolved -eq 'Blocklist') {
        $blocklistRaw = Get-NinjaProperty 'blocklistedWifiNetworks'
        if ([string]::IsNullOrWhiteSpace($blocklistRaw)) { $blocklistRaw = Get-NinjaProperty 'guestWifiNetwork' }
        $blocklist = Get-NormalizedSsidList $blocklistRaw
        if ($blocklist.Count -eq 0) {
            Set-NinjaProperty 'wifiCheckStatus' 'OK'
            exit 0
        }
        foreach ($ssid in $toCheck) {
            if ($blocklist -contains $ssid) {
                Set-NinjaProperty 'wifiCheckStatus' 'Blocklisted'
                Write-Host 'Blocklisted Wi-Fi network detected.'
                exit 1
            }
        }
        Set-NinjaProperty 'wifiCheckStatus' 'OK'
        exit 0
    }

    # Allowlist (deviation) mode
    $allowlistRaw = Get-NinjaProperty 'allowedWifiNetworks'
    $allowlist = Get-NormalizedSsidList $allowlistRaw
    if ($allowlist.Count -eq 0) {
        Set-NinjaProperty 'wifiCheckStatus' 'Not configured'
        exit 0
    }
    foreach ($ssid in $toCheck) {
        if ($allowlist -notcontains $ssid) {
            Set-NinjaProperty 'wifiCheckStatus' 'Deviation'
            Write-Host 'Wi-Fi deviation from allowlist detected.'
            exit 1
        }
    }
    Set-NinjaProperty 'wifiCheckStatus' 'OK'
    exit 0
} catch {
    try { Set-NinjaProperty 'wifiCheckStatus' "Error: $($_.Exception.Message)" } catch { }
    Write-Warning $_.Exception.Message
    exit 2
}

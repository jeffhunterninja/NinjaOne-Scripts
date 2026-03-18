<#
.SYNOPSIS
    Checks whether the device's current public IP is in the authorized list (from NinjaOne custom fields) and alerts on deviation.

.DESCRIPTION
    Reads authorized IPs from four device custom fields: authorizedIPorg, authorizedIPloc, authorizedIPuser, authorizedIPdevice.
    Combines them, fetches the device public IP via ipify, and compares. On deviation, writes status to device custom fields and exits 1
    so NinjaOne can alert (script failure or Condition Based Alerting on publicIPStatus).

.PARAMETER NoNinjaWrite
    When set, does not call Ninja-Property-Set. Use for local/testing without writing to NinjaOne.

.EXAMPLE
    .\Check-PublicIP.ps1
    .\Check-PublicIP.ps1 -NoNinjaWrite
#>
[CmdletBinding()]
param(
    [switch]$NoNinjaWrite
)

function Write-NinjaStatus {
    param([string]$StatusField, [string]$Value, [string]$MessageField, [string]$MessageValue)
    if ($NoNinjaWrite) { return }
    try {
        if (Get-Command Ninja-Property-Set -ErrorAction SilentlyContinue) {
            Ninja-Property-Set $StatusField $Value
            if ($MessageField -and $MessageValue) {
                Ninja-Property-Set $MessageField $MessageValue
            }
        }
    } catch {
        Write-Warning "Ninja-Property-Set failed: $($_.Exception.Message)"
    }
}

try {
    # Get authorized IPs from four NinjaOne device custom fields
    $authorizedIPorgRaw = $null
    $authorizedIPlocRaw = $null
    $authorizedIPuserRaw = $null
    $authorizedIPdeviceRaw = $null
    if (Get-Command Ninja-Property-Get -ErrorAction SilentlyContinue) {
        try { $authorizedIPorgRaw = Ninja-Property-Get authorizedIPorg } catch { }
        try { $authorizedIPlocRaw = Ninja-Property-Get authorizedIPloc } catch { }
        try { $authorizedIPuserRaw = Ninja-Property-Get authorizedIPuser } catch { }
        try { $authorizedIPdeviceRaw = Ninja-Property-Get authorizedIPdevice } catch { }
    }

    # Parse and trim each entry; combine and deduplicate
    $orgIPs = @(); $locIPs = @(); $userIPs = @(); $deviceIPs = @()
    if (![string]::IsNullOrWhiteSpace($authorizedIPorgRaw)) { $orgIPs = ($authorizedIPorgRaw -split ",\s*").ForEach({ $_.Trim() }) | Where-Object { $_ -ne "" } }
    if (![string]::IsNullOrWhiteSpace($authorizedIPlocRaw)) { $locIPs = ($authorizedIPlocRaw -split ",\s*").ForEach({ $_.Trim() }) | Where-Object { $_ -ne "" } }
    if (![string]::IsNullOrWhiteSpace($authorizedIPuserRaw)) { $userIPs = ($authorizedIPuserRaw -split ",\s*").ForEach({ $_.Trim() }) | Where-Object { $_ -ne "" } }
    if (![string]::IsNullOrWhiteSpace($authorizedIPdeviceRaw)) { $deviceIPs = ($authorizedIPdeviceRaw -split ",\s*").ForEach({ $_.Trim() }) | Where-Object { $_ -ne "" } }

    $authorizedIPs = ($orgIPs + $locIPs + $userIPs + $deviceIPs) | Sort-Object -Unique

    if ($authorizedIPs.Count -eq 0) {
        Write-Host "No authorized IPs configured; skipping check."
        Write-NinjaStatus -StatusField "publicIPStatus" -Value "No authorized IPs configured"
        exit 0
    }

    # Get own public IP address
    $publicIP = (Invoke-WebRequest -Uri "https://api.ipify.org" -UseBasicParsing).Content.Trim()
    Write-Host "Public IP Address: $publicIP"

    if (-not $NoNinjaWrite -and (Get-Command Ninja-Property-Set -ErrorAction SilentlyContinue)) {
        try { Ninja-Property-Set "currentPublicIP" $publicIP } catch { }
    }

    if ($authorizedIPs -contains $publicIP) {
        Write-Host "Public IP address is authorized."
        Write-NinjaStatus -StatusField "publicIPStatus" -Value "OK"
        exit 0
    }

    Write-Host "Public IP address is NOT authorized (deviation)."
    $msg = "Current IP $publicIP is not in authorized list."
    Write-NinjaStatus -StatusField "publicIPStatus" -Value "Deviation" -MessageField "publicIPDeviationMessage" -MessageValue $msg
    exit 1
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)"
    $errValue = "Error: $($_.Exception.Message)"
    Write-NinjaStatus -StatusField "publicIPStatus" -Value $errValue
    exit 2
}

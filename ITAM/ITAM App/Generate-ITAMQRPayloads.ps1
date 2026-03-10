<#
.SYNOPSIS
    Outputs JSON strings for NinjaOne ITAM app QR codes (user and devices).
.DESCRIPTION
    Standalone script. No dot-sourcing. Use the output strings in any QR code generator
    to produce QRs that the NinjaOne ITAM iPhone app can scan.
    Optionally reads from NinjaOne API (device list) or from a CSV.
.EXAMPLE
    # One user by email, two devices by ID
    .\Generate-ITAMQRPayloads.ps1 -UserEmail "jane@company.com" -DeviceIds 1001,1002
.EXAMPLE
    # User QR only (print and use for scanning "user" first)
    .\Generate-ITAMQRPayloads.ps1 -UserEmail "jane@company.com" -UserOnly
.EXAMPLE
    # Device QRs only (after user is already chosen in app)
    .\Generate-ITAMQRPayloads.ps1 -DeviceIds 1001,1002,1003 -DeviceOnly
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $UserEmail,

    [Parameter(Mandatory = $false)]
    [string] $UserUid,

    [Parameter(Mandatory = $false)]
    [int[]] $DeviceIds,

    [Parameter(Mandatory = $false)]
    [switch] $UserOnly,

    [Parameter(Mandatory = $false)]
    [switch] $DeviceOnly
)

# --- User QR ---
if (-not $DeviceOnly) {
    if ($UserEmail) {
        $userJson = @{ type = "user"; email = $UserEmail } | ConvertTo-Json -Compress
        Write-Host "User QR (email): $userJson"
    }
    elseif ($UserUid) {
        $userJson = @{ type = "user"; uid = $UserUid } | ConvertTo-Json -Compress
        Write-Host "User QR (uid):   $userJson"
    }
    else {
        Write-Warning "Provide -UserEmail or -UserUid to generate user QR."
    }
    if ($UserOnly) { exit 0 }
}

# --- Device QRs ---
if (-not $UserOnly -and $DeviceIds) {
    foreach ($id in $DeviceIds) {
        $deviceJson = @{ type = "device"; id = $id } | ConvertTo-Json -Compress
        Write-Host "Device QR:       $deviceJson"
    }
}
elseif (-not $UserOnly -and -not $DeviceIds) {
    Write-Warning "Provide -DeviceIds to generate device QR(s). Example: -DeviceIds 1001,1002"
}

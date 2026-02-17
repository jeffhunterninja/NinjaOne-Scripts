#Requires -Version 5.1
<#
.SYNOPSIS
  Determines if a Windows device is MDM-enrolled and identifies the provider.

.DESCRIPTION
  Parses dsregcmd /status, registry keys under Microsoft Enrollments and OMADM\Accounts,
  and infers MDM vendor from URLs/provider IDs. Outputs a summary object. Optionally
  updates NinjaOne custom fields mdmStatus and mdmProvider. Works in PS 5.1+; admin
  recommended for registry access.

.PARAMETER OutputFormat
  Output format: List (default, human-readable) or Json (machine-consumable).

.PARAMETER UpdateNinjaProperties
  When set, writes mdmStatus and mdmProvider to NinjaOne custom fields via
  Ninja-Property-Set (only when cmdlet is available).

.EXAMPLE
  .\Get-MDMEnrollment.ps1

.EXAMPLE
  .\Get-MDMEnrollment.ps1 -OutputFormat Json

.EXAMPLE
  .\Get-MDMEnrollment.ps1 -UpdateNinjaProperties

.NOTES
  Run context: Device script (runs on each managed Windows endpoint).
  Custom fields used: mdmStatus, mdmProvider (create in NinjaOne before using -UpdateNinjaProperties).
  Requires dsregcmd (Windows 10 1507+ / Server 2016+).

.EXIT CODES
  0 = Success
  1 = dsregcmd not found or fatal error
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet('List', 'Json')]
    [string]$OutputFormat = 'List',

    [Parameter()]
    [switch]$UpdateNinjaProperties
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-FirstLineValue {
    param([string]$Text, [string]$Label)
    if (-not $Text) { return $null }
    $pattern = "^\s*$([regex]::Escape($Label))\s*:\s*(.+?)\s*$"
    foreach ($line in $Text -split "`r?`n") {
        $m = [regex]::Match($line, $pattern)
        if ($m.Success) { return $m.Groups[1].Value.Trim() }
    }
    return $null
}

function Get-DsRegInfo {
    $dsregcmd = Get-Command dsregcmd -ErrorAction SilentlyContinue
    if (-not $dsregcmd) {
        Write-Warning "dsregcmd not found (Windows 10 1507+ / Server 2016+ required)"
        return [pscustomobject]@{
            Raw         = $null
            AADJoined   = $null
            MDMUrl      = $null
            MDMUserUPN  = $null
        }
    }
    try {
        $out = (& $dsregcmd.Path /status) -join "`n"
    } catch {
        return [pscustomobject]@{
            Raw         = $null
            AADJoined   = $null
            MDMUrl      = $null
            MDMUserUPN  = $null
        }
    }
    [pscustomobject]@{
        Raw         = $out
        AADJoined   = (Get-FirstLineValue -Text $out -Label 'AzureAdJoined')
        MDMUrl      = (Get-FirstLineValue -Text $out -Label 'MDMUrl')
        MDMUserUPN  = (Get-FirstLineValue -Text $out -Label 'MDMUserUPN')
    }
}

function Try-GetRegKeyProps {
    param([string]$Path, [string[]]$Names)
    try {
        $item = Get-ItemProperty -Path $Path -ErrorAction Stop
        $o = @{}
        foreach ($n in $Names) { $o[$n] = $item.$n }
        return [pscustomobject]$o
    } catch { return $null }
}

function Get-EnrollmentRegInfo {
    $base = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (-not (Test-Path $base)) { return @() }

    Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object {
        $p = $_.PSPath
        $vals = Try-GetRegKeyProps -Path $p -Names @(
            'ProviderID','ProviderName','UPN','EnrollmentType',
            'DiscoveryServiceFullURL','MdmEnrollmentUrl','TenantId'
        )
        if ($vals) {
            $vals | Add-Member NoteProperty 'KeyPath' $p
            $vals
        }
    } | Where-Object { $_ }
}

function Get-OmadmAccountsInfo {
    $base = 'HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts'
    if (-not (Test-Path $base)) { return @() }

    Get-ChildItem $base -ErrorAction SilentlyContinue | ForEach-Object {
        $p = $_.PSPath
        $vals = Try-GetRegKeyProps -Path $p -Names @(
            'ServerAddress','FriendlyName','DefaultProfile','Flags'
        )
        if ($vals) {
            $vals | Add-Member NoteProperty 'KeyPath' $p
            $vals
        }
    } | Where-Object { $_ }
}

function Guess-MdmVendor {
    param(
        [string]$Url,
        [string]$ProviderId,
        [string]$ProviderName
    )
    $haystack = (@($Url,$ProviderId,$ProviderName) | Where-Object { $_ } | ForEach-Object { $_.ToLowerInvariant() }) -join ' | '

    if (-not $haystack) { return $null }

    switch -Regex ($haystack) {
        'manage\.microsoft\.com|microsoftonline|enrollment\.manage|intune' { return 'Microsoft Intune' }
        'awmdm|airwatch|workspace one|vmware'                            { return 'VMware Workspace ONE (AirWatch)' }
        'maas360|fiberlink|ibm'                                          { return 'IBM MaaS360' }
        'mobileiron|ivanti'                                              { return 'Ivanti (MobileIron)' }
        'soti|mobicontrol'                                               { return 'SOTI MobiControl' }
        'citrix|xenmobile|endpointmanagement'                            { return 'Citrix Endpoint Management' }
        '42gears|suremdm'                                                { return '42Gears SureMDM' }
        default                                                          { return $null }
    }
}

function Get-MdmEnrollment {
    $ds   = Get-DsRegInfo
    $enr  = @(Get-EnrollmentRegInfo)
    $oma  = @(Get-OmadmAccountsInfo)

    # Best available URL source
    $urlCandidate = $ds.MDMUrl
    if (-not $urlCandidate) {
        $urlCandidate = ($enr | Where-Object { $_.MdmEnrollmentUrl } | Select-Object -First 1 -ExpandProperty MdmEnrollmentUrl)
        if (-not $urlCandidate) {
            $urlCandidate = ($enr | Where-Object { $_.DiscoveryServiceFullURL } | Select-Object -First 1 -ExpandProperty DiscoveryServiceFullURL)
        }
        if (-not $urlCandidate) {
            $urlCandidate = ($oma | Where-Object { $_.ServerAddress } | Select-Object -First 1 -ExpandProperty ServerAddress)
        }
    }

    # Other identity hints
    $provId   = ($enr | Where-Object { $_.ProviderID }   | Select-Object -First 1 -ExpandProperty ProviderID)
    $provName = ($enr | Where-Object { $_.ProviderName } | Select-Object -First 1 -ExpandProperty ProviderName)
    $upn      = $ds.MDMUserUPN
    if (-not $upn) { $upn = ($enr | Where-Object { $_.UPN } | Select-Object -First 1 -ExpandProperty UPN) }

    $vendor   = Guess-MdmVendor -Url $urlCandidate -ProviderId $provId -ProviderName $provName

    # If any of these show up, we consider it enrolled
    $isEnrolled = $false
    if ($urlCandidate) { $isEnrolled = $true }
    elseif ($provId -or $provName) { $isEnrolled = $true }
    elseif ($oma.Count -gt 0) { $isEnrolled = $true }

    # Helpful details
    $aadJoined = $null
    if ($ds.AADJoined) {
        $aadJoined = ($ds.AADJoined -eq 'YES')
    }

    # Roll up multiple enrollments (rare but possible) as details
    $details = [pscustomobject]@{
        DsReg = [pscustomobject]@{
            AADJoined  = $aadJoined
            MDMUrl     = $ds.MDMUrl
            MDMUserUPN = $ds.MDMUserUPN
        }
        EnrollmentRegistry = $enr
        OmadmAccounts      = $oma
    }

    # Primary summary
    [pscustomobject]@{
        IsMdmEnrolled  = $isEnrolled
        VendorGuess    = $vendor
        ProviderID     = $provId
        ProviderName   = $provName
        MDMUrl         = $urlCandidate
        MDMUserUPN     = $upn
        AzureAdJoined  = $aadJoined
        Details        = $details
    }
}

try {
    $result = Get-MdmEnrollment
} catch {
    Write-Error "Failed to get MDM enrollment status: $_"
    exit 1
}

# Output
if ($OutputFormat -eq 'Json') {
    $result | ConvertTo-Json -Depth 6
} else {
    $result | Format-List
}

# Optional NinjaOne custom field updates
if ($UpdateNinjaProperties) {
    $ninjaCmd = Get-Command Ninja-Property-Set -ErrorAction SilentlyContinue
    if (-not $ninjaCmd) {
        Write-Warning "Ninja-Property-Set not found; skipping custom field updates"
    } else {
        try {
            $status = if ($result.IsMdmEnrolled) { 'Enrolled' } else { 'Not Enrolled' }
            $provider = $result.VendorGuess
            if (-not $provider) { $provider = $result.ProviderID }
            if (-not $provider) { $provider = $result.ProviderName }
            if (-not $provider) { $provider = $result.MDMUrl }
            if (-not $provider) { $provider = 'Unknown' }
            & $ninjaCmd.Name -Name 'mdmStatus' -Value $status
            & $ninjaCmd.Name -Name 'mdmProvider' -Value $provider
        } catch {
            Write-Warning "Failed to update NinjaOne custom fields: $_"
        }
    }
}

exit 0

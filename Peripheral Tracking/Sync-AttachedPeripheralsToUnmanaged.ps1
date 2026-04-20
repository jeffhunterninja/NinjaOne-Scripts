#Requires -Version 5.1
<#
.SYNOPSIS
    Discovers attached peripherals and syncs them to NinjaOne unmanaged devices.

.DESCRIPTION
    Enumerates local Windows peripherals with Get-PnpDevice, normalizes hardware details,
    and then creates or updates related NinjaOne unmanaged devices using endpoint cmdlets:
    New-NinjaUnmanagedDevice, Get-NinjaUnmanagedDevice, and Set-NinjaUnmanagedDevice.

    Matching priority for updates:
      1) Asset ID (derived from short hardware ID) when MatchOnHardwareIdFirst is enabled
      2) Asset Serial Number (when available)
      3) Name (generated deterministic display name)

    The script is fully standalone and does not use dot-sourcing.

.PARAMETER Relation
    Directional relation name used when creating/updating unmanaged devices.

.PARAMETER AssignToDeviceUser
    Sets assignToDeviceUser = $true in create/update payloads.

.PARAMETER IncludeClasses
    Optional PnP classes to include. Uses defaults if not specified.

.PARAMETER ExcludeClasses
    Optional PnP classes to exclude.

.PARAMETER IncludeVirtualDevices
    Include virtual/infrastructure devices that are excluded by default.

.PARAMETER MaxRows
    Maximum number of discovered peripherals to process. 0 means no limit.

.PARAMETER RoleMap
    Optional hashtable mapping PnP class -> unmanaged role name.
    Example: @{ Monitor = 'Displays'; Mouse = 'Mouse'; Keyboard = 'Keyboard' }

.PARAMETER MatchOnHardwareIdFirst
    If set (default), match by generated assetId before other keys.

.PARAMETER NoWrite
    Discovery and matching only. Does not create/update NinjaOne unmanaged devices.

.PARAMETER SkipErrors
    Continue processing other peripherals if one create/update operation fails.

.EXAMPLE
    .\Sync-AttachedPeripheralsToUnmanaged.ps1 -NoWrite -Verbose

.EXAMPLE
    .\Sync-AttachedPeripheralsToUnmanaged.ps1 -Relation "Connected To" -AssignToDeviceUser -WhatIf

.EXAMPLE
    .\Sync-AttachedPeripheralsToUnmanaged.ps1 -RoleMap @{ Monitor = 'Displays'; USB = 'Accessory' }
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [string]$Relation = 'Connected To',
    [switch]$AssignToDeviceUser,
    [string[]]$IncludeClasses = @(),
    [string[]]$ExcludeClasses = @(),
    [switch]$IncludeVirtualDevices,
    [int]$MaxRows = 500,
    [hashtable]$RoleMap = @{},
    [bool]$MatchOnHardwareIdFirst = $true,
    [switch]$NoWrite,
    [bool]$SkipErrors = $true
)

$ErrorActionPreference = 'Stop'

$DefaultIncludeClasses = @(
    'Keyboard', 'Mouse', 'Monitor', 'USB', 'Biometric', 'Camera', 'Bluetooth',
    'Sound', 'Media', 'Net', 'Sensor', 'SmartCardReader', 'Printer', 'Image', 'BluetoothLE'
)

$VirtualDevicePatterns = @(
    '^WAN Miniport'
    '^Microsoft Kernel Debug'
    '^Hyper-V Virtual'
    '^USB Root Hub'
    '^Generic USB Hub$'
    '^USB Composite Device$'
    'Host Controller'
)

$DefaultRoleMap = @{
    Keyboard        = 'Keyboard'
    Mouse           = 'Mouse'
    Monitor         = 'Monitor'
    USB             = 'Peripheral'
    Biometric       = 'Biometric'
    Camera          = 'Camera'
    Bluetooth       = 'Peripheral'
    BluetoothLE     = 'Peripheral'
    Sound           = 'Audio Device'
    Media           = 'Audio Device'
    Net             = 'Network Device'
    Sensor          = 'Sensor'
    SmartCardReader = 'Smart Card Reader'
    Printer         = 'Printer'
    Image           = 'Imaging Device'
}

function Get-ConnectionType {
    param([string]$InstanceId)
    if ([string]::IsNullOrWhiteSpace($InstanceId)) { return '' }
    $prefix = ($InstanceId -split '\\')[0]
    switch ($prefix) {
        'USB'       { 'USB' }
        'BTHENUM'   { 'Bluetooth' }
        'BTHHFENUM' { 'Bluetooth' }
        'PCI'       { 'PCI' }
        'HDAUDIO'   { 'HD Audio' }
        'SWD'       { 'Software' }
        'HID'       { 'HID' }
        'DISPLAY'   { 'Display' }
        'MONITOR'   { 'Display' }
        'ACPI'      { 'Internal' }
        'ROOT'      { 'System' }
        default     { $prefix }
    }
}

function Get-ShortHardwareId {
    param([string[]]$HardwareIDs)
    if ($null -eq $HardwareIDs -or $HardwareIDs.Count -eq 0) { return '' }
    foreach ($id in $HardwareIDs) {
        if ($id -match 'VID_[0-9A-Fa-f]{4}&PID_[0-9A-Fa-f]{4}') {
            return ($Matches[0]).ToUpper()
        }
    }
    $segments = $HardwareIDs[0] -split '\\'
    if ($segments.Count -ge 2) { return $segments[1] }
    return $HardwareIDs[0]
}

function Convert-ByteArrayToAsciiString {
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Count -eq 0) { return '' }
    $chars = foreach ($b in $Bytes) {
        if ($b -gt 0) { [char]$b }
    }
    return ((-join $chars).Trim())
}

function Test-UsableSerial {
    param([string]$Serial)
    if ([string]::IsNullOrWhiteSpace($Serial)) { return $false }
    $value = $Serial.Trim()
    if ($value.Length -lt 4) { return $false }
    if ($value -match '^(0+|1+|9+|F+|X+)$') { return $false }
    if ($value -match '^(1234|12345|123456|1234567|12345678|123456789)$') { return $false }
    if ($value -match '^(.)\1{5,}$') { return $false }
    return $true
}

function Get-UsbSerialFromInstanceId {
    param([string]$InstanceId)
    if ([string]::IsNullOrWhiteSpace($InstanceId)) { return '' }
    if ($InstanceId -notmatch '^USB\\') { return '' }
    $parts = $InstanceId -split '\\'
    if ($parts.Count -lt 3) { return '' }
    $candidate = ($parts[2]).Trim()
    if ([string]::IsNullOrWhiteSpace($candidate)) { return '' }
    if ($candidate -match '^MI_\d{2}$') { return '' }
    if ($candidate -match '^&') { return '' }
    return $candidate
}

function Get-MonitorSerialMap {
    $map = @{}
    try {
        $monitorIds = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction Stop
        foreach ($monitor in $monitorIds) {
            $serial = Convert-ByteArrayToAsciiString -Bytes $monitor.SerialNumberID
            if (-not (Test-UsableSerial -Serial $serial)) { continue }
            $instanceKey = [string]$monitor.InstanceName
            if ([string]::IsNullOrWhiteSpace($instanceKey)) { continue }
            $instanceKey = $instanceKey -replace '_\d+$', ''
            $instanceKey = $instanceKey.ToUpperInvariant()
            if (-not $map.ContainsKey($instanceKey)) {
                $map[$instanceKey] = $serial
            }
        }
    } catch {
        Write-Verbose "WmiMonitorID query failed: $($_.Exception.Message)"
    }
    return $map
}

function Get-MonitorSerialFromInstanceId {
    param(
        [string]$InstanceId,
        [hashtable]$MonitorSerialMap
    )
    if ([string]::IsNullOrWhiteSpace($InstanceId)) { return '' }
    if ($null -eq $MonitorSerialMap -or $MonitorSerialMap.Count -eq 0) { return '' }
    $instanceUpper = $InstanceId.ToUpperInvariant()
    if ($MonitorSerialMap.ContainsKey($instanceUpper)) { return $MonitorSerialMap[$instanceUpper] }
    foreach ($k in $MonitorSerialMap.Keys) {
        if ($instanceUpper.StartsWith($k) -or $k.StartsWith($instanceUpper)) {
            return $MonitorSerialMap[$k]
        }
    }
    return ''
}

function Get-SafeSlug {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $value = $Text.ToLowerInvariant()
    $value = $value -replace '[^a-z0-9\-]+', '-'
    $value = $value.Trim('-')
    return $value
}

function Get-DeviceContextName {
    $name = $env:NINJA_DEVICE_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { $name = $env:COMPUTERNAME }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'managed-device' }
    return $name
}

function Get-PeripheralInventory {
    param(
        [string[]]$IncludeClasses,
        [string[]]$ExcludeClasses,
        [switch]$IncludeVirtualDevices
    )

    $classesToInclude = if ($IncludeClasses.Count -gt 0) { $IncludeClasses } else { $DefaultIncludeClasses }

    $allDevices = @()
    try {
        $allDevices = Get-PnpDevice -ErrorAction Stop | Where-Object { $_.Status -eq 'OK' }
    } catch {
        throw "Get-PnpDevice failed: $($_.Exception.Message)"
    }

    $filteredDevices = @($allDevices | Where-Object {
        $_.Class -in $classesToInclude -and ($ExcludeClasses.Count -eq 0 -or $_.Class -notin $ExcludeClasses)
    })

    if (-not $IncludeVirtualDevices) {
        $filteredDevices = @($filteredDevices | Where-Object {
            $name = $_.FriendlyName
            $isVirtual = $false
            foreach ($pattern in $VirtualDevicePatterns) {
                if ($name -match $pattern) { $isVirtual = $true; break }
            }
            -not $isVirtual
        })
    }

    $pnpByDeviceId = @{}
    try {
        $pnpEntities = Get-CimInstance -ClassName Win32_PnPEntity -ErrorAction Stop
        foreach ($e in $pnpEntities) { $pnpByDeviceId[$e.DeviceID] = $e }
    } catch {
        Write-Warning "Win32_PnPEntity query failed: $($_.Exception.Message)"
    }

    $busReportedNames = @{}
    $monitorSerialMap = Get-MonitorSerialMap
    foreach ($dev in $filteredDevices) {
        try {
            $prop = Get-PnpDeviceProperty -InstanceId $dev.InstanceId `
                -KeyName 'DEVPKEY_Device_BusReportedDeviceDesc' -ErrorAction SilentlyContinue
            if ($prop -and $prop.Data -and $prop.Data -ne '') {
                $busReportedNames[$dev.InstanceId] = $prop.Data
            }
        } catch { }
    }

    $managedName = Get-DeviceContextName
    return @($filteredDevices | ForEach-Object {
        $cls = if ([string]::IsNullOrWhiteSpace($_.Class)) { 'Unknown' } else { $_.Class }
        $friendlyName = if ([string]::IsNullOrWhiteSpace($_.FriendlyName)) { $cls } else { $_.FriendlyName }

        $entity = if ($pnpByDeviceId.ContainsKey($_.InstanceId)) { $pnpByDeviceId[$_.InstanceId] } else { $null }
        $desc = if ($entity -and $entity.Description) { $entity.Description } else { '' }
        $hwIds = if ($entity -and $entity.HardwareID) { $entity.HardwareID } else { @() }

        $busName = if ($busReportedNames.ContainsKey($_.InstanceId)) { $busReportedNames[$_.InstanceId] } else { '' }
        $deviceNamePart = if ($busName) { $busName } elseif ($friendlyName -ne $cls) { $friendlyName } elseif ($desc) { $desc } else { $friendlyName }
        $hardwareId = Get-ShortHardwareId -HardwareIDs $hwIds
        $connectionType = Get-ConnectionType -InstanceId $_.InstanceId
        $generatedName = "$deviceNamePart on $managedName"
        $discoveredSerial = ''
        $serialSource = 'None'
        $serialConfidence = 'None'

        if ($cls -in @('Monitor', 'Display')) {
            $monitorSerial = Get-MonitorSerialFromInstanceId -InstanceId $_.InstanceId -MonitorSerialMap $monitorSerialMap
            if (Test-UsableSerial -Serial $monitorSerial) {
                $discoveredSerial = $monitorSerial
                $serialSource = 'WmiMonitorID'
                $serialConfidence = 'High'
            }
        }
        if (-not $discoveredSerial) {
            $usbSerial = Get-UsbSerialFromInstanceId -InstanceId $_.InstanceId
            if (Test-UsableSerial -Serial $usbSerial) {
                $discoveredSerial = $usbSerial
                $serialSource = 'InstanceId'
                $serialConfidence = 'Medium'
            }
        }

        [PSCustomObject]@{
            PnpClass       = $cls
            FriendlyName   = $friendlyName
            DeviceNamePart = $deviceNamePart
            GeneratedName  = $generatedName
            HardwareId     = $hardwareId
            DiscoveredSerial = $discoveredSerial
            SerialSource   = $serialSource
            SerialConfidence = $serialConfidence
            ConnectionType = $connectionType
            InstanceId     = $_.InstanceId
        }
    } | Sort-Object -Property PnpClass, GeneratedName)
}

function New-PeripheralPayload {
    param(
        [pscustomobject]$Peripheral,
        [hashtable]$MergedRoleMap,
        [string]$Relation,
        [switch]$AssignToDeviceUser
    )
    $role = if ($MergedRoleMap.ContainsKey($Peripheral.PnpClass)) { $MergedRoleMap[$Peripheral.PnpClass] } else { 'Peripheral' }
    $assetId = if ([string]::IsNullOrWhiteSpace($Peripheral.HardwareId)) { $null } else { $Peripheral.HardwareId }
    $serial = if ([string]::IsNullOrWhiteSpace($Peripheral.DiscoveredSerial)) { $null } else { $Peripheral.DiscoveredSerial }

    $customFields = @{
        discoveredConnectionType = $Peripheral.ConnectionType
        discoveredPnpClass       = $Peripheral.PnpClass
        discoveredFriendlyName   = $Peripheral.FriendlyName
        discoveredSerial         = $Peripheral.DiscoveredSerial
        discoveredSerialSource   = $Peripheral.SerialSource
        discoveredSerialConfidence = $Peripheral.SerialConfidence
    }

    return [ordered]@{
        role               = $role
        name               = $Peripheral.GeneratedName
        relation           = $Relation
        assignToDeviceUser = [bool]$AssignToDeviceUser
        assetId            = $assetId
        assetSerialNumber  = $serial
        customFieldHashMap = $customFields
    }
}

function Get-ExistingLookupForRoleRelation {
    param([string]$Role, [string]$Relation)
    $existing = @(Get-NinjaUnmanagedDevice -Role $Role -Relation $Relation -ErrorAction Stop)
    $byAssetId = @{}
    $bySerial = @{}
    $byName = @{}

    foreach ($item in $existing) {
        $assetId = $item.assetId
        if ($assetId -and -not $byAssetId.ContainsKey($assetId)) { $byAssetId[$assetId] = $item }

        $serial = $item.assetSerialNumber
        if ($serial -and -not $bySerial.ContainsKey($serial)) { $bySerial[$serial] = $item }

        $name = $item.name
        if ($name -and -not $byName.ContainsKey($name)) { $byName[$name] = $item }
    }

    [PSCustomObject]@{
        ByAssetId = $byAssetId
        BySerial  = $bySerial
        ByName    = $byName
    }
}

function Resolve-ExistingMatch {
    param(
        [hashtable]$LookupByAssetId,
        [hashtable]$LookupBySerial,
        [hashtable]$LookupByName,
        [hashtable]$Payload,
        [bool]$MatchOnHardwareIdFirst
    )

    $assetId = $Payload.assetId
    $serial = $Payload.assetSerialNumber
    $name = $Payload.name

    if ($serial -and $LookupBySerial.ContainsKey($serial)) {
        return $LookupBySerial[$serial]
    }
    if ($MatchOnHardwareIdFirst -and $assetId -and $LookupByAssetId.ContainsKey($assetId)) {
        return $LookupByAssetId[$assetId]
    }
    if ($name -and $LookupByName.ContainsKey($name)) {
        return $LookupByName[$name]
    }
    if (-not $MatchOnHardwareIdFirst -and $assetId -and $LookupByAssetId.ContainsKey($assetId)) {
        return $LookupByAssetId[$assetId]
    }
    return $null
}

if (-not (Get-Command New-NinjaUnmanagedDevice -ErrorAction SilentlyContinue)) {
    throw 'New-NinjaUnmanagedDevice command is not available on this endpoint.'
}
if (-not (Get-Command Get-NinjaUnmanagedDevice -ErrorAction SilentlyContinue)) {
    throw 'Get-NinjaUnmanagedDevice command is not available on this endpoint.'
}
if (-not (Get-Command Set-NinjaUnmanagedDevice -ErrorAction SilentlyContinue)) {
    throw 'Set-NinjaUnmanagedDevice command is not available on this endpoint.'
}

$mergedRoleMap = @{}
foreach ($k in $DefaultRoleMap.Keys) { $mergedRoleMap[$k] = $DefaultRoleMap[$k] }
foreach ($k in $RoleMap.Keys) { $mergedRoleMap[$k] = [string]$RoleMap[$k] }

$peripherals = Get-PeripheralInventory -IncludeClasses $IncludeClasses -ExcludeClasses $ExcludeClasses -IncludeVirtualDevices:$IncludeVirtualDevices
if ($MaxRows -gt 0 -and $peripherals.Count -gt $MaxRows) {
    $peripherals = @($peripherals[0..($MaxRows - 1)])
}

$created = 0
$updated = 0
$skipped = 0
$failed = 0

$lookupByRoleRelation = @{}

foreach ($peripheral in $peripherals) {
    $payload = New-PeripheralPayload -Peripheral $peripheral -MergedRoleMap $mergedRoleMap -Relation $Relation -AssignToDeviceUser:$AssignToDeviceUser
    $lookupKey = "$($payload.role)|$($payload.relation)"

    try {
        if (-not $lookupByRoleRelation.ContainsKey($lookupKey)) {
            $lookupByRoleRelation[$lookupKey] = Get-ExistingLookupForRoleRelation -Role $payload.role -Relation $payload.relation
        }
        $lookup = $lookupByRoleRelation[$lookupKey]
        $match = Resolve-ExistingMatch `
            -LookupByAssetId $lookup.ByAssetId `
            -LookupBySerial $lookup.BySerial `
            -LookupByName $lookup.ByName `
            -Payload $payload `
            -MatchOnHardwareIdFirst $MatchOnHardwareIdFirst

        if ($NoWrite) {
            if ($null -eq $match) {
                Write-Host "[NoWrite] CREATE role='$($payload.role)' name='$($payload.name)' assetId='$($payload.assetId)' serial='$($payload.assetSerialNumber)' source='$($peripheral.SerialSource)'"
            } else {
                Write-Host "[NoWrite] UPDATE role='$($payload.role)' name='$($payload.name)' matchId='$($match.id)' serial='$($payload.assetSerialNumber)' source='$($peripheral.SerialSource)'"
            }
            $skipped++
            continue
        }

        if ($null -eq $match) {
            if ($PSCmdlet.ShouldProcess($payload.name, "Create unmanaged device (role '$($payload.role)')")) {
                $null = New-NinjaUnmanagedDevice $payload
                $created++
            }
        } else {
            $updateObject = $match.PSObject.Copy()
            $updateObject.name = $payload.name
            $updateObject.role = $payload.role
            $updateObject.relation = $payload.relation
            $updateObject.assignToDeviceUser = $payload.assignToDeviceUser
            if ($payload.assetId) { $updateObject.assetId = $payload.assetId }
            if ($payload.assetSerialNumber) { $updateObject.assetSerialNumber = $payload.assetSerialNumber }
            $updateObject.customFieldHashMap = $payload.customFieldHashMap

            if ($PSCmdlet.ShouldProcess($payload.name, "Update unmanaged device id '$($match.id)'")) {
                $null = Set-NinjaUnmanagedDevice $updateObject
                $updated++
            }
        }
    } catch {
        $failed++
        Write-Warning "Failed for '$($payload.name)': $($_.Exception.Message)"
        if (-not $SkipErrors) { throw }
    }
}

Write-Host "Peripheral sync complete. Discovered: $($peripherals.Count), Created: $created, Updated: $updated, Skipped: $skipped, Failed: $failed."

#Requires -Version 5.1
<#
.SYNOPSIS
    Enumerates attached peripherals and writes a detailed HTML table to a NinjaOne WYSIWYG device custom field.

.DESCRIPTION
    Scans the local Windows device for Plug and Play peripherals (keyboards, mice, monitors, USB devices,
    cameras, etc.), enriches each entry with manufacturer, bus-reported product name, connection type, and
    hardware ID, builds an HTML table, and writes it to a NinjaOne device custom field via Ninja-Property-Set.

    The target custom field must be of type WYSIWYG (Rich text) at Device scope so the HTML renders as a
    table in the NinjaOne UI; a plain text field will display raw HTML.

    By default, virtual/infrastructure devices (WAN Miniports, Root Hubs, Host Controllers, Hyper-V
    virtual adapters) are excluded to keep the report focused on real peripherals.

.PARAMETER NoNinjaWrite
    When set, does not call Ninja-Property-Set. Use for local/testing without writing to NinjaOne.

.PARAMETER CustomFieldName
    NinjaOne custom field API name to write the HTML table to. Default: attachedPeripherals.

.PARAMETER IncludeClasses
    Optional array of PnP device class names to include (e.g. Keyboard, Mouse, Monitor, USB).
    If not specified, a default set of common peripheral classes is used.

.PARAMETER ExcludeClasses
    Optional array of PnP device class names to exclude from the results.

.PARAMETER IncludeVirtualDevices
    When set, includes virtual/infrastructure devices (WAN Miniports, Root Hubs, Host Controllers,
    Hyper-V virtual adapters) that are excluded by default.

.PARAMETER MaxRows
    Maximum number of device rows to include in the table. Default: 500. Set to 0 for no limit.

.PARAMETER Simple
    Outputs only three columns (Friendly Name, Class, Status) instead of the full detail view.

.PARAMETER Detailed
    Accepted for backward compatibility. The enriched view is now the default; this switch is a no-op.

.EXAMPLE
    .\Report-AttachedPeripherals.ps1 -NoNinjaWrite
    Enumerates peripherals with full detail (Device Name, Type, Manufacturer, Connection, Status, Hardware ID, Serial Number, Serial Source).
.EXAMPLE
    .\Report-AttachedPeripherals.ps1 -NoNinjaWrite -IncludeVirtualDevices
    Includes virtual adapters, WAN miniports, root hubs, host controllers, etc.
.EXAMPLE
    .\Report-AttachedPeripherals.ps1 -NoNinjaWrite -Simple
    Compact 3-column table (Friendly Name, Class, Status) matching the legacy output format.
.EXAMPLE
    .\Report-AttachedPeripherals.ps1 -IncludeClasses Keyboard,Mouse,Monitor,USB
    Limits the report to specific device classes.
.NOTES
    Requires Get-PnpDevice (Windows 10+). PowerShell 5.1+.
    Bus-reported device descriptions (the real product name a device advertises) are retrieved via
    Get-PnpDeviceProperty when available; falls back to FriendlyName/Description otherwise.
#>
[CmdletBinding()]
param(
    [switch]$NoNinjaWrite,
    [string]$CustomFieldName = 'attachedPeripherals',
    [string[]]$IncludeClasses = @(),
    [string[]]$ExcludeClasses = @(),
    [switch]$IncludeVirtualDevices,
    [int]$MaxRows = 500,
    [switch]$Simple,
    [switch]$Detailed
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

$ClassDisplayNames = @{
    'Keyboard'       = 'Keyboard'
    'Mouse'          = 'Mouse'
    'Monitor'        = 'Monitor'
    'USB'            = 'USB Device'
    'Biometric'      = 'Biometric'
    'Camera'         = 'Camera'
    'Bluetooth'      = 'Bluetooth'
    'BluetoothLE'    = 'Bluetooth LE'
    'Sound'          = 'Audio'
    'Media'          = 'Audio / Media'
    'Net'            = 'Network Adapter'
    'Sensor'         = 'Sensor'
    'SmartCardReader' = 'Smart Card Reader'
    'Printer'        = 'Printer'
    'Image'          = 'Imaging Device'
    'HIDClass'       = 'HID Device'
    'DiskDrive'      = 'Disk Drive'
    'CDROM'          = 'CD/DVD Drive'
    'Display'        = 'Display Adapter'
}

function Escape-HtmlFragment {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = $Text
    $t = $t -replace '&', '&amp;'
    $t = $t -replace '<', '&lt;'
    $t = $t -replace '>', '&gt;'
    $t = $t -replace '"', '&quot;'
    return $t
}

function Get-ConnectionType {
    param([string]$InstanceId)
    if ([string]::IsNullOrEmpty($InstanceId)) { return '' }
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

$classesToInclude = if ($IncludeClasses.Count -gt 0) { $IncludeClasses } else { $DefaultIncludeClasses }

try {
    $allDevices = Get-PnpDevice -ErrorAction Stop | Where-Object { $_.Status -eq 'OK' }
} catch {
    Write-Warning "Get-PnpDevice failed: $($_.Exception.Message)"
    $allDevices = @()
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

if ($Simple) {
    $peripherals = $filteredDevices | ForEach-Object {
        $cls = if ([string]::IsNullOrEmpty($_.Class)) { 'Unknown' } else { $_.Class }
        $name = if ([string]::IsNullOrWhiteSpace($_.FriendlyName)) { $cls } else { $_.FriendlyName }
        [PSCustomObject]@{
            FriendlyName = $name
            Class        = $cls
            Status       = $_.Status
        }
    } | Sort-Object -Property Class, FriendlyName -Unique
} else {
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

    $peripherals = $filteredDevices | ForEach-Object {
        $cls = if ([string]::IsNullOrEmpty($_.Class)) { 'Unknown' } else { $_.Class }
        $friendlyName = if ([string]::IsNullOrWhiteSpace($_.FriendlyName)) { $cls } else { $_.FriendlyName }

        $entity = if ($pnpByDeviceId.ContainsKey($_.InstanceId)) { $pnpByDeviceId[$_.InstanceId] } else { $null }
        $mfr  = if ($entity -and $entity.Manufacturer) { $entity.Manufacturer } else { '' }
        $desc = if ($entity -and $entity.Description)  { $entity.Description }  else { '' }
        $hwIds = if ($entity -and $entity.HardwareID)  { $entity.HardwareID }   else { @() }

        $busName = if ($busReportedNames.ContainsKey($_.InstanceId)) { $busReportedNames[$_.InstanceId] } else { '' }
        $deviceName = if ($busName -ne '')              { $busName }
                      elseif ($friendlyName -ne $cls)   { $friendlyName }
                      elseif ($desc -ne '')             { $desc }
                      else                              { $friendlyName }

        $type       = if ($ClassDisplayNames.ContainsKey($cls)) { $ClassDisplayNames[$cls] } else { $cls }
        $connection = Get-ConnectionType -InstanceId $_.InstanceId
        $hardwareId = Get-ShortHardwareId -HardwareIDs $hwIds
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
            DeviceName   = $deviceName
            Type         = $type
            Manufacturer = $mfr
            Connection   = $connection
            Status       = $_.Status
            HardwareId   = $hardwareId
            SerialNumber = $discoveredSerial
            SerialSource = $serialSource
            SerialConfidence = $serialConfidence
        }
    } | Sort-Object -Property Type, DeviceName
}

$applyMax = $MaxRows -gt 0
$rows = @($peripherals)
if ($applyMax -and $rows.Count -gt $MaxRows) {
    $rows = $rows[0..($MaxRows - 1)]
}

$colCount  = if ($Simple) { 3 } else { 8 }
$cellStyle = 'border: 1px solid #ccc; padding: 4px 8px;'
$thStyle   = 'border: 1px solid #ccc; padding: 6px 8px; text-align: left;'

$sb = [System.Text.StringBuilder]::new()
$null = $sb.AppendLine('<table style="border-collapse: collapse; border: 1px solid #ccc; width: 100%;">')
$null = $sb.AppendLine('  <thead>')
$null = $sb.AppendLine('    <tr style="background: #f0f0f0;">')
if ($Simple) {
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Friendly Name</th>")
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Class</th>")
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Status</th>")
} else {
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Device Name</th>")
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Type</th>")
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Manufacturer</th>")
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Connection</th>")
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Status</th>")
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Hardware ID</th>")
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Serial Number</th>")
    $null = $sb.AppendLine("      <th style=`"$thStyle`">Serial Source</th>")
}
$null = $sb.AppendLine('    </tr>')
$null = $sb.AppendLine('  </thead>')
$null = $sb.AppendLine('  <tbody>')

if ($rows.Count -eq 0) {
    $null = $sb.AppendLine("    <tr><td colspan=`"$colCount`" style=`"$cellStyle`">No peripherals found.</td></tr>")
} else {
    foreach ($row in $rows) {
        if ($Simple) {
            $fn = Escape-HtmlFragment -Text $row.FriendlyName
            $cl = Escape-HtmlFragment -Text $row.Class
            $st = Escape-HtmlFragment -Text $row.Status
            $null = $sb.AppendLine("    <tr><td style=`"$cellStyle`">$fn</td><td style=`"$cellStyle`">$cl</td><td style=`"$cellStyle`">$st</td></tr>")
        } else {
            $dn  = Escape-HtmlFragment -Text $row.DeviceName
            $tp  = Escape-HtmlFragment -Text $row.Type
            $mfr = Escape-HtmlFragment -Text $row.Manufacturer
            $cn  = Escape-HtmlFragment -Text $row.Connection
            $st  = Escape-HtmlFragment -Text $row.Status
            $hid = Escape-HtmlFragment -Text $row.HardwareId
            $srn = Escape-HtmlFragment -Text $row.SerialNumber
            $srs = Escape-HtmlFragment -Text $row.SerialSource
            $null = $sb.AppendLine("    <tr><td style=`"$cellStyle`">$dn</td><td style=`"$cellStyle`">$tp</td><td style=`"$cellStyle`">$mfr</td><td style=`"$cellStyle`">$cn</td><td style=`"$cellStyle`">$st</td><td style=`"$cellStyle`">$hid</td><td style=`"$cellStyle`">$srn</td><td style=`"$cellStyle`">$srs</td></tr>")
        }
    }
    if ($applyMax -and @($peripherals).Count -gt $MaxRows) {
        $totalCount = @($peripherals).Count
        $null = $sb.AppendLine("    <tr><td colspan=`"$colCount`" style=`"$cellStyle font-style: italic;`">… ($totalCount total; list truncated)</td></tr>")
    }
}

$null = $sb.AppendLine('  </tbody>')
$null = $sb.AppendLine('</table>')
$html = $sb.ToString()

if (-not $NoNinjaWrite) {
    try {
        if (Get-Command Ninja-Property-Set -ErrorAction SilentlyContinue) {
            Ninja-Property-Set $CustomFieldName $html
        }
    } catch {
        Write-Warning "Ninja-Property-Set failed ($CustomFieldName): $($_.Exception.Message)"
    }
}

$html

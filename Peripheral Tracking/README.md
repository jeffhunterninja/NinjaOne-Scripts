# Peripheral Tracking (NinjaOne WYSIWYG)

This script enumerates attached peripherals on a Windows device (keyboards, mice, monitors, USB devices, cameras, etc.) and writes a detailed HTML table to a NinjaOne device custom field. When the field is a **WYSIWYG** (Rich text) type, the table renders in the NinjaOne UI.

The default output now includes eight columns — **Device Name**, **Type**, **Manufacturer**, **Connection**, **Status**, **Hardware ID**, **Serial Number**, and **Serial Source** — with bus-reported product names (the real name a device advertises, e.g. "SanDisk Ultra USB 3.0" instead of "USB Mass Storage Device") and automatic filtering of virtual/infrastructure devices.

## Prerequisites

- Windows 10 or later (uses `Get-PnpDevice`, `Get-PnpDeviceProperty`).
- PowerShell 5.1 or later.
- Script run as a **NinjaOne device script** so `Ninja-Property-Set` is available when writing to the custom field.

## NinjaOne setup

### 1. Create the custom field

1. In NinjaOne go to **Administration** > **Custom Fields**.
2. Add a **Device**-scope custom field.
3. Set the **type** to **WYSIWYG** (or **Rich text**). A plain text field will show raw HTML.
4. Set the **API name** to match the script default or your `-CustomFieldName` (e.g. `attachedPeripherals`).
5. Set API permissions as needed so script-written values are allowed.

### 2. Assign the script

- Add **Report-AttachedPeripherals.ps1** as a script in NinjaOne and assign it to the desired devices or policies.
- Run on a schedule (e.g. daily) or on-demand. The script writes the HTML table to the custom field and also outputs it so you can inspect script output in NinjaOne if needed.

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| **NoNinjaWrite** | switch | — | Do not call `Ninja-Property-Set`; only output the HTML (for local/testing). |
| **CustomFieldName** | string | `attachedPeripherals` | NinjaOne custom field API name to write the HTML table to. |
| **IncludeClasses** | string[] | (default set) | PnP device classes to include. If not specified, a default set of peripheral classes is used (Keyboard, Mouse, Monitor, USB, Biometric, Camera, Bluetooth, Sound, Media, Net, etc.). |
| **ExcludeClasses** | string[] | — | PnP device classes to exclude (e.g. `System`). |
| **IncludeVirtualDevices** | switch | — | Include virtual/infrastructure devices (WAN Miniports, Root Hubs, Host Controllers, Hyper-V virtual adapters) that are excluded by default. |
| **MaxRows** | int | `500` | Maximum rows in the table; excess is truncated with a note. Set to `0` for no limit. |
| **Simple** | switch | — | Output only three columns (Friendly Name, Class, Status) matching the legacy compact format. |
| **Detailed** | switch | — | Accepted for backward compatibility; the enriched view is now always used unless `-Simple` is set. |

## Examples

- **Default (write to NinjaOne with full detail):**
  `.\Report-AttachedPeripherals.ps1`

- **Test locally without writing:**
  `.\Report-AttachedPeripherals.ps1 -NoNinjaWrite`

- **Include virtual/infrastructure devices:**
  `.\Report-AttachedPeripherals.ps1 -NoNinjaWrite -IncludeVirtualDevices`

- **Legacy compact output:**
  `.\Report-AttachedPeripherals.ps1 -NoNinjaWrite -Simple`

- **Limit to specific classes:**
  `.\Report-AttachedPeripherals.ps1 -IncludeClasses Keyboard,Mouse,Monitor,USB`

## Default columns (enriched view)

| Column | Source |
|--------|--------|
| **Device Name** | Bus-reported description (`DEVPKEY_Device_BusReportedDeviceDesc`) if available, then `FriendlyName`, then `Description` from `Win32_PnPEntity`. This is the real product name the device advertises. |
| **Type** | PnP class mapped to a readable label (e.g. `MEDIA` → `Audio / Media`). |
| **Manufacturer** | From `Win32_PnPEntity.Manufacturer`. |
| **Connection** | Parsed from the `InstanceId` prefix (USB, PCI, Bluetooth, HD Audio, HID, etc.). |
| **Status** | PnP device status (filtered to `OK` by default). |
| **Hardware ID** | Short hardware identifier — `VID_xxxx&PID_xxxx` for USB devices, or the second segment of the first hardware ID. |
| **Serial Number** | Best-effort per-device serial when Windows exposes one. |
| **Serial Source** | Where serial came from: `WmiMonitorID`, `InstanceId`, or `None`. |

## Serial number discovery

Serial collection is **best effort** and hardware/driver dependent. The script keeps `Hardware ID` as a fallback identifier even when serial is missing.

Serial precedence:

1. **Monitor serial from `root\wmi:WmiMonitorID`** (highest confidence)
2. **USB instance-derived serial** from `InstanceId` (`USB\VID_xxxx&PID_yyyy\<serial>`)
3. **No serial** when neither source is usable

Notes and caveats:

- Some monitors do not expose a usable EDID serial through `WmiMonitorID`.
- Some docks enumerate as multiple child devices; serial may represent a child function rather than the whole dock.
- USB re-enumeration and generic hub/function nodes can yield unstable or missing instance-derived serials.
- Placeholder serials (all zeros, repeated characters, etc.) are filtered out.

## Filtered devices

By default, the following device name patterns are excluded to keep the report focused on real peripherals:

- WAN Miniport (*)
- Microsoft Kernel Debug Network Adapter
- Hyper-V Virtual *
- USB Root Hub (*)
- Generic USB Hub
- USB Composite Device
- * Host Controller *

Use `-IncludeVirtualDevices` to include them.

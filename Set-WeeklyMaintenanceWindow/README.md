# Set Weekly Maintenance Windows in NinjaOne from CSV

This folder contains two scripts that work together:

1. **Set-WeeklyMaintenanceWindow.ps1** – Sets weekly maintenance windows for organizations, locations, and devices in NinjaOne by importing values from a CSV via the NinjaOne API.
2. **Check-WeeklyMaintenanceWindow.ps1** – Checks whether the device's current local time is within or outside a recurring weekly maintenance window (for use in NinjaOne compound conditions).

---

## Set-WeeklyMaintenanceWindow.ps1

This script **sets** weekly maintenance windows by importing custom field values from a CSV. You define maintenance day, start time, and end time (and any other custom fields) in the CSV; the script updates the corresponding NinjaOne custom fields for each entity.

***Important Note*** that this script does not directly leverage the maintenance mode feature of NinjaOne. Instead, it imports the intended maintenance window into custom fields where the values can be used to logically determine script behavior. For instance, a pre-script could run to ensure that patching isn't conducted outside of a maintenance window.

Under the hood it imports custom field values at **organization**, **location**, and **device** levels. Each row is routed by a `level` column; the `name` column identifies the target. All other columns are treated as custom field name = value (e.g. `maintenanceDay`, `maintenanceStart`, `maintenanceEnd`).

### Requirements (Set script)

- **CSV columns**
  - `level` (required): One of `organization`, `location`, or `device`.
  - `name` (required): Identifier for the target.
  - All other columns: Custom field names. For maintenance windows, use columns such as `maintenanceDay`, `maintenanceStart`, `maintenanceEnd` (or whatever you named your NinjaOne custom fields). Cell values are written to those custom fields.

- **Location rows**: For `level = location`, `name` must be in the form **"organizationname,locationname"** (comma-separated, one column). Use quotes in the CSV if the value contains commas (e.g. `"Acme Corp,Main Office"`).

- **Organization rows**: `name` = organization name (matched to NinjaOne organization name).

- **Device rows**: `name` = device system name, or numeric device ID.

Custom fields must already exist in NinjaOne at the appropriate level (Administration > Organizations/Locations/Devices > Custom Fields). Column headers in the CSV must match the exact custom field names in NinjaOne (e.g. `maintenanceStart`, `maintenanceEnd`, `maintenanceDay`).

### Parameters (Set script)

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | Yes | Full path to the CSV file. |
| `-NinjaOneInstance` | No* | Instance host (e.g. `app.ninjarmm.com`). Defaults to env `NinjaOneInstance`. |
| `-NinjaOneClientId` | No* | OAuth client ID. Defaults to env `NinjaOneClientId`. |
| `-NinjaOneClientSecret` | No* | OAuth client secret. Defaults to env `NinjaOneClientSecret`. |
| `-OverwriteEmptyValues` | No | If `$true`, empty CSV values are sent so existing custom field values are cleared. Default: `$false`. |

\* Either pass credentials as parameters or set environment variables.

### Usage (Set script)

```powershell
.\Set-WeeklyMaintenanceWindow.ps1 -CsvPath "C:\data\maintenance-windows.csv" `
  -NinjaOneInstance "app.ninjarmm.com" `
  -NinjaOneClientId "your-client-id" `
  -NinjaOneClientSecret "your-client-secret"
```

With environment variables set:

```powershell
$env:NinjaOneInstance = "app.ninjarmm.com"
$env:NinjaOneClientId = "your-client-id"
$env:NinjaOneClientSecret = "your-client-secret"
.\Set-WeeklyMaintenanceWindow.ps1 -CsvPath ".\Import-CustomFields-Example.csv"
```

To clear existing values when a cell is empty:

```powershell
.\Set-WeeklyMaintenanceWindow.ps1 -CsvPath ".\maintenance-windows.csv" -OverwriteEmptyValues $true
```

### Example CSV (weekly maintenance window)

See `Import-CustomFields-Example.csv` for a sample that sets weekly maintenance windows:

- **organization**: `name` = organization name; include columns like `maintenanceDay`, `maintenanceStart`, `maintenanceEnd` for org-level maintenance windows.
- **location**: `name` = `"Acme Corp,Main Office"` (organization name, then location name, comma-separated); other columns are location-level custom fields (e.g. maintenance window).
- **device**: `name` = device system name (or device ID); other columns are device-level custom fields (e.g. maintenance window).

Rename the example columns to match your NinjaOne custom field names (`maintenanceStart`, `maintenanceEnd`, `maintenanceDay`, or whatever you created in NinjaOne).

### Behavior (Set script)

- Rows with invalid `level` or unresolved `name` are skipped and reported.
- Rows with no custom field columns (only `level` and `name`) are skipped.
- The script fetches all organizations, locations, and devices (with pagination where supported), then for each row resolves the target ID and sends a PATCH to the appropriate custom-fields endpoint.
- A short delay is used between PATCH requests to reduce rate-limit risk.
- At the end, a summary shows Updated, Skipped, and Failed counts.

---

## Check-WeeklyMaintenanceWindow.ps1

This script determines whether the device's current local time is **within** or **outside** a recurring weekly maintenance window. It is designed to work with maintenance windows set by **Set-WeeklyMaintenanceWindow.ps1** and can be used in NinjaOne compound conditions to control when automations run (e.g. only during maintenance, or only outside maintenance).

### Purpose (Check script)

- **Check in/out of window**: Returns exit code 0 when inside the window and exit code 1 when outside (configurable).
- **NinjaOne custom fields only**: Reads maintenance window values from NinjaOne custom fields (injected as environment variables when the script runs on a device). No script parameters—values are set by Set-WeeklyMaintenanceWindow.ps1 or the NinjaOne UI.
- **Two input formats**: HH:mm (e.g. `02:00`) or Unix milliseconds (as stored by Set-WeeklyMaintenanceWindow).

### Input (Check script – environment variables from NinjaOne custom fields)

| Environment variable | Required | Description                                                                 |
|----------------------|----------|-----------------------------------------------------------------------------|
| `maintenanceDay`     | Yes      | Day of week (e.g. Sunday, Monday).                                          |
| `maintenanceStart`   | Yes      | Start time as `HH:mm` or Unix milliseconds.                                |
| `maintenanceEnd`     | Yes      | End time as `HH:mm` or Unix milliseconds.                                  |
| `exitWhenInside`     | No       | If "true"/"1" (default), exit 0 when inside window. "false"/"0" to invert.  |

### Exit codes (Check script)

| Exit code | Meaning (when `ExitWhenInside` is true) |
|-----------|-----------------------------------------|
| 0         | Current time is **within** the maintenance window |
| 1         | Current time is **outside** the maintenance window |
| 2         | Error (missing/invalid parameters)                 |

When `ExitWhenInside` is false, exit 0 and 1 are inverted (0 = outside, 1 = inside).

### Usage (Check script)

The script takes no parameters. Run it on a device where NinjaOne supplies `maintenanceDay`, `maintenanceStart`, and `maintenanceEnd` from device, organization, or location custom fields (set by Set-WeeklyMaintenanceWindow.ps1 or the NinjaOne UI):

```powershell
.\Check-WeeklyMaintenanceWindow.ps1
```

To run when **not** in the maintenance window (exit 0 when outside, exit 1 when inside), set the `exitWhenInside` custom field to `false` or `0` in NinjaOne, or set the environment variable before running:

```powershell
$env:exitWhenInside = "false"; .\Check-WeeklyMaintenanceWindow.ps1
```

### Integration

Maintenance windows are typically populated by **Set-WeeklyMaintenanceWindow.ps1**. That script reads a CSV with columns `maintenanceDay`, `maintenanceStart`, and `maintenanceEnd` (e.g. `02:00`, `04:00`, `Sunday`) and writes them to NinjaOne custom fields at organization, location, or device level. On import, `maintenanceStart` and `maintenanceEnd` are converted to Unix milliseconds. Check-WeeklyMaintenanceWindow.ps1 accepts both the original HH:mm format and the Unix ms format returned by NinjaOne custom fields.

### Overnight windows (Check script)

Overnight windows (e.g. Sunday 22:00 - Monday 06:00) are supported. The script correctly treats the end time as the following calendar day when it is earlier than the start time.

---

## Prerequisites

- NinjaOne API OAuth application with **monitoring** and **management** scope (for Set-WeeklyMaintenanceWindow.ps1).
- Custom fields created in NinjaOne at the desired level(s) for maintenance window (and any other data), with names matching the CSV column headers (excluding `level` and `name`).



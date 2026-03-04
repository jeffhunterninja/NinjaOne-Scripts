# Set Maintenance Windows in NinjaOne from CSV

This folder contains two scripts that work together:

1. **Set-CustomFieldMaintenanceWindow.ps1** – Sets maintenance windows for organizations, locations, and devices in NinjaOne by importing values from a CSV via the NinjaOne API. Supports **Daily**, **Weekly**, and **Monthly** (nth weekday of month) recurrence. Run this locally on your workstation, or via an API server.
2. **Check-CustomFieldMaintenanceWindow.ps1** – Checks whether the device's current local time is within or outside a recurring maintenance window (Daily, Weekly, or Monthly) for use in NinjaOne compound conditions or with a pre-script.

---

## Custom fields to create in NinjaOne

All custom fields listed below must exist in NinjaOne at the desired level(s) before running the Set script or the Check script. Create them under **Administration > Organizations** (or **Locations** or **Devices**) **> Custom Fields**. CSV column headers and script expectations must match the exact custom field names in NinjaOne.

| Custom field name         | NinjaOne type   | Level(s)        | Allowed values / format                                                                 | Notes |
|---------------------------|-----------------|-----------------|------------------------------------------------------------------------------------------|-------|
| `maintenanceRecurrence`   | DROPDOWN| Org, Location, Device | `Daily`, `Weekly`, `Monthly` | For Daily/Weekly/Monthly logic. |
| `maintenanceDay`          | TEXT| Org, Location, Device | Day name: Sunday, Monday, Tuesday, etc. | Not used for Daily. |
| `maintenanceOccurrence`   | DROPDOWN| Org, Location, Device | `1`, `2`, `3`, `4`, or `Last` (e.g. 2 = 2nd Tuesday of month) | Only when recurrence is Monthly. |
| `maintenanceStart`        | TEXT        | Org, Location, Device | HH:mm in CSV (e.g. `02:00`); script sends Unix ms to API. Check accepts HH:mm, seconds from midnight UTC, or Unix ms. | Create as TIME in NinjaOne. |
| `maintenanceEnd`          | TEXT        | Org, Location, Device | HH:mm in CSV (e.g. `22:00`); script sends Unix ms to API. | Create as TIME in NinjaOne. |
| `exitWhenInside`          | DROPDOWN| Org, Location, Device | `true`/`1` (default): exit 0 when inside window. `false`/`0`: invert. | Check script only |

Create each field at every level (Organization, Location, and/or Device) where you will use it.


---

## Set-CustomFieldMaintenanceWindow.ps1

This script **sets** maintenance windows by importing custom field values from a CSV. You define recurrence type, optional day/occurrence, start time, and end time (and any other custom fields) in the CSV; the script updates the corresponding NinjaOne custom fields for each entity.

***Important Note*** that this script does not directly leverage the maintenance mode feature of NinjaOne. Instead, it imports the intended maintenance window into custom fields where the values can be used to logically determine script behavior. For instance, a pre-script could run to ensure that patching isn't conducted outside of a maintenance window.

Under the hood it imports custom field values at **organization**, **location**, and **device** levels. Each row is routed by a `level` column; the `name` column identifies the target. All other columns are treated as custom field name = value.

### Recurrence types (Set script)

| Recurrence | Description | CSV columns |
|------------|-------------|-------------|
| **Daily** | Window every day at the same start/end time | `maintenanceRecurrence=Daily`, `maintenanceStart`, `maintenanceEnd` (HH:mm). No `maintenanceDay`. |
| **Weekly** | Window on a specific day of week (default) | `maintenanceRecurrence=Weekly` or omit; `maintenanceDay`, `maintenanceStart`, `maintenanceEnd`. |
| **Monthly** | Window on nth weekday of month (e.g. 2nd Tuesday) | `maintenanceRecurrence=Monthly`, `maintenanceDay` (weekday), `maintenanceOccurrence` (1, 2, 3, 4, or Last), `maintenanceStart`, `maintenanceEnd`. |

If `maintenanceRecurrence` is missing, the Check script defaults to **Daily** when `maintenanceDay` is not set, and to **Weekly** when `maintenanceDay` is set.

### Requirements (Set script)

- **CSV columns**
  - `level` (required): One of `organization`, `location`, or `device`.
  - `name` (required): Identifier for the target.
  - `maintenanceRecurrence` (optional): `Daily`, `Weekly`, or `Monthly`. Default: `Weekly`.
  - `maintenanceDay` (Weekly/Monthly): Day of week (e.g. Sunday, Tuesday). Not used for Daily.
  - `maintenanceOccurrence` (Monthly only): `1`, `2`, `3`, `4`, or `Last` (e.g. 2 = 2nd occurrence of that weekday in the month).
  - `maintenanceStart`, `maintenanceEnd`: Start and end time as HH:mm (e.g. `02:00`, `22:00`). For all recurrence types (Daily, Weekly, Monthly), HH:mm values are converted to Unix milliseconds (time-of-day only, 1970-01-01 base) on import; NinjaOne TIME custom fields accept Unix ms.
  - All other columns: Any additional custom field names. Cell values are written to those custom fields.

- **Location rows**: For `level = location`, `name` must be in the form **"organizationname,locationname"** (comma-separated, one column). Use quotes in the CSV if the value contains commas (e.g. `"Acme Corp,Main Office"`).

- **Organization rows**: `name` = organization name (matched to NinjaOne organization name).

- **Device rows**: `name` = device system name, or numeric device ID.

Custom fields must already exist in NinjaOne at the appropriate level (Administration > Organizations/Locations/Devices > Custom Fields). Column headers in the CSV must match the exact custom field names in NinjaOne. For the new recurrence options, create **maintenanceRecurrence** (text/dropdown) and, for Monthly only, **maintenanceOccurrence** (text/dropdown or number) at the same level(s) where you use maintenance windows.

### Example CSV

See `Import-MaintenanceWindows-Example.csv` for samples of Daily, Weekly, and Monthly maintenance windows:

- **Daily**: `maintenanceRecurrence=Daily`, `maintenanceStart`, `maintenanceEnd` (no `maintenanceDay`).
- **Weekly**: `maintenanceRecurrence=Weekly` (or leave blank), `maintenanceDay`, `maintenanceStart`, `maintenanceEnd`.
- **Monthly**: `maintenanceRecurrence=Monthly`, `maintenanceDay` (weekday), `maintenanceOccurrence` (1, 2, 3, 4, or Last), `maintenanceStart`, `maintenanceEnd`.
- **location**: `name` must be `"organizationname,locationname"` (comma-separated).
- **device**: `name` = device system name or numeric device ID.

### Behavior (Set script)

- Rows with invalid `level` or unresolved `name` are skipped and reported.
- Rows with no custom field columns (only `level` and `name`) are skipped.
- For **Monthly** recurrence, if `maintenanceOccurrence` is missing or not 1, 2, 3, 4, or Last, the row is skipped with a warning.
- The script fetches all organizations, locations, and devices (with pagination where supported), then for each row resolves the target ID and sends a PATCH to the appropriate custom-fields endpoint.
- A short delay is used between PATCH requests to reduce rate-limit risk.
- At the end, a summary shows Updated, Skipped, and Failed counts.

---

## Check-CustomFieldMaintenanceWindow.ps1

This script determines whether the device's current local time is **within** or **outside** a recurring maintenance window (Daily, Weekly, or Monthly). It works with maintenance windows set by **Set-CustomFieldMaintenanceWindow.ps1** and can be used in NinjaOne compound conditions to control when automations run (e.g. only during maintenance, or only outside maintenance).

### Purpose (Check script)

- **Check in/out of window**: Returns exit code 0 when inside the window and exit code 1 when outside (configurable).
- **NinjaOne custom fields only**: Reads maintenance window values from NinjaOne custom fields (injected as environment variables when the script runs on a device). No script parameters—values are set by Set-CustomFieldMaintenanceWindow.ps1 or the NinjaOne UI.

### Input (Check script – environment variables from NinjaOne custom fields)

| Environment variable      | Required              | Description                                                                 |
|---------------------------|-----------------------|-----------------------------------------------------------------------------|
| `maintenanceRecurrence`   | No (when missing: Daily if no maintenanceDay; Weekly if maintenanceDay set) | `Daily`, `Weekly`, or `Monthly`.                                           |
| `maintenanceDay`          | Yes for Weekly/Monthly| Day of week (e.g. Sunday, Tuesday). Not used for Daily.                   |
| `maintenanceOccurrence`   | Yes for Monthly       | `1`, `2`, `3`, `4`, or `Last` (e.g. 2 = 2nd Tuesday of month).              |
| `maintenanceStart`        | Yes                   | Start time as `HH:mm`, seconds from midnight UTC, or Unix milliseconds.   |
| `maintenanceEnd`          | Yes                   | End time as `HH:mm`, seconds from midnight UTC, or Unix milliseconds.      |
| `exitWhenInside`          | No                    | If "true"/"1" (default), exit 0 when inside window. "false"/"0" to invert.  |

### Exit codes (Check script)

| Exit code | Meaning (when `ExitWhenInside` is true) |
|-----------|-----------------------------------------|
| 0         | Current time is **within** the maintenance window |
| 1         | Current time is **outside** the maintenance window |
| 2         | Error (missing/invalid parameters)                 |

When `ExitWhenInside` is false, exit 0 and 1 are inverted (0 = outside, 1 = inside).

### Usage (Check script)

The script takes no parameters. Run it on a device where NinjaOne supplies the maintenance window custom fields (e.g. `maintenanceRecurrence`, `maintenanceDay`, `maintenanceStart`, `maintenanceEnd`, and for Monthly `maintenanceOccurrence`) from device, organization, or location as set by Set-CustomFieldMaintenanceWindow.ps1.

### Overnight windows (Check script)

Overnight windows (e.g. Sunday 22:00 - Monday 06:00) are supported. The script correctly treats the end time as the following calendar day when it is earlier than the start time.

---

## Prerequisites

- NinjaOne API OAuth application with **monitoring** and **management** scope and **client credentials** grant type (for Set-CustomFieldMaintenanceWindow.ps1).
- Custom fields created in NinjaOne at the desired level(s) for maintenance window (and any other data), with names matching the CSV column headers (excluding `level` and `name`). For Daily/Monthly support, create **maintenanceRecurrence** and, for Monthly, **maintenanceOccurrence** at the same level(s).

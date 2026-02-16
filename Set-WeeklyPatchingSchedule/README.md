# Set Weekly Patching Schedules in NinjaOne from CSV

This folder contains two scripts that work together:

1. **Set-WeeklyPatchingSchedule.ps1** – Sets weekly patching schedules for organizations, locations, and devices in NinjaOne by importing values from a CSV via the NinjaOne API.
2. **Check-WeeklyPatchingSchedule.ps1** – Determines whether the device should patch now based on the weekly patching schedule. If the script runs shortly before the patch time (within the hold window), it waits until the exact start time, then exits 0. For use in NinjaOne compound conditions.

---

## Set-WeeklyPatchingSchedule.ps1

This script **sets** weekly patching schedules by importing custom field values from a CSV. You define patching day and start time (and any other custom fields) in the CSV; the script updates the corresponding NinjaOne custom fields for each entity.

Under the hood it imports custom field values at **organization**, **location**, and **device** levels. Each row is routed by a `level` column; the `name` column identifies the target. All other columns are treated as custom field name = value (e.g. `patchingDay`, `patchingStart`).

### Requirements (Set script)

- **CSV columns**
  - `level` (required): One of `organization`, `location`, or `device`.
  - `name` (required): Identifier for the target.
  - All other columns: Custom field names. For patching schedules, use columns such as `patchingDay`, `patchingStart` (or whatever you named your NinjaOne custom fields). Cell values are written to those custom fields.

- **Location rows**: For `level = location`, `name` must be in the form **"organizationname,locationname"** (comma-separated, one column). Use quotes in the CSV if the value contains commas (e.g. `"Acme Corp,Main Office"`).

- **Organization rows**: `name` = organization name (matched to NinjaOne organization name).

- **Device rows**: `name` = device system name, or numeric device ID.

Custom fields must already exist in NinjaOne at the appropriate level (Administration > Organizations/Locations/Devices > Custom Fields). Column headers in the CSV must match the exact custom field names in NinjaOne (e.g. `patchingStart`, `patchingDay`).

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
.\Set-WeeklyPatchingSchedule.ps1 -CsvPath "C:\data\patching-schedule.csv" `
  -NinjaOneInstance "app.ninjarmm.com" `
  -NinjaOneClientId "your-client-id" `
  -NinjaOneClientSecret "your-client-secret"
```

With environment variables set:

```powershell
$env:NinjaOneInstance = "app.ninjarmm.com"
$env:NinjaOneClientId = "your-client-id"
$env:NinjaOneClientSecret = "your-client-secret"
.\Set-WeeklyPatchingSchedule.ps1 -CsvPath ".\Import-PatchSchedules-Example.csv"
```

To clear existing values when a cell is empty:

```powershell
.\Set-WeeklyPatchingSchedule.ps1 -CsvPath ".\patching-schedule.csv" -OverwriteEmptyValues $true
```

### Example CSV (weekly patching schedule)

See `Import-PatchSchedules-Example.csv` for a sample:

- **organization**: `name` = organization name; include columns like `patchingDay`, `patchingStart` for org-level patching schedules.
- **location**: `name` = `"Acme Corp,Main Office"` (organization name, then location name, comma-separated); other columns are location-level custom fields.
- **device**: `name` = device system name (or device ID); other columns are device-level custom fields.

Rename the example columns to match your NinjaOne custom field names (`patchingStart`, `patchingDay`, or whatever you created in NinjaOne). No end time is used; patching occurs at the exact start time.

### Behavior (Set script)

- Rows with invalid `level` or unresolved `name` are skipped and reported.
- Rows with no custom field columns (only `level` and `name`) are skipped.
- The script fetches all organizations, locations, and devices (with pagination where supported), then for each row resolves the target ID and sends a PATCH to the appropriate custom-fields endpoint.
- `patchingStart` values in HH:mm format are converted to Unix milliseconds on import.
- A short delay is used between PATCH requests to reduce rate-limit risk.
- At the end, a summary shows Updated, Skipped, and Failed counts.

---

## Check-WeeklyPatchingSchedule.ps1

This script determines whether the device **should patch now** based on a weekly patching schedule. It is designed to work with schedules set by **Set-WeeklyPatchingSchedule.ps1** and can be used in NinjaOne compound conditions to control when patching runs.

### Purpose (Check script)

- **Exact-time patching**: Patching should occur at the exact start time. If the script runs shortly before that time (within the hold window), it waits until the exact time, then exits 0.
- **Hold mechanism**: Because NinjaOne runs scripts on a schedule (e.g., every 15 minutes), the script may run before the patch time. Set `holdWindowMinutes` to match your schedule interval. For example, if the script runs every 15 min and patch is at 02:00, a run at 01:45 will wait 15 min until 02:00, then exit 0.
- **NinjaOne custom fields only**: Reads values from NinjaOne custom fields (injected as environment variables when the script runs on a device). No script parameters.
- **Two input formats**: HH:mm (e.g. `02:00`) or Unix milliseconds (as stored by Set-WeeklyPatchingSchedule).

### Input (Check script – environment variables from NinjaOne custom fields)

| Environment variable | Required | Description |
|----------------------|----------|-------------|
| `patchingDay`        | Yes      | Day of week (e.g. Sunday, Monday). |
| `patchingStart`      | Yes      | Exact start time as `HH:mm` or Unix milliseconds. |
| `holdWindowMinutes`  | No       | Max minutes before patching start to wait. Should match NinjaOne schedule interval (e.g., 15 if script runs every 15 min). Default: 15. |
| `exitWhenShouldPatch`| No       | If "true"/"1" (default), exit 0 when device should patch. "false"/"0" to invert. |

### Exit codes (Check script)

| Exit code | Meaning (when `exitWhenShouldPatch` is true) |
|-----------|-----------------------------------------------|
| 0         | Device **should patch** (at or past start time, or held until start) |
| 1         | Device **should not patch** (wrong day, or before start and outside hold window) |
| 2         | Error (missing/invalid parameters)            |

When `exitWhenShouldPatch` is false, exit 0 and 1 are inverted (0 = should not patch, 1 = should patch).

### Logic

1. **Not patching day** → exit 1 (do not patch)
2. **Patching day, current time >= patchingStart** → exit 0 immediately (proceed to patch)
3. **Patching day, current time < patchingStart, within hold window** → `Start-Sleep` until exact patching start time, then exit 0 (proceed)
4. **Patching day, current time < patchingStart, outside hold window** → exit 1 (do not patch; next scheduled run will be closer)

### Usage (Check script)

The script takes no parameters. Run it on a device where NinjaOne supplies `patchingDay` and `patchingStart` from device, organization, or location custom fields (set by Set-WeeklyPatchingSchedule.ps1 or the NinjaOne UI):

```powershell
.\Check-WeeklyPatchingSchedule.ps1
```

Set `holdWindowMinutes` in NinjaOne to match your script schedule interval. For example, if Check-WeeklyPatchingSchedule runs every 15 minutes:

```powershell
$env:holdWindowMinutes = "15"; .\Check-WeeklyPatchingSchedule.ps1
```

To run when **not** supposed to patch (exit 0 when should not patch, exit 1 when should patch), set the `exitWhenShouldPatch` custom field to `false` or `0` in NinjaOne:

```powershell
$env:exitWhenShouldPatch = "false"; .\Check-WeeklyPatchingSchedule.ps1
```

### Integration

Patching schedules are typically populated by **Set-WeeklyPatchingSchedule.ps1**. That script reads a CSV with columns `patchingDay` and `patchingStart` (e.g. `02:00`, `Sunday`) and writes them to NinjaOne custom fields at organization, location, or device level. On import, `patchingStart` is converted to Unix milliseconds. Check-WeeklyPatchingSchedule.ps1 accepts both the original HH:mm format and the Unix ms format returned by NinjaOne custom fields.

---

## Prerequisites

- NinjaOne API OAuth application with **monitoring** and **management** scope (for Set-WeeklyPatchingSchedule.ps1).
- Custom fields created in NinjaOne at the desired level(s) for patching schedule (`patchingDay`, `patchingStart`, and optionally `holdWindowMinutes`), with names matching the CSV column headers (excluding `level` and `name`).

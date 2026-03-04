# Set Patching Schedules (Daily, Weekly, Monthly) in NinjaOne from CSV

This folder contains two scripts that work together:

1. **Set-CustomFieldPatchingSchedule.ps1** – Sets patching schedules for organizations, locations, and devices in NinjaOne by importing values from a CSV via the NinjaOne API. Supports **Daily**, **Weekly**, and **Monthly** (nth weekday of month) recurrence.
2. **Check-CustomFieldPatchingSchedule.ps1** – Determines whether the device should patch now based on the patching schedule (Daily, Weekly, or Monthly). If the script runs shortly before the patch time (within the hold window), it waits until the exact start time, then exits 0. For use in NinjaOne as a script result condition

***Important Note***: **Set-CustomFieldPatchingSchedule.ps1** is not intended to be executed or scheduled through NinjaOne. Instead, run it on your workstation for the initial setup and any subsequent changes, or leverage the API server framework if you need a shared or automated method for updating the patching times more regularly.

---

## Custom fields to create in NinjaOne

All custom fields listed below must exist in NinjaOne at the desired level(s) before running the Set script or the Check script. Create them under **Administration > Organizations** (or **Locations** or **Devices**) **> Custom Fields**. CSV column headers and script expectations must match the exact custom field names in NinjaOne.

| Custom field name      | Type   | Inheritance        | Allowed values / format                                                                 | Notes |
|-----------------------|-----------------|-----------------|------------------------------------------------------------------------------------------|-------|
| `patchingRecurrence`  | TEXT or DROPDOWN| Org, Location, Device | `Daily`, `Weekly`, `Monthly` | For Daily/Weekly/Monthly logic. |
| `patchingDay`         | TEXT or DROPDOWN| Org, Location, Device | Day name: Sunday, Monday, Tuesday, Wednesday, Thursday, Friday, Saturday | Not used for Daily. |
| `patchingOccurrence`  | TEXT or DROPDOWN | Org, Location, Device | `1`, `2`, `3`, `4`, or `Last` (e.g. 2 = 2nd Tuesday of month) | Only when recurrence is Monthly. |
| `patchingStart`       | **TEXT** | Org, Location, Device | Store HH:mm (e.g. `15:33`); run Set script with `-PatchingStartAsLocalTime`; Check uses it as device local time (same wall-clock time on every device). | Use TIME for one global patch time; use TEXT + `-PatchingStartAsLocalTime` for "3:33 PM in each device's timezone". |
| `disablePatching`     | **CHECKBOX**    | Org, Location, Device | In CSV: `true`/`false`, `1`/`0`, or `yes`/`no`. When checked/true, Check script exits 1 and patching is skipped. | Create as CHECKBOX if used. |

Column headers in your CSV must match the exact custom field names (not label) in the table (excluding the reserved `level` and `name` columns).

### API

Create a client app ID with **monitoring** and **management** scopes that uses the **client credentials** grant type.

---

## Set-CustomFieldPatchingSchedule.ps1

This script **sets** patching schedules by importing custom field values from a CSV. You define recurrence type, optional day/occurrence, start time (and any other custom fields) in the CSV; the script updates the corresponding NinjaOne custom fields for each entity.

Under the hood it imports custom field values at **organization**, **location**, and **device** levels. Each row is routed by a `level` column; the `name` column identifies the target. All other columns are treated as custom field name = value (e.g. `patchingRecurrence`, `patchingDay`, `patchingStart`).

### Recurrence types (Set script)

| Recurrence | Description | CSV columns |
|------------|--------------|--------------|
| **Daily** | Patch every day at the same start time | `patchingRecurrence=Daily`, `patchingStart` (HH:mm). No `patchingDay`. |
| **Weekly** | Patch on a specific day of week (default) | `patchingRecurrence=Weekly` or omit; `patchingDay`, `patchingStart`. |
| **Monthly** | Patch on nth weekday of month (e.g. 2nd Tuesday) | `patchingRecurrence=Monthly`, `patchingDay` (weekday), `patchingOccurrence` (1, 2, 3, 4, or Last), `patchingStart`. |

If `patchingRecurrence` is missing, the Check script defaults to **Daily** when `patchingDay` is not set, and to **Weekly** when `patchingDay` is set.

### Requirements (Set script)

- **CSV columns**
  - `level` (required): One of `organization`, `location`, or `device`.
  - `name` (required): Identifier for the target.
  - All other columns: Custom field names. For patching schedules, use columns such as `patchingRecurrence`, `patchingDay`, `patchingOccurrence` (Monthly only), `patchingStart`. Cell values are written to those custom fields.

- **Location rows**: For `level = location`, `name` must be in the form **"organizationname,locationname"** (comma-separated, one column). Use quotes in the CSV if the value contains commas (e.g. `"Acme Corp,Main Office"`).

- **Organization rows**: `name` = organization name (matched to NinjaOne organization name).

- **Device rows**: `name` = device system name, or numeric device ID.

Custom fields must already exist in NinjaOne at the appropriate level. Column headers in the CSV must match the exact custom field names in NinjaOne. See the [Custom fields to create in NinjaOne](#custom-fields-to-create-in-ninjaone) table for types, levels, and permissions.

### Parameters (Set script)

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | Yes | Full path to the CSV file. |
| `-NinjaOneInstance` | No* | Instance host (e.g. `app.ninjarmm.com`). Defaults to env `NinjaOneInstance`. |
| `-NinjaOneClientId` | No* | OAuth client ID. Defaults to env `NinjaOneClientId`. |
| `-NinjaOneClientSecret` | No* | OAuth client secret. Defaults to env `NinjaOneClientSecret`. |
| `-OverwriteEmptyValues` | No | If `$true`, empty CSV values are sent so existing custom field values are cleared. Default: `$false`. |
| `-PatchingStartAsLocalTime` | No | If set, `patchingStart` is sent as HH:MM string (no conversion to Unix ms). Use with a **TEXT** custom field so each device patches at the same local time (e.g. 3:33 PM in each device's timezone). When not set (default), HH:MM is converted to Unix ms for a TIME field (same UTC moment globally). |

\* Either pass credentials as parameters or set environment variables.


### Local device time (same wall-clock time on every device)

To have each device patch at the **same local time** (e.g. 3:33 PM in each device's timezone) instead of one UTC moment:

1. Create `patchingStart` as a **TEXT** custom field in NinjaOne (not TIME).
2. Put the start time as HH:MM in the CSV (e.g. `15:33` for 3:33 PM).
3. Run the Set script with **`-PatchingStartAsLocalTime`** so the value is sent as-is (no conversion to Unix ms).

The Check script treats HH:mm as device local time, so patching will run at that clock time on each device.

```powershell
.\Set-CustomFieldPatchingSchedule.ps1 -CsvPath ".\patching-schedule.csv" -PatchingStartAsLocalTime $true
```

Existing setups that use a TIME field and omit the switch are unchanged (HH:MM is still converted to Unix ms).

### Example CSV (patching schedule with recurrence)

See `Import-PatchSchedules-Example.csv` for samples:

- **organization**: `name` = organization name; include columns like `patchingRecurrence`, `patchingDay`, `patchingOccurrence` (Monthly), `patchingStart`.
- **location**: `name` = `"Acme Corp,Main Office"` (organization name, then location name, comma-separated); other columns are location-level custom fields.
- **device**: `name` = device system name (or device ID); other columns are device-level custom fields.

Rename the example columns to match your NinjaOne custom field names. No end time is used; patching occurs at the exact start time. **TIME fields** (e.g. `patchingStart`): provide HH:MM (e.g. 02:00) in the CSV; the script converts them to Unix milliseconds before sending (NinjaOne TIME type expects Unix ms). Values that are already numeric (Unix ms) are passed through unchanged.

### Behavior (Set script)

- Rows with invalid `level` or unresolved `name` are skipped and reported.
- Rows with no custom field columns (only `level` and `name`) are skipped.
- For **Monthly** recurrence, if `patchingOccurrence` is missing or not 1, 2, 3, 4, or Last, the row is skipped with a warning.
- **patchingStart**: When `-PatchingStartAsLocalTime` is not set, HH:MM is converted to Unix milliseconds for TIME fields; numeric values are sent as-is. When `-PatchingStartAsLocalTime` is set, the value is sent as HH:MM string for use with a TEXT custom field (device local time).
- The script fetches all organizations, locations, and devices (with pagination where supported), then for each row resolves the target ID and sends a PATCH to the appropriate custom-fields endpoint.
- A short delay is used between PATCH requests to reduce rate-limit risk.
- At the end, a summary shows Updated, Skipped, and Failed counts.

---

## Check-CustomFieldPatchingSchedule.ps1

This script determines whether the device **should patch now** based on a patching schedule (Daily, Weekly, or Monthly). It is designed to work with schedules set by **Set-CustomFieldPatchingSchedule.ps1** and can be used in NinjaOne script result conditions to control when patching runs.

It is extremely important to align your `holdWindowMinutes` property in the script with the interval and timeout of the script result condition. Because NinjaOne runs scripts on a schedule (e.g., every 15 minutes), the script may run before the patch time. Set `holdWindowMinutes` to match your schedule interval, and set your timeout to be one minute longer. For example, if the script runs every 15 min and patch is at 02:00, a run at 01:45 or later will wait until 02:00, then exit 0. The script must have enough of a timeout period to prevent the hold window from being compromised.

### Exit codes (Check script)

| Exit code | Meaning (when `exitWhenShouldPatch` is true) |
|-----------|-----------------------------------------------|
| 0         | Device **should patch** (at or past start time, or held until start) |
| 1         | Device **should not patch** (wrong day, or before start and outside hold window) |
| 2         | Error (missing/invalid parameters)            |

When `exitWhenShouldPatch` is false, exit 0 and 1 are inverted (0 = should not patch, 1 = should patch).

### Logic

The script computes the **next patching occurrence** from the recurrence type (Daily = today/tomorrow at start time; Weekly = next occurrence of that weekday; Monthly = this or next month's nth weekday). Then:

1. **Current time >= next patching occurrence** → exit 1 (past patch time; run before start within hold window to trigger).
2. **Current time < next patching occurrence, within hold window** → `Start-Sleep` until exact patching start time, then exit 0 (proceed).
3. **Current time < next patching occurrence, outside hold window** → exit 1 (do not patch; next scheduled run will be closer).

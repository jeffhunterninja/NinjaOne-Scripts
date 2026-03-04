# Recurring Maintenance Mode (NinjaOne)

These scripts schedule recurring maintenance windows for NinjaOne devices using device, location, and organization custom fields and the NinjaOne API. The main script evaluates schedules, computes the next window, and sets maintenance mode (start/end, disabled features). The import script bulk-updates those custom fields from CSV.

This workflow, specifcally **Set-RecurringMaintenanceMode**, is designed to use the API server framework. These are not device scripts and should not be deployed to run on each managed endpoint. For API server setup, see [Getting started with an API server](https://docs.mspp.io/ninjaone-auto-documentation/getting-started); the first part of [this video](https://www.youtube.com/watch?v=Qy9g6-KVfbo) walks through setting up an API server.

## Prerequisites

- **PowerShell:** 5.1 or later (both scripts use `#Requires -Version 5.1`).
- **API server (recommended):** Set up an API server/automated documentation server per the link above.
- **NinjaOne API:** A Client ID and Client Secret with **monitoring** and **management** scopes (machine-to-machine, client credentials). Create under **Administration > Apps > API > Client App IDs**: add an app, choose API Services (machine-to-machine), select the scopes above, and allow client credentials only. The **client secret is shown only once**—copy and store it securely.
- **Custom fields:** All recurring-maintenance custom fields must exist at **Device**, **Location**, and **Organization** scope with API **Read/Write**. See [Custom fields](#custom-fields) below for the full list.

## Scripts overview

| Script | Role |
|--------|------|
| **Set-RecurringMaintenanceMode.ps1** | Run on a schedule (e.g. from NinjaOne or Task Scheduler). Connects to NinjaOne, finds devices/orgs/locations with recurring maintenance enabled, computes the next maintenance window per schedule type, and sets maintenance mode via the API. Writes the last result to `recurringMaintenanceLastResult`. |
| **Import-RecurringMaintenanceFromCsv.ps1** | One-off or bulk: imports a CSV and PATCHes recurring maintenance custom fields on devices, organizations, or locations. Uses the same credential variables as the main script. |

## Configuration

Set these variables at the top of each script (or ensure they are available where the script runs):

| Variable | Script(s) | Description |
|----------|-----------|-------------|
| `$NinjaOneInstance` | Both | NinjaOne hostname only (e.g. `app.ninjarmm.com`), no path. |
| `$NinjaOneClientId` | Both | API Client ID from NinjaOne. |
| `$NinjaOneClientSecret` | Both | API Client Secret from NinjaOne. |
| `$SkipCustomFieldTest` | Set-RecurringMaintenanceMode.ps1 only | If `$true`, skips validation that required custom fields exist. Optional; default `$false`. |

## Custom fields

All listed fields must exist in NinjaOne at **Device**, **Location**, and **Organization** scope with **API Permissions** set to Read/Write. Which fields you *fill in* depends on the schedule type—see [Examples by schedule type](#examples-by-schedule-type) below. For exact labels and tooltips to create them in NinjaOne, see the script header in `Set-RecurringMaintenanceMode.ps1` (lines 30–78).

| Purpose | Field name | Type | Required for | Notes |
|--------|------------|------|---------------|--------|
| Enable schedule | `recurringMaintenanceEnableRecurringSchedule` | Checkbox | All | Must be true for the entity to be processed. |
| Start time (24h) | `recurringMaintenanceTimeToStart24hFormat` | Time | All | 24-hour format (NinjaOne stores as Unix ms). |
| Duration (minutes) | `recurringMaintenanceTotalMinutesForMaintenanceMode` | Integer | All | 1–1440 (max 24 hours). |
| Stop applying after | `recurringMaintenanceDateToStopApplyingRecurringSchedule` | Date | All | Optional to fill; after this date the script will not set new windows. |
| Suppress scripting/tasks | `recurringMaintenanceSuppressScriptingAndTasks` | Checkbox | All | Disable tasks during maintenance. |
| Suppress patching | `recurringMaintenanceSuppressPatching` | Checkbox | All | Disable patching during maintenance. |
| Suppress AV scans | `recurringMaintenanceSuppressAvScans` | Checkbox | All | Disable AV scans during maintenance. |
| Suppress condition-based alerting | `recurringMaintenanceSuppressConditionBasedAlerting` | Checkbox | All | Disable condition-based alerts during maintenance. |
| Last result (output) | `recurringMaintenanceLastResult` | Text (Multiline) | All | Set by the script; do not configure manually. |
| Schedule type | `recurringMaintenanceScheduleType` | Single-select | All | Daily, Weekly, Monthly, or MonthlyDayOfWeek. |
| Legacy weekly days | `recurringMaintenanceSelectDay` | Multi-select | Weekly | "Every Sunday", "Every Monday", … "Every Saturday". Use this or `recurringMaintenanceDayOfWeek` for Weekly. |
| Day of week (weekly) | `recurringMaintenanceDayOfWeek` | Multi-select | Weekly | Sunday, Monday, … Saturday. Use this or `recurringMaintenanceSelectDay` for Weekly. |
| Day of month (monthly) | `recurringMaintenanceDayOfMonth` | Integer | Monthly | 1–31. |
| Monthly weekday | `recurringMaintenanceMonthlyDayOfWeek` | Single-select | MonthlyDayOfWeek | Sunday … Saturday. |
| Monthly occurrence | `recurringMaintenanceMonthlyOccurrence` | Single-select | MonthlyDayOfWeek | 1, 2, 3, 4, or Last (e.g. 2nd Tuesday). |

## Schedule types

| Type | Behavior |
|------|----------|
| **Daily** | Every day at the configured start time. |
| **Weekly** | On selected weekdays (via `recurringMaintenanceDayOfWeek` or legacy `recurringMaintenanceSelectDay`). |
| **Monthly** | Same day of month (1–31) each month. |
| **MonthlyDayOfWeek** | Nth occurrence of a weekday in the month (e.g. 2nd Tuesday) or "Last". |

Values are resolved with inheritance: **device ← location ← organization**. The main script uses scoped custom fields and device-level resolution with inheritance. See [Examples by schedule type](#examples-by-schedule-type) for which fields to set for each type.

### Examples by schedule type

For each schedule type, set the fields listed below. All modes also require the common fields: `recurringMaintenanceEnableRecurringSchedule` (true), `recurringMaintenanceTimeToStart24hFormat`, `recurringMaintenanceTotalMinutesForMaintenanceMode`, and any suppression checkboxes you want. Optionally set `recurringMaintenanceDateToStopApplyingRecurringSchedule` to stop applying the schedule after a date.

**Daily**

- **Set:** `recurringMaintenanceScheduleType` = `Daily`. No day-of-week or day-of-month fields.
- **Example:** Start time 02:00 (2 AM), duration 60 minutes. Maintenance runs every day at 2 AM for 1 hour.

**Weekly**

- **Set:** `recurringMaintenanceScheduleType` = `Weekly`, and **either** `recurringMaintenanceDayOfWeek` **or** `recurringMaintenanceSelectDay`.
- **Example:** Schedule type `Weekly`; day of week `Monday`, `Wednesday` (or legacy `Every Monday`, `Every Wednesday`); start time 03:00; duration 120. Maintenance runs every Monday and Wednesday at 3 AM for 2 hours.

**Monthly**

- **Set:** `recurringMaintenanceScheduleType` = `Monthly`, and `recurringMaintenanceDayOfMonth` (1–31).
- **Example:** Schedule type `Monthly`; day of month `1` or `15`; start time 01:00; duration 180. Maintenance runs on the 1st (or 15th) of each month at 1 AM for 3 hours.

**MonthlyDayOfWeek**

- **Set:** `recurringMaintenanceScheduleType` = `MonthlyDayOfWeek`, `recurringMaintenanceMonthlyDayOfWeek` (weekday name), and `recurringMaintenanceMonthlyOccurrence` (1, 2, 3, 4, or Last).
- **Example:** Schedule type `MonthlyDayOfWeek`; monthly day of week `Tuesday`; monthly occurrence `2` (second Tuesday); start time 00:00; duration 240. Maintenance runs on the second Tuesday of each month at midnight for 4 hours.

## Running the main script

- **Manually:** Set `$NinjaOneInstance`, `$NinjaOneClientId`, and `$NinjaOneClientSecret` in the script (or dot-source a config), then run `Set-RecurringMaintenanceMode.ps1`. It has no parameters.
- **Scheduling:** Run it frequently (e.g. every 5–15 minutes) so the next maintenance window is set before the start time. You can run it as a NinjaOne script assignment or from Task Scheduler; credentials must be available where the script runs (e.g. script variables or secure storage).
- **Output:** At the end the script prints counts (devices marked for maintenance, already in maintenance, already scheduled, actually set, skipped past apply date) and any warning messages from the error log.

## CSV import (Import-RecurringMaintenanceFromCsv.ps1)

Bulk-update recurring maintenance custom fields from a CSV.

### Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | Yes | Path to the CSV file. |
| `-OverwriteEmptyValues` | No | If set, empty cells are sent as `$null` and clear existing values. |
| `-WhatIf` | No | Report what would be updated without calling the API. |

### CSV format

- **Headers** = custom field API names (e.g. `recurringMaintenanceEnableRecurringSchedule`, `recurringMaintenanceScheduleType`).
- **Reserved columns** (not sent as custom fields): `id` or `deviceId` for devices; `scope` (Device, Organization, or Location); `organizationId` for Organization scope; `locationId` for Location scope; `systemName`, `organizationName`, `locationName` for name-based lookups.
- **Name-based identification:** You can supply `systemName` (device), `organizationName` (organization), or `locationName` (location) instead of the corresponding ID. The script resolves names to IDs via the NinjaOne API (case-insensitive). If multiple entities share the same name, the script reports an error for that row. For large tenants, ID-based rows are faster than name-based device rows.
- **Multi-select fields:** Use comma- or semicolon-separated values in a single cell; the script sends them as arrays.
- **Time field:** For `recurringMaintenanceTimeToStart24hFormat`, use the same format NinjaOne stores (e.g. Unix milliseconds). The example file uses a placeholder—replace or confirm against an existing device value in NinjaOne.

See `recurring-maintenance-import-example.csv` for an example with device and optional organization/location rows, including rows that use IDs and rows that use systemName, organizationName, or locationName (Weekly, Daily, Monthly, MonthlyDayOfWeek).

### Examples

Device-only CSV with `id` and custom field columns:

```powershell
.\Import-RecurringMaintenanceFromCsv.ps1 -CsvPath .\maintenance-import.csv
```

Preview updates and clear empty values:

```powershell
.\Import-RecurringMaintenanceFromCsv.ps1 -CsvPath .\maintenance-import.csv -OverwriteEmptyValues -WhatIf
```

For mixed scope (Device, Organization, Location), include a `scope` column and the appropriate ID or name column (`organizationId`/`organizationName`, `locationId`/`locationName`, or `deviceId`/`id`/`systemName`). See the comment-based help in `Import-RecurringMaintenanceFromCsv.ps1` for the full synopsis and examples.

## Attribution

Recurring maintenance logic and authentication: **Gavin Stone** (NinjaOne); authentication functions: **Luke Whitelock** (NinjaOne); PS 5.1 support: **Kyle Bohlander** (NinjaOne). See the header in `Set-RecurringMaintenanceMode.ps1` for version and date.

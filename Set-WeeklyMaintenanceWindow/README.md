# Set Weekly Maintenance Windows in NinjaOne from CSV

This script **sets weekly maintenance windows** for organizations, locations, and devices in NinjaOne by importing values from a CSV. You define maintenance day, start time, and end time (and any other custom fields) in the CSV; the script updates the corresponding NinjaOne custom fields for each entity.

Under the hood it imports custom field values at **organization**, **location**, and **device** levels. Each row is routed by a `level` column; the `name` column identifies the target. All other columns are treated as custom field name = value (e.g. `maintenanceDay`, `maintenanceStart`, `maintenanceEnd`).

## Requirements

- **CSV columns**
  - `level` (required): One of `organization`, `location`, or `device`.
  - `name` (required): Identifier for the target.
  - All other columns: Custom field names. For maintenance windows, use columns such as `maintenanceDay`, `maintenanceStart`, `maintenanceEnd` (or whatever you named your NinjaOne custom fields). Cell values are written to those custom fields.

- **Location rows**: For `level = location`, `name` must be in the form **"organizationname,locationname"** (comma-separated, one column). Use quotes in the CSV if the value contains commas (e.g. `"Acme Corp,Main Office"`).

- **Organization rows**: `name` = organization name (matched to NinjaOne organization name).

- **Device rows**: `name` = device system name, or numeric device ID.

Custom fields must already exist in NinjaOne at the appropriate level (Administration > Organizations/Locations/Devices > Custom Fields). Column headers in the CSV must match the exact custom field names in NinjaOne (e.g. `maintenanceStart`, `maintenanceEnd`, `maintenanceDay`).

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | Yes | Full path to the CSV file. |
| `-NinjaOneInstance` | No* | Instance host (e.g. `app.ninjarmm.com`). Defaults to env `NinjaOneInstance`. |
| `-NinjaOneClientId` | No* | OAuth client ID. Defaults to env `NinjaOneClientId`. |
| `-NinjaOneClientSecret` | No* | OAuth client secret. Defaults to env `NinjaOneClientSecret`. |
| `-OverwriteEmptyValues` | No | If `$true`, empty CSV values are sent so existing custom field values are cleared. Default: `$false`. |

\* Either pass credentials as parameters or set environment variables.

## Usage

```powershell
.\Import-NinjaOneCustomFieldsFromCsv.ps1 -CsvPath "C:\data\maintenance-windows.csv" `
  -NinjaOneInstance "app.ninjarmm.com" `
  -NinjaOneClientId "your-client-id" `
  -NinjaOneClientSecret "your-client-secret"
```

With environment variables set:

```powershell
$env:NinjaOneInstance = "app.ninjarmm.com"
$env:NinjaOneClientId = "your-client-id"
$env:NinjaOneClientSecret = "your-client-secret"
.\Import-NinjaOneCustomFieldsFromCsv.ps1 -CsvPath ".\Import-CustomFields-Example.csv"
```

To clear existing values when a cell is empty:

```powershell
.\Import-NinjaOneCustomFieldsFromCsv.ps1 -CsvPath ".\maintenance-windows.csv" -OverwriteEmptyValues $true
```

## Example CSV (weekly maintenance window)

See `Import-CustomFields-Example.csv` for a sample that sets weekly maintenance windows:

- **organization**: `name` = organization name; include columns like `maintenanceDay`, `maintenanceStart`, `maintenanceEnd` for org-level maintenance windows.
- **location**: `name` = `"Acme Corp,Main Office"` (organization name, then location name, comma-separated); other columns are location-level custom fields (e.g. maintenance window).
- **device**: `name` = device system name (or device ID); other columns are device-level custom fields (e.g. maintenance window).

Rename the example columns to match your NinjaOne custom field names (`maintenanceStart`, `maintenanceEnd`, `maintenanceDay`, or whatever you created in NinjaOne).

## Behavior

- Rows with invalid `level` or unresolved `name` are skipped and reported.
- Rows with no custom field columns (only `level` and `name`) are skipped.
- The script fetches all organizations, locations, and devices (with pagination where supported), then for each row resolves the target ID and sends a PATCH to the appropriate custom-fields endpoint.
- A short delay is used between PATCH requests to reduce rate-limit risk.
- At the end, a summary shows Updated, Skipped, and Failed counts.

## Prerequisites

- NinjaOne API OAuth application with **monitoring** and **management** scope.
- Custom fields created in NinjaOne at the desired level(s) for maintenance window (and any other data), with names matching the CSV column headers (excluding `level` and `name`).

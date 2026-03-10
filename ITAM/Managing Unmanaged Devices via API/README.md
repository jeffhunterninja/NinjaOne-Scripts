# Unmanaged Device Import

Bulk-create NinjaOne ITAM unmanaged devices from a CSV file. Use this workflow after receiving and reconciling equipment (e.g. mouse, keyboard, headset, dock, monitors) to import items into NinjaOne from a simple spreadsheet.

## Prerequisites

- **PowerShell 5.1 or later**
- **NinjaOne API application** (e.g. Machine-to-Machine) with:
  - **Grant type:** Client credentials
  - **Scopes:** `monitoring` and `management`
- **Unmanaged device roles** already created in NinjaOne for each equipment type you import. Role names in the CSV must match exactly (case-sensitive). Common examples:
  - Mouse, Keyboard, Displays (or Monitor), Headset, Dock
  - Create or rename roles in NinjaOne under the relevant ITAM / device role configuration so the names match your CSV.
- **Base URL** for your NinjaOne instance, e.g.:
  - `app.ninjarmm.com`
  - `eu.ninjarmm.com`
  - `ca.ninjarmm.com`
  - `oc.ninjarmm.com`
  - `us2.ninjarmm.com`

## CSV format

### Required columns

Use **either** of these sets:

- **By name:** `Name`, `RoleName`, `OrganizationName`, `LocationName`
- **By ID:** `Name`, `RoleName`, `OrganizationId`, `LocationId`

- **Name** – Display name for the unmanaged device (e.g. "Logitech MX Master 3", "Dell P2422H"). If empty, the script may derive a name from Make + Model or generate one.
- **RoleName** – Must match an existing unmanaged device role in NinjaOne (e.g. Mouse, Keyboard, Displays, Headset, Dock). Note: NinjaOne may use "Displays" instead of "Monitor"; check your instance.

### Optional columns

- **SerialNumber** – Serial number for the device (sent in the create request).
- **WarrantyStartDate** / **WarrantyEndDate** – Parsed as dates; if omitted, defaults are today and +3 years.
- **Make**, **Model**, **PurchaseDate**, **PurchaseAmount** – If present and your instance has the corresponding device custom fields (e.g. manufacturer, model, itamAssetSerialNumber, itamAssetPurchaseDate, itamAssetPurchaseAmount), the script will PATCH them after creating the device.

Column names are matched case-insensitively. Use UTF-8 encoding for the CSV if you need special characters.

## Usage

### Basic import

```powershell
.\Import-NinjaUnmanagedDevicesFromCsv.ps1 -CsvPath ".\Import-UnmanagedDevices-Example.csv" -BaseUrl ca.ninjarmm.com
```

You will be prompted for Client ID and Client Secret if not provided via parameters or environment variables.

### Using environment variables for credentials

```powershell
$env:NinjaOneClientId = "your-client-id"
$env:NinjaOneClientSecret = "your-client-secret"
.\Import-NinjaUnmanagedDevicesFromCsv.ps1 -CsvPath ".\equipment.csv" -BaseUrl app.ninjarmm.com
```

### Preview only (no devices created)

```powershell
.\Import-NinjaUnmanagedDevicesFromCsv.ps1 -CsvPath ".\equipment.csv" -BaseUrl ca.ninjarmm.com -WhatIf
```

### Continue on errors and report at the end

```powershell
.\Import-NinjaUnmanagedDevicesFromCsv.ps1 -CsvPath ".\equipment.csv" -BaseUrl ca.ninjarmm.com -SkipErrors
```

### Verbose output

```powershell
.\Import-NinjaUnmanagedDevicesFromCsv.ps1 -CsvPath ".\equipment.csv" -BaseUrl ca.ninjarmm.com -Verbose
```

## Example CSV

See **Import-UnmanagedDevices-Example.csv** in this folder for sample rows (Mouse, Keyboard, Displays, Headset, Dock) with optional columns.

## Security

- Do not commit or log API credentials. Prefer `$env:NinjaOneClientId` and `$env:NinjaOneClientSecret` or a secure credential store.
- Run the script from a secure device if it has access to production NinjaOne data.

## Script behavior

- The script is **standalone** (no dot-sourcing); all API and OAuth logic is inlined.
- It obtains an OAuth token, caches organizations, locations, and unmanaged device roles, then processes each CSV row: resolves role/org/location, creates the device via `v2/itam/unmanaged-device`, and optionally PATCHes custom fields.
- Output: summary of created and failed counts; with `-SkipErrors`, all row errors are listed at the end.

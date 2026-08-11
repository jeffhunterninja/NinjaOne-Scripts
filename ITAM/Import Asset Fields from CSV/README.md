# Import Asset Fields from CSV (Unmanaged Devices)

Update-only: set NinjaOne **core ITAM asset fields** on existing **unmanaged** devices from a CSV. This does **not** create devices — use [Managing Unmanaged Devices via API](../Managing%20Unmanaged%20Devices%20via%20API/) for create.

**Script:** [Import-NinjaUnmanagedAssetFieldsFromCsv.ps1](Import-NinjaUnmanagedAssetFieldsFromCsv.ps1)

## Prerequisites

- **PowerShell 5.1 or later**
- **NinjaOne API application** (e.g. Machine-to-Machine) with:
  - **Grant type:** Client credentials
  - **Scopes:** `monitoring` and `management`
- Devices already exist in NinjaOne as unmanaged (`nodeClass` = `UNMANAGED_DEVICE`)
- **Base URL** for your instance (e.g. `app.ninjarmm.com`, `ca.ninjarmm.com`)

## CSV format

### Required

| Column | Notes |
|--------|--------|
| `Id` | Ninja device / node id (positive integer) |

### Optional (logging)

| Column | Notes |
|--------|--------|
| `Name` | Display/logging only — not written to NinjaOne |

### Optional asset fields

Use friendly names **or** internal API names as headers (case-insensitive). Only columns present in the CSV are considered.

| Friendly column | API field | Notes |
|-----------------|-----------|--------|
| `AssetId` | `itamAssetId` | Text |
| `AssetStatus` | `itamAssetStatus` | Text (e.g. `In use`, `inventory`, `Provisioning`) |
| `SerialNumber` | `itamAssetSerialNumber` | Text |
| `PurchaseDate` | `itamAssetPurchaseDate` | Date (e.g. `YYYY-MM-DD`) → UTC milliseconds |
| `PurchaseAmount` | `itamAssetPurchaseAmount` | Numeric (integer) |
| `ExpectedLifetime` | `itamAssetExpectedLifetime` | e.g. `3 years` (stored lowercased) |
| `EndOfLifeDate` | `itamAssetEndOfLifeDate` | Date → UTC milliseconds |

Empty cells are skipped by default (existing values left alone). Use `-OverwriteEmptyValues` to clear fields by sending `null`.

**Out of scope:** `Make` / `Model` / `manufacturer` (not core asset fields).

UTF-8 encoding recommended.

## Usage

```powershell
.\Import-NinjaUnmanagedAssetFieldsFromCsv.ps1 -CsvPath ".\Import-UnmanagedAssetFields-Example.csv" -BaseUrl ca.ninjarmm.com
```

Credentials from parameters, or `$env:NinjaOneClientId` / `$env:NinjaOneClientSecret`, or interactive prompt.

### Dry run

```powershell
.\Import-NinjaUnmanagedAssetFieldsFromCsv.ps1 -CsvPath ".\asset-updates.csv" -BaseUrl app.ninjarmm.com -WhatIf
```

### Clear empty cells

```powershell
.\Import-NinjaUnmanagedAssetFieldsFromCsv.ps1 -CsvPath ".\asset-updates.csv" -BaseUrl app.ninjarmm.com -OverwriteEmptyValues
```

## Behavior

1. Authenticates with client credentials.
2. Loads devices (paged) and indexes by id.
3. For each CSV row: resolve `Id`, require `nodeClass` = `UNMANAGED_DEVICE` (otherwise skip with warning).
4. Builds a PATCH body from present asset columns (with date/amount/lifetime coercion).
5. `PATCH /v2/device/{id}/custom-fields` (skipped when `-WhatIf`).

Summary counts: Updated, Skipped, Failed. With `-SkipErrors` (default), per-row errors are collected and listed at the end.

## Example CSV

See [Import-UnmanagedAssetFields-Example.csv](Import-UnmanagedAssetFields-Example.csv). Replace `Id` values with real unmanaged device ids from your tenant.

## Security

Do not commit or log API credentials. Prefer environment variables or a secure store.

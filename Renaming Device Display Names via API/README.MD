This script presumes you have a CSV with two columns: `systemName` and `displayName`. Pass the CSV path via `-CsvPath`. The script matches devices in NinjaOne by system name and updates the display name via the API.

## Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | Auth failure or device fetch failure |
| `2` | Validation error (CSV missing, invalid, or duplicate systemName in CSV) |

## Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | Yes | Path to the CSV file with `systemName` and `displayName` columns |
| `-NinjaOneInstance` | No | Instance hostname (default: ca.ninjarmm.com) |
| `-NinjaOneClientId` | No* | OAuth client ID (or set `NINJA_CLIENT_ID` env var) |
| `-NinjaOneClientSecret` | No* | OAuth client secret (or set `NINJA_CLIENT_SECRET` env var) |
| `-WhatIf` | No | Preview changes without applying them |

*Credentials are required; pass as parameters or via environment variables.

## Example

```powershell
.\Rename-NinjaDisplayName.ps1 -CsvPath "C:\path\to\devices.csv" -WhatIf
.\Rename-NinjaDisplayName.ps1 -CsvPath "C:\path\to\devices.csv" -NinjaOneClientId "your-id" -NinjaOneClientSecret "your-secret"
```

## Duplicate systemName

Each `systemName` in the CSV must appear only once. Duplicate entries cause exit code 2. If multiple devices in NinjaOne share the same system name, the last match wins (a warning is emitted). Consider using device ID instead of system name if you have duplicates.

## API Credentials

Create API credentials using client credentials with `monitoring` and `management` scopes, machine-to-machine application type.

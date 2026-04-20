# Generate Ninja Installer Tokens

Generates Windows MSI installer tokens for each location and Windows role in NinjaOne, then exports the results to CSV.

## What it does

- Authenticates to NinjaOne using OAuth client credentials.
- Retrieves organizations, locations, and roles from the API.
- Filters roles to Windows node classes.
- Generates a Windows MSI installer for each location/role combination.
- Extracts the token from each installer URL.
- Exports results to a CSV file.

The script continues when individual location/role generations fail and logs warnings, then exports all successful rows.

## Prerequisites

- A NinjaOne API application with valid `client_id` and `client_secret`.
- API scope: `monitoring management`.
- Permission to read organizations, locations, and roles, and generate installers.
- PowerShell 5.1+ (or PowerShell 7+).

## Parameters

| Parameter | Required | Default | Description |
| --- | --- | --- | --- |
| `NinjaOneClientID` | Yes | None | NinjaOne API client ID. |
| `NinjaOneClientSecret` | Yes | None | NinjaOne API client secret. |
| `NinjaOneInstance` | No | `ca.ninjarmm.com` | NinjaOne instance hostname. |
| `OutPath` | No | `c:\temp\InstallerTokens.csv` | Output CSV path. Parent folder must exist. |
| `PageSize` | No | `1000` | Pagination page size for org/location retrieval. |

## Usage

Basic usage:

```powershell
.\Generate-NinjaInstallerTokens.ps1 `
  -NinjaOneClientID 'your-client-id' `
  -NinjaOneClientSecret 'your-client-secret'
```

Custom instance and output path:

```powershell
.\Generate-NinjaInstallerTokens.ps1 `
  -NinjaOneClientID 'your-client-id' `
  -NinjaOneClientSecret 'your-client-secret' `
  -NinjaOneInstance 'eu.ninjarmm.com' `
  -OutPath 'C:\temp\InstallerTokens-EU.csv'
```

## Output

The script exports CSV rows with the following columns:

- `OrgName`
- `LocName`
- `OrgID`
- `LocID`
- `RoleName`
- `RoleID`
- `Token`

Each row represents one location + one Windows role token.

## Security notes

- Do not hardcode credentials directly in source files.
- Prefer secure variable injection from your automation platform or secret store.
- Treat exported token CSV files as sensitive operational data.

## Known limitations

- Token parsing expects installer URLs in the `NinjaOneAgent_<token>.msi` or `.pkg` format.
- Very large tenants may take time depending on number of locations and Windows roles.
- API/network failures can still stop execution if they affect core discovery/authentication steps.

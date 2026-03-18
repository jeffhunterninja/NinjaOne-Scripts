# Policy Mapping

This folder holds a **CSV-driven** way to map NinjaOne organizations to policy assignments by device role. The script reads the CSV and updates each organization’s policy assignments via the NinjaOne API (no hardcoded credentials).

## CSV format

Use a CSV with exactly these columns (required):

| Column             | Required | Description                                                                 |
|--------------------|----------|-----------------------------------------------------------------------------|
| **OrganizationName** | Yes      | Exact NinjaOne organization name.                                           |
| **PolicyName**       | Yes      | Exact NinjaOne policy name.                                                 |
| **DeviceRole**       | Yes      | Role name or role ID. Resolved via **GET /api/v2/roles** (match by name or ID). |

- One row = one assignment: “For this org, assign this policy to this device role.”
- **DeviceRole** must be non-empty on every row. The script resolves it using the API roles endpoint (by role name, case-insensitive, or by numeric ID).
- Multiple rows per org are allowed: use one row per (org, policy, role). Same org can have different policies for different roles (e.g. Windows Desktop vs Windows Server).
- If the same org and role appear in more than one row, the **last** row wins for that role.

### Example

```csv
OrganizationName,PolicyName,DeviceRole
Contoso,Contoso Workstations,Windows Desktop
Contoso,Contoso Workstations,Windows Laptop
Contoso,Contoso Servers,Windows Server
```

## Script

**Set-NinjaOrganizationPolicyFromCsv.ps1** reads the CSV, resolves organization names, policy names, and device roles to IDs (using `/v2/organizations`, `/v2/policies`, and **/api/v2/roles**), then **PUT**s the policy assignments per organization. Only organizations that appear in the CSV are updated.

### Credentials

The script uses **client credentials** (machine-to-machine). No browser or redirect.

- **Parameters:** `-NinjaOneInstance`, `-NinjaOneClientId`, `-NinjaOneClientSecret`
- **Environment variables:** `NINJA_BASE_URL`, `NINJA_CLIENT_ID`, `NINJA_CLIENT_SECRET`

Create an API client in NinjaOne set to **API Services (machine-to-machine)** with **Client Credentials** grant and the **monitoring** and **management** scopes. Do not hardcode secrets in the script.

### Usage

```powershell
# Required: path to CSV
.\Set-NinjaOrganizationPolicyFromCsv.ps1 -CsvPath .\policymapping.csv

# With parameters
.\Set-NinjaOrganizationPolicyFromCsv.ps1 -CsvPath .\policymapping.csv -NinjaOneClientId "..." -NinjaOneClientSecret "..."

# Using environment variables
$env:NINJA_CLIENT_ID = "..."
$env:NINJA_CLIENT_SECRET = "..."
.\Set-NinjaOrganizationPolicyFromCsv.ps1 -CsvPath .\policymapping.csv

# Preview changes without applying (WhatIf)
.\Set-NinjaOrganizationPolicyFromCsv.ps1 -CsvPath .\policymapping.csv -WhatIf

# If your instance uses /ws/oauth/token
.\Set-NinjaOrganizationPolicyFromCsv.ps1 -CsvPath .\policymapping.csv -UseWsPaths
```

### Parameters

| Parameter               | Required | Description                                                                 |
|-------------------------|----------|-----------------------------------------------------------------------------|
| **CsvPath**             | Yes      | Path to the CSV file.                                                      |
| NinjaOneInstance        | No       | Base URL (e.g. https://app.ninjarmm.com). Default: env `NINJA_BASE_URL` or https://app.ninjarmm.com. |
| NinjaOneClientId        | No       | OAuth client ID. Default: env `NINJA_CLIENT_ID`.                            |
| NinjaOneClientSecret    | No       | OAuth client secret. Default: env `NINJA_CLIENT_SECRET`.                    |
| WhatIf                  | No       | List intended changes per org without calling the API.                     |
| UseWsPaths              | No       | Use `/ws/oauth/token` for token endpoint (required for some instances).    |

### Exit codes

- **0** – Success.
- **1** – Auth or API error (e.g. token failure, PUT failed).
- **2** – Validation error (missing credentials, missing/invalid CSV, missing org/policy/role in NinjaOne).

### Error handling

- Missing or empty **DeviceRole** in a row: script **fails** with a clear message.
- **DeviceRole** not found in GET /api/v2/roles: script **fails** and reports the row.
- Organization or policy name not found in NinjaOne: script **fails** and lists the unresolved names.

## Files

- **policymapping.csv** – Example CSV with the required columns. Backfill **DeviceRole** for every row (e.g. `Windows Desktop`, `Windows Laptop` from your instance’s roles).
- **Set-NinjaOrganizationPolicyFromCsv.ps1** – Script that updates organization policy assignments from the CSV.

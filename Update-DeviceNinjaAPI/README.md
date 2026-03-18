# Update-DeviceNinjaAPI

Scripts that update NinjaOne devices in a group via the API: set policy, set role, or move devices to a different organization and location.

## Requirements

- PowerShell 5.1 or later (`#Requires -Version 5.1`).

## Scripts

| Script | Purpose |
|--------|---------|
| **Set-NinjaDevicePolicyForGroup.ps1** | Set `policyId` for all devices in a group |
| **Set-NinjaDeviceRoleForGroup.ps1** | Set `roleId` for all devices in a group (role IDs from `GET /api/v2/roles`) |
| **Move-NinjaDevicesToOrgAndLocation.ps1** | Set `organizationId` and `locationId` for all devices in a group |

## Credentials and base URL

- **Required:** OAuth client ID and client secret with `monitoring` and `management` scopes (client credentials, machine-to-machine).
- Pass credentials via parameters `-NinjaOneClientId` and `-NinjaOneClientSecret`, or set environment variables:
  - `NINJA_CLIENT_ID`
  - `NINJA_CLIENT_SECRET`
- **Optional:** Base URL / instance. Default is `https://app.ninjarmm.com`. Override with:
  - Parameter: `-NinjaOneInstance` (e.g. `app.ninjarmm.com` or `ca.ninjarmm.com`)
  - Environment: `NINJA_BASE_URL`

## OAuth and API paths

- By default, scripts use **`/oauth/token`** for the token endpoint (e.g. `https://app.ninjarmm.com/oauth/token`). This works for many US/global instances.
- Some instances (e.g. certain regional or hosted environments) require the **`/ws/`** prefix; without it, `/oauth/token` may return 405. For those, use the **`-UseWsPaths`** switch. That switches the token URL to **`/ws/oauth/token`** (e.g. `https://ca.ninjarmm.com/ws/oauth/token`).
- Group and device endpoints used: `GET /v2/group/{id}/device-ids`, `PATCH /api/v2/device/{id}`. These are the same regardless of `-UseWsPaths`.

## Troubleshooting

- **Token request returns 405 (Method Not Allowed):** Some instances require the `/ws/` prefix for the token endpoint. Use the **`-UseWsPaths`** switch (e.g. `-NinjaOneInstance ca.ninjarmm.com -UseWsPaths`).

## Exit codes

| Code | Meaning |
|------|---------|
| 0 | Success (all devices in the group updated, or group has no devices) |
| 1 | Auth failure, API failure (e.g. failed to get group device list), or one or more device PATCH calls failed |
| 2 | Validation error (missing/invalid parameters or credentials) |

## Parameters (common)

- **GroupId** (mandatory) – NinjaOne group ID.
- **NinjaOneInstance** – Instance hostname or base URL; default from `NINJA_BASE_URL` or `https://app.ninjarmm.com`.
- **NinjaOneClientId** / **NinjaOneClientSecret** – Default from `NINJA_CLIENT_ID` / `NINJA_CLIENT_SECRET`.
- **WhatIf** – Preview which devices would be updated without making changes.
- **UseWsPaths** – Use `/ws/oauth/token` for token endpoint (see OAuth and API paths above).
- **ThrottleMs** – Optional. Delay in milliseconds between each device PATCH (default 0). For large groups, a value such as 200–500 can help avoid rate limits.

Script-specific mandatory parameters: **PolicyId** (policy script), **RoleId** (role script), **OrganizationId** and **LocationId** (move script, integers).

Each script writes a result object to the pipeline: **TotalDevices**, **UpdatedCount**, **FailedCount**, **FailedDeviceIds**. When any device update fails, the script also writes failed device IDs to the error stream and exits with code 1.

## Examples

```powershell
# Set env vars (or pass -NinjaOneClientId / -NinjaOneClientSecret)
$env:NINJA_CLIENT_ID = 'your-client-id'
$env:NINJA_CLIENT_SECRET = 'your-client-secret'

# Set policy 66 for all devices in group 41
.\Set-NinjaDevicePolicyForGroup.ps1 -GroupId 41 -PolicyId 66

# Preview role change (no API writes)
.\Set-NinjaDeviceRoleForGroup.ps1 -GroupId 222 -RoleId 201 -WhatIf

# Move devices to org 8, location 18
.\Move-NinjaDevicesToOrgAndLocation.ps1 -GroupId 222 -OrganizationId 8 -LocationId 18

# Instance that requires /ws/oauth/token
.\Set-NinjaDevicePolicyForGroup.ps1 -GroupId 41 -PolicyId 66 -NinjaOneInstance ca.ninjarmm.com -UseWsPaths

# Large group: add delay between PATCH calls to avoid rate limits
.\Set-NinjaDevicePolicyForGroup.ps1 -GroupId 41 -PolicyId 66 -ThrottleMs 300
```

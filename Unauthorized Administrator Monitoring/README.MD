# Unauthorized Administrator Monitoring

## Overview

This solution detects local administrators on Windows devices that are not on an allowed list, writes the current admin list and any violations to NinjaOne custom fields, and can trigger automation (e.g. alerts or remediation) when unauthorized admins are found. You can optionally run remediation scripts to disable or log off unauthorized accounts.

## Prerequisites

- **NinjaOne** with custom fields and (optionally) org docs configured as below.
- **Report Local Admins**: Import "Report Local Admins" from the NinjaOne Template Library so it runs on the same schedule as `Test-UnauthorizedAdministrators.ps1` and updates the "Local Admins" custom field. Alternatively use the in-repo `Report-LocalAdmins.ps1` if you prefer a standalone script.
- **Authorized admins**: Set allowed accounts via org docs, org property, or device custom field (see Custom fields).

## Custom fields

Create three multi-line custom fields in NinjaOne, all with read/write from automations:

| Display name       | Ninja field name (script reference) | Purpose |
|--------------------|--------------------------------------|---------|
| Local Admins       | `LocalAdmins`                        | Current list of local administrators (updated by Report Local Admins or by Test-UnauthorizedAdministrators). |
| Authorized Admins  | `authorizedAdmins` (device), `authorizedAdminsOrg` (org), org docs | Allowed accounts; see below. |
| Unauthorized Admins| `unauthorizedAdmins`, `unauthorizedLocalAdmins` | Set by Test-UnauthorizedAdministrators: violations (or "No unauthorized..." / error message). |

**Where the allowed list is built from (in order):**

- Built-in: `Administrator`, `Domain Admins` (always allowed).
- **Org docs**: Section "User Management", subsection "Admins", field `authorizedOrgAdmins` (newline-separated list).
- **Org property**: `authorizedAdminsOrg` (comma- or newline-separated).
- **Device custom field**: `authorizedAdmins` (comma- or newline-separated).

Allowed entries are case-insensitive. You can use either `DOMAIN\user` or `user`; both match.

## Scripts

| Script | Description |
|--------|-------------|
| **Report Local Admins** (Template Library) | Enumerates local admins and updates the **Local Admins** custom field. Schedule this with Test-UnauthorizedAdministrators. |
| **Report-LocalAdmins.ps1** | Standalone report-only script that enumerates local admins and writes to `LocalAdmins` (optional in-repo alternative to the template). |
| **Test-UnauthorizedAdministrators.ps1** | Compares local admins to the allowed list, writes `LocalAdmins` and `unauthorizedAdmins` / `unauthorizedLocalAdmins`. Run as a **script result condition** in NinjaOne (e.g. trigger automation when exit code is 1). |
| **Create-NinjaAdminAccount.ps1** | **Test/demo only.** Creates or updates a named local or domain admin (parameters or `$env:adminusername` / `$env:adminpassword`). By default registers a one-time task to **disable that account after 24 hours** (`-AutoDisableAfterHours`; use `0` to skip). Run on a DC for domain admins. |

Set authorized administrators using org docs (User Management / Admins), org-level custom property `authorizedAdminsOrg`, and/or device-level custom field `authorizedAdmins`.

## Exit codes (Test-UnauthorizedAdministrators.ps1)

| Code | Meaning |
|------|---------|
| 0 | No unauthorized admins found. |
| 1 | One or more unauthorized admins found. |
| 2 | Script error (e.g. ADSI failure, Ninja property error). |

Use exit code 1 in NinjaOne as the script result condition to trigger alerts or remediation.

## Remediation

In response to unauthorized admins (exit code 1), you can run remediation scripts. **Test in a lab first and confirm the user identity before disabling or forcing logoff.**

| Script | Description | Link |
|--------|-------------|------|
| **Disable Bad User** | Reads `unauthorizedAdmins` / `unauthorizedLocalAdmins` and disables each listed local user with `Disable-LocalUser`. | [Pastebin](https://pastebin.com/vjw95bkg); local copy: `Disable-UnauthorizedLocalUser.ps1` |
| **Force Bad User Logoff** | Reads the list, finds each user's session with `query user`, and runs `logoff` for that session. | [Pastebin](https://pastebin.com/08gMv2T8); local copy: `Force-UnauthorizedUserLogoff.ps1` |

Remediation scripts are destructive. Use only after validating the list and your NinjaOne automation conditions.

## Allowed-list format

- **Org docs** (`authorizedOrgAdmins`): newline-separated.
- **Org property** (`authorizedAdminsOrg`): comma- or newline-separated.
- **Device field** (`authorizedAdmins`): comma- or newline-separated.

Entries are trimmed; empty lines are ignored. Comparison is case-insensitive.

## Troubleshooting

- **Org docs or custom fields missing**: Test-UnauthorizedAdministrators catches missing org docs and continues; missing device/org properties are skipped. Ensure custom field names match exactly (`LocalAdmins`, `authorizedAdmins`, `authorizedAdminsOrg`, `unauthorizedAdmins`, `unauthorizedLocalAdmins`).
- **Allowed list not matching**: Use consistent format (e.g. `DOMAIN\user` or `user`). Both full and bare account names are matched case-insensitively.
- **Domain Admins reported as unauthorized**: Fixed in current script version; "Domain Admins" is always allowed and name normalization handles `DOMAIN\Domain Admins`.
- **Remediation script uses wrong separator**: The pastebin scripts split on newline. Test-UnauthorizedAdministrators writes `unauthorizedLocalAdmins` with newline separation; use that field for remediation, or use the in-repo remediation scripts which accept both comma and newline.

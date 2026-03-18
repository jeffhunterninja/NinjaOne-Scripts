# Guest Wifi Check

This script checks a device’s Wi‑Fi (saved profiles and/or current connection) against either a **blocklist** or an **allowlist** defined in NinjaOne device custom fields. You can alert when blocklisted networks (e.g. guest Wi‑Fi) appear, or when the device deviates from an allowed list.

## Modes

| Mode | Purpose |
|------|--------|
| **Blocklist** | Alert when the device has saved or is connected to any network in the blocklist (e.g. guest Wi‑Fi, unauthorized networks). |
| **Allowlist** | Alert when any saved or connected SSID is *not* in the allowed list (deviation from approved networks only). |

Mode is set by the device custom field `wifiCheckMode` (`Blocklist` or `Allowlist`). If missing or invalid, the script defaults to **Blocklist** for backward compatibility.

## Scope

What gets checked is controlled by `wifiCheckScope` (or the `-Scope` parameter):

| Value | Meaning |
|-------|--------|
| `SavedOnly` | Only SSIDs from saved Wi‑Fi profiles (`netsh wlan show profiles`). |
| `ConnectedOnly` | Only the currently connected SSID(s) (`netsh wlan show interfaces`). |
| `Both` | Both saved profiles and current connection (default). |

## Custom fields

### Input

| Purpose | Field name | Used in mode | Format |
|--------|------------|--------------|--------|
| Mode | `wifiCheckMode` | Both | `Blocklist` or `Allowlist` (case-insensitive). |
| Blocklisted SSIDs | `guestWifiNetwork` or `blocklistedWifiNetworks` | Blocklist | Comma-separated SSIDs, e.g. `Guest-WiFi, Cafe` |
| Allowed SSIDs | `allowedWifiNetworks` | Allowlist | Comma-separated SSIDs, e.g. `Corporate, Office-5G` |
| Scope | `wifiCheckScope` | Both | `SavedOnly`, `ConnectedOnly`, or `Both` (optional; default `Both`). |

The script reads blocklist from `blocklistedWifiNetworks` first, then falls back to `guestWifiNetwork`. SSID comparison is case-insensitive and trims whitespace.

### Output (written by the script)

| Purpose | Field name | Example values |
|--------|------------|----------------|
| Saved SSIDs | `wifinetworks` | Newline-separated list of saved profile SSIDs, or `wlansvc is not running` / `N/A` |
| Current SSID(s) | `currentWifiNetwork` | Connected SSID(s), or `Disconnected`, `wlansvc is not running`, or `N/A` |
| Status | `wifiCheckStatus` | `OK`, `Blocklisted`, `Deviation`, `Not configured`, or `Error: ...` |

Create these device custom fields in NinjaOne if you want Condition Based Alerting or reporting.

### Bulk setting custom fields

To set the input custom fields for many organizations, locations, or devices at once, use `Set-WifiCheckCustomFieldsFromCsv.ps1`. It reads a CSV and updates NinjaOne via the API. You can set values at **organization**, **location**, or **device** level (custom fields inherit from org → location → device).

- **Required CSV columns**: `level` (organization | location | device) and `name` (identifier).
- **Organization**: `name` = organization name (e.g. `Contoso`).
- **Location**: `name` = `"organizationname,locationname"` (e.g. `Contoso,Branch A`).
- **Device**: `name` = device system name or numeric device ID (e.g. `LAPTOP-01`).
- All other columns are custom field API names: `wifiCheckMode`, `wifiCheckScope`, `blocklistedWifiNetworks`, `guestWifiNetwork`, `allowedWifiNetworks`.

NinjaOne API credentials are required (`NinjaOneInstance`, `NinjaOneClientId`, `NinjaOneClientSecret` as parameters or environment variables). Use `WifiCheckCustomFields-Example.csv` as a template.

```powershell
.\Set-WifiCheckCustomFieldsFromCsv.ps1 -CsvPath ".\WifiCheckCustomFields-Example.csv" -NinjaOneInstance "app.ninjarmm.com" -NinjaOneClientId "..." -NinjaOneClientSecret "..."
.\Set-WifiCheckCustomFieldsFromCsv.ps1 -CsvPath ".\my-import.csv" -WhatIf
```

## Removing Wi‑Fi profiles by policy

`Remove-WifiProfilesByPolicy.ps1` removes saved Wi‑Fi profiles according to the same mode and lists as the check script (`wifiCheckMode`, `blocklistedWifiNetworks`/`guestWifiNetwork`, `allowedWifiNetworks`).

| Mode | Behavior |
|------|----------|
| **Blocklist** | Deletes every saved profile whose SSID is in the blocklist. |
| **Allowlist** | If the device is currently connected to an SSID that is *not* in the allowlist, deletes that profile only. |

Optional parameters:

- **`-Mode`** — Override mode (`Blocklist` or `Allowlist`). Otherwise the script reads `wifiCheckMode` from NinjaOne (default `Blocklist`).
- **`-WhatIf`** — Report which profile(s) would be deleted without running `netsh wlan delete profile`.
- **`-NoNinjaWrite`** — Reserved for testing; the script does not write to NinjaOne and still reads mode/lists when available.

**Privileges:** Deleting Wi‑Fi profiles usually requires elevated rights (admin or running as SYSTEM). When scheduling this script in NinjaOne, ensure the script task runs with sufficient privileges to delete profiles.

```powershell
.\Remove-WifiProfilesByPolicy.ps1
.\Remove-WifiProfilesByPolicy.ps1 -WhatIf -Mode Blocklist
```

Exit codes: **0** = success (profiles removed or nothing to do); **2** = error (WLAN unavailable, netsh failure, or exception).

## Alerting

### Option 1: Alert on script failure

1. Schedule `10 - Check for blacklisted wifi network.ps1` as a script in NinjaOne (e.g. daily or on login).
2. Configure the script task to **alert when the script fails** (non-zero exit code).
3. The script exits **1** when a blocklisted network is found or when there is an allowlist deviation; it exits **2** on errors. Both trigger failure-based alerting.

### Option 2: Condition Based Alerting

1. Create the device custom fields above (including `wifiCheckStatus`).
2. In NinjaOne, create a **Condition Based Alert** that triggers when `wifiCheckStatus` is `Blocklisted` or `Deviation` (or is not equal to `OK`).
3. Schedule the script to run regularly; when it writes `Blocklisted` or `Deviation`, the condition will fire.

You can use both options together.

## Exit codes

| Code | Meaning |
|------|--------|
| 0 | OK: no blocklisted network found (Blocklist mode), or no deviation (Allowlist mode), or check not configured / no WLAN. |
| 1 | Alert: blocklisted network found (Blocklist) or Wi‑Fi deviates from allowed list (Allowlist). |
| 2 | Error: netsh/WLAN failure, unrecoverable Ninja error, or invalid configuration. |

## Testing without NinjaOne

Run with `-NoNinjaWrite` to avoid writing to NinjaOne custom fields:

```powershell
.\10 - Check for blacklisted wifi network.ps1 -NoNinjaWrite
```

You can override mode and scope for testing:

```powershell
.\10 - Check for blacklisted wifi network.ps1 -NoNinjaWrite -Mode Blocklist -Scope Both
.\10 - Check for blacklisted wifi network.ps1 -NoNinjaWrite -Mode Allowlist -Scope ConnectedOnly
```

The script still reads `wifiCheckMode`, `wifiCheckScope`, and the blocklist/allowlist from NinjaOne when the Ninja cmdlets are available; if not, it uses defaults (Blocklist, Both) and empty lists, and exits 0.

# Public IP Deviation Alert

This workflow alerts when a device's current public IP is **not** in the configured list of authorized IPs. Authorized IPs are defined in four NinjaOne device custom fields; the script runs on the device, gets its public IP, compares it to the combined list, and reports status so NinjaOne can alert on deviation.

## Custom fields

### Input (authorized IPs)

| Purpose | Field name | Scope | Format |
|--------|------------|--------|--------|
| Authorized IPs (org) | `authorizedIPorg` | Device | Comma-separated IPs, e.g. `203.0.113.1, 198.51.100.0` |
| Authorized IPs (location) | `authorizedIPloc` | Device | Comma-separated IPs, e.g. `203.0.113.1, 198.51.100.0` |
| Authorized IPs (user) | `authorizedIPuser` | Device | Comma-separated IPs, e.g. `203.0.113.1, 198.51.100.0` |
| Authorized IPs (device) | `authorizedIPdevice` | Device | Comma-separated IPs, e.g. `203.0.113.1, 198.51.100.0` |

The script merges all four lists and removes duplicates. You can populate any combination of the four fields. Whitespace around each IP is trimmed.

### Output (written by the script)

| Purpose | Field name | Scope | Example values |
|--------|------------|--------|----------------|
| Current public IP | `currentPublicIP` | Device | `203.0.113.50` |
| Status | `publicIPStatus` | Device | `OK`, `Deviation`, `No authorized IPs configured`, or `Error: ...` |
| Deviation message | `publicIPDeviationMessage` | Device | `Current IP 203.0.113.50 is not in authorized list.` |

Create these device custom fields in NinjaOne if you want Condition Based Alerting or reporting; the script will write to them when run.

## Alerting

### Option 1: Alert on script failure

1. Schedule `Check-PublicIP.ps1` as a script in NinjaOne (e.g. daily or hourly).
2. Configure the script task to **alert when the script fails** (non-zero exit code).
3. On deviation the script exits with code **1**; on other errors it exits **2**. Both will trigger failure-based alerting.

### Option 2: Condition Based Alerting

1. Create device custom fields above (e.g. `publicIPStatus`).
2. In NinjaOne, create a **Condition Based Alert** that triggers when the device custom field `publicIPStatus` contains `Deviation` (or is not equal to `OK`).
3. Schedule the script so it runs regularly; when it writes `Deviation`, the condition will fire.

You can use both options together.

## Exit codes

| Code | Meaning |
|------|--------|
| 0 | Public IP is authorized, or no authorized IPs configured (check skipped). |
| 1 | Deviation: current public IP is not in the authorized list. |
| 2 | Error (e.g. unable to reach ipify or Ninja cmdlet failure). |

## Testing without NinjaOne

Run with `-NoNinjaWrite` to avoid writing to NinjaOne custom fields (useful when testing locally or outside the agent):

```powershell
.\Check-PublicIP.ps1 -NoNinjaWrite
```

The script will still read authorized IPs from the four custom fields if the Ninja cmdlets are available; if not, it will treat the device as having no authorized IPs and exit 0.

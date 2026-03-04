# Set-QuietHours

## Overview

This PowerShell script collects end-user **Quiet Hours** (Do-Not-Disturb) preferences via a WPF popup. It saves preferences to a local JSON file and, when run in a NinjaOne context, can write the same JSON to a NinjaOne custom field so policies or other scripts can respect the user's quiet windows.

**Important:** The script must run as the **logged-on user** (not SYSTEM) so the WPF dialog displays for the user. Preferences are saved to `C:\RMM\NinjaOne-QuietHours\quiet_hours.json` so both Set-QuietHours (user) and Test-QuietHours (e.g. SYSTEM) use the same file. The folder `C:\RMM\NinjaOne-QuietHours` may need to be created once with write access for users (e.g. grant Users Modify); see Troubleshooting if preferences do not save.

---

## Exit codes

| Code | Meaning |
|------|---------|
| `0`  | Script completed (user closed the dialog). Save vs Cancel is not distinguished. This is a UI-only script. |

---

## Prerequisites

- **PowerShell 5.1** or later
- **.NET WPF** (PresentationFramework, PresentationCore) for the survey UI
- **Optional:** NinjaOne with `Ninja-Property-Set` available and a device (or org) custom field for storing the JSON (e.g. `quietHours`)

---

## Parameters and environment variables

| Parameter           | Type   | Default | Description |
|--------------------|--------|---------|-------------|
| **PreferencePath** | string | `C:\RMM\NinjaOne-QuietHours\quiet_hours.json` | Full path to the JSON file where preferences are saved. |
| **NinjaCustomField** | string | `quietHours` (or `$env:quietHoursCustomField` if set) | Name of the NinjaOne custom field to write the JSON to when `Ninja-Property-Set` is available. |

When run from NinjaOne, you can set a script variable that populates `$env:quietHoursCustomField` to override the custom field name without changing the script.

---

## Setup in NinjaOne

1. **Create a device (or organization) custom field** (e.g. name: `quietHours`, type: Text) to store the quiet hours JSON.
2. **Add the script** to NinjaOne:
   - **Type:** PowerShell
   - **OS:** Windows
   - **Run As:** **Logged-on user** (not SYSTEM)
3. Ensure the custom field name in NinjaOne matches the script parameter or `$env:quietHoursCustomField` if you use it.

---

## Running the test and log script (Test-QuietHours.ps1)

**Test-QuietHours.ps1** runs the quiet-hours check once and can log the result for scheduled or manual runs. It does not show the survey UI.

| Parameter         | Default | Description |
|-------------------|---------|-------------|
| **PreferencePath** | `C:\RMM\NinjaOne-QuietHours\quiet_hours.json` | Path to the quiet hours JSON file. |
| **LogPath**       | (none)  | If set, append a one-line status (timestamp and `InQuietHours=True/False` or `NoPrefs`) to this file. Directory is created if needed. |
| **Quiet**         | false   | Suppress host output; use with **LogPath** when you only want file logging. |
| **Mode**          | (none, or `$env:quietHoursMode` in NinjaOne) | **Alert Within** = exit 1 when current time is within quiet hours; **Alert Outside** = exit 1 when outside quiet hours. When not set, script always exits 0. In NinjaOne, set script variable `quietHoursMode` to `Alert Within` or `Alert Outside` to drive alerting from exit code. |

**Exit codes (when Mode is set):** `0` = no alert; `1` = alert condition met (so NinjaOne can treat non-zero as failure and alert). If no preferences file exists, the script always exits 0.

**Examples:**

```powershell
# Run test and write status to host
.\Test-QuietHours.ps1

# Also append to a log file
.\Test-QuietHours.ps1 -LogPath "C:\RMM\NinjaOne-QuietHours\quiet_hours_check.log"

# Log to file only (e.g. for a scheduled task)
.\Test-QuietHours.ps1 -LogPath "C:\Logs\quiet_hours.log" -Quiet

# Alert when device is outside quiet hours (exit 1 = alert in NinjaOne)
.\Test-QuietHours.ps1 -Mode "Alert Outside"
```

---

## Using Test-QuietHours from other scripts

Other scripts can dot-source `Set-QuietHours.ps1` to use the `Test-QuietHours` and `Get-QuietPrefs` helpers, then skip disruptive actions when the current time is inside the user's quiet window.

Example:

```powershell
. "$PSScriptRoot\Set-QuietHours.ps1"
$prefs = Get-QuietPrefs -Path "C:\RMM\NinjaOne-QuietHours\quiet_hours.json"
if (Test-QuietHours -Prefs $prefs) {
  Write-Verbose "[QuietHours] Skipping disruptive actions at $(Get-Date)."
  return
}
# ... continue with normal work ...
```

Run the calling script with `-Verbose` to see when quiet hours are in effect.

---

## Examples

**Run the survey with defaults:**

```powershell
.\Set-QuietHours.ps1
```

**Run with a custom preference path and verbose output:**

```powershell
.\Set-QuietHours.ps1 -PreferencePath "C:\Data\quiet_hours.json" -Verbose
```

---

## Troubleshooting

| Issue | Likely cause | Solution |
|-------|--------------|----------|
| Preferences not saving | Path not writable or wrong user | Run as the logged-on user; ensure `C:\RMM\NinjaOne-QuietHours` exists and that Users have Modify permission (create the folder and set ACLs if needed). |
| Ninja custom field not updating | Run As, field name, or cmdlet missing | Run as **logged-on user**; ensure the custom field exists and name matches (e.g. `quietHours` or env var); confirm `Ninja-Property-Set` is available in the NinjaOne script context. |
| WPF dialog does not show | Running as SYSTEM or non-interactive | Run as the **logged-on user** in an interactive session so the dialog can display. |
| Per-day override rejected | Wrong format | Use **HH:mm-HH:mm** (24-hour), e.g. `21:00-07:00`. |

---

## Technical note: time

The `Test-TimeInRange` and `Test-QuietHours` logic use the machine's **local time** (`Get-Date`) to decide if the current time is inside a quiet window. No timezone conversion is performed.

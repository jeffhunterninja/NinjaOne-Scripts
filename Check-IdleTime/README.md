## Overview

***Note on 2026-02-20*** version 2 of this script normalizes the exit codes used in v1 - please adjust your script executions if updating this script in place.

This PowerShell script measures **per-user idle time** on Windows endpoints when executed as **SYSTEM**.

It works by launching a lightweight PowerShell helper **inside each logged-in user’s session**, which calls `GetLastInputInfo` to determine how long the user has been idle.

### Key Features

- Measures **per-user** idle time via Windows API  
- Runs as **SYSTEM** with `CreateProcessAsUser` for each session  
- Selects the most relevant session (Console > most-idle Active > any)  
- Writes to **NinjaOne custom fields**  
- Supports configurable idle time thresholds  
- Returns standardized **exit codes** for policy automation - i.e. only patching when the idle time has been above a certain threshold, or using if a device has been idle for a certain period of time as a condition in a compound condition.

### ⚙️ Exit Codes

Meaning of 0 and 1 depends on the **Alert On** setting (parameter or script variable):

| Code | Meaning |
|------|----------|
| `0`  | OK — With "Idle Time Greater": no threshold or idle &lt; threshold. With "Idle Time Lesser": no threshold or idle ≥ threshold. |
| `1`  | ALERT — With "Idle Time Greater": idle ≥ threshold. With "Idle Time Lesser": idle &lt; threshold. |
| `2`  | Not elevated (must run as SYSTEM) |

---

##  How It Works

### 1. Elevation Check

Ensures the script is running with Administrator privileges.  
If not, it exits immediately with code **2**.

### 2. Result Collection

The script collects results for all active sessions (`WTSActive`, `WTSConnected`, or `WTSIdle`):

| Property | Description |
|-----------|--------------|
| `SessionId` | Windows session ID |
| `WinStation` | Session name (e.g. Console, RDP-Tcp#5) |
| `State` | Session state |
| `IdleMinutes` | Calculated idle minutes |
| `IdleSeconds` | Calculated idle seconds |
| `MeasuredVia` | Method or status (e.g. `CreateProcessAsUser:GetLastInputInfo` or `Failed`) |

### 3. Session Selection

The script prioritizes which session to evaluate:

1. Console session (if available)  
2. Most-idle active session  
3. Any other measured session (fallback)

### 4. NinjaOne Custom Field Updates

Three custom fields are updated:

| Field | Type | Example Value | Description |
|--------|------|----------------|--------------|
| `idleTime` | Text | `1 hour, 20 minutes` | Human-readable idle duration |
| `idleTimeStatus` | Text | `ALERT: Idle 85 min (>= 60)` or `85` | Numeric minutes or alert text |
| `idleTimeMinutes` | Integer | `85` | Idle duration in minutes (integer, for filtering/sorting) |

### 5. Threshold Handling

If a threshold is defined (`ThresholdMinutes` or `thresholdminutes` env var), behavior depends on **Alert On**:

- **Idle Time Greater** (default): When idle ≥ threshold → alert in `idleTimeStatus`, exit **1**. Otherwise → exit **0**.
- **Idle Time Lesser**: When idle &lt; threshold → alert in `idleTimeStatus`, exit **1**. When idle ≥ threshold → exit **0**.

---

##  Parameters and Environment Variables

- **UserName** (parameter, optional): When set, only sessions for users matching this name (via `query user`) are measured. If `query user` fails or returns no data (e.g. non-English Windows), the user filter is skipped and **all** sessions are measured.
- **ThresholdMinutes** (parameter, default `0`): Idle threshold in minutes. With default Alert On "Idle Time Greater", exit code 1 when idle ≥ this value. Must be ≥ 0.
- **PerProcessTimeoutSeconds** (parameter, default `10`): Timeout in seconds for the helper run in each session. Valid range 1–300.
- **AlertOn** (parameter, default `Idle Time Greater`): When to alert (exit 1). `Idle Time Greater` = alert when idle ≥ threshold. `Idle Time Lesser` = alert when idle &lt; threshold.
- **NinjaOne:** Create a script variable "Threshold Minutes" (integer) to set the threshold; the script reads `$env:thresholdminutes` when present and valid. Optionally create "Alert On" (text); the script reads `$env:alerton` and accepts `Idle Time Greater` or `Idle Time Lesser` (case-insensitive). Parameter defaults are used when env vars are missing or invalid.

---

##  Setup in NinjaOne

### 1. Create Device Custom Fields

Create three custom fields in NinjaOne under **Devices → Custom Fields**:

| Name | Type | Purpose |
|------|------|----------|
| `idleTime` | Text | Stores the human-readable idle duration |
| `idleTimeStatus` | Text | Stores either numeric minutes or an alert string |
| `idleTimeMinutes` | Integer | Stores idle duration in minutes (integer) for filtering/sorting |

### 2. Add the Script

| Setting | Value |
|----------|--------|
| **Type** | PowerShell |
| **OS** | Windows |
| **Run As** | SYSTEM |

### 3. Configure Thresholds

#### Create script variables
- **Threshold Minutes** (Integer): Sets the idle threshold in minutes.
- **Alert On** (optional, Text): Set to `Idle Time Greater` (alert when idle ≥ threshold) or `Idle Time Lesser` (alert when idle &lt; threshold). Default is "Idle Time Greater" if not set.

---

## 🧾 Example Outputs

### Example 1 — No Threshold
```
=== Summary ===
ComputerName       : DESKTOP123
IdleMinutes        : 38
IdleTime           : 38 minutes
ThresholdMinutes   : 0
ThresholdExceeded  : False
UsedFallback       : False
```

Custom Fields:
```
idleTime: 38 minutes
idleTimeStatus: 38
idleTimeMinutes: 38

```

---

### Example 2 — Threshold Exceeded
```
Idle time threshold exceeded: 85 minutes (threshold: 60).
```

Custom Fields:
```
idleTime: 1 hour, 25 minutes
idleTimeStatus: ALERT: Idle 85 min (>= 60)
idleTimeMinutes: 85
Exit Code: 1
```

---

## 🔍 Troubleshooting

| Issue | Likely Cause | Solution |
|--------|--------------|-----------|
| `Access Denied` / Exit Code 2 | Script not elevated | Run as **SYSTEM** |
| `(No sessions measured or all failed)` | No interactive users or helper timed out | Confirm a user is logged in; increase `PerProcessTimeoutSeconds` if sessions are slow |
| Idle time incorrect | Different session evaluated | Check per-session table; run with `-Verbose` to see which session was evaluated |
| Threshold ignored | Env not set or invalid | Set NinjaOne script variable "Threshold Minutes" (integer); param default is used when env is missing |
| Custom fields not updating | CFs missing, misnamed, or cmdlet absent | Verify exact field names; script warns if `Ninja-Property-Set` is not available |

---

## 🧠 Technical Details

- **Windows API:** Uses `GetLastInputInfo` for precise idle tracking.
- **Session Management:** Via `WTSEnumerateSessions` and `CreateProcessAsUser`.
- **Supported States:** `WTSActive`, `WTSConnected`, `WTSIdle`.
- **Run Context:** Must be **SYSTEM** to access other sessions.
- **TickCount Handling:** Uses unsigned arithmetic; wraps at ~49 days (idle may be wrong after long uptime).

- **`query user`:** Used for reference/display and for optional `-UserName` filtering; column headers are localized—parsing may fail on non-English Windows. When `query user` returns no data and `-UserName` was specified, the script skips user filtering and measures all sessions.

---

> **Author’s Note:**  
> This script is provided as-is and does not fall under normal scope of NinjaOne support.

## 📘 Overview

This PowerShell script measures **per-user idle time** on Windows endpoints, even when executed as **SYSTEM** - which is necessary to interface with NinjaOne custom fields.

It works by launching a lightweight PowerShell helper **inside each logged-in user’s session**, which calls `GetLastInputInfo` to determine how long the user has been idle.

### ✅ Key Features

- Measures **per-user** idle time via Windows API  
- Runs as **SYSTEM** with `CreateProcessAsUser` for each session  
- Selects the most relevant session (Console > most-idle Active > any)  
- Writes to **NinjaOne custom fields**  
- Supports configurable idle time thresholds  
- Returns standardized **exit codes** for policy automation - i.e. only patching when the idle time has been above a certain threshold.

### ⚙️ Exit Codes

| Code | Meaning |
|------|----------|
| `0`  | OK — no threshold set or idle below threshold |
| `1`  | Not elevated (must run as SYSTEM) |
| `2`  | ALERT — idle time ≥ threshold |

---

## 🧩 How It Works

### 1. Elevation Check

Ensures the script is running with Administrator privileges.  
If not, it exits immediately with code **1**.

### 2. Result Collection

The main script collects results for all active sessions (`WTSActive`, `WTSConnected`, or `WTSIdle`):

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

Two custom fields are updated:

| Field | Type | Example Value | Description |
|--------|------|----------------|--------------|
| `idleTime` | Text | `1 hour(s), 20 minute(s)` | Human-readable idle duration |
| `idleTimeStatus` | Text | `ALERT: Idle 85 min (>= 60)` or `85` | Numeric minutes or alert text |

### 5. Threshold Handling

If a threshold is defined (`ThresholdMinutes` or `thresholdminutes` env var):

- When idle time ≥ threshold:  
  → Writes an alert to `idleTimeStatus` and exits with code **2**
- Otherwise:  
  → Writes numeric idle time and exits **0**

---

## 🔧 Parameters and Environment Variables

Create a Script Form Variable called "Threshold Minutes" if you want to specifiy a certain timeframe that will constitute an idle device.

```powershell
$ThresholdMinutes = $env:thresholdminutes
```

---

## 🧱 Setup in NinjaOne

### 1. Create Device Custom Fields

Create two custom fields in NinjaOne under **Devices → Custom Fields**:

| Name | Type | Purpose |
|------|------|----------|
| `idleTime` | Text | Stores the human-readable idle duration |
| `idleTimeStatus` | Text | Stores either numeric minutes or an alert string |

### 2. Add the Script

| Setting | Value |
|----------|--------|
| **Type** | PowerShell |
| **OS** | Windows |
| **Run As** | SYSTEM |
| **PowerShell Version** | 5.1 |
| **Execution Policy** | Bypass |
| **Timeout** | ≥ 60 seconds recommended |

Paste the full original script into the script body.

### 3. Configure Thresholds

#### Create script variable
Set a script variable in the script called "Threshold Minutes" that uses the "Integer" data type.

---

## 🧾 Example Outputs

### Example 1 — No Threshold
```
=== Summary ===
ComputerName       : DESKTOP123
IdleMinutes        : 38
IdleTime           : 38 minute(s)
ThresholdMinutes   : 0
ThresholdExceeded  : False
UsedFallback       : False
```

Custom Fields:
```
idleTime: 38 minute(s)
idleTimeStatus: 38
Exit Code: 0
```

---

### Example 2 — Threshold Exceeded
```
Idle time threshold exceeded: 85 minute(s) (threshold: 60).
```

Custom Fields:
```
idleTime: 1 hour(s), 25 minute(s)
idleTimeStatus: ALERT: Idle 85 min (>= 60)
Exit Code: 2
```

---

## 🔍 Troubleshooting

| Issue | Likely Cause | Solution |
|--------|--------------|-----------|
| `Access Denied` / Exit Code 1 | Script not elevated | Run as **SYSTEM** |
| `(No sessions measured or all failed)` | No interactive users | Confirm a user is logged in |
| Idle time incorrect | Different session evaluated | Check per-session table |
| Threshold ignored | Env var override | Remove or update `thresholdminutes` |
| Custom fields not updating | CFs missing or misnamed | Verify exact field names |

---

## 🧠 Technical Details

- **Windows API:** Uses `GetLastInputInfo` for precise idle tracking.
- **Session Management:** Via `WTSEnumerateSessions` and `CreateProcessAsUser`.
- **Supported States:** `WTSActive`, `WTSConnected`, `WTSIdle`.
- **Run Context:** Must be **SYSTEM** to access other sessions.
- **TickCount Handling:** Uses unsigned arithmetic to avoid overflow.
- **Error Handling:** All `Ninja-Property-Set` calls wrapped in `try/catch`.

---

> 🧩 **Author’s Note:**  
> This script is provided as-is and does not fall under normal scope of NinjaOne support.

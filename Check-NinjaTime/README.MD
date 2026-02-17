# Script Variables: Schedule Modes & Parameters

This document outlines the configurable script variables used to define scheduling behavior in the script. The scheduling is controlled through a combination of execution modes (`Mode`) and supporting parameters that vary depending on the selected mode.

---

## Exit Codes

| Code | Meaning |
|------|---------|
| `1` | Condition met (within window or at target time; NinjaOne triggers linked automation) |
| `2` | Condition not met (outside window, too far from target, or validation/parse error) |

---

## 📋 Script Variables Overview

| Variable                  | Used For                  | Type      | Notes                                                                 |
|---------------------------|---------------------------|-----------|-----------------------------------------------------------------------|
| `Mode`                   | Execution type             | DROPDOWN  | Required. Options: `Once`, `Daily`, `Weekly`, `Monthly`, `Window`    |
| `Target Time`            | Time or DateTime           | DATETIME  | Full datetime for `Once`, time only for `Daily`, `Weekly`, and `Monthly` |
| `Weekly Day Of Week`     | Weekly mode                | TEXT      | Comma-separated days (e.g., `Monday,Wednesday`)                      |
| `Monthly Day Of Month`   | Monthly mode               | INTEGER   | Numeric day of the month (1–31)                                      |
| `Window Recurrence`      | Window recurrence          | DROPDOWN  | Required for `Window` mode. Options: `Daily`, `Weekly`               |
| `Window Start`           | Start of window            | DATETIME  | Time only used                                                       |
| `Window End`             | End of window              | DATETIME  | Time only used                                                       |
| `Window Day of Week`     | Weekly window mode         | TEXT      | Required if `Window Recurrence` is `Weekly`; comma-separated         |
| `Time Window Minutes`    | Grace period (minutes)     | INTEGER   | Optional. Set to interval + 1 (e.g., 16 for 15-min interval). Default 5. |

---

## 🔄 Mode Behavior and Field Requirements

### `Once`
- **Uses**: Full `Target Time` (date and time)
- **Required**: `Target Time`

### `Daily`
- **Uses**: Time from `Target Time`
- **Required**: `Target Time`

### `Weekly`
- **Uses**: Time from `Target Time` and `Weekly Day Of Week`
- **Required**: `Target Time`, `Weekly Day Of Week`

### `Monthly`
- **Uses**: Time from `Target Time` and `Monthly Day Of Month`
- **Required**: `Target Time`, `Monthly Day Of Month`

### `Window`
- **Uses**: `Window Recurrence`, `Window Start`, `Window End`, optionally `Window Day of Week`
- **Required**:
  - `Window Recurrence`
  - `Window Start`
  - `Window End`
  - `Window Day of Week` (if recurrence is `Weekly`)

---
**Important Note About Condition Resets and Script Timeouts**

It is not recommended to use "When no longer" for any mode - only use time-based resets.

For any mode, set the `Time Window Minutes` script variable (or `timeWindowMinutes` env var) to be in sync with your interval and script timeout. For example, if you run the compound condition every 15 minutes, set `Time Window Minutes` to 16. This allows the script to sleep until the desired target time is reached. Ensure the script timeout exceeds `TimeWindowMinutes` (e.g., 17 minutes for a 15-minute interval).

- Script interval
- TimeWindowMinutes = interval + 1 minute
- Script timeout = interval + 2 minutes

When using Window mode, it also becomes important to ensure the auto-reset is longer than the maintenance window - i.e. a 4 hour maintenance window should have a 4 hour or longer auto-reset time. If the reset is too short, the automation could potentially trigger again if there is not an adequate state change monitor as a part of the compound condition.

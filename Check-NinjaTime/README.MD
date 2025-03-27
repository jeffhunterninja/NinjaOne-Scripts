# Script Variables: Schedule Modes & Parameters

This document outlines the configurable script variables used to define scheduling behavior in the script. The scheduling is controlled through a combination of execution modes (`Mode`) and supporting parameters that vary depending on the selected mode.

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

For any mode, it's important to adjust the $TimeWindowMinutes variable in the script to be in sync with your interval and script timeout. For example, if I'm running my compound condition every 15 minutes, I would set my $TimeWindowMinutes to 16 minutes. This allows the script to sleep until the desired $TargetTime is reached. I need to ensure that the timeout for the script exceeds my $TimeWindowMinutes - so in my case I would make it 17 minutes.
-Script interval
-$TimeWindowMinutes = interval + 1 minute
-Script timeout = interval + 2 minutes

When using Window mode, it also becomes important to ensure the auto-reset is longer than the maintenance window - i.e. a 4 hour maintenance window should have a 4 hour or longer auto-reset time. If the reset is too short, the automation could potentially trigger again if there is not an adequate state change monitor as a part of the compound condition.

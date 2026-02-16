# Script Documentation: NinjaOne Tag Presence Evaluation

This script checks if specified tags are present on the current device using the NinjaOne API. It supports two evaluation modes: `ANY` and `ALL`.

---

## Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Match found (tag condition satisfied) |
| `1` | No match (tag condition not satisfied) |
| `2` | Error (validation failure, Get-NinjaTag failure, etc.) |

---

## 📋 Script Variables Overview

| Variable          | Used For                             | Type     | Environment Variable | Notes                                                                 |
|-------------------|--------------------------------------|----------|----------------------|-----------------------------------------------------------------------|
| `Tags To Search` | Tags to look for                     | TEXT     | `tagsToSearch`       | Required. Accepts a single tag or a comma-separated list of tags      |
| `Mode`           | Matching behavior                    | DROPDOWN | `mode`               | Optional. Values: `ANY` or `ALL` (case-insensitive). Defaults to `ANY` if omitted |

---

## 🔎 Tag Matching Logic

### `Tags To Search`
- Accepts a comma-separated list of tags to evaluate (e.g., `Critical,Production,VIP`)
- Whitespace is trimmed from each tag
- Empty entries (e.g. from `tag1,,tag2`) are ignored
- Tags are evaluated against the current device’s assigned tags (via `Get-NinjaTag`)

### `Mode`
- Determines how the tag match is evaluated:
  - **`ANY`**: Script exits successfully if **any** of the specified tags are present
  - **`ALL`**: Script exits successfully only if **all** specified tags are present
- Case-insensitive comparison is applied

---

# Script Documentation: NinjaOne Tag Presence Evaluation

This script checks if specified tags are present on the current device using the NinjaOne API. It supports two evaluation modes: `ANY` and `ALL`.

---

## 📋 Script Variables Overview

| Variable          | Used For                             | Type     | Notes                                                                 |
|-------------------|--------------------------------------|----------|-----------------------------------------------------------------------|
| `Tags To Search` | Tags to look for                     | TEXT     | Required. Accepts a single tag or a comma-separated list of tags      |
| `Mode`           | Matching behavior                    | DROPDOWN | Optional. Values: `ANY` or `ALL` (case-insensitive). Defaults to `ANY` if omitted |

---

## 🔎 Tag Matching Logic

### `Tags To Search`
- Accepts a comma-separated list of tags to evaluate (e.g., `Critical,Production,VIP`)
- Whitespace is trimmed from each tag
- Tags are evaluated against the current device’s assigned tags (via `Get-NinjaTag`)

### `Mode`
- Determines how the tag match is evaluated:
  - **`ANY`**: Script exits successfully if **any** of the specified tags are present
  - **`ALL`**: Script exits successfully only if **all** specified tags are present
- Case-insensitive comparison is applied

---

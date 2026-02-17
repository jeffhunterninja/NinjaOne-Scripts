# Printer Inventory Collection via User + SYSTEM Scripts (NinjaOne)

## Overview

This document describes a **two-stage printer inventory collection process** in NinjaOne designed to reliably capture **both per-user and machine-wide printers** on Windows endpoints.

Because Windows printers can be installed:
- **Per-user (HKCU)** – only visible in the user context
- **Machine-wide (HKLM)** – visible to all users and SYSTEM

a **single SYSTEM script cannot reliably detect all printers**. This solution uses:
1. A **user-context script** to collect printer data and store it in `C:\Windows\Temp\printer_info.json`. This shows printers specific to the user, plus anything that is also machine-wide.
2. A **SYSTEM-context script** to ingest that data and populate NinjaOne custom fields. Most non-admin user accounts cannot directly interface with custom fields, which is why this second script running at the System level is required.

### Scheduling

Configure both scripts to run **sequentially at user login** in NinjaOne (user script first, then system script). The user script writes to `C:\Windows\Temp`, which exists by default and is writable by standard users. The system script reads the file, updates custom fields, and removes the temp file.

### Custom Fields

Only the `printers` custom field is required - `printerdrivers` is optional depending on how you'd like to view the information.

| Field Name        | Type             | Scope  | Purpose |
|------------------|------------------|--------|---------|
| `printers`        | Multi-line | Device | Stores printer queue names |
| `printerDrivers`  | Multi-line | Device | Stores printer driver names |

> **Important:**  
> Field names are **case-sensitive** when referenced by scripts. The names above must match exactly.

### Path

`C:\Windows\Temp\printer_info.json` is used for the handoff because it exists on all Windows installs and is writable by standard (non-admin) users. ProgramData and System32 are not used because standard users typically cannot write to those locations.

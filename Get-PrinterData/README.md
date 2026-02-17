# Printer Inventory Collection (NinjaOne)

## Overview

This document describes printer inventory collection for NinjaOne, designed to reliably capture **both per-user and machine-wide printers** on Windows endpoints.

Because Windows printers can be installed:
- **Per-user (HKCU)** – only visible in the user context
- **Machine-wide (HKLM)** – visible to all users and SYSTEM

two approaches are available: a **single-script CreateProcessAsUser** method (recommended) or a **two-stage user + SYSTEM** method.

### Custom Fields

Only the `printers` custom field is required - `printerDrivers` is optional depending on how you'd like to view the information.

| Field Name       | Type      | Scope  | Purpose                  |
|------------------|-----------|--------|--------------------------|
| `printers`       | Multi-line| Device | Stores printer queue names |
| `printerDrivers` | Multi-line| Device | Stores printer driver names |

> **Important:** Field names are **case-sensitive** when referenced by scripts. The names above must match exactly.

---

## Single-Script Approach: Get-PrinterDataAsSystem.ps1

`Get-PrinterDataAsSystem.ps1` runs as SYSTEM and uses **CreateProcessAsUser** to launch a helper in each logged-in user session. Each helper runs `Get-Printer` in user context (capturing per-user HKCU and machine-wide HKLM printers), writes JSON to a per-session temp file, and the main script merges results, deduplicates, and updates NinjaOne custom fields via `Set-NinjaProperty`. When no users are logged in, it falls back to `Get-Printer` in SYSTEM context (HKLM printers only).

### Scheduling

Configure **one schedule** to run as SYSTEM (e.g. at login or on a recurring interval). No user-context script or sequential scheduling is required.

---

## Two-Stage Approach: User + SYSTEM Scripts

This approach uses:
1. A **user-context script** (`Get-Printer User Printer Collection.ps1`) to collect printer data and store it in `C:\Windows\Temp\printer_info.json`.
2. A **SYSTEM-context script** (`Get-Printer Custom Field Update.ps1`) to ingest that data and populate NinjaOne custom fields.

### Scheduling

Configure both scripts to run **sequentially at user login** in NinjaOne (user script first, then system script). The user script writes to `C:\Windows\Temp`, which exists by default and is writable by standard users. The system script reads the file, updates custom fields, and removes the temp file.

### Path (Two-Stage Only)

`C:\Windows\Temp\printer_info.json` is used for the handoff because it exists on all Windows installs and is writable by standard (non-admin) users. ProgramData and System32 are not used because standard users typically cannot write to those locations.

### Scripts

| File | Approach | Purpose |
|------|----------|---------|
| `Get-PrinterDataAsSystem.ps1` | Single-script | CreateProcessAsUser; one schedule as SYSTEM |
| `Get-Printer User Printer Collection.ps1` | Two-stage | User context; collects and writes JSON |
| `Get-Printer Custom Field Update.ps1` | Two-stage | SYSTEM context; reads JSON and updates custom fields |

# README

This is a collection of useful workflows and scripts for endpoint management using NinjaOne, as well as examples of accessing the NinjaOne API.

This repository does not fall under the scope of the NinjaOne support or solutions engineering teams - if you encounter any issues, please open an issue here on Github.

## Updates

Major update and re-write of several scripts in late February 2026

---

## New Scripts / Features

- **Install AutoCAD** ([Install AutoCAD/](Install%20AutoCAD/)) — Runs the AutoCAD/ODIS bootstrapper (Deploy.exe), waits for Summary.txt, then runs Installer.exe. See [Install AutoCAD/README.md](Install%20AutoCAD/README.md).

- **Export-WindowsPatchReportToHtml.ps1** ([Windows OS Patch Reporting/](Windows%20OS%20Patch%20Reporting/)) — Exports the Windows patch report to HTML. Optional PSWriteHTML for styled reports; per-organization output supported. See [Windows OS Patch Reporting/README.md](Windows%20OS%20Patch%20Reporting/README.md).

- **Export-WindowsPatchReportToCsv.ps1** ([Windows OS Patch Reporting/](Windows%20OS%20Patch%20Reporting/)) — Exports the Windows patch report to CSV. Per-organization output supported. See [Windows OS Patch Reporting/README.md](Windows%20OS%20Patch%20Reporting/README.md).

---

## Recent Updates

- **Check-IdleTime** ([Check-IdleTime/](Check-IdleTime/)) — **v2**: Normalized exit codes (0 OK, 1 Alert, 2 not elevated); support for an additional integer-only custom field. See [Check-IdleTime/README.md](Check-IdleTime/README.md).

- **Check-NinjaTime** ([Check-NinjaTime/](Check-NinjaTime/)) — **v3**: New `mode` and `Window` recurrence option, *Monthly Day of Week*; exit code structure standardized; added `InvertExitCode` support for flexible alerting. See [Check-NinjaTime/README.MD](Check-NinjaTime/README.MD).

- **Check-NinjaTag** ([Check-NinjaTag/](Check-NinjaTag/)) — Minor bugfixes and documentation updates See [Check-NinjaTag/README.MD](Check-NinjaTag/README.MD).

- **Sync-PatchStatus.ps1** ([Patch Status Sync/](Patch%20Status%20Sync/)) — Standardized exit codes, bugfixes, documentation updates. See [Patch Status Sync/README.md](Patch%20Status%20Sync/README.md).

- **Rename-NinjaDisplayName.ps1** ([Renaming Device Display Names via API/](Renaming%20Device%20Display%20Names%20via%20API/)) — Standardized code structure, added a WhatIf mode. See [Renaming Device Display Names via API/README.MD](Renaming%20Device%20Display%20Names%20via%20API/README.MD).

---

## Beta Scripts

- **Get-PrinterDataAsSystem.ps1** ([Get-PrinterData/](Get-PrinterData/)) — Single-script printer inventory that runs as SYSTEM using CreateProcessAsUser. Captures per-user and machine-wide printers and writes to NinjaOne custom fields. I don't have a method of testing broadly, please let me know what you see in the field. See [Get-PrinterData/README.md](Get-PrinterData/README.md).

## Other Changes

- **Get-DeviceAlertAnalysisReport.ps1** — Minor updates.
- **Get-Printer Custom Field Update.ps1** and **Get-Printer User Printer Collection.ps1** — Updates for reliability and consistency.

---
# Install AutoCAD (NinjaOne)

Runs the AutoCAD/ODIS bootstrapper (Deploy.exe) transferred by NinjaOne, waits for the deployment image to be ready, then runs the installer.

## Flow

1. Start bootstrapper (`Deploy.exe`) with `/q /p`.
2. The process doesn't seem to end gracefully, so poll for `Summary.txt` at the deploy path as the indicator that files are staged and ready for install.
3. Optional post-wait for cleanup.
4. Run `Installer.exe` with the deployment `Collection.xml`; script exit code reflects installer result.

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Installer completed successfully. |
| 1 | Bootstrap/validation error (e.g. Deploy.exe missing, invalid paths). |
| 2 | Installer failed or could not be started. |
| 3 | Timeout waiting for bootstrap (Summary.txt not found). |

## Script Variables (NinjaOne)

| Variable | Env var | Purpose |
|----------|---------|---------|
| Bootstrapper Path | `bootstrapperPath` | Path to Deploy.exe (default `C:\RMM\Deploy.exe`). |
| Bootstrapper Arguments | `bootstrapperArguments` | Arguments for bootstrapper (default `/q /p`). |
| Summary Path | `summaryPath` | Path to Summary.txt (default `C:\Autodesk\Deploy\AutoCADLT2025\Summary.txt`). |
| Installer Path | `installerPath` | Path to Installer.exe. |
| Collection XML Path | `collectionXmlPath` | Path to Collection.xml. |
| Installer Version | `installerVersion` | Value for `--installer_version` (default `2.21.0.623`). |
| Bootstrap Timeout Minutes | `bootstrapTimeoutMinutes` | Max minutes to wait for Summary.txt (default 60). |
| Bootstrap Poll Seconds | `bootstrapPollSeconds` | Seconds between checks (default 15). |
| Post Wait Minutes | `postWaitMinutes` | Minutes to wait after Summary.txt before install (default 10). |
| Install Only | `installOnly` | Set to `1` or `true` to skip bootstrap and only run the installer (image already staged). |

## Prerequisites

- NinjaOne transfers the bootstrapper (e.g. `Deploy.exe`) to the device before running this script.
- Script runs with appropriate privileges for install.

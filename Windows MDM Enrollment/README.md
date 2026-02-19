# Get-MDMEnrollment.ps1

## Overview

This PowerShell script determines if a Windows device is MDM-enrolled and identifies the MDM provider by parsing `dsregcmd /status`, registry keys under `HKLM:\SOFTWARE\Microsoft\Enrollments` and `OMADM\Accounts`, and inferring vendor from URLs and provider IDs. It optionally updates NinjaOne custom fields with enrollment status and provider name.

**Run context:** This is a **device script** intended to run on each managed Windows endpoint (e.g. via NinjaOne script deployment). It does not call the NinjaOne API; it runs locally on the device.

**dsregcmd parsing:** The script reads MDM-related fields from `dsregcmd /status` using case-insensitive and variant-tolerant label matching (e.g. `MdmUrl` or `MDMUrl`, `MDMUserUPN` or `MDM User UPN`) so behavior is robust across Windows versions and builds.

## Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | dsregcmd not found or fatal error |

## Prerequisites

1. **PowerShell 5.1+**: The script runs on Windows PowerShell 5.1 or later (PowerShell 7 not required).

2. **Windows version**: `dsregcmd` is available on Windows 10 1507+ and Windows Server 2016+. On older systems, the script will report that dsregcmd was not found.

3. **Administrator rights** (recommended): Registry access to `HKLM:\SOFTWARE\Microsoft\Enrollments` and `OMADM\Accounts` may require elevated privileges for full results.

4. **NinjaOne custom fields** (when using `-UpdateNinjaProperties`):
   - `mdmStatus` – Will be set to "Enrolled" or "Not Enrolled"
   - `mdmProvider` – Will be set to the detected vendor (e.g. "Microsoft Intune") or fallback values

   Create these device-level custom fields in NinjaOne before enabling `-UpdateNinjaProperties`.

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-OutputFormat` | `List` (default) or `Json`. Use Json for machine-consumable output. |
| `-UpdateNinjaProperties` | When set, writes `mdmStatus` and `mdmProvider` to NinjaOne custom fields via Ninja-Property-Set. Skips if the cmdlet is not available. |

## Examples

```powershell
# Default: human-readable list output
.\Get-MDMEnrollment.ps1

# JSON output for automation
.\Get-MDMEnrollment.ps1 -OutputFormat Json

# Update NinjaOne custom fields
.\Get-MDMEnrollment.ps1 -UpdateNinjaProperties
```

## Detected MDM Vendors

The script attempts to identify common MDM platforms from URLs and provider metadata:

- Microsoft Intune
- VMware Workspace ONE (AirWatch)
- IBM MaaS360
- Ivanti (MobileIron)
- SOTI MobiControl
- Citrix Endpoint Management
- 42Gears SureMDM

If no match is found, `VendorGuess` is null and fallback values (ProviderID, ProviderName, MDMUrl) are used for the NinjaOne `mdmProvider` field.

## Detection details

- **EnrollmentType:** The script exposes `EnrollmentType` from the enrollment registry (user vs device) in the summary object when available.
- **Multiple enrollments:** When more than one enrollment exists (e.g. Intune and another MDM), the script prefers the Microsoft Intune enrollment for URL, ProviderID, and ProviderName so `VendorGuess` and `mdmProvider` are stable and Intune is reported when present.

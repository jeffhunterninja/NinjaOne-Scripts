# **NinjaOne Windows Patch Installation Report Script**

Video overview of settings up an API server and this script: https://www.youtube.com/watch?v=Qy9g6-KVfbo

## **Overview**
This PowerShell script generates a report of Windows device patch installations for the specified month and year using the **NinjaOne API**. It can generate:
- **Documentation** in the NinjaOne platform
- **Knowledge Base (KB)** articles for organizations
- A **global KB** article for all organizations
- **Local CSV or HTML** exports via companion scripts (see below)

---

## **Features**
- Automatically retrieves patch installation details for **Windows Workstations** and **Servers**.
- Allows reporting for:
  - The **current month** (default).
  - A **specific month and year** using the `-ReportMonth` parameter (e.g., "December 2024").
- Creates or updates:
  - Organization-specific documentation.
  - Organization-specific Knowledge Base articles in NinjaOne.
  - Global KB articles for aggregated patch data.
- Microsoft Defender Updates are excluded from this report
---

## **CSV Export Script: Export-WindowsPatchReportToCsv.ps1**

A companion script that uses the **same patch data** (one line per KB per device) as `Publish-WindowsPatchReport.ps1` but exports to CSV instead of NinjaOne Documents or Knowledge Base. Use it when you need the report in a spreadsheet or for external processing.

**Parameters:**

| Parameter | Description |
|-----------|-------------|
| `-ReportMonth` | Optional. Month and year for the report (e.g. `"December 2024"`). Omit for the current month. |
| `-OutputPath` | Optional. Full path for the CSV file. If omitted, writes `WindowsPatchReport_<YYYYMM>.csv` in the current directory. |
| `-PerOrganization` | Optional switch. When set, writes one CSV per organization (e.g. `WindowsPatchReport_<OrgName>_<YYYYMM>.csv`) instead of a single combined file. |

**CSV columns:** OrganizationName, DeviceName, PatchName, KBNumber, InstalledAt, Timestamp, DeviceId.

**Example:**

```powershell
# Current month, single CSV in current directory
.\Export-WindowsPatchReportToCsv.ps1

# Specific month, custom path
.\Export-WindowsPatchReportToCsv.ps1 -ReportMonth "December 2024" -OutputPath "C:\Reports\patches.csv"

# One CSV per organization
.\Export-WindowsPatchReportToCsv.ps1 -ReportMonth "January 2025" -PerOrganization
```

Credentials and prerequisites are the same as for `Publish-WindowsPatchReport.ps1` (NinjaOne API credentials, PowerShell 7+, NinjaOneDocs module). NinjaOne Documentation does not need to be enabled for the CSV export script.

---

## Prerequisites

1. **Set Up an API Server/Automated Documentation Server**
   Follow the instructions here:  
   [https://docs.mspp.io/ninjaone/getting-started](https://docs.mspp.io/ninjaone/getting-started)

2. **PowerShell 7+**  
   If PowerShell 7 is not installed, the script will prompt to restart itself in PowerShell 7.

3. **NinjaOne Documentation**  
   You will need to have NinjaOne Documentation enabled in your NinjaOne instance.

4. **Script Variables**  
   Add these in the NinjaOne script editor after importing the script.


| Name   | Pretty Name            | Type   |
|------------------------|------------------------|--------|
| `sendToKnowledgeBase`  | Send To Knowledge Base | Checkbox |
| `sendToDocumentation`  | Send To Documentation  | Checkbox |
| `globalOverview`       | Global Overview        | Checkbox |
| `reportMonth`          | Report Month           | String |

**Report Month** (`reportMonth`): Leave blank for the current month. For a specific month use full month name and year, e.g. `December 2024`. Historical reports use the correct date range for patch and activity data.


![Patch Installations](https://github.com/jeffhunterninja/NinjaOne-Scripts/blob/main/Windows%20OS%20Patch%20Reporting/patchinstallations.png)

---

## Troubleshooting

### Error: `NO_ALLOWED_DOCUMENT_TEMPLATE_FOUND`

This error can occur when the **"Patch Installation Reports"** document template has been **archived** in NinjaOne. Archived templates are not allowed for creating or updating documents via the API.

**To resolve:**

- **Restore the template:** In NinjaOne, go to **Administration** → **Documentation** (or **API** → **API Documentation**), find the archived "Patch Installation Reports" template, and restore it.
- **Or delete the archived template** so the script can create a fresh one on the next run (the script creates the template automatically if it doesn't exist).

---

## **HTML Export Script: Export-WindowsPatchReportToHtml.ps1**

A companion script that uses the **same patch data and logic** as `Publish-WindowsPatchReport.ps1` but exports to **local HTML** instead of NinjaOne Documents or Knowledge Base. Use it when you need a shareable or printable report in a browser. By default it produces simple styled HTML with device links to NinjaOne; optionally use the PSWriteHTML module for richer styled reports.

**Parameters:**

| Parameter | Description |
|-----------|-------------|
| `-ReportMonth` | Optional. Month and year for the report (e.g. `"December 2024"`). Omit for the current month. |
| `-OutputPath` | Optional. Full path for the HTML file. If omitted, writes `WindowsPatchReport_<YYYYMM>.html` in the current directory. |
| `-PerOrganization` | Optional switch. When set, writes one HTML file per organization (e.g. `WindowsPatchReport_<OrgName>_<YYYYMM>.html`) instead of a single combined file. Per-org files include patch statistics (scan cycles, apply cycles, installations, failures). |
| `-UsePSWriteHTML` | Optional switch. When set, uses the PSWriteHTML module for styled reports (installed automatically if missing). Otherwise uses built-in simple HTML. |

**Output:** Single file by default, or one file per organization with `-PerOrganization`. Table columns: OrganizationName, DeviceName, PatchName, KBNumber, InstalledAt. In simple HTML mode, device names link to the NinjaOne device dashboard when the instance is available.

**Example:**

```powershell
# Current month, single HTML in current directory
.\Export-WindowsPatchReportToHtml.ps1

# Specific month, custom path
.\Export-WindowsPatchReportToHtml.ps1 -ReportMonth "December 2024" -OutputPath "C:\Reports\patches.html"

# One HTML per organization
.\Export-WindowsPatchReportToHtml.ps1 -ReportMonth "January 2025" -PerOrganization

# Styled report using PSWriteHTML module
.\Export-WindowsPatchReportToHtml.ps1 -ReportMonth "January 2025" -UsePSWriteHTML
```

Credentials and prerequisites are the same as for `Publish-WindowsPatchReport.ps1` (NinjaOne API credentials, PowerShell 7+, NinjaOneDocs module). NinjaOne Documentation does not need to be enabled for the HTML export script.

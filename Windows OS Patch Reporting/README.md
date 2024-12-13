# **NinjaOne Windows Patch Installation Report Script**

## **Overview**
This PowerShell script generates a report of Windows device patch installations for the specified month and year using the **NinjaOne API**. It can generate:
- **Documentation** in the NinjaOne platform
- **Knowledge Base (KB)** articles for organizations
- A **global KB** article for all organizations

---

## **Features**
- Automatically retrieves patch installation details for **Windows Workstations** and **Servers**.
- Allows reporting for:
  - The **current month** (default).
  - A **specific month and year** using the `-ReportMonth` parameter (e.g., "December 2024").
- Creates or updates:
  - Organization-specific documentation.
  - Knowledge Base articles in NinjaOne.
  - Global KB articles for aggregated patch data.
- Microsoft Defender Updates are excluded from this report
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


![Patch Installations](https://github.com/jeffhunterninja/NinjaOne-Scripts/blob/main/Windows%20OS%20Patch%20Reporting/patchinstallations.png))

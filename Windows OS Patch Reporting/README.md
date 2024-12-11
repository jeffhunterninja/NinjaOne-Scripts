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

---

## Prerequisites

1. **PowerShell 7+**  
   If PowerShell 7 is not installed, the script will prompt to restart itself in PowerShell 7.

2. **NinjaOne Documentation**  
   You will need to have NinjaOne Documentation enabled in your NinjaOne instance.

3. **Script Variables**  
   Add these in the NinjaOne script editor after importing the script.


| Environment Variable   | Pretty Name            | Type   |
|------------------------|------------------------|--------|
| `sendToKnowledgeBase`  | Send To Knowledge Base | Switch |
| `sendToDocumentation`  | Send To Documentation  | Switch |
| `globalOverview`       | Global Overview        | Switch |
| `reportMonth`          | Report Month           | String |




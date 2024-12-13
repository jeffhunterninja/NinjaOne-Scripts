# **NinjaOne Warranty Report Script**

## **Overview**
This PowerShell script generates a report of Windows device patch installations for the specified month and year using the **NinjaOne API**. It can generate:
- **Documentation** in the NinjaOne platform
- **Knowledge Base (KB)** articles for organizations
- A **global KB** article for all organizations

---

## **Features**
- Automatically retrieves warranty information from the NinjaOne API.
- Creates or updates:
  - Organization-specific warranty documentation in NinjaOne Apps & Services. (This is useful for exporting as a PDF in NinjaOne reporting)
  - Organization-specific warranty Knowledge Base articles in NinjaOne. (This is useful for creating reports/dashboards that end users with specific roles can access)
  - Global KB articles for aggregated patch data. (Displays all devices regardless of organization)

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


| Environment Variable   | Pretty Name            | Type   |
|------------------------|------------------------|--------|
| `sendToKnowledgeBase`  | Send To Knowledge Base | Switch |
| `sendToDocumentation`  | Send To Documentation  | Switch |
| `globalOverview`       | Global Overview        | Switch |

I'd recommend running this on a monthly basis, or weekly at the most frequent, as warranty updates in NinjaOne do not occur on a more frequenct cadence than weekly.

![Warranty Report](https://github.com/jeffhunterninja/NinjaOne-Scripts/blob/main/Warranty%20Reporting/warrantyreport.png?raw=true)


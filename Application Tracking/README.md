# Application Tracking

These scripts form the basis of a framework that allows you to use **NinjaOne** to establish a whitelist of authorized applications in an **organization custom field** and compare installed applications on a device against that whitelist. You can also add per-device overrides in an optional **device custom field**. Unauthorized applications are written into a custom field for alerting/reporting purposes.

This may not be a feasible workflow depending on your environment. For example:

- If every user has admin or installation rights, the alerts will eventually reach an unsustainable volume.
- Users used to having install rights may not react positively to having them withdrawn
- However, if users are otherwise locked down already, this could be a useful tool to identify unexpected application installations.

**Only Windows devices are supported today**

> **Recommendation:** Use this process primarily for tracking application usage across endpoints as a **periodic report**, rather than for security/operational processes requiring immediate responses.

---

## Steps

Before utilizing these scripts, there are some **prerequisites**:

### 1. Set Up an API Server/Automated Documentation Server  
Follow the instructions here:  
[https://docs.mspp.io/ninjaone/getting-started](https://docs.mspp.io/ninjaone/getting-started)

---

### 2. Custom Fields Configuration  

Create the following custom fields with the specified permissions. You can rename these fields, but you must modify the scripts accordingly.

| **Name**                   | **Display Name**             | **Permissions**                                                                 | **Scope**               | **Type**                                |
|----------------------------|------------------------------|---------------------------------------------------------------------------------|-------------------------|-----------------------------------------------|
| **softwareList**           | Software List                | - Read-Only to Technicians  <br> - Read-Only to Automations <br> - Read/Write to API | Organization and Device | WYSIWYG  |
| **deviceSoftwareList**     | Device Software List         | - Read-Only to Technicians <br> - Read-Only to Automations <br> - Read/Write to API | Device                  | WYSIWYG |
| **unauthorizedApplications** | Unauthorized Applications   | - Read-Only to Technicians <br> - Read/Write to Automations <br> - Read-Only to API | Device                  | Multi-line |


---

### 3. Import Scripts  

Import the scripts in this repository into **NinjaOne**. Each script has script variables that need to be created within the NinjaOne script editor.

---

## Script: Check-AuthorizedApplications

### Execution  
Run as **script result condition**.

### Script Variables Required  

Create these variables in the NinjaOne script editor:

| **Name**              | **Pretty Name**                          | **Script Variable Type** |
|-----------------------|------------------------------------------|--------------------------|
| **matchingCriteria**  | Matching Criteria                       | Drop-Down                |

This variable controls the matching mode of the script. Enter all options in the **Mode** column as options for the drop-down script variable.

## Matching Mode

| **Mode**            | **Case-Sensitive** | **Matching Behavior**                  | **Use Case**                      |
|----------------------|--------------------|----------------------------------------|-----------------------------------|
| **Exact**           | Yes                | Full, identical string matches only    | When precision is critical.       |
| **CaseInsensitive** | No                 | Matches identical strings, ignores case| When case differences exist.      |
| **Partial**         | Yes                | Checks if authorized app is a substring| For loose or fuzzy comparisons.   |

---

## Script: Update-AuthorizedApplications

### Execution  
Run as needed from the API Server/Automated Documentation Server to update which applications are authorized at the organization or device levels.

### Script Variables Required  

| **Name**                                    | **Pretty Name**                                | **Script Variable Type** |
|-------------------------------------------|--------------------------------------------|--------------------------|
| **commaSeparatedListOfOrganizationsToUpdate** | Comma Separated List Of Organizations To Update | String/Text              |
| **updateOrganizationsBasedOnCurrentSoftwareInventory** | Update Organizations Based On Current Software Inventory | Checkbox                  |
| **appendToOrganizations**                   | Append To Organizations                     | String/Text              |
| **softwareToAppend**                        | Software To Append                          | String/Text              |
| **removeFromOrganizations**                 | Remove From Organizations                   | String/Text              |
| **softwareToRemove**                        | Software To Remove                          | String/Text              |
| **appendToDevices**                         | Append To Devices                           | String/Text              |
| **deviceSoftwareToAppend**                  | Device Software To Append                   | String/Text              |
| **removeFromDevices**                       | Remove From Devices                         | String/Text              |
| **deviceSoftwareToRemove**                  | Device Software To Remove                   | String/Text              |

---

### Main Functions of the Script  

1. **Authorize Installed Applications:**  
   - Authorize all applications currently installed across all organizations.  
   - Authorize applications for selected organizations using a comma-separated list.
   - **Using this option will overwrite data already present in the organization custom fields**

2. **Append Applications to Organization(s):**  
   - Append software to authorized applications for all or specific organizations.

3. **Remove Applications from Organization(s):**  
   - Remove software from authorized applications for all or specific organizations.

4. **Per-Device Overrides:**  
   - Append or remove software for specific devices.

---

## Script: Recover-AuthorizedApplications

### Execution  
Run on a **daily cadence** to back up or restore as needed.

### Script Variables Required  
For backups, only the **Action** variable must be entered.
A **BackupFile** or **BackupDirectory** must be specified for restorations, and a **TargetType** must be selected.

| **Name**              | **Pretty Name**                          | **Script Variable Type** |
|-----------------------|------------------------------------------|--------------------------|
| **Action**            | Action                                  | Drop-Down                |
| **BackupFile**        | Backup File                             | String/Text              |
| **BackupDirectory**   | Backup Directory                        | String/Text              |
| **TargetType**        | Target Type                             | Drop-Down                |
| **RestoreTargets**    | Restore Targets                         | String/Text              |

> **Recommendation:** Run the backup script **daily** or **weekly**.

---

## Script: Report-UnauthorizedApplications  

> **Requirement:** This script utilizes the NinjaOneDocs Powershell module located here: [https://github.com/lwhitelock/NinjaOneDocs/tree/main/Public](https://github.com/lwhitelock/NinjaOneDocs/tree/main/Public).

This script outputs:  

- A comma-separated list of unauthorized applications across the entire environment.  
- A CSV export of unauthorized applications by device and listing each location and organization into C:\temp.

> **Note:** This depends on the "Unauthorized Applications" custom field. Offline agents may not update the field until they come back online.

---

## Caveats and Disclaimers  

- **User-context Installations:**  
  NinjaOne does not track user-context installations by default.

- **Not Zero-Trust:**  
  This framework does not prevent privilege escalation but provides insights into unexpected installations.

- **False Positives:**  
  Expect false positives, especially from application name changes. **Partial matching** can help reduce these but increases the risk of missed unauthorized installations.

- **Response Strategy:**  
  Responding to unauthorized applications requires time. Without adequate user controls, alerts can quickly spiral.  

> **Recommendation:** Review unauthorized applications on a **weekly** or **monthly** cadence for sustainability.

---

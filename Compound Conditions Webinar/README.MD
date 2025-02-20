Thank you to all who attended the Compound Conditions webinar, which you can find on YouTube [here]{https://youtu.be/y7WvImYobJ0}

## First Example: Software Deployment with Compound Conditions

The first example revolved around software deployment by utilizing a compound condition to check for the absence of software (Google Chrome in this case), while also looking for the text string **"Google Chrome"** in a custom field.

To update more easily, you can leverage an automated documentation server or API server to bulk update and append to custom fields.

### Steps:
- **Read more about setting up an API server here:** [NinjaOne API Setup](https://docs.mspp.io/ninjaone/getting-started)
- **Set up a multi-line custom field** with 'Organization' scope, with Read/Write capabilities for the API and Read-Only for automations.
  - Whether you want technicians to be able to modify this field is up to you—otherwise, only system administrators would be able to update, remove, and add software for deployment using the API server method.
  - If technicians can modify this, it's recommended to regularly back up the list of applications to be installed via API.
- **Set up a second multi-line custom field** with 'Device' scope so that applications can be specified per device.

### Script for Updating Deployed Software
Access the script to update deployed software here, and import it into NinjaOne.

#### Required Script Variables:
| Name | Pretty Name | Script Variable Type |
|------|------------|----------------------|
| `appendToOrganizations` | Append To Organizations | String/Text |
| `softwareToAppend` | Software To Append | String/Text |
| `removeFromOrganizations` | Remove From Organizations | String/Text |
| `softwareToRemove` | Software To Remove | String/Text |
| `appendToDevices` | Append To Devices | String/Text |
| `deviceSoftwareToAppend` | Device Software To Append | String/Text |
| `removeFromDevices` | Remove From Devices | String/Text |
| `deviceSoftwareToRemove` | Device Software To Remove | String/Text |

You can now use the API server and the `Update-SoftwareDeployment` script to start creating a comma-separated list of applications that should be installed for an organization. Then, create a **compound condition** looking for:
1. The absence of that software.
2. The text string identifying your application that is found in your custom field(s).

---

## Second Example: Authorized Applications Tracking

The second example was an **inversion of the first**—instead of setting up applications for deployment, this approach establishes a **list of authorized applications** in an organization scope and/or device scope WYSIWYG custom field.

📌 **Full documentation and scripts for this workflow:**  
[GitHub - NinjaOne Application Tracking](https://github.com/jeffhunterninja/NinjaOne-Scripts/tree/main/Application%20Tracking)

### Workflow Overview:
- **`Check-AuthorizedApplications`** runs as a **script result condition** as part of a compound condition.
  - A **custom field condition** (e.g., a checkbox) can be used as an enablement switch if you don’t want these checks to apply to all organizations.
- **`Update-AuthorizedApplications`** runs on your API server and is used to:
  - Update, remove, or append application names to a **comma-separated string** stored in a WYSIWYG field at either the organization or device level.
- **`Recover-AuthorizedApplications`** backs up your stored comma-separated string with an **option to restore** if necessary.
- **`Report-UnauthorizedApplications`** retrieves all currently reported unauthorized applications for:
  - A **global view** of unauthorized applications.
  - Easier **whitelisting at scale**.

## Third Example: Using NinjaOne Documentation for Device Configuration

The third example utilized **NinjaOne Documentation** to create a document with **custom field checkboxes** that indicate whether certain device configurations should be enabled or not.

### Process:
1. **Custom Field Checkboxes**  
   - A document is created in **NinjaOne Documentation** with checkboxes that specify if certain device configurations should be enabled (e.g., "Enable Firewall").
   
2. **Script Result Condition (Check Configuration Intent)**  
   - A **script result condition** is used as part of a **compound condition** to check if the checkbox is marked in the documentation.

3. **Script Result Condition (Verify Device State)**  
   - A second **script result condition** evaluates the actual **state of the device**.
   - For example, if the **Device Configuration** document has the checkbox marked for **"Enable Firewall"**, the template library script **"Audit Firewall Status"** is used to confirm whether the firewall is actually enabled on the device.

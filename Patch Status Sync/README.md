# Sync-PatchStatus.ps1

## Overview

This PowerShell script retrieves patch status information (pending, failed, and approved patches) from the NinjaOne  API and updates corresponding custom fields on devices within the NinjaOne platform. It ensures that the custom fields `pendingPatches`, `approvedPatches`, and `failedPatches` always reflect the current state of each device's patching. Recommend running this once per hour. Only values that have changed will be modified.

## Prerequisites

1. **PowerShell 7+**:  
   The script requires PowerShell 7 or later.

2. **Set Up an API Server/Automated Documentation Server**
   Follow the instructions here:  
   [https://docs.mspp.io/ninjaone/getting-started](https://docs.mspp.io/ninjaone/getting-started)

3. **Custom Fields in NinjaOne**:
   Prior to executing this script, you must create three custom text fields at the device level in NinjaOne:
   - `pendingPatches`
   - `approvedPatches`
   - `failedPatches`

   These fields will be updated by the script to reflect the current patch states of each device.

   **Note:**  
   - To create a custom field in NinjaOne:
     1. Navigate to **Administration** > **Devices** > **Custom Fields** (role or global).
     2. Create a new custom field with the name corresponding to each required field.  
   - Ensure the fields are **Multi-line** fields, and if role custom fields are properly applied to `Windows` workstation and server classes.

## What the Script Does

1. **Checks PowerShell Version**:  
   If the script is not running in PowerShell 7 or later, it tries to restart itself in `pwsh`.

2. **Loads the NinjaOneDocs Module**:  
   Installs and imports the `NinjaOneDocs` module if necessary.

3. **Retrieves API Credentials**:  
   Uses `Ninja-Property-Get` to pull `ninjaoneInstance`, `ninjaoneClientId`, and `ninjaoneClientSecret`.

4. **Connects to the NinjaOne API**:  
   Uses `Connect-NinjaOne` to establish a session.

5. **Fetches Device and Patch Information**:  
   Queries the NinjaOne API for devices and their associated patch status (pending, failed, and approved patches).

6. **Updates Custom Fields**:  
   For each device, the script compares the current value of the `pendingPatches`, `approvedPatches`, and `failedPatches` fields to the actual patch data retrieved from the API. If there's a difference:
   - The script updates the custom field with the latest patch names.

7. **Removes Stale Data**:  
   If a device no longer has pending, failed, or approved patches (based on current API data) but the custom field still contains old values, the script clears them out.

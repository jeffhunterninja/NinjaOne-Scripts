# Set-DeviceTimezoneCustomField

Retrieves the Windows device timezone and writes it to a NinjaOne device custom field.

## Purpose

When run on a Windows device via NinjaOne (script assignment or scheduled task), the script reads the system timezone (e.g. `Eastern Standard Time`) and writes the timezone ID to a device custom field. The value is consistent and machine-parseable for reporting, filtering, or automation.

## Prerequisites

- **PowerShell 5.1 or later** (Windows)
- **NinjaOne device custom field**: Create a device custom field in NinjaOne under **Devices → Custom Fields**. Use type **TEXT**. The default field name used by the script is `timezone`; you can create a field with that name or pass a different name via `-CustomFieldName` or the `timezoneCustomField` script variable in NinjaOne.

## Deployment

1. In NinjaOne, create the device custom field (e.g. name: `timezone`, type: TEXT).
2. Add **Set-DeviceTimezoneCustomField.ps1** as a script in NinjaOne.
3. Assign the script to the desired devices or policies (e.g. run on a schedule or on demand).
4. When the script runs on a device, the NinjaOne agent provides `Ninja-Property-Set`; the script uses it to write the timezone ID to the device’s custom field.

## Usage

- **Default (write to field "timezone")**: Run the script as deployed; no parameters needed.
- **Custom field name**: Use script parameter `-CustomFieldName deviceTimezone` or set NinjaOne script variable `timezoneCustomField` to the field name.
- **Local testing (no Ninja write)**: Run with `-NoNinjaWrite` to detect and display the timezone without calling Ninja-Property-Set.

## Exit codes

- **0** – Success (timezone detected and, unless `-NoNinjaWrite`, written to the custom field).
- **1** – Error (e.g. could not get timezone, or Ninja-Property-Set failed / not available).

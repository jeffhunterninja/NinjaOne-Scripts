# NinjaOne API Device Import and Custom Field Update Script

This script demonstrates how to interact with the NinjaOne API, import device data from a CSV, match it with devices from NinjaOne, and update custom fields dynamically based on CSV headers.  
It includes enhanced error handling, logging, and dynamic request body generation. This has only been tested with text, multiline, and checkbox types of custom fields. Only field types that accept a string value are likely to work, anything involving attachments or WYSIWYG fields is not likely to function.

## Parameters

### `$OverwriteEmptyValues`
- **Description**: Determines if empty CSV values are included as `$null` (overwriting existing data) or excluded from the update payload.  
- **Default Value**: `$false`

## Prerequisites

Before running this script, ensure the following:

1. The CSV file (`csvexample.csv`) must have at least the following columns:
   - `Id`
   - `Name`
2. Additional columns (e.g., `assetOwner`, `location`, etc.) will be used as custom fields.
3. Replace the following placeholders in the script with your actual values:
   - `$NinjaOneClientId`
   - `$NinjaOneClientSecret`
   - `$csvPath` (Path to the CSV file)

## Usage

Run the script in PowerShell with the required parameters:

```powershell
.\ScriptName.ps1 -OverwriteEmptyValues $true

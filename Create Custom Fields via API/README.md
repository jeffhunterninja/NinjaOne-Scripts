# NinjaOne Custom Fields Creator

Creates NinjaOne custom fields (node attributes) via the Public API using **OAuth 2.0 Authorization Code flow**. The script currently supports **bulk creation from a CSV file** only.

## Scripts

| Script | Purpose |
|--------|---------|
| **New-NinjaOneCustomField.ps1** | Bulk-creates custom fields from a CSV file. Uses Authorization Code flow: opens a browser for sign-in, captures the redirect via a local HTTP listener, then exchanges the code for an access token and calls the bulk API. |
| **Convert-NinjaOneCustomFieldsJsonToCsv.ps1** | Converts a NinjaOne custom-fields JSON export to the CSV format expected by `New-NinjaOneCustomField.ps1 -CsvPath`. You must supply your own JSON file (e.g. from NinjaOne or a backup). |

## Prerequisites

- PowerShell 5.1 or later
- A NinjaOne API app registered as a **Regular Web Application** with **Authorization Code** grant type

## App registration

1. In NinjaOne go to **Administration** > **Apps** > **API**.
2. Create or edit an app and choose **Regular Web Application** (or equivalent that supports Authorization Code).
3. Grant types: enable **Authorization Code**.
4. Scopes: include at least **monitoring** and **management** (required for custom field creation).
5. **Redirect URI**: add exactly the URL the script will listen on. The script’s default `-RedirectUri` is `http://localhost:8888/`; use the same value here (including trailing slash).
6. Save and note the **Client ID** and **Client Secret** (secret is shown only once).

## Usage

### Bulk create from CSV

Pass a CSV file with `-CsvPath`. The CSV must include a **Label** column (required). The API field name is derived automatically from each label using camelCase. Other columns are optional: **Description**, **Type**, **DefinitionScope**, **TechnicianPermission**, **ScriptPermission**, **ApiPermission**, **DefaultValue**, **DropdownValues**.

- **DefinitionScope**: use semicolons for multiple scopes (e.g. `NODE;END_USER`). Defaults to `NODE` if empty.
- **DropdownValues**: for DROPDOWN or MULTI_SELECT types, list options separated by semicolons (e.g. `Low;Medium;High`).
- Rows with blank Label are skipped. Invalid **Type** values fall back to `TEXT`.

Use UTF-8 encoding for the CSV if you need special characters in labels or descriptions.

```powershell
.\New-NinjaOneCustomField.ps1 `
  -NinjaOneInstance 'app.ninjarmm.com' `
  -ClientId 'your-client-id' `
  -RedirectUri 'http://localhost:8888/' `
  -CsvPath '.\NinjaOne-CustomFields-Example.csv'
```

An example CSV, **NinjaOne-CustomFields-Example.csv**, is included in this folder. You can use it as a template or generate a similar CSV from a NinjaOne JSON export using **Convert-NinjaOneCustomFieldsJsonToCsv.ps1** (see below).

If you do not pass `-ClientSecret`, the script uses `$env:NinjaOneClientSecret`. On first run (without `-AccessToken` or `-TokenFile`), a browser opens for NinjaOne sign-in; after you approve, the script receives the authorization code and creates the fields.

### Optional: `-UseTestConfig`

After editing the `$TestConfig` hashtable at the top of **New-NinjaOneCustomField.ps1** (placeholders only—never commit a real secret), you can run:

```powershell
.\New-NinjaOneCustomField.ps1 -UseTestConfig
```

Prefer `$env:NinjaOneClientSecret` when using test config so the secret is not stored in the file.

### Using a cached token

To skip the browser flow when you already have an access token:

```powershell
.\New-NinjaOneCustomField.ps1 `
  -NinjaOneInstance 'app.ninjarmm.com' `
  -AccessToken 'your-access-token' `
  -RedirectUri 'http://localhost:8888/' `
  -CsvPath '.\NinjaOne-CustomFields-Example.csv'
```

Or save a token to a file and use `-TokenFile 'C:\path\to\token.txt'`.

### Converting a NinjaOne JSON export to CSV

If you have a custom-fields export from NinjaOne as JSON (either a root-level array of field objects or an object with a `customFields` array), use **Convert-NinjaOneCustomFieldsJsonToCsv.ps1** to convert it to the CSV format expected by `-CsvPath`. You must supply your own JSON file (e.g. `customfieldstemplate.json`); it is not included in this folder.

Default run (reads `customfieldstemplate.json` in the script directory, writes **NinjaOne-CustomFields-Example.csv**):

```powershell
.\Convert-NinjaOneCustomFieldsJsonToCsv.ps1
```

With custom paths:

```powershell
.\Convert-NinjaOneCustomFieldsJsonToCsv.ps1 -JsonPath 'C:\export.json' -CsvPath '.\NinjaOne-CustomFields-Example.csv'
```

The resulting CSV can be passed to `New-NinjaOneCustomField.ps1 -CsvPath` for bulk create.

### Planned / not yet implemented

- **Single custom field** (e.g. `-Label`, `-Type`, `-DefinitionScope` without CSV).
- **Bulk create from hashtable array** (`-FieldDefinitions`).
- **Bulk create from JSON file** (`-JsonPath`).

## Field types and options

- **Type**: e.g. `TEXT`, `TEXT_MULTILINE`, `WYSIWYG`, `DROPDOWN`, `MULTI_SELECT`, `CHECKBOX`, `NUMERIC`, `DATE`, `EMAIL`, `URL`. See script help or the API docs for the full list.
- **DefinitionScope**: `NODE` (devices), `END_USER`, `LOCATION`, `ORGANIZATION`. Pass one or more in the CSV (semicolon-separated).
- **Permissions**: TechnicianPermission, ScriptPermission, ApiPermission (e.g. `NONE`, `READ_ONLY`, `READ_WRITE`) as CSV columns.
- **Dropdown/Multi-select**: use the **DropdownValues** column with options separated by semicolons.

## Script-specific custom field definitions

Custom field definitions for other scripts in this repo (e.g. maintenance windows, patching schedules) live in those scripts’ folders. Use the CSV in those folders with `-CsvPath` when creating fields for that workflow. See [Set-CustomFieldMaintenanceWindow/](../Set-CustomFieldMaintenanceWindow/) and [Set-CustomFieldPatchingSchedule/](../Set-CustomFieldPatchingSchedule/) for the maintenance-window and patching-schedule field definitions.

## API references

- [createNodeAttribute](https://ca.ninjarmm.com/apidocs-beta/core-resources/operations/createNodeAttribute) – single field
- [bulkCreateNodeAttributes](https://ca.ninjarmm.com/apidocs-beta/core-resources/operations/bulkCreateNodeAttributes) – multiple fields (all-or-nothing)

## Security

- Do not commit or log the client secret. Prefer `$env:NinjaOneClientSecret` or a secure store.
- Redirect URI must match exactly between NinjaOne and the script (including trailing slash for `http://localhost:8888/`).
- The script listens only on the host/port from `-RedirectUri` (e.g. `localhost:8888`) and stops after receiving one request.
- For OAuth troubleshooting, run **New-NinjaOneCustomField.ps1** with `-Verbose` to show authorize URL and host (client secret is never printed).

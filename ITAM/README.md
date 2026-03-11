# NinjaOne ITAM Manager

Standalone PowerShell WPF application that combines equipment import (CSV or manual entry), device QR code generation, QR upload to devices as related items, and scan-and-assign (user + device QR codes → set device owner). Uses OAuth Authorization Code + PKCE; session state flows between tabs (e.g. imported device IDs pre-fill QR generation; QR output directory pre-fills the upload tab).

**Script:** [Invoke-NinjaITAMManager.ps1](Invoke-NinjaITAMManager.ps1)

## Prerequisites

- **PowerShell:** 5.1 or later (script uses `#Requires -Version 5.1`).
- **Platform:** Windows (WPF: `PresentationFramework`, `PresentationCore`, `WindowsBase`, `System.Windows.Forms`). Not compatible with PowerShell Core on non-Windows.
- **NinjaOne OAuth app:** Configure in NinjaOne as **Native** (Authorization Code). Redirect URI must be `http://localhost` (port wildcard for native apps per RFC 8252). Scopes: `monitoring`, `management`.
- **Optional:** Tab 2 (Generate QR Codes) calls **api.qrserver.com** over HTTPS; no API key required. Outbound HTTPS is needed for that tab.

## Parameters and environment variables

| Parameter | Description | Default |
|-----------|-------------|---------|
| `NinjaOneInstance` | Instance hostname or base URL | `$env:NINJA_BASE_URL` or `ca.ninjarmm.com` |
| `ClientId` | OAuth application Client ID | `$env:NinjaOneClientId` |
| `AllowInsecureHttp` | Allows `http://` instance URLs for local testing only | Off (HTTPS required) |

Connection settings can also be entered in the UI (Instance, Client ID) and persist for the session.

## Tabs (workflow)

**Connection:** Expand "Connection Settings", enter Instance and Client ID, then click "Sign In to NinjaOne". A browser opens for OAuth; after sign-in, return to the app.

### Tab 1 — Import Equipment

- **CSV:** Browse to a CSV file; a preview loads. Click "Import" to create unmanaged devices in NinjaOne.
  - **Required:** `RoleName` (must match an unmanaged-device role in NinjaOne).
  - **Org/Location:** Provide either `OrganizationId` + `LocationId` **or** `OrganizationName` + `LocationName`.
  - **Optional:** Name, Make, Model, SerialNumber, WarrantyStartDate, WarrantyEndDate, PurchaseDate, PurchaseAmount. If Name is omitted, it is derived from Make/Model or "Unmanaged {RoleName} {row}".
  - Devices are created via `itam/unmanaged-device`; then custom fields (manufacturer, model, itamAssetSerialNumber, itamAssetPurchaseDate, itamAssetPurchaseAmount) are patched when provided.
- **Manual:** Select Organization (locations load), Role, and Location; fill Name, Serial, Make, Model, Purchase Date/Amount, Warranty Start/End; click "Add Device". Same API as CSV.
- The list of imported devices is used by Tab 2 ("From Import").

### Tab 2 — Generate QR Codes

- Add device IDs manually (text box + "Add") or click "From Import" to load IDs from Tab 1.
- Choose output directory (default `.\DeviceQRCodes`) and QR size (100–600px).
- Click "Generate QR Codes". The script calls api.qrserver.com with the device dashboard URL `{baseUrl}/#/deviceDashboard/{id}/overview` and saves **Device_{id}.png**.
- The generated output directory is remembered and pre-fills the Upload tab (Tab 3).

### Tab 3 — Upload QR Codes

- Set "Image Dir" (or use the path pre-filled from Tab 2). Click "Scan Directory" to find `Device_*.png` files and parse device ID from each filename.
- Set description (default "Device dashboard QR code") and optionally check "Replace existing".
- Click "Upload All" to attach each PNG to the matching device as a related item (multipart upload). If "Replace existing" is checked, existing related items with the same description are removed before uploading.

### Tab 4 — Scan & Assign

- Same flow as the [ITAM Scanner](ITAM%20Scanner/README.md): focus the scanner input; scan or paste a **user** dashboard URL, then one or more **device** dashboard URLs (press Enter after each).
- User is resolved from NinjaOne `users`/`contacts`; devices from `device/{id}`. Click "Assign All to User" to set the scanned user as owner of all listed devices via `POST device/{id}/owner/{ownerUid}`. "Reset" clears the current user and device list.
- QR content must match `userDashboard/(\d+)` or `deviceDashboard/(\d+)`.

## CSV format (Import Equipment)

- **Encoding:** UTF-8 (script uses `Import-Csv -Encoding UTF8`).
- **Columns (case-insensitive):**
  - **Required:** `RoleName` (must match an unmanaged-device role in NinjaOne).
  - **Org/Location (one of):** `OrganizationId` + `LocationId` **or** `OrganizationName` + `LocationName`.
  - **Optional:** `Name`, `Make`, `Model`, `SerialNumber`, `WarrantyStartDate`, `WarrantyEndDate`, `PurchaseDate`, `PurchaseAmount`.
- **Dates:** Parsed with standard .NET date parsing (e.g. YYYY-MM-DD).
- **PurchaseAmount:** Sent as integer (script uses `[int][double]$amount` for numeric values).

Example:

```csv
Name,RoleName,OrganizationName,LocationName,SerialNumber,Make,Model,PurchaseDate,PurchaseAmount
Conference Room TV,Display,Main Org,Building A,SN123,Acme,Pro 55,2024-01-15,1200
```

## Usage

**Run:**

```powershell
.\Invoke-NinjaITAMManager.ps1
```

With parameters:

```powershell
.\Invoke-NinjaITAMManager.ps1 -NinjaOneInstance app.ninjarmm.com -ClientId "your-client-id"
```

**Typical flow:** Sign in → Import Equipment (CSV or manual) → Generate QR Codes ("From Import", choose dir/size, Generate) → Upload QR Codes (Scan Directory, Upload All) and/or Scan & Assign (scan user + devices, Assign).

### Branded / partner portals

For branded or partner portals (for example, `rcs-sales.rmmservice.ca`), set `NinjaOneInstance` (or `NINJA_BASE_URL`) to the branded host so that the entire OAuth Authorization Code + PKCE flow (authorize → consent → redirect back to `http://localhost`) stays on the same host. The script always uses the instance you provide for both `/ws/oauth/authorize` and `/ws/oauth/token`. If the browser is redirected to a regional host such as `https://ca.ninjarmm.com/ws/oauth/consent` and shows `Missing or empty sessionKey.`, that redirect and error are coming from the NinjaOne web application, not from this script.

## API and behavior

- **Auth:** OAuth 2.0 Authorization Code + PKCE; token refresh when needed.
- **APIs used:** organizations, locations, noderole/list (UNMANAGED_DEVICE), users, contacts, device/{id}, itam/unmanaged-device (POST), device/{id}/custom-fields (PATCH), related-items (list/delete/attachment upload), device/{id}/owner/{uid} (POST).
- **QR generation:** External api.qrserver.com; no API key; URL encoded in query.
- **Standalone:** The script is self-contained (no dot-sourcing or shared helpers), per project conventions.

## Notes

- **User QR codes:** For Scan & Assign (Tab 4), user QR codes can be generated with [New-NinjaUserQRCode.ps1](QR%20Codes/New-NinjaUserQRCode.ps1). Device QR codes are produced by Tab 2.
- **ITAM Scanner:** The [ITAM Scanner](ITAM%20Scanner/README.md) is a separate WPF app that does only the scan-and-assign workflow; the Manager includes that workflow in Tab 4.
- **STA:** The script ensures WPF runs on an STA thread (spawns one if needed).

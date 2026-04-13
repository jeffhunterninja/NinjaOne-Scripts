# NinjaOne ITAM Desktop Companion

This workflow is an example of how NinjaOne ITAM might be used for an IT team. As equipment is ordered, delivered, and reconciled, it also needs to be added into your ITAM system. With this workflow, assets can be added in by CSV import. QR codes linking to the assets in NinjaOne can be created, then the QR codes are uploaded as related items to the device. I imagine that people might want to print out the QR code on the asset label, which I haven't worked on yet - I did see this project (https://github.com/t2-schreiner/NinjaOneLabelPrinter) from @t2-schreiner, which I'll be investigating further once I have a printer.

**Script:** [Invoke-NinjaITAMManager.ps1](Invoke-NinjaITAMManager.ps1)

## Prerequisites

- **PowerShell:** 5.1 or later (script uses `#Requires -Version 5.1`).
- **Platform:** Windows (WPF: `PresentationFramework`, `PresentationCore`, `WindowsBase`, `System.Windows.Forms`). Not compatible with PowerShell Core on non-Windows.
- **NinjaOne OAuth app:** Configure in NinjaOne as **Native** application platform. Scopes: `monitoring`, `management`. Grant Types: Authorization Code with Refresh Token. Redirect URI must be `http://localhost:8888/` - you'll need to create the client app ID first, then delete the default redirect URI and add the correct one. (I'll note that this will likely change in the future to be more standardized with how PKCE authentication is supposed to operate.)

## Parameters and environment variables

| Parameter | Description | Default |
|-----------|-------------|---------|
| `NinjaOneInstance` | Instance hostname or base URL | `$env:NINJA_BASE_URL` or `ca.ninjarmm.com` |
| `ClientId` | OAuth application Client ID | `$env:NinjaOneClientId` |
| `AllowInsecureHttp` | Allows `http://` instance URLs for local testing only | Off (HTTPS required) |

Connection settings can also be entered in the UI (Instance, Client ID) and persist for the session.

## Tabs (workflow)

**Connection:** Expand "Connection Settings", enter Instance and Client ID, then click "Sign In to NinjaOne". A browser opens for OAuth; after sign-in, return to the app. You'll have the option to store your API instance, client, and token locally using the secure method described below for ease of use.

### Tab 1 — Import Equipment

You can create **unmanaged devices** (e.g. docks, monitors, peripherals) or **staged devices** (e.g. laptops created before the agent is installed). Staged devices use the `staged-device` API so the record exists in NinjaOne before the agent is deployed.

- **CSV:** Browse to a CSV file; a preview loads. Click "Import" to create devices in NinjaOne.
  - **Device type (optional):** Column `DeviceType` with values `Unmanaged` or `Staged`. If missing or invalid, the row is treated as Unmanaged.
  - **Required:** `RoleName` — must match an **unmanaged-device** role (e.g. Dock, Keyboard, Monitor) when `DeviceType` is Unmanaged, or a **managed/staged** role (e.g. Windows Workstation) when `DeviceType` is Staged.
  - **Org/Location:** Provide either `OrganizationId` + `LocationId` **or** `OrganizationName` + `LocationName`.
  - **Optional:** Name, Make, Model, SerialNumber, WarrantyStartDate, WarrantyEndDate, PurchaseDate, PurchaseAmount, AssetStatus, ExpectedLifetime, EndOfLifeDate. If Name is omitted, it is derived from Make/Model or "Unmanaged {RoleName} {row number}" / "Staged {RoleName} {row number}".
  - **Unmanaged rows:** Created via `itam/unmanaged-device`; then custom fields are patched when provided.
  - **Staged rows:** Created via `staged-device` with the same org/location/role and optional itam fields in the request body.
- **Manual:** Choose **Device type** (Unmanaged device or Staged device). The Role list shows unmanaged roles or staged/managed roles depending on the selection. Select Organization (locations load), Role, and Location; fill Name, Serial, Make, Model, Purchase Date/Amount, Warranty Start/End, Asset Status; click "Add Device".
- The list of imported devices is used by Tab 2 ("From Import").

### Tab 2 — Generate QR Codes

- Add device IDs manually (text box + "Add") or click "From Import" to load IDs from Tab 1.
- Choose output directory (default `.\DeviceQRCodes`) and QR size (100–600px, preset sizes).
- Click "Generate QR Codes". The script calls api.qrserver.com with the device dashboard URL `{baseUrl}/#/deviceDashboard/{id}/overview` and saves **Device_{id}.png**.
- The generated output directory is remembered and pre-fills the Upload tab (Tab 3).

### Tab 3 — Upload QR Codes

- Set "Image Dir" (or use the path pre-filled from Tab 2). Click "Scan Directory" to find `Device_*.png` files and parse device ID from each filename.
- Set description (default "Device dashboard QR code") and optionally check "Replace existing".
- Click "Upload All" to attach each PNG to the matching device as a related item (multipart upload). If "Replace existing" is checked, existing related items whose name matches the filename (without extension), e.g. `Device_123`, are removed before uploading. The description is the label shown for the upload.

### Tab 5 — Scan & Assign

- Choose **Assign to user** or **Check in devices**. Focus the scanner input; press Enter after each scan or paste.
- **Assign:** Select an end user (type-ahead, or scan a **user** dashboard URL), then scan one or more **device** URLs. Click **Assign All to User** to set the owner (`POST device/{id}/owner/{ownerUid}`) and set `itamAssetStatus` to **In use** via `PATCH device/{id}/custom-fields`.
- **Check in:** Scan device URLs only (no user). Click **Check in all** to remove the assigned user (`DELETE device/{id}/owner`), delete ITAM asset relationships for that device (`GET` / `DELETE` `itam/asset-relationship`), and set `itamAssetStatus` to **inventory**. This does not remove QR-related **related items** (Tab 3 uploads).
- "Reset" clears the device list (and user in assign mode). QR content must match `userDashboard/(\d+)` or `deviceDashboard/(\d+)` (or asset search URLs with `assetId`).

## CSV format (Import Equipment)

- **Encoding:** UTF-8 (script uses `Import-Csv -Encoding UTF8`).
- **Columns (case-insensitive):**
  - **Optional:** `DeviceType` — `Unmanaged` or `Staged`. Defaults to Unmanaged when missing or invalid. Use Staged for devices (e.g. laptops) that will get the agent later.
  - **Required:** `RoleName` — must match an unmanaged-device role (Unmanaged rows) or a managed/staged role (Staged rows) in NinjaOne.
  - **Org/Location (one of):** `OrganizationId` + `LocationId` **or** `OrganizationName` + `LocationName`.
  - **Optional:** `Name`, `Make`, `Model`, `SerialNumber`, `WarrantyStartDate`, `WarrantyEndDate`, `PurchaseDate`, `PurchaseAmount`, `AssetStatus`, `ExpectedLifetime`, `EndOfLifeDate`.
- **Dates:** Parsed with standard .NET date parsing (e.g. YYYY-MM-DD).
- **PurchaseAmount:** Sent as integer (script uses `[int][double]$amount` for numeric values).

**Typical flow:** Sign in → Import Equipment (CSV or manual) → Generate QR Codes ("From Import", choose dir/size, Generate) → Upload QR Codes (Scan Directory, Upload All) and/or Scan & Assign (scan user + devices, Assign).

### Encryption

This concept is taken from @t2-schreiner's Label Printer script, and I've copied his documentation below: 

API credentials are **never stored in plaintext** on disk. They are protected by the following methods:

| Component | Method |
|---|---|
| **Encryption** | AES-256-CBC with random IV per save operation |
| **Key Derivation** | PBKDF2-HMAC-SHA256, 100,000 iterations, 32-byte salt |
| **Password Verifier** | PBKDF2-HMAC-SHA256 hash (separate from encryption key) |

### Master Password

- The master password must be **at least 8 characters** long.
- It is **not** stored – only a cryptographic verifier (PBKDF2 hash).
- To change the master password: **Change Master Password** in the right column (Connection Settings).

### Reset on Forgotten Master Password

If the master password is forgotten, the credentials **cannot** be recovered (this is by design). Instead:

1. Click **Clear Saved Session** (right column, Connection Settings).
2. All encrypted credentials are permanently deleted.
3. Enter a new master password and your API credentials the next time you connect.


### Branded / partner portals

For branded or partner portals (for example, `rcs-sales.rmmservice.ca`), set `NinjaOneInstance` (or `NINJA_BASE_URL`) to the branded host so that the entire OAuth Authorization Code + PKCE flow (authorize → consent → redirect back to `http://localhost:8888/`) stays on the same host. The script always uses the instance you provide for both `/ws/oauth/authorize` and `/ws/oauth/token`.

## API and behavior

- **Auth:** OAuth 2.0 Authorization Code + PKCE; token refresh when needed.
- **QR generation:** External api.qrserver.com; no API key; URL encoded in query.

## Notes

- **User QR codes:** For Scan & Assign (Tab 4), user QR codes can be generated with [New-NinjaUserQRCode.ps1](QR%20Codes/New-NinjaUserQRCode.ps1). Device QR codes are produced by Tab 2, or you can create QR codes for assets that already exist in NinjaOne by using [New-NinjaDeviceQRCode.ps1](QR%20Codes/New-NinjaDeviceQRCode.ps1).
- **STA:** The script ensures WPF runs on an STA thread (spawns one if needed).

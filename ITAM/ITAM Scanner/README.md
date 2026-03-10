# NinjaOne ITAM Scanner

Standalone PowerShell WPF application that uses a USB barcode scanner (or paste + Enter) to scan NinjaOne user and device QR codes, then assigns the scanned devices to the scanned user in NinjaOne via the Set Device Owner API.

## Prerequisites

- **PowerShell:** 5.1 or later (script uses `#Requires -Version 5.1`).
- **Platform:** Windows (WPF: `PresentationFramework`, `PresentationCore`, `WindowsBase`). Not compatible with PowerShell Core on non-Windows.
- **NinjaOne OAuth app:** Configure in NinjaOne as **Native** (Authorization Code). Redirect URI must be `http://localhost` (port wildcard for native apps per RFC 8252). Scopes: `monitoring`, `management`.
- **QR codes:** User and device QR codes must encode NinjaOne dashboard URLs. Generate them with the sibling scripts in the parent ITAM folder (see [Generating QR codes](#generating-qr-codes) below).

## Parameters and environment variables

| Parameter | Description | Default |
|-----------|-------------|---------|
| `NinjaOneInstance` | Instance hostname or base URL | `$env:NINJA_BASE_URL` or `ca.ninjarmm.com` |
| `ClientId` | OAuth application Client ID | `$env:NinjaOneClientId` |

Connection settings can also be entered in the UI (Instance, Client ID) and persist for the session.

## Workflow

1. **Sign in** — Enter instance and Client ID (or rely on defaults), click "Sign In to NinjaOne". Browser opens for OAuth; after sign-in, return to the app.
2. **Step 1 — User** — Scan a **user** QR code (or paste a user dashboard URL and press Enter). The app resolves the user from NinjaOne users/contacts and displays name and UID.
3. **Step 2 — Devices** — Scan one or more **device** QR codes (or paste device dashboard URLs). Each device is looked up and added to the list. Optionally remove entries with "Remove Selected".
4. **Assign** — Click "Assign All to User" to set the scanned user as owner of all listed devices via the NinjaOne API. Status and any errors are shown.
5. **Reset** — "Reset" clears the current user and device list for the next round.

## Generating QR codes

- **User QR codes:** Use [New-NinjaUserQRCode.ps1](../New-NinjaUserQRCode.ps1). Encoded URL pattern: `/#/userDashboard/{id}/overview`.
- **Device QR codes:** Use [New-NinjaDeviceQRCode.ps1](../New-NinjaDeviceQRCode.ps1). Encoded URL pattern: `/#/deviceDashboard/{id}/overview`.

The scanner only accepts these URL patterns; other QR content is rejected with a short message.

## Usage

**Run:**

```powershell
.\Invoke-NinjaITAMScanner.ps1
```

With parameters:

```powershell
.\Invoke-NinjaITAMScanner.ps1 -NinjaOneInstance app.ninjarmm.com -ClientId "your-client-id"
```

**Input:** Focus the "Scanner Input" text box. Use a USB barcode scanner (e.g. Eyoyo EYH2) that types the URL and sends Enter, or paste a URL and press Enter.

## API behavior

- **Auth:** OAuth 2.0 Authorization Code with PKCE; token refresh when needed.
- **Lookups:** Users via `users` and `contacts` endpoints (by ID); devices via `device/{id}`.
- **Assignment:** `POST device/{id}/owner/{ownerUid}` (owner UID from the resolved user/contact).

## Notes

- **Alternative:** The [ITAM App](../ITAM%20App/) (iOS) provides a similar flow (OAuth, scan user + devices, assign); this scanner is the Windows/WPF option.
- **Combined workflow:** For import, QR generation, QR upload, and scan-and-assign in one app, see the [NinjaOne ITAM Manager](../README.md).
- **Standalone:** The script is self-contained (no dot-sourcing or shared helpers), per project conventions.

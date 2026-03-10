# NinjaOne ITAM iPhone App

Native iOS app that signs in with NinjaOne (OAuth PKCE), scans one user QR and multiple device QR codes, then assigns those devices to the user in NinjaOne via the Set Device Owner API.

## Requirements

- Xcode 15+ (iOS 16+)
- NinjaOne instance and a **Native** OAuth application with Authorization Code + PKCE
- Redirect URI: `ninjaone-itam://callback`

## Build and run

1. Open `ITAM App/NinjaOneITAM.xcodeproj` in Xcode.
2. Select your team under Signing & Capabilities.
3. Run on a device or simulator (camera required for QR; use a real device for full flow).

## QR code format

The app expects JSON payloads in QR codes.

### User QR (scan once)

Identifies the NinjaOne user who will own the devices.

**By email (recommended):**

```json
{"type":"user","email":"jane@company.com"}
```

The app looks up this user in NinjaOne `/users` and `/contacts` and uses the matched UID as the device owner.

**By UID:**

```json
{"type":"user","uid":"12345"}
```

Use if you pre-print QRs with known NinjaOne user/contact UIDs.

### Device QR (scan multiple times)

Each QR represents one device to assign to the user.

```json
{"type":"device","id":12345}
```

or with optional display name:

```json
{"type":"device","id":12345,"name":"DESKTOP-ABC"}
```

- `id` = NinjaOne device ID (integer)
- `name` = optional; used only for display in the app

## Generating QR codes

Use the PowerShell script in this folder to output JSON strings for users and devices; then encode each string with any QR generator.

See [Generate-ITAMQRPayloads.ps1](Generate-ITAMQRPayloads.ps1) for usage.

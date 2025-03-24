# ✋ NinjaOne Custom Field Notification Script

This PowerShell script sets a custom field in NinjaOne with an emoji (`✋`) to indicate a status or notify users. After a brief delay, the field is automatically cleared. This is useful for visual indicators or time-sensitive flags in NinjaOne. Adminstrators can also create a condition that monitors for the presence of the hand emoji to streamline the ticket submission/"call for help" process.

---

## 🧰 Requirements

- **Script form variable** `customFieldName` must be set, containing the name of the custom field to target. Alternately the name of the custom field can be hardcoded into the script.

---

## ⚙️ What It Does

1. Verifies that the `Ninja-Property-Set` cmdlet exists.
2. Sets the specified custom field to `✋`.
3. Waits for **3 minutes**.
4. Clears the custom field.

---

## 🛠️ Usage

### 1. Set the Required Environment Variable

Before running the script, define the custom field name as an environment variable:

```powershell
$env:customFieldName = "YourCustomFieldName"
```
---

## 🧼 Cleanup

The script automatically clears the custom field after the delay, ensuring temporary notifications don’t persist.

---

## 📌 Notes

- You can change the emoji or message by modifying the `$Message` variable.
- You can adjust the delay by changing the `Start-Sleep -Seconds` value (currently set to 180 seconds = 3 minutes).

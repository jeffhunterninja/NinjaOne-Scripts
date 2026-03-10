When creating unmanaged devices via CLI, you can detect items such as monitors, USB printers, and other peripherals, and define the relationship these items have with managed devices in NinjaOne. Additionally, you can assign the unmanaged device to a user via CLI, and NinjaOne will populate the data in the console where applicable.

When you install the NinjaOne RMM agent on the device, NinjaOne sends the system and BIOS serial numbers to the initial payload to optimize data mapping.

---

# Index

Select a category to learn more:

- [New Command Line Options](#new-command-line-options)
- [Using PowerShell Commands](#using-powershell-commands)
- Additional Resources

---

# New Command Line Options

We created new command lines and added options to existing command lines to facilitate this process.

## Command Line Options

| Command Line Name | Description |
|-------------------|------------|
| `unmanaged-new` | Create an unmanaged device and pass in a JSON object to define it. When you create an unmanaged device via the CLI, the device will be automatically assigned to the same organization and location as its related managed device. |
| `unmanaged-get` | Obtain the values of unmanaged devices related to the managed device, including ID numbers for each asset (excluding secure fields). The unmanaged device must belong to the same organization and location as its related managed device. |
| `unmanaged-set` | Update values for existing unmanaged devices using the device ID. This command line only works if the unmanaged device is related to an existing managed device and belongs to the same organization and location. |
| `user-set` | Set the currently assigned user of the device. You can use an email address as the parameter, and NinjaOne will search end users, technicians, and contacts. If an account or contact is limited to an organization, NinjaOne will only match devices that belong to the same organization as the user. Technicians must have permission to access the device. |

---

## Fields for `unmanaged-new`

The following table lists the fields you can use with the `unmanaged-new` command line to define unmanaged devices.

| Field Name | Required | Type | Notes |
|------------|----------|------|-------|
| Role | Yes | String |  |
| Name | Yes | String |  |
| Relation | Yes | String | The directional asset relation name between the creating and created device. |
| Warranty Start Date | No | String | Accepts dates in `YYYY-MM-DD` format. |
| Warranty End Date | No | String | Accepts dates in `YYYY-MM-DD` format. |
| Assign to Device User | No | Boolean | This option is `false` by default. |
| Asset ID | No | String |  |
| Asset Status | No | String |  |
| Purchase Date | No | String | Accepts dates in `YYYY-MM-DD` format. |
| Purchase Amount | No | Float |  |
| Expected Lifetime | No | String | Accepts `1`, `2`, `3`, `4`, or `5 years`. |
| End of Life Date | No | String | Accepts dates in `YYYY-MM-DD` format. |
| Asset Serial Number | No | String |  |
| Custom Field Values | No | Multiple | As defined when created. |

---

## CLI Example

```bash
.\ninjarmm-clie.exe unmanaged-new '{ 
  "assignToDeviceUser": true,
  "assetId": "asset-001",
  "assetSerialNumber": "asset-serial-001",
  "assetStatus": "in use",
  "customFieldValues": {
    "exampleCustomFieldName": "Custom Field Value",
    "typeipaddress": "255.255.255.255"
  },
  "endOfLifeDate": "2025-12-31",
  "expectedLifetime": "5 years",
  "name": "Test Device with CF",
  "purchaseAmount": 1000,
  "purchaseDate": "2020-01-01",
  "relation": "related to",
  "role": "Another Role",
  "warrantyEndDate": "2023-12-31",
  "warrantyStartDate": "2021-01-01"
}'
```

---

# Using PowerShell Commands

NinjaOne supports the following PowerShell commands:

- `New-NinjaUnmanagedDevice`
- `Get-NinjaUnmanagedDevice`
- `Set-NinjaUnmanagedDevice`
- `Set-NinjaUser`

## PowerShell Example

```powershell
$DiscoveredMonitor = @{
   role               = "Monitor"
   name               = "AEG-12345 on Lukes Desktop"
   relation           = "Connected To"
   assignToDeviceUser = $true
   assetId            = "123456"
   assetSerialNumber  = "123456"
   customFieldHashMap = @{
       monitorResolution   = "1920 x 1080"
       monitorManufacturer = "AEG"
       monitorModel        = "12345"
   }
}

$Null = New-NinjaUnmanagedDevice $DiscoveredMonitor

$Devices = Get-NinjaUnmanagedDevice -Role 'Monitor' -Relation 'Connected To'

$UpdateDevice = $Devices[0]
$UpdateDevice.monitorResolution = '5120 x 1440'

$Null = Set-NinjaUnmanagedDevice $UpdateDevice
```
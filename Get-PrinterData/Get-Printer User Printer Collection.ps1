#Requires -Version 5.1
<#
.SYNOPSIS
  Collects printer data in user context and writes to a temp file for the SYSTEM script to consume.

.DESCRIPTION
  Runs as the current logged-on user to capture per-user (HKCU) and machine-wide (HKLM) printers
  via Get-Printer. Writes results to C:\Windows\Temp for the companion SYSTEM script to read.
  Designed to run at user login, immediately before the Custom Field Update script.

.EXIT CODES
  0 = Success
  2 = Error (Get-Printer failure, write failure)
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$filePath = Join-Path $env:SystemRoot "Temp\printer_info.json"

try {
    $printers = Get-Printer |
        Select-Object Name, DriverName

    $json = $printers | ConvertTo-Json -Depth 2 -Compress
    [System.IO.File]::WriteAllText($filePath, $json, [System.Text.UTF8Encoding]::new($false))
}
catch {
    Write-Error "Failed to collect or write printer data: $_"
    exit 2
}

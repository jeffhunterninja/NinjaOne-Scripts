#Requires -Version 5.1
<#
.SYNOPSIS
  Reads printer data from temp file and writes to NinjaOne custom fields.

.DESCRIPTION
  Runs as SYSTEM to update NinjaOne custom fields (requires elevated context).
  Reads JSON written by the User Printer Collection script, normalizes the data,
  and populates the printers and printerDrivers custom fields.

.EXIT CODES
  0 = Success
  1 = No data (file missing or empty)
  2 = Error (parse failure, Ninja-Property-Set failure)
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$filePath = Join-Path $env:SystemRoot "Temp\printer_info.json"

if (-not (Test-Path -Path $filePath -PathType Leaf)) {
    Write-Host "Printer data file not found at $filePath. User collection script may not have run."
    exit 1
}

try {
    $raw = Get-Content -Raw -Path $filePath
    $printers = $raw | ConvertFrom-Json
}
catch {
    Write-Error "Could not parse printer data from $filePath : $_"
    exit 2
}

# Normalize to an array (ConvertFrom-Json returns a PSCustomObject for single item)
if ($printers -isnot [System.Collections.IEnumerable] -or $printers -is [string]) {
    $printers = @($printers)
}

# Extract and sanitize
$printerNames = $printers | ForEach-Object { $_.Name } | Where-Object { $_ } | Sort-Object -Unique
$driverNames  = $printers | ForEach-Object { $_.DriverName } | Where-Object { $_ } | Sort-Object -Unique

try {
    Ninja-Property-Set "printers"       ($printerNames -join "`r`n")
    Ninja-Property-Set "printerDrivers" ($driverNames  -join "`r`n")
}
catch {
    Write-Error "Could not write to NinjaOne custom fields: $_"
    exit 2
}

# Clean up temp file
try {
    Remove-Item -Path $filePath -Force -ErrorAction SilentlyContinue
}
catch {
    # Non-fatal; log and continue
    Write-Verbose "Could not remove temp file: $_"
}

exit 0

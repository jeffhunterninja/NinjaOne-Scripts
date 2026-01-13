# Run as System

$filePath = "C:\temp\printer_info.json"

if (-not (Test-Path -Path $filePath -PathType Leaf)) {
    Write-Host "Could not find json file at $filePath, exiting script"
    throw
}

try {
    $printers = Get-Content -Raw -Path $filePath | ConvertFrom-Json
}
catch {
    Write-Host "Could not retrieve/parse data from json file at $filePath"
    throw
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
    Ninja-Property-Set "printerdrivers" ($driverNames  -join "`r`n")
}
catch {
    Write-Host "Could not write into custom field(s)"
    throw
}

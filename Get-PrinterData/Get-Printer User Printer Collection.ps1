#Run as current logged on user

$filePath   = "C:\temp\printer_info.json"
$folderPath = "C:\temp"

# Check if the folder path exists
if (Test-Path -Path $folderPath -PathType Container) {

    # Collect printer information
    $printers = Get-Printer |
        Select-Object Name, DriverName

    # Write JSON to disk (UTF-8 for compatibility)
    $printers |
        ConvertTo-Json -Depth 2 |
        Out-File -FilePath $filePath -Encoding UTF8 -Force

    # Wait 60 seconds to allow the SYSTEM script to consume the file
    Start-Sleep -Seconds 60
}
else {
    Write-Error "Folder path does not exist. Script cannot continue."
    throw
}

$JsonFilePath = "c:\ProgramData\NinjaRMMAgent\jsonoutput\jsonoutput-agent.txt"

# Check that the file exists
if (-not (Test-Path $JsonFilePath)) {
    Write-Error "File not found: $JsonFilePath"
    exit 1
}

# Read and convert the JSON content
$jsonContent = Get-Content $JsonFilePath -Raw | ConvertFrom-Json

# Extract the OS dataset (assuming the dataset has dataspecName "os")
$osDataset = $jsonContent.node.datasets | Where-Object { $_.dataspecName -eq "os" }

if (-not $osDataset) {
    Write-Output "OS dataset not found in the JSON file."
    exit 1
}

# Assuming there is one OS datapoint, extract its data object
$osData = $osDataset.datapoints[0].data

Write-Output "Operating System Information:"
Write-Output "--------------------------------"
Write-Output ("Name           : {0}" -f $osData.name)
Write-Output ("Short Name     : {0}" -f $osData.shortName)
Write-Output ("Build Number   : {0}" -f $osData.buildNumber)
Write-Output ("Install Date   : {0}" -f $osData.installDate)
Write-Output ("Release ID     : {0}" -f $osData.releaseId)
Write-Output ("Architecture   : {0}" -f $osData.osArchitecture)
Write-Output ""

# Determine if the OS name indicates Windows 10 or Windows 11
if ($osData.name -match "10") {
    Write-Output "This device appears to be running Windows 10."
} elseif ($osData.name -match "11") {
    Write-Output "This device appears to be running Windows 11."
} else {
    Write-Output "The OS version does not clearly indicate Windows 10 or 11."
}

# For repurposing (e.g. server analysis), check if the OS name includes "Server"
if ($osData.name -match "Server") {
    Write-Output "This appears to be a server operating system."
} else {
    Write-Output "This does not appear to be a server operating system."
}

# Path to the JSON file
$jsonFilePath = "C:\ProgramData\NinjaRMMAgent\jsonoutput\jsonoutput-agent.txt"

# Define authorized applications in separate objects (example data)
$orgAuthorizedApps = Ninja-Property-Get softwareList | ConvertFrom-Json | Select-Object -ExpandProperty 'text' -EA 0
$deviceAuthorizedApps = Ninja-Property-Get deviceSoftwareList | ConvertFrom-Json | Select-Object -ExpandProperty 'text' -EA 0
$authorizedApps = $orgAuthorizedApps + $deviceAuthorizedApps

$authorizedApps = $authorizedApps -split ','
$authorizedApps = $authorizedApps | ForEach-Object { $_.Trim() }

# Define a parameter for the match mode
# Options: "Exact", "CaseInsensitive", "Partial"
$matchMode = $env:matchingCriteria # Change to "Exact" or "CaseInsensitive" as needed

# Function to perform matching based on the mode
function Compare-Application {
    param (
        [string]$InstalledApp,
        [string[]]$AuthorizedApps,
        [string]$Mode
    )
    switch ($Mode) {
        "Exact" {
            return $AuthorizedApps -contains $InstalledApp
        }
        "CaseInsensitive" {
            return $AuthorizedApps | ForEach-Object { $_ -ieq $InstalledApp }
        }
        "Partial" {
            foreach ($authApp in $AuthorizedApps) {
                if ($InstalledApp -match [regex]::Escape($authApp)) {
                    return $true
                }
            }
            return $false
        }
        default {
            throw "Invalid match mode specified: $Mode. Valid options are 'Exact', 'CaseInsensitive', or 'Partial'."
        }
    }
}

# Check if the file exists
if (Test-Path $jsonFilePath) {
    # Read the JSON file
    $jsonContent = Get-Content -Path $jsonFilePath -Raw
    
    # Parse the JSON content
    $jsonObject = $jsonContent | ConvertFrom-Json
    
    # Extract the software inventory data
    $softwareInventory = $jsonObject.node.datasets | Where-Object { $_.dataspecName -eq "softwareInventory" }
    
    # Create a list of installed applications
    $installedApps = @()
    foreach ($datapoint in $softwareInventory.datapoints) {
        foreach ($software in $datapoint.data) {
            $installedApps += $software.name
        }
    }
    
    # Compare installed applications with authorized applications using the selected mode
    $discrepancies = @()
    foreach ($app in $installedApps) {
        if (-not (Compare-Application -InstalledApp $app -AuthorizedApps $authorizedApps -Mode $matchMode)) {
            $discrepancies += [PSCustomObject]@{
                DiscrepantApplication = $app
            }
        }
    }
    
        # Output results
        if ($discrepancies.Count -gt 0) {
            # Extract discrepant application names as a comma-separated string
            $discrepantAppsString = ($discrepancies.DiscrepantApplication) -join ', '
            Write-Output "WARNING - Discrepancies found: $discrepantAppsString"
            Ninja-Property-Set unauthorizedApplications $discrepantAppsString
        } else {
            Write-Output "No discrepancies found. All installed applications are authorized."
            Ninja-Property-Set unauthorizedApplicatiions $null
        }

} else {
    Write-Error "The file '$jsonFilePath' does not exist."
}

$CustomFieldName = $env:customFieldName

$Message = '✋'

# Function to check if a command exists
function Test-CommandExists {
    param (
        [Parameter(Mandatory = $true)]
        [string]$CommandName
    )
    return (Get-Command $CommandName -ErrorAction SilentlyContinue) -ne $null
}

# Check if Ninja-Property-Set cmdlet exists
if (-not (Test-CommandExists -CommandName 'Ninja-Property-Set')) {
    Write-Error "Ninja-Property-Set cmdlet not found. Please ensure it is installed and accessible."
    exit 1
}

# Function to set custom field
function Set-CustomField {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FieldName,

        [Parameter(Mandatory = $true)]
        [string]$Value
    )
    try {
        Ninja-Property-Set -Name $FieldName -Value $Value -ErrorAction Stop
        Write-Host "Custom field '$FieldName' set to '$Value'."
    } catch {
        Write-Error "Failed to set custom field '$FieldName'. Error: $_"
        exit 1
    }
}

# Function to clear custom field
function Clear-CustomField {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FieldName
    )
    try {
        Ninja-Property-Set -Name $FieldName -Value "" -ErrorAction Stop
        Write-Host "Custom field '$FieldName' cleared."
    } catch {
        Write-Error "Failed to clear custom field '$FieldName'. Error: $_"
        exit 1
    }
}

# Set the custom field with the message
Set-CustomField -FieldName $CustomFieldName -Value $Message

# Wait for 5 minutes
Write-Host "Waiting for 5 minutes..."
Start-Sleep -Seconds 180

# Clear the custom field
Clear-CustomField -FieldName $CustomFieldName

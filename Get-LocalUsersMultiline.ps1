# Modified version of "Local User Report" in the NinjaOne Template Library.
## Outputs to a multi-line custom field instead of a WYSIWYG field - useful for viewing users at scale from the Devices menu


[CmdletBinding()]
param (
    # Will return disabled users as well as enabled users
    [Parameter()]
    [Switch]$AllUsers = [System.Convert]::ToBoolean($env:includeDisabledUsers)
)

begin {
    Write-Output "Starting List Users"
    # Array to store user objects
    $UserList = @()
}
process {
    # If we are running powershell under a 32bit OS or running a 32bit version of powershell
    # or we don't have access to Get-LocalUser
    if (
        [Environment]::Is64BitProcess -ne [Environment]::Is64BitOperatingSystem -or
            (Get-Command -Name "Get-LocalUser" -ErrorAction SilentlyContinue).Count -eq 0
    ) {
        # Get users from net.exe user
        $Data = $(net.exe user) | Select-Object -Skip 4
        # Check if the command ran the way we wanted and the exit code is 0
        if ($($Data | Select-Object -Last 2 | Select-Object -First 1) -like "*The command completed successfully.*" -and $LASTEXITCODE -eq 0) {
            # Process the output and get only the users
            $Users = $Data[0..($Data.Count - 3)] -split '\s+' | Where-Object { -not $([String]::IsNullOrEmpty($_) -or [String]::IsNullOrWhiteSpace($_)) }
            # Loop through each user
            foreach ($UserName in $Users) {
                # Get the Account active property look for a Yes
                $Enabled = $(net.exe user $UserName) | Where-Object {
                    $_ -like "Account active*" -and
                    $($_ -split '\s+' | Select-Object -Last 1) -like "Yes"
                }
                # Create a custom object for the user and add it to the user list
                $UserList += [PSCustomObject]@{
                    Name    = $UserName
                    Enabled = if ($Enabled -like "*Yes*") { $true } else { $false }
                }
            }
        }
        else {
            exit 1
        }
    }
    else {
        try {
            if ($AllUsers) {
                $UserList += Get-LocalUser
            }
            else {
                $UserList += Get-LocalUser | Where-Object { $_.Enabled -eq $true }
            }
        }
        catch {
            Write-Error $_
            exit 1
        }
    }
}
end {
    # Output just the names of the users
    $UserList.Name
    $Delimiter = ', '
    Ninja-Property-Set -Name localUsers -Value $($UserList.Name -join $Delimiter)
}

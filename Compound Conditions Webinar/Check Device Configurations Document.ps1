param (
    [string]$Setting
)

if ($env:configuration -notlike "null") { 
    $Setting = $env:configuration 
}

# Retrieve security settings
$firewallValue = Ninja-Property-Docs-Get 'Device Configuration' 'Device Configuration' enforceFirewallEnablement
$disablePowershell = Ninja-Property-Docs-Get 'Device Configuration' 'Device Configuration' disablePowershell20
$smbV1 = Ninja-Property-Docs-Get 'Device Configuration' 'Device Configuration' disableSmbV1
$inactiveusers = Ninja-Property-Docs-Get 'Device Configuration' 'Device Configuration' inactiveUsersAlert

# Determine which variable to output
switch ($Setting.ToLower()) {
    "firewall" { Write-Output "Firewall Enforcement: $firewallValue" }
    "powershell" { Write-Output "PowerShell v2 Disabled: $disablePowershell" }
    "smb" { Write-Output "SMBv1 Disabled: $smbV1" }
    "inactiveusers" { Write-Output "Inactive Users Alert: $inactiveusers" }
    default { Write-Output "Invalid input. Please enter 'firewall', 'powershell', 'smb', or 'inactiveusers'." }
}

#Requires -Version 5.1
<#
.SYNOPSIS
    (Test/Demo only) Creates a named admin account for local or domain environments.

.DESCRIPTION
    TESTING AND DEMONSTRATION ONLY. Do not use in production. Creates or updates a privileged
    account so you can exercise unauthorized-admin detection (e.g. Test-UnauthorizedAdministrators).

    Plain-text -Password is intentional for test/demo only; use secure patterns if you ever adapt
    this for real use.

    Use -Username and -Password when calling directly; from NinjaOne or automation you can use
    $env:adminusername and $env:adminpassword. Behavior depends on machine role:
    ProductType 1 = Workstation (local admin), 2 = Domain Controller (domain admin), 3 = Server (local admin).

    When -AutoDisableAfterHours is greater than 0 (default 24), registers a one-time scheduled
    task on this machine (SYSTEM) to disable the account after that many hours—local users use
    Disable-LocalUser; on a DC, Disable-ADAccount. Requires rights to create scheduled tasks.
    Set -AutoDisableAfterHours 0 to leave the account enabled.

.NOTES
    Test/demo only. Creates real administrators. Auto-disable is lab cleanup, not a security control.

.EXIT CODES
    0 = Success (user created or password set; optional disable task registered)
    1 = Validation error (missing Username/Password), unknown ProductType, or invalid -AutoDisableAfterHours

.PARAMETER Username
    Admin account name. Defaults to $env:adminusername.

.PARAMETER Password
    Plain-text password (test/demo only). Defaults to $env:adminpassword.

.PARAMETER AutoDisableAfterHours
    Hours until a one-shot task disables this account (0 = skip). Default 24.

.EXAMPLE
    .\Create-NinjaAdminAccount.ps1 -Username "CorpAdmin" -Password "SecureP@ss"

.EXAMPLE
    .\Create-NinjaAdminAccount.ps1 -Username "CorpAdmin" -Password "SecureP@ss" -AutoDisableAfterHours 0

.EXAMPLE
    $env:adminusername = "CorpAdmin"; $env:adminpassword = "SecureP@ss"
    .\Create-NinjaAdminAccount.ps1
#>
param(
    [string] $Username = $env:adminusername,
    [string] $Password = $env:adminpassword,
    [int] $AutoDisableAfterHours = 24
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Username) -or [string]::IsNullOrWhiteSpace($Password)) {
    Write-Error "Username and Password are required. Provide -Username and -Password, or set env:adminusername and env:adminpassword."
    exit 1
}

if ($AutoDisableAfterHours -lt 0) {
    Write-Error "AutoDisableAfterHours must be 0 or greater."
    exit 1
}

function Get-ADUserBySam {
    param([string] $Sam)
    $escaped = $Sam.Replace("'", "''")
    Get-ADUser -Filter "SamAccountName -eq '$escaped'" -ErrorAction SilentlyContinue
}

function Set-NamedAccount ($Username, $Password, $type) {
    switch ($type) {
        'Local' {
            $ExistingUser = Get-LocalUser $Username -ErrorAction SilentlyContinue
            if (!$ExistingUser) {
                Write-Host "Creating new user admin $Username" -ForegroundColor Green
                New-LocalUser -Name $Username -Password $Password -PasswordNeverExpires
                # S-1-5-32-544 = well-known SID for local Administrators (language-independent).
                Add-LocalGroupMember -Member $Username -SID 'S-1-5-32-544'
            }
            else {
                Write-Host "Setting password for admin $Username" -ForegroundColor Green
                Set-LocalUser -Name $Username -Password $Password
            }
        }
        'Domain' {
            $ExistingUser = Get-ADUserBySam -Sam $Username
            if (!$ExistingUser) {
                Write-Host "Creating new domain admin for $Username" -ForegroundColor Green
                New-ADUser -Name $Username -SamAccountName $Username -AccountPassword $Password -Enabled $True -PasswordNeverExpires $true
                $ExistingUser = Get-ADUserBySam -Sam $Username
                if (!$ExistingUser) {
                    throw "New-ADUser succeeded but user '$Username' was not found for group membership."
                }
                $Groups = @("Domain Admins", "Administrators", "Schema Admins", "Enterprise Admins")
                $Groups | ForEach-Object { Add-ADGroupMember -Members $ExistingUser -Identity $_ -ErrorAction SilentlyContinue }
            }
            else {
                Write-Host "Setting password for admin $Username" -ForegroundColor Green
                Set-ADAccountPassword -Identity $ExistingUser -NewPassword $Password -Reset
            }
        }
    }
}

function Register-NinjaDemoAutoDisableSchedule {
    param(
        [string] $UserName,
        [int] $Hours,
        [ValidateSet('Local', 'Domain')]
        [string] $Scope
    )
    $safe = ($UserName -replace '[^a-zA-Z0-9]', '')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = ([Guid]::NewGuid().ToString('N')).Substring(0, 8)
    }
    if ($safe.Length -gt 40) {
        $safe = $safe.Substring(0, 40)
    }
    $taskName = "NinjaDemo_DisableAdmin_$safe"
    $inner = $UserName -replace '"', '`"'
    if ($Scope -eq 'Local') {
        $scriptText = "Disable-LocalUser -Name `"$inner`" -ErrorAction SilentlyContinue"
    }
    else {
        $scriptText = "Import-Module ActiveDirectory -ErrorAction Stop; Disable-ADAccount -Identity `"$inner`" -ErrorAction SilentlyContinue"
    }
    $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($scriptText))
    $arg = "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encoded"
    $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddHours($Hours))
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $arg
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
    Write-Host "Registered scheduled task '$taskName' to disable account in $Hours hour(s)." -ForegroundColor Cyan
}

$securePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
$DomainCheck = Get-CimInstance -ClassName Win32_OperatingSystem
$scope = $null
switch ($DomainCheck.ProductType) {
    1 {
        $scope = 'Local'
        Set-NamedAccount -Username $Username -Password $securePassword -type 'Local'
    }
    2 {
        $scope = 'Domain'
        Set-NamedAccount -Username $Username -Password $securePassword -type 'Domain'
    }
    3 {
        $scope = 'Local'
        Set-NamedAccount -Username $Username -Password $securePassword -type 'Local'
    }
    default {
        Write-Warning "Could not get Server Type. Quitting script."
        exit 1
    }
}

if ($AutoDisableAfterHours -gt 0) {
    Register-NinjaDemoAutoDisableSchedule -UserName $Username -Hours $AutoDisableAfterHours -Scope $scope
}

exit 0

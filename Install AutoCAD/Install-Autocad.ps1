#Requires -Version 5.1
<#
.SYNOPSIS
  Runs the AutoCAD/ODIS bootstrapper (Deploy.exe), waits for the deployment image, then runs the installer.

.DESCRIPTION
  Intended for use with NinjaOne: transfer the bootstrapper (Deploy.exe) first, then run this script.
  The script starts the bootstrapper, polls for Summary.txt (bootstrap complete), optionally waits
  for cleanup, then runs Installer.exe with the deployment Collection.xml. Supports InstallOnly mode
  to skip the bootstrap when the image is already staged (e.g. retries).

.EXIT CODES
  0 = Installer completed successfully.
  1 = Bootstrap/validation error (e.g. Deploy.exe missing, invalid paths).
  2 = Installer failed or could not be started.
  3 = Timeout waiting for bootstrap (Summary.txt not found).

.PARAMETER BootstrapperPath
  Path to the deployment bootstrapper (Deploy.exe). Defaults to C:\RMM\Deploy.exe.
  (NinjaOne: env var bootstrapperPath.)

.PARAMETER BootstrapperArguments
  Arguments passed to the bootstrapper. Defaults to '/q /p'.
  (NinjaOne: env var bootstrapperArguments.)

.PARAMETER SummaryPath
  Path to Summary.txt whose presence indicates bootstrap completion. Defaults to
  C:\Autodesk\Deploy\AutoCADLT2025\Summary.txt.
  (NinjaOne: env var summaryPath.)

.PARAMETER InstallerPath
  Path to Installer.exe. Defaults to C:\Autodesk\Deploy\AutoCADLT2025\image\Installer.exe.
  (NinjaOne: env var installerPath.)

.PARAMETER CollectionXmlPath
  Path to the deployment Collection.xml. Defaults to C:\Autodesk\Deploy\AutoCADLT2025\image\Collection.xml.
  (NinjaOne: env var collectionXmlPath.)

.PARAMETER InstallerVersion
  Value for --installer_version. Defaults to 2.21.0.623.
  (NinjaOne: env var installerVersion.)

.PARAMETER BootstrapTimeoutMinutes
  Maximum minutes to wait for Summary.txt. Default 60.
  (NinjaOne: env var bootstrapTimeoutMinutes.)

.PARAMETER BootstrapPollSeconds
  Seconds between checks for Summary.txt. Default 15.
  (NinjaOne: env var bootstrapPollSeconds.)

.PARAMETER PostWaitMinutes
  Minutes to wait after Summary.txt appears before running the installer. Default 10.
  (NinjaOne: env var postWaitMinutes.)

.PARAMETER InstallOnly
  Skip bootstrap; only run the installer if the image exists. Use when image is already staged.
  (NinjaOne: env var installOnly - set to 1 or true to enable.)
#>

[CmdletBinding()]
param(
    [string]$BootstrapperPath = 'C:\RMM\Deploy.exe',
    [string]$BootstrapperArguments = '/q /p',
    [string]$SummaryPath = 'C:\Deploy\AutoCAD\Summary.txt',
    [string]$InstallerPath = 'C:\Deploy\AutoCAD\image\Installer.exe',
    [string]$CollectionXmlPath = 'C:\Deploy\AutoCAD\image\Collection.xml',
    [string]$InstallerVersion = '2.21.0.623',
    [int]$BootstrapTimeoutMinutes = 60,
    [int]$BootstrapPollSeconds = 15,
    [int]$PostWaitMinutes = 10,
    [switch]$InstallOnly
)

$ErrorActionPreference = 'Stop'

# Override from environment (NinjaOne script variables)
if ($null -ne $env:bootstrapperPath -and -not [string]::IsNullOrWhiteSpace($env:bootstrapperPath)) { $BootstrapperPath = $env:bootstrapperPath.Trim() }
if ($null -ne $env:bootstrapperArguments -and -not [string]::IsNullOrWhiteSpace($env:bootstrapperArguments)) { $BootstrapperArguments = $env:bootstrapperArguments.Trim() }
if ($null -ne $env:summaryPath -and -not [string]::IsNullOrWhiteSpace($env:summaryPath)) { $SummaryPath = $env:summaryPath.Trim() }
if ($null -ne $env:installerPath -and -not [string]::IsNullOrWhiteSpace($env:installerPath)) { $InstallerPath = $env:installerPath.Trim() }
if ($null -ne $env:collectionXmlPath -and -not [string]::IsNullOrWhiteSpace($env:collectionXmlPath)) { $CollectionXmlPath = $env:collectionXmlPath.Trim() }
if ($null -ne $env:installerVersion -and -not [string]::IsNullOrWhiteSpace($env:installerVersion)) { $InstallerVersion = $env:installerVersion.Trim() }
if ($null -ne $env:bootstrapTimeoutMinutes -and -not [string]::IsNullOrWhiteSpace($env:bootstrapTimeoutMinutes)) {
    $parsed = 0
    if ([int]::TryParse($env:bootstrapTimeoutMinutes.Trim(), [ref]$parsed) -and $parsed -gt 0) { $BootstrapTimeoutMinutes = $parsed }
}
if ($null -ne $env:bootstrapPollSeconds -and -not [string]::IsNullOrWhiteSpace($env:bootstrapPollSeconds)) {
    $parsed = 0
    if ([int]::TryParse($env:bootstrapPollSeconds.Trim(), [ref]$parsed) -and $parsed -gt 0) { $BootstrapPollSeconds = $parsed }
}
if ($null -ne $env:postWaitMinutes -and -not [string]::IsNullOrWhiteSpace($env:postWaitMinutes)) {
    $parsed = 0
    if ([int]::TryParse($env:postWaitMinutes.Trim(), [ref]$parsed) -and $parsed -ge 0) { $PostWaitMinutes = $parsed }
}
if ($env:installOnly -eq '1' -or [string]::Equals($env:installOnly, 'true', [StringComparison]::OrdinalIgnoreCase)) { $InstallOnly = $true }

function Test-BootstrapComplete {
    if (-not (Test-Path -LiteralPath $SummaryPath -PathType Leaf)) { return $false }
    Write-Output "Summary.txt detected at $SummaryPath"
    return $true
}

function Invoke-Installer {
    if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) {
        Write-Error "Installer not found: $InstallerPath"
        exit 1
    }
    if (-not (Test-Path -LiteralPath $CollectionXmlPath -PathType Leaf)) {
        Write-Error "Collection.xml not found: $CollectionXmlPath"
        exit 1
    }
    $installerArgs = "-i deploy --offline_mode -q -o `"$CollectionXmlPath`" --installer_version `"$InstallerVersion`""
    Write-Verbose "Running installer: $InstallerPath with args: $installerArgs"
    Write-Output "Running AutoCAD installer..."
    try {
        $proc = Start-Process -FilePath $InstallerPath -ArgumentList $installerArgs -Wait -NoNewWindow -PassThru
        if ($proc.ExitCode -ne 0) {
            Write-Error "Installer exited with code $($proc.ExitCode)."
            exit 2
        }
        Write-Output "Installer completed successfully."
        exit 0
    } catch {
        Write-Error "Failed to start installer: $_"
        exit 2
    }
}

# InstallOnly: skip bootstrap, run installer if image exists
if ($InstallOnly) {
    Write-Verbose "InstallOnly mode: skipping bootstrap."
    Invoke-Installer
}

# Validate bootstrapper before starting
if (-not (Test-Path -LiteralPath $BootstrapperPath -PathType Leaf)) {
    Write-Error "Bootstrapper not found: $BootstrapperPath"
    exit 1
}

# Start the deployment stub (do NOT wait; it hands off to ODIS)
Write-Verbose "Starting bootstrapper: $BootstrapperPath $BootstrapperArguments"
try {
    Start-Process -FilePath $BootstrapperPath -ArgumentList $BootstrapperArguments -NoNewWindow | Out-Null
} catch {
    Write-Error "Failed to start bootstrapper: $_"
    exit 1
}

$deadline = (Get-Date).AddMinutes($BootstrapTimeoutMinutes)
Write-Verbose "Waiting for Summary.txt (timeout $BootstrapTimeoutMinutes minutes, poll every $BootstrapPollSeconds seconds)."

while ((Get-Date) -lt $deadline) {
    if (Test-BootstrapComplete) {
        if ($PostWaitMinutes -gt 0) {
            Write-Output "Waiting an additional $PostWaitMinutes minutes for cleanup to complete..."
            Start-Sleep -Seconds ($PostWaitMinutes * 60)
            Write-Output "Post-wait complete."
        }
        Invoke-Installer
    }
    Start-Sleep -Seconds $BootstrapPollSeconds
}

Write-Output "Timed out after $BootstrapTimeoutMinutes minutes waiting for Summary.txt at $SummaryPath"
exit 3

#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

$email = $null

# Stagger `dsregcmd /status` calls across machines to avoid endpoint hammering.
# Max delay is intentionally conservative; enable with `-Verbose` for troubleshooting.
$MaxDelaySeconds = 180
$DelaySeconds = Get-Random -Minimum 0 -Maximum ($MaxDelaySeconds + 1)
Write-Verbose "dsregcmd delay: sleeping $DelaySeconds second(s) before /status"
Start-Sleep -Seconds $DelaySeconds

$dsregOut = dsregcmd /status 2>$null
if ($dsregOut) {
    $labels = @('User Identity', 'UserEmail', 'MDMUserUPN', 'MDM User UPN')
    foreach ($label in $labels) {
        $match = $dsregOut | Select-String -Pattern ('^\s*' + [regex]::Escape($label) + '\s*:\s*(.+?)\s*$')
        if ($match) {
            $candidate = $match.Matches[0].Groups[1].Value.Trim()
            if ($candidate -match '^[\w\.\-\+]+@[\w\-]+(?:\.[\w\-]+)+$') {
                $email = $Matches[0]
                break
            }
        }
    }
}

if (-not $email) {
    $base = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    if (Test-Path $base) {
        Get-ChildItem -LiteralPath $base -ErrorAction SilentlyContinue | ForEach-Object {
            if (-not $email) {
                try {
                    $upn = (Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction Stop).UPN
                    if ($upn -and $upn -match '^[\w\.\-\+]+@[\w\-]+(?:\.[\w\-]+)+$') {
                        $email = $Matches[0]
                    }
                } catch {}
            }
        }
    }
}

if ($email) {
    $result = [PSCustomObject]@{
        DeviceName   = $env:COMPUTERNAME
        EmailAddress = $email
    }
    $result
    Ninja-Property-Set mdmUser $($result.EmailAddress)
    Set-NinjaUser $($result.EmailAddress)
} else {
    Write-Warning "Could not determine MDM user email for device $env:COMPUTERNAME."
}

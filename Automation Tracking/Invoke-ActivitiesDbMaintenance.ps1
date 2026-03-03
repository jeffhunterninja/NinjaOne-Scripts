<#
.SYNOPSIS
    Deletes Activities rows older than a configurable retention period and optionally runs VACUUM to reclaim space.
.DESCRIPTION
    Reads the same SQLite Activities database used by Get-AutomationActivities, Invoke-ScriptTracker, Invoke-ScriptStatusSync, and Invoke-AutomationDashboard.
    Deletes rows from the Activities table where activityTime (epoch seconds or milliseconds) is older than -RetentionDays.
    Handles both epoch-second and epoch-millisecond values in activityTime. Optionally runs VACUUM to shrink the database file.
    Uses sqlite3.exe: place it in the script directory, in PATH, or pass -SqliteExePath. All logic is in-line; no dot-sourcing.
.PARAMETER DbPath
    Path to the SQLite database file. Defaults to C:\RMM\Activities.db (same as other Automation Tracking scripts).
.PARAMETER RetentionDays
    Number of days to retain; rows older than this are deleted. Default 90. Set to 0 to skip deletion (e.g. only run VACUUM).
.PARAMETER SqliteExePath
    Full path to sqlite3.exe. If not set, script directory, PATH, then C:\RMM\sqlite3.exe are tried.
.PARAMETER Vacuum
    If set, runs VACUUM after deletion to reclaim disk space. Can be run without deletion (RetentionDays = 0) to compact the file.
.LINK
    https://www.sqlite.org/download.html
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$DbPath = 'C:\RMM\Activities.db',
    [Parameter()]
    [int]$RetentionDays = 90,
    [Parameter()]
    [string]$SqliteExePath = '',
    [Parameter()]
    [switch]$Vacuum
)

$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrWhiteSpace($DbPath)) { $DbPath = 'C:\RMM\Activities.db' }

# --- Resolve sqlite3.exe (same as Get-AutomationActivities / Invoke-ScriptTracker) ---
$sqliteExe = $null
if (-not [string]::IsNullOrWhiteSpace($SqliteExePath)) {
    if ((Test-Path -LiteralPath $SqliteExePath -PathType Leaf)) { $sqliteExe = $SqliteExePath }
    else { throw "SqliteExePath specified but file not found: $SqliteExePath. Download sqlite3.exe from https://www.sqlite.org/download.html." }
}
if (-not $sqliteExe -and $scriptDir) {
    $candidate = Join-Path $scriptDir 'sqlite3.exe'
    if (Test-Path -LiteralPath $candidate -PathType Leaf) { $sqliteExe = $candidate }
}
if (-not $sqliteExe) {
    $cmd = Get-Command sqlite3 -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -and (Test-Path -LiteralPath $cmd.Source -PathType Leaf)) { $sqliteExe = $cmd.Source }
}
if (-not $sqliteExe -and (Test-Path -LiteralPath 'C:\RMM\sqlite3.exe' -PathType Leaf)) { $sqliteExe = 'C:\RMM\sqlite3.exe' }
if (-not $sqliteExe) {
    throw "sqlite3.exe not found. Place it in the script directory, add to PATH, set -SqliteExePath, or install to C:\RMM\sqlite3.exe. Download from https://www.sqlite.org/download.html."
}

if (-not (Test-Path -LiteralPath $DbPath -PathType Leaf)) {
    throw "SQLite database not found: $DbPath. Run Get-AutomationActivities.ps1 first to create the database."
}

# --- SQLite helpers (in-line, no dot-sourcing) ---
function Invoke-SqliteNonQuery {
    param([string]$SqliteExe, [string]$DataSource, [string]$Sql)
    $errFile = [System.IO.Path]::GetTempFileName()
    try {
        $Sql | & $SqliteExe $DataSource 2> $errFile
        if ($LASTEXITCODE -ne 0) {
            $errText = Get-Content -LiteralPath $errFile -Raw -ErrorAction SilentlyContinue
            throw "sqlite3.exe exited with code $LASTEXITCODE. $errText"
        }
    } finally {
        if (Test-Path -LiteralPath $errFile) { Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue }
    }
}

function Invoke-SqliteScalar {
    param([string]$SqliteExe, [string]$DataSource, [string]$Sql)
    $out = & $SqliteExe $DataSource $Sql 2>&1
    $line = if ($out -is [string]) { $out.Trim() } else { ($out | Out-String).Trim() }
    $first = ($line -split "`n")[0]
    return $first.Trim()
}

Write-Host "[Info] === Activities DB maintenance starting ==="
Write-Host "[Info] DbPath: $DbPath | RetentionDays: $RetentionDays | Vacuum: $Vacuum"

$countBefore = [int](Invoke-SqliteScalar -SqliteExe $sqliteExe -DataSource $DbPath -Sql "SELECT COUNT(*) FROM Activities;")
Write-Host "[Info] Rows before: $countBefore"

$deleted = 0
if ($RetentionDays -gt 0) {
    $cutoffSec = [long][System.DateTimeOffset]::UtcNow.AddDays(-1 * [math]::Abs($RetentionDays)).ToUnixTimeSeconds()
    $cutoffMs = $cutoffSec * 1000
    $deleteSql = "DELETE FROM Activities WHERE (activityTime < $cutoffSec) OR (activityTime >= 1000000000000 AND activityTime < $cutoffMs);"
    Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $DbPath -Sql $deleteSql
    $countAfter = [int](Invoke-SqliteScalar -SqliteExe $sqliteExe -DataSource $DbPath -Sql "SELECT COUNT(*) FROM Activities;")
    $deleted = $countBefore - $countAfter
    Write-Host "[Info] Deleted rows older than $RetentionDays days: $deleted | Rows after: $countAfter"
} else {
    Write-Host "[Info] RetentionDays is 0; skipping deletion."
}

if ($Vacuum) {
    Write-Host "[Info] Running VACUUM..."
    Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $DbPath -Sql "VACUUM;"
    Write-Host "[Info] VACUUM complete."
}

Write-Host "[Info] === Activities DB maintenance complete ==="

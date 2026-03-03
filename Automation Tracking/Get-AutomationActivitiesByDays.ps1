<#
.SYNOPSIS
  Retrieves NinjaOne DEVICE/ACTION activities for a configurable number of days (1-90) from the API and inserts them into the same SQLite database used by Get-AutomationActivities. No duplicate entries (INSERT OR IGNORE by id).

.DESCRIPTION
  Fetches activities from the NinjaOne API for the last N days using after/before epoch parameters, pages with olderThan, enriches with device/org/location, and writes to SQLite. Uses the same DB path, table schema, and INSERT OR IGNORE so activities already present are skipped. Requires NinjaOneDocs module and sqlite3.exe. Requires PowerShell 7.

.PARAMETER Days
  Number of days to fetch (1-90). Default is 30. Script computes afterUnixEpoch and beforeUnixEpoch (UTC) from this.

.PARAMETER DbPath
  Path to the SQLite database file. Defaults to C:\RMM\Activities.db (same as Get-AutomationActivities).

.PARAMETER SqliteExePath
  Full path to sqlite3.exe. Same resolution as Get-AutomationActivities (script dir, PATH, C:\RMM\sqlite3.exe).

.PARAMETER NinjaOneInstance
  NinjaOne instance. Used when not running in NinjaOne; else Ninja-Property-Get or env NINJAONE_INSTANCE.

.PARAMETER NinjaOneClientId
  NinjaOne API client ID. Same resolution as Get-AutomationActivities.

.PARAMETER NinjaOneClientSecret
  NinjaOne API client secret. Same resolution as Get-AutomationActivities.

.LINK
  NinjaOneDocs: https://github.com/lwhitelock/NinjaOneDocs
.LINK
  https://www.sqlite.org/download.html
#>

[CmdletBinding()]
param (
    [Parameter()]
    [ValidateRange(1, 90)]
    [int]$Days = 30,
    [Parameter()]
    [string]$DbPath = 'C:\RMM\Activities.db',
    [Parameter()]
    [string]$SqliteExePath = 'C:\ProgramData\chocolatey\bin\sqlite3.exe',
    [Parameter()]
    [string]$NinjaOneInstance = '',
    [Parameter()]
    [string]$NinjaOneClientId = '',
    [Parameter()]
    [string]$NinjaOneClientSecret = ''
)

# Resolve path defaults from script directory when not supplied
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrWhiteSpace($DbPath)) { $DbPath = Join-Path $scriptDir 'Activities.db' }
$dbFile = $DbPath

# --- Resolve sqlite3.exe (no DLLs) ---
$sqliteExe = $null
if (-not [string]::IsNullOrWhiteSpace($SqliteExePath)) {
    if ((Test-Path -LiteralPath $SqliteExePath -PathType Leaf)) { $sqliteExe = $SqliteExePath }
    else { throw "SqliteExePath specified but file not found: $SqliteExePath. Download sqlite3.exe from https://www.sqlite.org/download.html and place it in the script directory or pass a valid -SqliteExePath." }
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
    throw "sqlite3.exe not found. Place it in the script directory, add it to PATH, set -SqliteExePath, or install to C:\RMM\sqlite3.exe. Download from https://www.sqlite.org/download.html (Precompiled Binaries for Windows)."
}

# SQLite helpers (sqlite3.exe; all in-line, no dot-sourcing)
function Escape-SqlString {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Replace("'", "''")
}
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
function Build-ActivityInsertSql {
    param([psobject]$Activity)
    $ingestTs = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $id = if ($null -ne $Activity.id) { [int]$Activity.id } else { 0 }
    $activityTime = if ($null -ne $Activity.activityTime) { [double]$Activity.activityTime } else { 'NULL' }
    $deviceId = if ($null -ne $Activity.deviceId) { [int]$Activity.deviceId } else { 0 }
    $orgId = if ($null -ne $Activity.OrgID) { [int]$Activity.OrgID } else { 0 }
    $locId = if ($null -ne $Activity.LocID) { [int]$Activity.LocID } else { 0 }
    $str = { param($v) if ($null -eq $v) { 'NULL' } else { "'" + (Escape-SqlString ([string]$v)) + "'" } }
    $dataJson = if ($null -ne $Activity.data) { & $str (ConvertTo-Json $Activity.data -Depth 10 -Compress) } else { 'NULL' }
    $activityTypeVal = if (-not [string]::IsNullOrWhiteSpace([string]$Activity.activityType)) { [string]$Activity.activityType } else { [string]$Activity.type }
    $statusCodeVal = if (-not [string]::IsNullOrWhiteSpace([string]$Activity.statusCode)) { [string]$Activity.statusCode } else { [string]$Activity.status }
    return "INSERT OR IGNORE INTO Activities (id, created_at, activityTime, deviceId, seriesUid, activityType, statusCode, status, activityResult, sourceConfigUid, sourceName, subject, message, type, data, OrgID, OrgName, LocID, LocName, DeviceName) VALUES ($id, $(& $str $ingestTs), $activityTime, $deviceId, $(& $str $Activity.seriesUid), $(& $str $activityTypeVal), $(& $str $statusCodeVal), $(& $str $Activity.status), $(& $str $Activity.activityResult), $(& $str $Activity.sourceConfigUid), $(& $str $Activity.sourceName), $(& $str $Activity.subject), $(& $str $Activity.message), $(& $str $Activity.type), $dataJson, $orgId, $(& $str $Activity.OrgName), $locId, $(& $str $Activity.LocName), $(& $str $Activity.DeviceName));"
}

# NinjaOneDocs module (must load before Ninja-Property-Get)
try {
    $moduleName = "NinjaOneDocs"
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Install-Module -Name $moduleName -Force -AllowClobber
    }
    Import-Module $moduleName -ErrorAction Stop
} catch {
    Write-Error "Failed to import NinjaOneDocs module. Error: $_"
    exit 1
}

# Credentials: Ninja-Property-Get (when in NinjaOne) -> env vars -> parameters
$N1Instance = $null
try { $fromNinja = Ninja-Property-Get ninjaoneInstance; if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $N1Instance = $fromNinja } } catch { }
if ([string]::IsNullOrWhiteSpace($N1Instance)) { $N1Instance = $env:NINJAONE_INSTANCE }
if ([string]::IsNullOrWhiteSpace($N1Instance) -and $PSBoundParameters.ContainsKey('NinjaOneInstance')) { $N1Instance = $NinjaOneInstance }

$N1ClientId = $null
try { $fromNinja = Ninja-Property-Get ninjaoneClientId; if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $N1ClientId = $fromNinja } } catch { }
if ([string]::IsNullOrWhiteSpace($N1ClientId)) { $N1ClientId = $env:NINJAONE_CLIENT_ID }
if ([string]::IsNullOrWhiteSpace($N1ClientId) -and $PSBoundParameters.ContainsKey('NinjaOneClientId')) { $N1ClientId = $NinjaOneClientId }

$N1ClientSecret = $null
try { $fromNinja = Ninja-Property-Get ninjaoneClientSecret; if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $N1ClientSecret = $fromNinja } } catch { }
if ([string]::IsNullOrWhiteSpace($N1ClientSecret)) { $N1ClientSecret = $env:NINJAONE_CLIENT_SECRET }
if ([string]::IsNullOrWhiteSpace($N1ClientSecret) -and $PSBoundParameters.ContainsKey('NinjaOneClientSecret')) { $N1ClientSecret = $NinjaOneClientSecret }

if ([string]::IsNullOrWhiteSpace($N1Instance) -or [string]::IsNullOrWhiteSpace($N1ClientId) -or [string]::IsNullOrWhiteSpace($N1ClientSecret)) {
    Write-Error "Missing required API credentials. Set ninjaoneInstance, ninjaoneClientId, ninjaoneClientSecret in NinjaOne custom properties, or use env vars NINJAONE_INSTANCE, NINJAONE_CLIENT_ID, NINJAONE_CLIENT_SECRET, or -NinjaOneInstance, -NinjaOneClientId, -NinjaOneClientSecret."
    exit 1
}

try {
    Connect-NinjaOne -NinjaOneInstance $N1Instance -NinjaOneClientID $N1ClientId -NinjaOneClientSecret $N1ClientSecret
} catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit 1
}

$startTime = Get-Date
Write-Host "[Info] === NinjaOne Activities by Days ($Days days) → SQLite sync starting ==="

# Compute date range (UTC epoch seconds)
$rangeStart = [DateTimeOffset]::UtcNow.AddDays(-$Days)
$rangeEnd = [DateTimeOffset]::UtcNow
$afterEpoch = [long]$rangeStart.ToUnixTimeSeconds()
$beforeEpoch = [long]$rangeEnd.ToUnixTimeSeconds()
Write-Host "[Info] Date range: after=$afterEpoch, before=$beforeEpoch (UTC epoch seconds)"

# Bootstrap DB and indexes (same as Get-AutomationActivities)
if (-not (Test-Path $dbFile)) { Write-Host "[Info] DB not found; creating $dbFile"; New-Item -Path $dbFile -ItemType File | Out-Null }
$createTableSql = "CREATE TABLE IF NOT EXISTS Activities ( id INTEGER PRIMARY KEY, created_at TEXT, activityTime REAL, deviceId INTEGER, seriesUid TEXT, activityType TEXT, statusCode TEXT, status TEXT, activityResult TEXT, sourceConfigUid TEXT, sourceName TEXT, subject TEXT, message TEXT, type TEXT, data TEXT, OrgID INTEGER, OrgName TEXT, LocID INTEGER, LocName TEXT, DeviceName TEXT );"
Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql $createTableSql
Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql "CREATE INDEX IF NOT EXISTS IX_Activities_activityTime ON Activities(activityTime DESC);"
Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql "CREATE INDEX IF NOT EXISTS IX_Activities_deviceId ON Activities(deviceId);"
Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql "CREATE INDEX IF NOT EXISTS IX_Activities_status ON Activities(status);"

$validStatuses = @('CANCEL_REQUESTED','CANCELLED','COMPLETED','IN_PROCESS','START_REQUESTED','STARTED')

# Ref data via NinjaOneDocs
try {
    $devices = Invoke-NinjaOneRequest -Method GET -Path 'devices-detailed'
    $orgs    = Invoke-NinjaOneRequest -Method GET -Path 'organizations'
    $locs    = Invoke-NinjaOneRequest -Method GET -Path 'locations'
} catch {
    throw "Ref data fetch failed: $($_.Exception.Message)"
}
$devById = @{}; foreach ($d in $devices) { $devById[$d.id] = $d }
$orgById = @{}; foreach ($o in $orgs)    { $orgById[$o.id] = $o }
$locById = @{}; foreach ($l in $locs)    { $locById[$l.id] = $l }
$maps = [pscustomobject]@{ Devices = $devById; Orgs = $orgById; Locs = $locById }

# Enrichment + insert (null-safe; same as Get-AutomationActivities)
function Enrich-And-Insert {
    param($items, $mapsRef, $db, $sqliteExePath)
    $countBefore = [int](Invoke-SqliteScalar -SqliteExe $sqliteExePath -DataSource $db -Sql "SELECT COUNT(*) FROM Activities;")
    $sqlBatch = [System.Text.StringBuilder]::new()
    [void]$sqlBatch.AppendLine("BEGIN;")
    $attempted = 0
    foreach ($act in $items) {
        $d = $mapsRef.Devices[$act.deviceId]
        $o = $null; $l = $null
        if ($d) { $o = $mapsRef.Orgs[$d.organizationId]; $l = $mapsRef.Locs[$d.locationId] }
        $deviceName = if ($d -and $d.systemName) { $d.systemName } else { '' }
        $orgId = if ($d -and $null -ne $d.organizationId) { $d.organizationId } else { 0 }
        $orgName = if ($o -and $o.name) { $o.name } else { '' }
        $locId = if ($d -and $null -ne $d.locationId) { $d.locationId } else { 0 }
        $locName = if ($l -and $l.name) { $l.name } else { '' }
        $act | Add-Member -NotePropertyName DeviceName -NotePropertyValue $deviceName -Force
        $act | Add-Member -NotePropertyName OrgID      -NotePropertyValue $orgId -Force
        $act | Add-Member -NotePropertyName OrgName    -NotePropertyValue $orgName -Force
        $act | Add-Member -NotePropertyName LocID      -NotePropertyValue $locId -Force
        $act | Add-Member -NotePropertyName LocName    -NotePropertyValue $locName -Force
        [void]$sqlBatch.AppendLine((Build-ActivityInsertSql -Activity $act))
        $attempted++
    }
    [void]$sqlBatch.AppendLine("COMMIT;")
    Invoke-SqliteNonQuery -SqliteExe $sqliteExePath -DataSource $db -Sql $sqlBatch.ToString()
    $countAfter = [int](Invoke-SqliteScalar -SqliteExe $sqliteExePath -DataSource $db -Sql "SELECT COUNT(*) FROM Activities;")
    $ins = $countAfter - $countBefore
    $skip = $attempted - $ins
    ,@($ins, $skip)
}

# Fetch by date range: API uses after, before (epoch seconds), and olderThan for pagination
$retrievedCount = 0
$filteredCount = 0
$insertedCount = 0
$skippedCount = 0
$olderThan = $null

while ($true) {
    $queryParams = "class=DEVICE&type=ACTION&pageSize=1000&after=$afterEpoch&before=$beforeEpoch"
    if ($null -ne $olderThan) { $queryParams += "&olderThan=$olderThan" }

    try {
        $resp = Invoke-NinjaOneRequest -Method GET -Path 'activities' -QueryParams $queryParams
    } catch {
        throw "Activities fetch failed: $($_.Exception.Message)"
    }

    $rawBatch = $resp.activities
    if (-not $rawBatch -or $rawBatch.Count -eq 0) { break }

    $retrievedCount += $rawBatch.Count
    $statusFiltered = $rawBatch | Where-Object { $validStatuses -contains $_.status }
    # Filter to activities within our window (activityTime can be sec or ms)
    $statusFiltered = $statusFiltered | Where-Object {
        $t = $_.activityTime
        if ($null -eq $t) { $false }
        elseif ($t -ge 1000000000000) { $t -ge ($afterEpoch * 1000) -and $t -le ($beforeEpoch * 1000) }
        else { $t -ge $afterEpoch -and $t -le $beforeEpoch }
    }
    $filteredCount += $statusFiltered.Count

    if ($statusFiltered.Count -gt 0) {
        $res = Enrich-And-Insert -items $statusFiltered -mapsRef $maps -db $dbFile -sqliteExePath $sqliteExe
        $insertedCount += $res[0]
        $skippedCount += $res[1]
    }

    # Next page: oldest in batch (last item; API returns newest first)
    $oldestInBatch = $rawBatch[-1]
    if ($null -eq $oldestInBatch -or $null -eq $oldestInBatch.id) { break }

    $oldestTime = $oldestInBatch.activityTime
    if ($null -ne $oldestTime) {
        $oldestEpoch = if ($oldestTime -ge 1000000000000) { [long]($oldestTime / 1000) } else { [long]$oldestTime }
        if ($oldestEpoch -lt $afterEpoch) { break }
    }

    $olderThan = $oldestInBatch.id
    if ($rawBatch.Count -lt 1000) { break }
}

Write-Host ("[Info] Sync complete: retrieved {0}, filtered {1}, inserted {2}, skipped (duplicates) {3}." -f $retrievedCount, $filteredCount, $insertedCount, $skippedCount)
$rowCount = Invoke-SqliteScalar -SqliteExe $sqliteExe -DataSource $dbFile -Sql "SELECT COUNT(*) FROM Activities;"
Write-Host "[Info] Verified row count in DB: $rowCount"
$duration = (Get-Date) - $startTime
Write-Host ("[Info] Total duration: {0} seconds" -f [math]::Round($duration.TotalSeconds, 2))

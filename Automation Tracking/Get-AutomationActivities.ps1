<#
.SYNOPSIS
  Syncs NinjaOne DEVICE/ACTION activities to a local SQLite database. First run seeds the most recent 1,000; subsequent runs are incremental using newerThan. Requires NinjaOneDocs module.

.DESCRIPTION
  Pulls NinjaOne activities via the API, enriches with device/org/location, and writes to SQLite. Uses NinjaOneDocs for authentication and API calls (Connect-NinjaOne, Invoke-NinjaOneRequest). Requires PowerShell 7. The SQLite command-line tool (sqlite3.exe) is required: place it in the script directory, in PATH, or pass -SqliteExePath. Download from https://www.sqlite.org/download.html (Precompiled Binaries for Windows).

.PARAMETER DbPath
  Path to the SQLite database file. Defaults to script directory\scriptruns.db.

.PARAMETER SqliteExePath
  Full path to sqlite3.exe. If not set, the script looks for sqlite3.exe in the script directory, then in PATH, then at C:\RMM\sqlite3.exe. Download from https://www.sqlite.org/download.html if needed.

.PARAMETER WindowMinutes
  Optional time-window filter for recent-only evaluation (0 = disabled).

.PARAMETER NinjaOneInstance
  NinjaOne instance (e.g. app.ninjaone.com). Used when not running in NinjaOne; else use Ninja-Property-Get or env NINJAONE_INSTANCE.

.PARAMETER NinjaOneClientId
  NinjaOne API client ID. Used when not running in NinjaOne; else use Ninja-Property-Get or env NINJAONE_CLIENT_ID.

.PARAMETER NinjaOneClientSecret
  NinjaOne API client secret. Used when not running in NinjaOne; else use Ninja-Property-Get or env NINJAONE_CLIENT_SECRET.

.LINK
  NinjaOneDocs: https://github.com/lwhitelock/NinjaOneDocs
.LINK
  https://www.sqlite.org/download.html
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$DbPath = 'C:\RMM\Activities.db',
    [Parameter()]
    [string]$SqliteExePath = 'C:\ProgramData\chocolatey\bin\sqlite3.exe',
    [Parameter()]
    [int]$WindowMinutes = 0,
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
if ([string]::IsNullOrWhiteSpace($DbPath)) { $DbPath = Join-Path $scriptDir 'scriptruns.db' }
$dbFile        = $DbPath
$windowMinutes = $WindowMinutes

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
function Invoke-SqliteQuery {
    param([string]$SqliteExe, [string]$DataSource, [string]$Sql)
    $out = & $SqliteExe -csv -header $DataSource $Sql 2>$null
    $text = if ($null -eq $out) { '' } elseif ($out -is [string]) { $out } else { $out -join "`n" }
    $text = $text.TrimStart([char]0xFEFF)
    $lines = $text -split "`r?`n"
    $lines = $lines | Where-Object { $_.Length -gt 0 }
    if ($lines.Count -lt 2) { return @() }
    $headers = $lines[0] -split ',(?=(?:[^"]*"[^"]*")*[^"]*$)'
    $headers = $headers | ForEach-Object { $_.Trim().Trim('"').TrimStart([char]0xFEFF) }
    $result = [System.Collections.Generic.List[pscustomobject]]::new()
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $fields = $lines[$i] -split ',(?=(?:[^"]*"[^"]*")*[^"]*$)'
        $fields = $fields | ForEach-Object { $_.Trim().Trim('"').Replace('""', '"') }
        $row = [ordered]@{}
        for ($j = 0; $j -lt [Math]::Min($headers.Count, $fields.Count); $j++) { $row[$headers[$j]] = $fields[$j] }
        $result.Add([pscustomobject]$row)
    }
    return $result
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
try {
    $fromNinja = Ninja-Property-Get ninjaoneInstance
    if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $N1Instance = $fromNinja }
} catch { }
if ([string]::IsNullOrWhiteSpace($N1Instance)) { $N1Instance = $env:NINJAONE_INSTANCE }
if ([string]::IsNullOrWhiteSpace($N1Instance) -and $PSBoundParameters.ContainsKey('NinjaOneInstance')) { $N1Instance = $NinjaOneInstance }

try {
    $fromNinja = Ninja-Property-Get ninjaoneClientId
    if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $N1ClientId = $fromNinja }
} catch { }
if ([string]::IsNullOrWhiteSpace($N1ClientId)) { $N1ClientId = $env:NINJAONE_CLIENT_ID }
if ([string]::IsNullOrWhiteSpace($N1ClientId) -and $PSBoundParameters.ContainsKey('NinjaOneClientId')) { $N1ClientId = $NinjaOneClientId }

try {
    $fromNinja = Ninja-Property-Get ninjaoneClientSecret
    if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $N1ClientSecret = $fromNinja }
} catch { }
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
Write-Host "[Info] === NinjaOne Activities → SQLite sync starting ==="

# Bootstrap DB and indexes (no activityTimeIso column). One statement per call to avoid sqlite3 "incomplete input" when passing multi-line SQL on Windows.
if (-not (Test-Path $dbFile)) { Write-Host "[Info] DB not found; creating $dbFile"; New-Item -Path $dbFile -ItemType File | Out-Null }
$createTableSql = "CREATE TABLE IF NOT EXISTS Activities ( id INTEGER PRIMARY KEY, created_at TEXT, activityTime REAL, deviceId INTEGER, seriesUid TEXT, activityType TEXT, statusCode TEXT, status TEXT, activityResult TEXT, sourceConfigUid TEXT, sourceName TEXT, subject TEXT, message TEXT, type TEXT, data TEXT, OrgID INTEGER, OrgName TEXT, LocID INTEGER, LocName TEXT, DeviceName TEXT );"
Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql $createTableSql
Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql "CREATE INDEX IF NOT EXISTS IX_Activities_activityTime ON Activities(activityTime DESC);"
Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql "CREATE INDEX IF NOT EXISTS IX_Activities_deviceId ON Activities(deviceId);"
Invoke-SqliteNonQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql "CREATE INDEX IF NOT EXISTS IX_Activities_status ON Activities(status);"

# 5) Statuses
$validStatuses = @('CANCEL_REQUESTED','CANCELLED','COMPLETED','IN_PROCESS','START_REQUESTED','STARTED')

# 6) Incremental seed: MAX(id)
[int]$maxId = [int](Invoke-SqliteScalar -SqliteExe $sqliteExe -DataSource $dbFile -Sql "SELECT COALESCE(MAX(id),0) FROM Activities;")
Write-Host "[Info] Seeding newerThan from DB MAX(id) = $maxId"

# Optional time window bound (epoch seconds, UTC)
$minEpoch = 0
if ($windowMinutes -gt 0) {
    $minEpoch = [int][DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - ($windowMinutes * 60)
    Write-Host "[Info] Time-window filter enabled: last $windowMinutes minutes (epoch >= $minEpoch)"
}

# 7) Ref data via NinjaOneDocs
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

# 8) Enrichment + insert (null-safe when device/org/location missing from ref data)
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
    # Return a flat 2-element array so callers can use $res[0] and $res[1].
    @($ins, $skip)
}

# 9) Fetch & write
$retrievedCount = 0; $filteredCount = 0; $insertedCount = 0; $skippedCount = 0
$newerThan = $maxId

if ($maxId -eq 0) {
    # FIRST RUN: seed only the most recent 1000
    try {
        $resp = Invoke-NinjaOneRequest -Method GET -Path 'activities' -QueryParams 'class=DEVICE&type=ACTION&pageSize=1000'
        $rawBatch = $resp.activities
    } catch { throw "Activities seed fetch failed: $($_.Exception.Message)" }
    if ($rawBatch) {
        $retrievedCount = $rawBatch.Count
        $statusFiltered = $rawBatch | Where-Object { $validStatuses -contains $_.status }
        if ($minEpoch -gt 0) { $statusFiltered = $statusFiltered | Where-Object { $_.activityTime -is [double] -and $_.activityTime -ge $minEpoch } }
        $filteredCount = $statusFiltered.Count
        $res = Enrich-And-Insert -items $statusFiltered -mapsRef $maps -db $dbFile -sqliteExePath $sqliteExe
        $insertedCount = $res[0]; $skippedCount = $res[1]
    }
} else {
    # INCREMENTAL: page forward with newerThan
    while ($true) {
        try {
            $resp = Invoke-NinjaOneRequest -Method GET -Path 'activities' -QueryParams "class=DEVICE&type=ACTION&pageSize=1000&newerThan=$newerThan"
            $rawBatch = $resp.activities
        } catch { throw "Activities fetch failed: $($_.Exception.Message)" }
        if (-not $rawBatch -or $rawBatch.Count -eq 0) { break }
        $retrievedCount += $rawBatch.Count
        $batchMaxId = ($rawBatch | Measure-Object -Property id -Maximum).Maximum
        $statusFiltered = $rawBatch | Where-Object { $validStatuses -contains $_.status }
        if ($minEpoch -gt 0) { $statusFiltered = $statusFiltered | Where-Object { $_.activityTime -is [double] -and $_.activityTime -ge $minEpoch } }
        $filteredCount += $statusFiltered.Count
        $res = Enrich-And-Insert -items $statusFiltered -mapsRef $maps -db $dbFile -sqliteExePath $sqliteExe
        $insertedCount += $res[0]; $skippedCount += $res[1]
        if ($batchMaxId -le $newerThan) { break }
        $newerThan = [int]$batchMaxId
    }
}

Write-Host ("[Info] Sync complete: retrieved {0}, filtered {1}, inserted {2}, skipped {3}." -f $retrievedCount, $filteredCount, $insertedCount, $skippedCount)
$rowCount = Invoke-SqliteScalar -SqliteExe $sqliteExe -DataSource $dbFile -Sql "SELECT COUNT(*) FROM Activities;"
Write-Host "[Info] Verified row count in DB: $rowCount"
$duration = (Get-Date) - $startTime
Write-Host ("[Info] Total duration: {0} seconds" -f [math]::Round($duration.TotalSeconds,2))

# 10) Diagnostic: show what is actually stored for status/type (so Sync can match)
$sampleQ = "SELECT id, statusCode, status, activityType, type, sourceName FROM Activities ORDER BY id DESC LIMIT 5;"
$sampleRows = Invoke-SqliteQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql $sampleQ
if ($sampleRows -and $sampleRows.Count -gt 0) {
    Write-Host "[Info] Sample of latest rows in DB (statusCode, status, activityType, type):"
    foreach ($r in $sampleRows) {
        Write-Host ("  id={0} statusCode='{1}' status='{2}' activityType='{3}' type='{4}' sourceName='{5}'" -f [string]$r.id, [string]$r.statusCode, [string]$r.status, [string]$r.activityType, [string]$r.type, [string]$r.sourceName)
    }
}

# 11) Recent preview (compute local time from epoch; handles ms or sec)
function Get-RecentActivities {
    param([string]$DbFile, [int]$Limit = 10, [string]$SqliteExePath)
    $q = "SELECT id, created_at, CASE WHEN activityTime IS NOT NULL AND activityTime >= 1000000000000 THEN datetime(CAST(activityTime/1000 AS INTEGER), 'unixepoch', 'localtime') WHEN activityTime IS NOT NULL AND activityTime >= 1000000000 THEN datetime(CAST(activityTime AS INTEGER), 'unixepoch', 'localtime') ELSE NULL END AS activityTimeLocal, activityType, status, OrgName, LocName, DeviceName FROM Activities ORDER BY created_at DESC LIMIT $Limit;"
    Invoke-SqliteQuery -SqliteExe $SqliteExePath -DataSource $DbFile -Sql $q
}
Write-Host "[Info] Recent activities (by created_at):"
$recent = Get-RecentActivities -DbFile $dbFile -Limit 10 -SqliteExePath $sqliteExe
if ($recent -and $recent.Count -gt 0) {
    $recent | Format-Table id, created_at, activityTimeLocal, OrgName, DeviceName, activityType, status -AutoSize
} else {
    Write-Host "  (none or query returned no rows)"
}

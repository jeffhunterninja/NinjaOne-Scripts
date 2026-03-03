<#
.SYNOPSIS
  Reads latest COMPLETED script runs from the Activities SQLite DB and syncs per-device script status to a NinjaOne device custom field (e.g. scriptStatus). Requires Get-AutomationActivities to have run first.

.DESCRIPTION
  Queries the same SQLite database populated by Get-AutomationActivities for the latest COMPLETED ACTION run per device and script name, maps activity result/status to SUCCESS/FAILURE, and writes a key:value map to a NinjaOne device custom field. Requires NinjaOneDocs for authentication and API calls. Requires PowerShell 7. Uses sqlite3.exe: place it in the script directory, in PATH, or pass -SqliteExePath. Download from https://www.sqlite.org/download.html (Precompiled Binaries for Windows).

.PARAMETER DbPath
  Path to the SQLite database file. Defaults to C:\RMM\Activities.db (same as Get-AutomationActivities).

.PARAMETER CustomFieldName
  NinjaOne device custom field name (label or key) to write script status to. Defaults to scriptStatus.

.PARAMETER SqliteExePath
  Full path to sqlite3.exe. If not set, the script looks for sqlite3.exe in the script directory, then in PATH, then at C:\RMM\sqlite3.exe. Download from https://www.sqlite.org/download.html if needed.

.PARAMETER ForceRebuild
  When true, restricts the query to activities within -LookbackDays. Defaults to true.

.PARAMETER LookbackDays
  Number of days to look back when ForceRebuild is true. Defaults to 10.

.PARAMETER PreserveExisting
  When true with ForceRebuild, merge existing custom field values with newly synced script status so scripts not in the lookback window are preserved.

.PARAMETER TrimRunPrefix
  When true, strip a leading "run " from script names when building the status map. Defaults to true.

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
    [string]$CustomFieldName = 'scriptStatus',
    [Parameter()]
    [string]$SqliteExePath = 'C:\ProgramData\chocolatey\bin\sqlite3.exe',
    [Parameter()]
    [bool]$ForceRebuild = $true,
    [Parameter()]
    [int]$LookbackDays = 10,
    [Parameter()]
    [bool]$PreserveExisting = $false,
    [Parameter()]
    [bool]$TrimRunPrefix = $true,
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

$CompletedStatuses = @('COMPLETED')
$ScriptNameFieldPreference = @('sourceName','subject','scriptName')

if (-not (Test-Path $DbPath)) { throw "SQLite DB not found: $DbPath" }

# --- Resolve sqlite3.exe (no DLLs; same workflow as Get-AutomationActivities) ---
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

# SQLite helpers (sqlite3.exe; all in-line, no dot-sourcing; same as Get-AutomationActivities)
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

function Convert-MapToScriptStatusText {
    param([Parameter(Mandatory)][hashtable]$Map)
    $keys = $Map.Keys | ForEach-Object { [string]$_ } | Sort-Object
    ($keys | ForEach-Object { '{0}:{1}' -f $_, [string]$Map[$_] }) -join "`r`n"
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

function Canonicalize-Name {
  param([string]$Name,[switch]$TrimRun)
  if (-not $Name) { return @{ key=''; display='' } }
  $n = $Name.Trim()
  if ($TrimRun) { $n = ($n -replace '^\s*run\s+','', 'IgnoreCase') }
  @{ key = ($n -replace '\s+',' ').ToLowerInvariant(); display = $n }
}
function Map-Status {
  param([string]$ActivityResult,[string]$Status,[string]$StatusCode)
  $ar = ([string]$ActivityResult).ToUpperInvariant()
  if ($ar -eq 'SUCCESS') { return 'SUCCESS' }
  if ($ar -eq 'FAILURE') { return 'FAILURE' }
  $s  = ([string]$Status).ToUpperInvariant()
  $sc = ([string]$StatusCode).ToUpperInvariant()
  if ($s -match 'SUCCESS|SUCCEEDED') { return 'SUCCESS' }
  if ($s -match 'FAIL|FAILED|FAILURE|ERROR') { return 'FAILURE' }
  if ($sc -match 'SUCCESS') { return 'SUCCESS' }
  if ($sc -match 'FAIL|ERROR') { return 'FAILURE' }
  'FAILURE'
}

# Cache resolved keys (label -> working key)
$script:FieldKeyCache = @{}

# Cache for resolved custom-field keys
if (-not $script:FieldKeyCache) { $script:FieldKeyCache = @{} }

function Get-FieldKeyVariants {
  param([Parameter(Mandatory)][string]$Name)

  # Clean and split words
  $clean = ($Name -replace '[^A-Za-z0-9 _-]', ' ').Trim()
  $parts = @()
  foreach ($p in ($clean -split '\s+')) { if ($p) { $parts += $p } }

  if ($parts.Count -eq 0) { return ,$Name }

  # PascalCase
  $pascalSB = New-Object System.Text.StringBuilder
  foreach ($p in $parts) {
    if ($p.Length -ge 1) {
      [void]$pascalSB.Append($p.Substring(0,1).ToUpper())
      if ($p.Length -gt 1) { [void]$pascalSB.Append($p.Substring(1)) }
    }
  }
  $pascal = $pascalSB.ToString()

  # camelCase
  $camelSB = New-Object System.Text.StringBuilder
  $first = $parts[0]
  if ($first.Length -ge 1) {
    [void]$camelSB.Append($first.Substring(0,1).ToLower())
    if ($first.Length -gt 1) { [void]$camelSB.Append($first.Substring(1)) }
  }
  if ($parts.Count -gt 1) {
    for ($i = 1; $i -lt $parts.Count; $i++) {
      $p = $parts[$i]
      if ($p.Length -ge 1) {
        [void]$camelSB.Append($p.Substring(0,1).ToUpper())
        if ($p.Length -gt 1) { [void]$camelSB.Append($p.Substring(1)) }
      }
    }
  }
  $camel = $camelSB.ToString()

  # other common variants
  $snake   = ($parts -join '_')
  $snakeL  = $snake.ToLower()
  $concat  = ($parts -join '')
  $concatL = $concat.ToLower()
  $lower   = $clean.ToLower().Replace(' ','')  # e.g. "script status" -> "scriptstatus"

  @($Name, $camel, $pascal, $snake, $snakeL, $concat, $concatL, $lower) | Select-Object -Unique
}

function Set-DeviceCustomFieldValue {
  param(
    [Parameter(Mandatory)][string]$DeviceId,
    [Parameter(Mandatory)][string]$DesiredKey,
    [Parameter(Mandatory)][string]$Value,
    [switch]$StrictFieldKey
  )

  $candidates = @()
  if ($script:FieldKeyCache.ContainsKey($DesiredKey)) { $candidates += $script:FieldKeyCache[$DesiredKey] }
  $candidates += $DesiredKey
  if (-not $StrictFieldKey) { $candidates += (Get-FieldKeyVariants -Name $DesiredKey) }
  $candidates = $candidates | Where-Object { $_ } | Select-Object -Unique

  foreach ($key in $candidates) {
    $body = @{ $key = $Value } | ConvertTo-Json -Compress
    try {
      Invoke-NinjaOneRequest -Method PATCH -Path ("device/{0}/custom-fields" -f $DeviceId) -Body $body | Out-Null
      $script:FieldKeyCache[$DesiredKey] = $key
      Write-Host ("Updated {0} for device {1} (key used: '{2}')" -f $DesiredKey, $DeviceId, $key)
      return $true
    } catch {
      $txt = ($_ | Out-String)
      if ($txt -match 'FIELD_NOT_FOUND') { continue }  # try next variant
      Write-Warning ("Update failed for device {0} using key '{1}': {2}" -f $DeviceId, $key, $txt.Trim())
      return $false
    }
  }

  Write-Warning ("No working key found for field '{0}' on device {1}. Tried: {2}" -f $DesiredKey, $DeviceId, ($candidates -join ', '))
  return $false
}


function Compare-And-UpdateCustomField {
  param(
    [Parameter(Mandatory)][string]$DeviceId,
    [Parameter(Mandatory)][string]$FieldName,       # the key you want (label or exact key)
    [Parameter(Mandatory)][hashtable]$MapToWrite,
    [switch]$StrictFieldKey                         # if set, only try FieldName as-is
  )

  # Deterministic JSON (stable property order) to store
  $ordered = [System.Collections.Specialized.OrderedDictionary]::new()
  foreach ($k in ($MapToWrite.Keys | ForEach-Object { [string]$_ } | Sort-Object)) { $ordered[$k] = [string]$MapToWrite[$k] }
  # Build human-readable text instead of JSON
  $newText = Convert-MapToScriptStatusText -Map $MapToWrite
  
  # Try to avoid a PATCH if we can read an identical value (field may be omitted if null)
  $currentValue = $null
  try {
    $cf = Invoke-NinjaOneRequest -Method GET -Path ("device/{0}/custom-fields" -f $DeviceId) -Paginate:$false
    if ($cf -and $cf.PSObject -and $cf.PSObject.Properties) {
      $allKeys = $cf.PSObject.Properties.Name
      $resolved = $FieldName
      if (-not ($allKeys -icontains $resolved) -and $script:FieldKeyCache.ContainsKey($FieldName)) {
        $resolved = $script:FieldKeyCache[$FieldName]
      }
      if ($allKeys -icontains $resolved) { $currentValue = [string]$cf."$resolved" }
    }
  } catch { }
  
  if ($currentValue -and $currentValue -ceq $newText) {
    $resolved = $(if ($script:FieldKeyCache.ContainsKey($FieldName)) { $script:FieldKeyCache[$FieldName] } else { $FieldName })
    Write-Host ("No change for {0} on device {1} (key '{2}')" -f $FieldName, $DeviceId, $resolved)
    return
  }
  
  # Always attempt the write even if GET omitted the key (null/empty)
  [void](Set-DeviceCustomFieldValue -DeviceId $DeviceId -DesiredKey $FieldName -Value $newText)
}



# --- SQL (TRIM + ACTION + optional window) ---
$timeCond = ''
if ($ForceRebuild) {
  $cutoffEpoch = [long][System.DateTimeOffset]::UtcNow.AddDays(-1 * [math]::Abs($LookbackDays)).ToUnixTimeSeconds()
  $timeCond = " AND activityTime >= $cutoffEpoch "
}

# Diagnostic: report row counts so we can confirm data is visible (same DB, correct filters)
$totalRows = [int](Invoke-SqliteScalar -SqliteExe $sqliteExe -DataSource $DbPath -Sql "SELECT COUNT(*) FROM Activities;")
$completedRows = [int](Invoke-SqliteScalar -SqliteExe $sqliteExe -DataSource $DbPath -Sql "SELECT COUNT(*) FROM Activities WHERE (statusCode IS NOT NULL AND UPPER(TRIM(statusCode)) = 'COMPLETED') OR (status IS NOT NULL AND UPPER(TRIM(status)) = 'COMPLETED');")
$actionRows = [int](Invoke-SqliteScalar -SqliteExe $sqliteExe -DataSource $DbPath -Sql "SELECT COUNT(*) FROM Activities WHERE (activityType IS NOT NULL AND UPPER(TRIM(activityType)) = 'ACTION') OR (type IS NOT NULL AND UPPER(TRIM(type)) = 'ACTION');")
$completedActionRows = [int](Invoke-SqliteScalar -SqliteExe $sqliteExe -DataSource $DbPath -Sql "SELECT COUNT(*) FROM Activities WHERE ((statusCode IS NOT NULL AND UPPER(TRIM(statusCode)) = 'COMPLETED') OR (status IS NOT NULL AND UPPER(TRIM(status)) = 'COMPLETED')) AND ((activityType IS NOT NULL AND UPPER(TRIM(activityType)) = 'ACTION') OR (type IS NOT NULL AND UPPER(TRIM(type)) = 'ACTION'));")
Write-Host ("[Info] DB: $DbPath | Total rows: $totalRows | COMPLETED: $completedRows | ACTION: $actionRows | COMPLETED+ACTION: $completedActionRows")

$sql = @"
WITH latest AS (
  SELECT deviceId, sourceName, MAX(activityTime) AS maxTime
  FROM Activities
  WHERE deviceId IS NOT NULL
    AND sourceName IS NOT NULL AND TRIM(sourceName) <> ''
    AND ( (statusCode IS NOT NULL AND UPPER(TRIM(statusCode)) = 'COMPLETED') OR (status IS NOT NULL AND UPPER(TRIM(status)) = 'COMPLETED') )
    AND ( (activityType IS NOT NULL AND UPPER(TRIM(activityType)) = 'ACTION') OR (type IS NOT NULL AND UPPER(TRIM(type)) = 'ACTION') )
    $timeCond
  GROUP BY deviceId, sourceName
),
latest2 AS (
  SELECT a.deviceId, a.sourceName, MAX(a.id) AS maxId
  FROM Activities a
  JOIN latest l
    ON a.deviceId = l.deviceId
   AND a.sourceName = l.sourceName
   AND a.activityTime = l.maxTime
  WHERE ( (a.statusCode IS NOT NULL AND UPPER(TRIM(a.statusCode)) = 'COMPLETED') OR (a.status IS NOT NULL AND UPPER(TRIM(a.status)) = 'COMPLETED') )
    AND ( (a.activityType IS NOT NULL AND UPPER(TRIM(a.activityType)) = 'ACTION') OR (a.type IS NOT NULL AND UPPER(TRIM(a.type)) = 'ACTION') )
  GROUP BY a.deviceId, a.sourceName
)
SELECT a.deviceId, a.sourceName, a.activityResult, a.status, a.statusCode, a.activityTime, a.id
FROM Activities a
JOIN latest2 t
  ON a.deviceId  = t.deviceId
 AND a.sourceName= t.sourceName
 AND a.id        = t.maxId
WHERE ( (a.statusCode IS NOT NULL AND UPPER(TRIM(a.statusCode)) = 'COMPLETED') OR (a.status IS NOT NULL AND UPPER(TRIM(a.status)) = 'COMPLETED') )
  AND ( (a.activityType IS NOT NULL AND UPPER(TRIM(a.activityType)) = 'ACTION') OR (a.type IS NOT NULL AND UPPER(TRIM(a.type)) = 'ACTION') )
ORDER BY a.deviceId, a.sourceName;
"@

$rawRows = Invoke-SqliteQuery -SqliteExe $sqliteExe -DataSource $DbPath -Sql $sql
$rxRun = [regex]::new('^\s*run\s+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
$items = [System.Collections.Generic.List[object]]::new()
foreach ($row in $rawRows) {
  $deviceId = if ($null -eq $row.deviceId) { '' } else { [string]$row.deviceId }
  $sourceName = if ($null -eq $row.sourceName) { '' } else { [string]$row.sourceName }
  if (-not $deviceId -or -not $sourceName) { continue }
  if ($TrimRunPrefix) { $sourceName = $rxRun.Replace($sourceName, '') }
  $activityResult = if ($null -eq $row.activityResult) { '' } else { [string]$row.activityResult }
  $status = if ($null -eq $row.status) { '' } else { [string]$row.status }
  $statusCode = if ($null -eq $row.statusCode) { '' } else { [string]$row.statusCode }
  $activityTime = if ($null -eq $row.activityTime) { '' } else { [string]$row.activityTime }
  $idval = if ($null -eq $row.id) { '' } else { [string]$row.id }
  $norm = Map-Status -ActivityResult $activityResult -Status $status -StatusCode $statusCode
  [void]$items.Add([pscustomobject]@{ deviceId = $deviceId; name = $sourceName; status = $norm; activityTime = $activityTime; id = $idval })
}
Write-Verbose ("Plain objects created: {0}" -f $items.Count)
if ($items.Count -eq 0) { Write-Host "Info: No COMPLETED script runs in the lookback period; nothing to sync. Ensure Get-AutomationActivities has run and that the DB has COMPLETED activities."; exit 0 }

# Group by device and build per-device maps
$deviceMaps = @{}
$groups = $items | Group-Object deviceId
Write-Verbose ("Devices grouped: {0}" -f $groups.Count)
foreach ($g in $groups) {
  $devId = [string]$g.Name
  $map = @{}
  foreach ($it in $g.Group) { $map[$it.name] = [string]$it.status }  # latest already chosen in SQL
  $deviceMaps[$devId] = $map
}

# Push to NinjaOne (your existing Compare-And-UpdateCustomField loop)
$processed = 0
foreach ($devId in $deviceMaps.Keys) {
  $map = $deviceMaps[$devId]
  if ($ForceRebuild -and $PreserveExisting) {
    try {
      $existing = Invoke-NinjaOneRequest -Method GET -Path ("device/{0}/custom-fields?fields={1}" -f $devId,$CustomFieldName) -Paginate:$false
      if ($existing -and $existing.customFields -and $existing.customFields.ContainsKey($CustomFieldName)) {
        $curr = [string]$existing.customFields[$CustomFieldName]
        if ($curr) {
          try {
            $obj = $curr | ConvertFrom-Json
            if ($obj -ne $null) {
              foreach ($p in $obj.PSObject.Properties) {
                if (-not $map.ContainsKey($p.Name)) { $map[$p.Name] = [string]$p.Value }
              }
            }
          } catch { }
        }
      }
    } catch { Write-Warning ("PreserveExisting read failed for device {0}: {1}" -f $devId, $_) }
  }
  Compare-And-UpdateCustomField -DeviceId $devId -FieldName $CustomFieldName -MapToWrite $map
  $processed++
}
Write-Host ("Completed. Devices processed: {0}" -f $processed)

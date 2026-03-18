<#
.SYNOPSIS
    Analyzes automation failures from the Activities SQLite DB using threshold-based clustering (TF-IDF + cosine similarity) and writes failure-clusters.json for use by Invoke-ScriptTracker.
.DESCRIPTION
    Queries the same database and time window as Invoke-ScriptTracker (COMPLETED + ACTION, Past 10 Days), filters to FAILURE rows only, builds combined text from message and data (verbose output), runs an embedded Python script that clusters failure messages per automation using TF-IDF and cosine similarity: two messages are grouped when their similarity is at or above SimilarityThreshold (no fixed cluster count). Output is written to BaseOutputFolder\TimeframeName\failure-clusters.json. Invoke-ScriptTracker.ps1 reads this file when present and adds "Failure message groupings" below the success/failure widget on each Automation Detail page. Requires Python with sklearn (scikit-learn). If Python or sklearn is missing, exits gracefully without overwriting any existing failure-clusters.json.
.PARAMETER DbPath
    Path to the SQLite database file. Defaults to C:\RMM\Activities.db (same as Get-AutomationActivities / Invoke-ScriptTracker).
.PARAMETER SqliteExePath
    Full path to sqlite3.exe. Same resolution as Invoke-ScriptTracker.
.PARAMETER BaseOutputFolder
    Top-level folder for report output. Defaults to C:\RMM\Reports\Script Tracking. Output is written to BaseOutputFolder\TimeframeName\failure-clusters.json.
.PARAMETER TimeframeDays
    Number of days to look back. Default 10. Must match Invoke-ScriptTracker's window so the same folder name is used.
.PARAMETER TimeframeName
    Name of the timeframe subfolder (e.g. "Past 10 Days"). Default "Past 10 Days". Must match Invoke-ScriptTracker.
.PARAMETER MaxTextLength
    Maximum combined text length per failure for clustering. Default 3000.
.PARAMETER SimilarityThreshold
    Minimum cosine similarity (0 to 1) for two failure messages to be grouped into the same cluster. Default 0.6. Lower values merge more; higher values keep clusters tighter.
.PARAMETER MaxClustersPerAutomation
    Deprecated. No longer used; clustering is threshold-based. Kept for backward compatibility.
.PARAMETER PythonPath
    Full path to python.exe. Use when Python is not on PATH in the process (e.g. when run by NinjaOne). If not set, the script tries Get-Command, then Machine PATH from registry, then common install locations.
.LINK
    Invoke-ScriptTracker.ps1
    https://www.sqlite.org/download.html
#>

[CmdletBinding()]
param (
    [Parameter()]
    [string]$DbPath = 'C:\RMM\Activities.db',
    [Parameter()]
    [string]$SqliteExePath = 'C:\ProgramData\chocolatey\bin\sqlite3.exe',
    [Parameter()]
    [string]$BaseOutputFolder = 'C:\RMM\Reports\Script Tracking',
    [Parameter()]
    [int]$TimeframeDays = 10,
    [Parameter()]
    [string]$TimeframeName = 'Past 10 Days',
    [Parameter()]
    [int]$MaxTextLength = 3000,
    [Parameter()]
    [ValidateRange(0.01, 1.0)]
    [double]$SimilarityThreshold = 0.6,
    [Parameter()]
    [int]$MaxClustersPerAutomation = 10,
    [Parameter()]
    [string]$PythonPath = 'c:\Program Files\Python314\python.exe'
)

$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrWhiteSpace($DbPath)) { $DbPath = 'C:\RMM\Activities.db' }
$dbFile = $DbPath
$baseOutputFolder = $BaseOutputFolder

# --- Resolve sqlite3.exe (same as Invoke-ScriptTracker) ---
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
    throw "sqlite3.exe not found. Place it in the script directory, add to PATH, or set -SqliteExePath."
}

if (-not (Test-Path -LiteralPath $dbFile -PathType Leaf)) {
    throw "SQLite database not found: $dbFile. Run Get-AutomationActivities.ps1 first."
}

# --- SQLite query helper (in-line, no dot-sourcing) ---
function Invoke-SqliteQuery {
    param([Parameter(Mandatory)] [string]$SqliteExe, [Parameter(Mandatory)] [string]$DataSource, [Parameter(Mandatory)] [string]$Sql)
    $out = & $SqliteExe -csv -header $DataSource $Sql 2>$null
    $text = if ($null -eq $out) { '' } elseif ($out -is [string]) { $out } else { $out -join "`n" }
    $text = $text.TrimStart([char]0xFEFF)
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    $lines = $text -split "`r?`n"
    $lines = $lines | Where-Object { $_.Length -gt 0 }
    if ($lines.Count -lt 2) { return @() }
    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText($tempFile, $text, $utf8NoBom)
        $result = Import-Csv -Path $tempFile -Encoding UTF8
        return @($result)
    } finally {
        if (Test-Path -LiteralPath $tempFile) { Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue }
    }
}

# --- Epoch conversion (same as Invoke-ScriptTracker) ---
$epoch0 = [datetime]'1970-01-01T00:00:00Z'
function Get-EpochSeconds { param([Parameter(Mandatory)][datetime]$Date) $utc = $Date.ToUniversalTime(); [int64][math]::Floor(($utc - $epoch0).TotalSeconds) }

# --- Extract readable text from activity data JSON ---
function Get-TextFromActivityData {
    param([string]$DataJson)
    if ([string]::IsNullOrWhiteSpace($DataJson)) { return '' }
    try {
        $obj = $DataJson | ConvertFrom-Json -ErrorAction Stop
        $parts = [System.Collections.Generic.List[string]]::new()
        $keys = @('output', 'stdout', 'stderr', 'message', 'error', 'result')
        foreach ($key in $keys) {
            if ($obj.PSObject.Properties[$key]) {
                $val = $obj.$key
                if ($null -ne $val) {
                    if ($val -is [string]) { [void]$parts.Add($val) }
                    elseif ($val -is [System.Collections.IEnumerable] -and $val -isnot [string]) {
                        foreach ($item in $val) { if ($null -ne $item -and [string]$item -ne '') { [void]$parts.Add([string]$item) } }
                    }
                    else { [void]$parts.Add([string]$val) }
                }
            }
        }
        if ($parts.Count -eq 0) { return $DataJson.Trim() }
        return ($parts -join " `n ").Trim()
    } catch {
        return $DataJson.Trim()
    }
}

# --- Build WHERE clause: COMPLETED + ACTION + time window + FAILURE ---
$statusCondition = "( (statusCode IS NOT NULL AND UPPER(TRIM(statusCode)) = 'COMPLETED') OR (status IS NOT NULL AND UPPER(TRIM(status)) = 'COMPLETED') )"
$typeCondition   = "( (activityType IS NOT NULL AND UPPER(TRIM(activityType)) = 'ACTION') OR (type IS NOT NULL AND UPPER(TRIM(type)) = 'ACTION') )"
$failureCondition = " AND ( activityResult IS NOT NULL AND UPPER(TRIM(activityResult)) = 'FAILURE' )"
$rangeStart = (Get-Date).AddDays(-$TimeframeDays)
$rangeEndEx = (Get-Date).AddSeconds(1)
$afterEpoch = Get-EpochSeconds -Date $rangeStart
$beforeEpochEx = Get-EpochSeconds -Date $rangeEndEx
$afterEpochMs = $afterEpoch * 1000
$beforeEpochExMs = $beforeEpochEx * 1000
$timeCondition = " AND ( (activityTime >= $afterEpoch AND activityTime < $beforeEpochEx) OR (activityTime >= $afterEpochMs AND activityTime < $beforeEpochExMs) )"
$whereClause = "WHERE $statusCondition AND $typeCondition $failureCondition $timeCondition"

$queryFailures = "SELECT id, sourceName, message, data, deviceId, activityTime FROM Activities $whereClause ORDER BY sourceName, id;"
$failureRows = Invoke-SqliteQuery -SqliteExe $sqliteExe -DataSource $dbFile -Sql $queryFailures

if (-not $failureRows -or $failureRows.Count -eq 0) {
    Write-Host "[Info] No FAILURE rows in the time window; skipping clustering."
    $outputFolder = Join-Path $baseOutputFolder $TimeframeName
    if (-not (Test-Path -LiteralPath $outputFolder -PathType Container)) { New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null }
    $outPath = Join-Path $outputFolder "failure-clusters.json"
    $emptyResult = @{ timeframeName = $TimeframeName; byAutomation = @{} } | ConvertTo-Json -Depth 10 -Compress
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($outPath, $emptyResult, $utf8NoBom)
    Write-Host "[Info] Wrote empty failure-clusters.json to $outPath"
    exit 0
}

# --- Build combined text per row and list for Python (include id, deviceId, activityTime for enrichment) ---
$maxLen = [Math]::Max(100, $MaxTextLength)
$records = [System.Collections.Generic.List[object]]::new()
foreach ($r in $failureRows) {
    $msg = if ($null -ne $r.message -and [string]$r.message -ne '') { [string]$r.message.Trim() } else { '' }
    $dataText = Get-TextFromActivityData -DataJson ([string]$r.data)
    $combined = if ([string]::IsNullOrWhiteSpace($dataText)) { $msg } else { "$msg $dataText" }
    $combined = $combined.Trim()
    if ([string]::IsNullOrWhiteSpace($combined)) { $combined = '(no message)' }
    if ($combined.Length -gt $maxLen) { $combined = $combined.Substring(0, $maxLen) }
    $sourceName = if ($null -ne $r.sourceName) { [string]$r.sourceName } else { 'Unknown' }
    $deviceId = if ($null -ne $r.deviceId -and [string]$r.deviceId -ne '') { [string]$r.deviceId } else { $null }
    $activityTime = if ($null -ne $r.activityTime -and [string]$r.activityTime -ne '') { [string]$r.activityTime } else { $null }
    $rowId = if ($null -ne $r.id -and [string]$r.id -ne '') { [string]$r.id } else { $null }
    [void]$records.Add([pscustomobject]@{ id = $rowId; sourceName = $sourceName; combinedText = $combined; deviceId = $deviceId; activityTime = $activityTime })
}

$inputJson = @($records) | ConvertTo-Json -Depth 5 -Compress
$tempInputPath = [System.IO.Path]::GetTempFileName()
$tempInputPath = $tempInputPath -replace '\.tmp$', '.json'
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($tempInputPath, $inputJson, $utf8NoBom)

# --- Embedded Python script: TF-IDF + cosine similarity threshold + connected components per automation ---
$pythonScript = @'
import json
import sys
import os
import re
import warnings

def _normalize_for_clustering(text):
    """Replace numbers and date/time patterns with placeholders so similar failures cluster together."""
    if not text or not text.strip():
        return text
    t = text
    # Strip trailing PowerShell metadata so it does not affect TF-IDF
    t = re.sub(r"\s*@\{[^}]*\}\s*$", "", t)
    t = re.sub(r"\s*@\{[^}]*\}\s*", " ", t)
    # Script paths (NinjaRMM scripting folder) so different gen IDs do not split same error
    t = re.sub(r"[A-Za-z]:\\.*?scripting\\.*?\.ps1", "<script_path>", t)
    # "X days ago" / "X.XX days ago" so numeric variation does not split clusters
    t = re.sub(r"\d+(?:\.\d+)?\s*days?\s+ago", "<num> days ago", t, flags=re.IGNORECASE)
    # "At path:line char:col" so different line/char numbers do not split
    t = re.sub(r"At\s+[^\n]+?:\d+\s+char:\d+", "At <path>:<line> char:<col>", t)
    # Process lock file path so only message type matters
    t = re.sub(r"Process lock file found at\s+'[^']*'", "Process lock file found at '<path>'", t)
    # Minutes until ... (e.g. "76 min until 03/04/2026 15:50:00")
    t = re.sub(r"\d{1,6}\s*min\s+until\s+[\d/:\s\-]+", "<minutes> min until <datetime>", t, flags=re.IGNORECASE)
    # Standalone epoch-like or large numbers (e.g. 10012, 40254)
    t = re.sub(r"\b\d{4,}\b", "<num>", t)
    # Dates like 03/04/2026, 04/01/2026
    t = re.sub(r"\d{1,2}/\d{1,2}/\d{2,4}", "<date>", t)
    # Times like 15:50:00, 13:05:00
    t = re.sub(r"\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?", "<time>", t, flags=re.IGNORECASE)
    # Version-like tokens (e.g. 10.0.26100.1150) so different builds do not split
    t = re.sub(r"\b\d+\.\d+\.\d+(?:\.\d+)?\b", "<version>", t)
    # "Error: 87" style so same Windows error code clusters
    t = re.sub(r"Error:\s*\d+", "Error: <code>", t, flags=re.IGNORECASE)
    # "ExitCode: N" so same exit code clusters (e.g. Local Users Report)
    t = re.sub(r"ExitCode:\s*\d+", "ExitCode: <n>", t, flags=re.IGNORECASE)
    # "Duplicate article name '...' (from ...)" so only message type matters
    t = re.sub(r"Duplicate article name\s+'[^']*'\s+\(from\s+[^)]+\)", "Duplicate article name '<name>' (from <path>)", t)
    return t

def _cluster_label(sample_message, max_len=100):
    """Produce a short human-readable label from the representative message."""
    if not sample_message or not sample_message.strip():
        return "No message"
    s = sample_message.strip()
    # Strip PowerShell metadata so label never shows @{code=...}
    s = re.sub(r"\s*@\{[^}]*\}\s*$", "", s)
    s = re.sub(r"\s*@\{[^}]*\}\s*", " ", s)
    s = s.strip()
    # Strip or shorten script paths so label shows the error, not the path
    s = re.sub(r"[A-Za-z]:\\.*?scripting\\.*?\.ps1", "...\\\\script.ps1", s)
    # Drop common prefixes to get to the actual error
    for prefix in ("Action completed:", "Result: FAILURE", "Output:", "Result: Failed", "Action:", "Failed\n"):
        if prefix in s:
            s = s.split(prefix, 1)[-1].strip()
    # Prefer first meaningful error-like line (across lines)
    lines = [ln.strip() for ln in s.split("\n") if ln.strip()]
    error_markers = ("[Error]", ".ps1 :", "Write-Error", "required for", "required.", "Exiting.", "not compatible", "not designed", "throw ")
    chosen = ""
    for line in lines:
        if any(m in line for m in error_markers) or (line and not any(line.startswith(p) for p in ("Result:", "Output:", "Action:"))):
            # Take first sentence of this line or full line up to max_len
            for sep in (". ", "\n", " At ", " + "):
                if sep in line:
                    line = line.split(sep)[0].strip()
                    break
            if len(line) > 15:
                chosen = line
                break
    if not chosen and lines:
        # Fallback: first line that is long enough
        for line in lines:
            if len(line) > 15:
                for sep in (". ", " At ", " + "):
                    if sep in line:
                        line = line.split(sep)[0].strip()
                        break
                chosen = line
                break
    if not chosen and s:
        # Final fallback: first 80 chars of remainder after strips
        chosen = s[:80].strip()
        for sep in (". ", "\n"):
            if sep in chosen:
                chosen = chosen.split(sep)[0].strip()
    s = chosen if chosen else s
    # First sentence or up to max_len
    for sep in (". ", "\n", " At ", " + "):
        if sep in s:
            s = s.split(sep)[0].strip()
    if len(s) > max_len:
        s = s[:max_len].rstrip() + "..."
    return s if s else "No message"

def main():
    warnings.filterwarnings("ignore", category=UserWarning, module="sklearn")
    try:
        from sklearn.exceptions import ConvergenceWarning
        warnings.filterwarnings("ignore", category=ConvergenceWarning)
    except Exception:
        pass

    input_path = (sys.argv[1] if len(sys.argv) > 1 else None) or os.environ.get("FAILURE_CLUSTER_INPUT")
    if not input_path or not os.path.isfile(input_path):
        print(json.dumps({"error": "Missing input file"}), file=sys.stderr)
        sys.exit(1)
    with open(input_path, "r", encoding="utf-8") as f:
        records = json.load(f)
    if not records:
        print(json.dumps({"timeframeName": os.environ.get("FAILURE_TIMEFRAME", "Past 10 Days"), "byAutomation": {}}))
        return

    try:
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics.pairwise import cosine_similarity
        import numpy as np
    except ImportError as e:
        print(json.dumps({"error": "sklearn required: " + str(e)}), file=sys.stderr)
        sys.exit(1)

    timeframe_name = os.environ.get("FAILURE_TIMEFRAME", "Past 10 Days")
    try:
        similarity_threshold = float(os.environ.get("FAILURE_SIMILARITY_THRESHOLD", "0.6"))
    except (TypeError, ValueError):
        similarity_threshold = 0.6
    similarity_threshold = max(0.01, min(1.0, similarity_threshold))
    max_sample_device_ids = 10

    # Group by automation; keep full records (id, deviceId, activityTime, combinedText) for enrichment
    by_automation = {}
    for rec in records:
        name = rec.get("sourceName") or "Unknown"
        if name not in by_automation:
            by_automation[name] = []
        by_automation[name].append(rec)

    result = {"timeframeName": timeframe_name, "byAutomation": {}}

    for automation_name, recs in by_automation.items():
        # Original text for output; normalized text for clustering
        texts_orig = []
        texts_norm = []
        for r in recs:
            t = (r.get("combinedText") or "").strip() or "(no message)"
            texts_orig.append(t)
            texts_norm.append(_normalize_for_clustering(t))

        if not texts_orig:
            continue
        if len(texts_orig) == 1:
            rec = recs[0]
            label = _cluster_label(texts_orig[0][:500])
            d = rec.get("deviceId")
            device_ids = [str(d)] if d is not None and str(d).strip() else []
            at = rec.get("activityTime")
            first_seen = last_seen = None
            if at is not None:
                try:
                    ts = float(at)
                    if ts > 1e12:
                        ts = ts / 1000.0
                    from datetime import datetime, timezone
                    dt = datetime.fromtimestamp(ts, tz=timezone.utc)
                    iso = dt.strftime("%Y-%m-%dT%H:%M:%SZ")
                    first_seen = last_seen = iso
                except Exception:
                    pass
            out_obj = {
                "clusterId": 0,
                "count": 1,
                "label": label,
                "sampleMessage": texts_orig[0][:500],
                "topMessages": [texts_orig[0][:300]],
                "affectedDeviceCount": len(set(device_ids)),
                "sampleDeviceIds": device_ids[:max_sample_device_ids]
            }
            if first_seen:
                out_obj["firstSeen"] = first_seen
            if last_seen:
                out_obj["lastSeen"] = last_seen
            result["byAutomation"][automation_name] = [out_obj]
            continue

        n = len(texts_orig)
        vectorizer = TfidfVectorizer(max_features=5000, stop_words="english", min_df=1, max_df=0.95, sublinear_tf=True)
        try:
            X = vectorizer.fit_transform(texts_norm)
        except Exception:
            rec = recs[0]
            label = _cluster_label(texts_orig[0][:500])
            device_ids = list(dict.fromkeys(str(r.get("deviceId")) for r in recs if r.get("deviceId") is not None))[:max_sample_device_ids]
            ats = [r.get("activityTime") for r in recs if r.get("activityTime") is not None]
            first_seen = last_seen = None
            if ats:
                try:
                    from datetime import datetime, timezone
                    vals = []
                    for at in ats:
                        ts = float(at)
                        if ts > 1e12:
                            ts = ts / 1000.0
                        vals.append(ts)
                    first_seen = datetime.fromtimestamp(min(vals), tz=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                    last_seen = datetime.fromtimestamp(max(vals), tz=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    pass
            out_obj = {
                "clusterId": 0,
                "count": n,
                "label": label,
                "sampleMessage": texts_orig[0][:500],
                "topMessages": [t[:300] for t in texts_orig[:5]],
                "affectedDeviceCount": len(set(str(r.get("deviceId")) for r in recs if r.get("deviceId") is not None)),
                "sampleDeviceIds": device_ids
            }
            if first_seen:
                out_obj["firstSeen"] = first_seen
            if last_seen:
                out_obj["lastSeen"] = last_seen
            result["byAutomation"][automation_name] = [out_obj]
            continue

        sim = cosine_similarity(X, X)
        # Union-find: merge indices i, j when similarity >= threshold
        parent = list(range(n))
        def find(x):
            if parent[x] != x:
                parent[x] = find(parent[x])
            return parent[x]
        def union(x, y):
            px, py = find(x), find(y)
            if px != py:
                parent[px] = py
        for i in range(n):
            for j in range(i + 1, n):
                if sim[i, j] >= similarity_threshold:
                    union(i, j)
        # Group indices by root (use min index in component as stable id for ordering)
        root_to_indices = {}
        for i in range(n):
            r = find(i)
            if r not in root_to_indices:
                root_to_indices[r] = []
            root_to_indices[r].append(i)
        # Sort components by minimum index so clusterId order is deterministic
        sorted_roots = sorted(root_to_indices.keys(), key=lambda r: min(root_to_indices[r]))

        out_list = []
        for cid, root in enumerate(sorted_roots):
            indices = root_to_indices[root]
            ctexts_orig = [texts_orig[i] for i in indices]
            crecs = [recs[i] for i in indices]
            rep = ctexts_orig[0]
            label = _cluster_label(rep[:500])
            device_ids = list(dict.fromkeys(str(r.get("deviceId")) for r in crecs if r.get("deviceId") is not None))[:max_sample_device_ids]
            ats = [r.get("activityTime") for r in crecs if r.get("activityTime") is not None]
            first_seen = last_seen = None
            if ats:
                try:
                    from datetime import datetime, timezone
                    vals = []
                    for at in ats:
                        ts = float(at)
                        if ts > 1e12:
                            ts = ts / 1000.0
                        vals.append(ts)
                    first_seen = datetime.fromtimestamp(min(vals), tz=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                    last_seen = datetime.fromtimestamp(max(vals), tz=timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
                except Exception:
                    pass
            if len(rep) > 500:
                rep = rep[:500] + "..."
            top = list(dict.fromkeys(ctexts_orig))[:5]
            top = [t[:300] + ("..." if len(t) > 300 else "") for t in top]
            out_obj = {
                "clusterId": int(cid),
                "count": len(ctexts_orig),
                "label": label,
                "sampleMessage": rep,
                "topMessages": top,
                "affectedDeviceCount": len(set(str(r.get("deviceId")) for r in crecs if r.get("deviceId") is not None)),
                "sampleDeviceIds": device_ids
            }
            if first_seen:
                out_obj["firstSeen"] = first_seen
            if last_seen:
                out_obj["lastSeen"] = last_seen
            out_list.append(out_obj)
        result["byAutomation"][automation_name] = out_list

    print(json.dumps(result))

if __name__ == "__main__":
    main()
'@

$tempPyPath = [System.IO.Path]::GetTempFileName()
$tempPyPath = $tempPyPath -replace '\.tmp$', '.py'
[System.IO.File]::WriteAllText($tempPyPath, $pythonScript, [System.Text.UTF8Encoding]::new($false))

$outputFolder = Join-Path $baseOutputFolder $TimeframeName
if (-not (Test-Path -LiteralPath $outputFolder -PathType Container)) {
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
}
$outPath = Join-Path $outputFolder "failure-clusters.json"

$env:FAILURE_CLUSTER_INPUT = $tempInputPath
$env:FAILURE_TIMEFRAME = $TimeframeName
$env:FAILURE_SIMILARITY_THRESHOLD = [string]$SimilarityThreshold

try {
    $pyExe = $null
    if (-not [string]::IsNullOrWhiteSpace($PythonPath) -and (Test-Path -LiteralPath $PythonPath -PathType Leaf)) {
        $pyExe = $PythonPath
    }
    if (-not $pyExe) {
        foreach ($exe in @('python', 'python3')) {
            $cmd = Get-Command $exe -ErrorAction SilentlyContinue
            if ($cmd -and $cmd.Source) { $pyExe = $cmd.Source; break }
        }
    }
    if (-not $pyExe) {
        $regEnv = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment' -Name Path -ErrorAction SilentlyContinue
        $machinePath = if ($regEnv -and $regEnv.PSObject.Properties['Path']) { $regEnv.Path } else { $null }
        if ($machinePath) {
            $dirs = $machinePath -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            foreach ($dir in $dirs) {
                $candidate = Join-Path -Path $dir -ChildPath 'python.exe'
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { $pyExe = $candidate; break }
            }
        }
    }
    if (-not $pyExe) {
        $commonRoots = @('C:\Program Files\Python*', 'C:\Python*')
        if ($env:ProgramFiles) { $commonRoots += Join-Path $env:ProgramFiles 'Python*' }
        $pf86 = [Environment]::GetFolderPath('ProgramFilesX86')
        if ($pf86) { $commonRoots += Join-Path $pf86 'Python*' }
        foreach ($root in $commonRoots) {
            if (-not $root) { continue }
            $dirs = Get-Item -Path $root -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }
            foreach ($d in $dirs) {
                $candidate = Join-Path -Path $d.FullName -ChildPath 'python.exe'
                if (Test-Path -LiteralPath $candidate -PathType Leaf) { $pyExe = $candidate; break }
            }
            if ($pyExe) { break }
        }
    }
    if (-not $pyExe) {
        Write-Warning "Python not found on PATH. Skipping failure clustering; failure-clusters.json will not be updated."
        exit 0
    }
    Write-Verbose "Using Python: $pyExe"
    Write-Verbose "Temp input JSON: $tempInputPath"
    Write-Verbose "Temp Python script: $tempPyPath"
    # Capture stderr to a temp file so we can show it on failure; stdout stays clean for JSON when exit code is 0
    $stderrCapture = [System.IO.Path]::GetTempFileName()
    try {
        $pyResult = & $pyExe $tempPyPath $tempInputPath 2>$stderrCapture
        $stderrContent = $null
        if (Test-Path -LiteralPath $stderrCapture -PathType Leaf) {
            $stderrContent = Get-Content -LiteralPath $stderrCapture -Raw -ErrorAction SilentlyContinue
        }
        $jsonOut = if ($pyResult -is [string]) { $pyResult } else { $pyResult -join "`n" }
        $jsonOut = $jsonOut.Trim()
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Python script exited with code $LASTEXITCODE. Skipping write of failure-clusters.json."
            if (-not [string]::IsNullOrWhiteSpace($stderrContent)) {
                Write-Host "Python stderr:"
                Write-Host $stderrContent.Trim()
            }
            if (-not [string]::IsNullOrWhiteSpace($jsonOut)) {
                Write-Host "Python stdout:"
                Write-Host $jsonOut
            }
            exit 0
        }
    } finally {
        if (Test-Path -LiteralPath $stderrCapture -PathType Leaf) { Remove-Item -LiteralPath $stderrCapture -Force -ErrorAction SilentlyContinue }
    }
    if ([string]::IsNullOrWhiteSpace($jsonOut)) {
        Write-Warning "Python script produced no output. Skipping write of failure-clusters.json."
        if (-not [string]::IsNullOrWhiteSpace($stderrContent)) {
            Write-Host "Python stderr:"
            Write-Host $stderrContent.Trim()
        }
        exit 0
    }
    if ($jsonOut -match '^\s*\{\s*"error"') {
        $errorMsg = $jsonOut
        try {
            $errObj = $jsonOut | ConvertFrom-Json -ErrorAction Stop
            if ($errObj.PSObject.Properties['error']) { $errorMsg = $errObj.error }
        } catch { }
        Write-Warning "Python script reported error. Skipping write of failure-clusters.json. Error: $errorMsg"
        if ($errorMsg -ne $jsonOut) { Write-Host "Full Python output: $jsonOut" }
        exit 0
    }
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($outPath, $jsonOut, $utf8NoBom)
    Write-Host "[Info] Wrote failure-clusters.json to $outPath"
} finally {
    $env:FAILURE_CLUSTER_INPUT = $null
    $env:FAILURE_TIMEFRAME = $null
    $env:FAILURE_SIMILARITY_THRESHOLD = $null
    if (Test-Path -LiteralPath $tempInputPath -PathType Leaf) { Remove-Item -LiteralPath $tempInputPath -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $tempPyPath -PathType Leaf) { Remove-Item -LiteralPath $tempPyPath -Force -ErrorAction SilentlyContinue }
}

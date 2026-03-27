# =============================================================================
# Patch Cache Health Monitor for NinjaOne
# Version: 2.0.0 (Custom Field Output)
#
# Builds a patch cache utilization report and writes the HTML
# directly to a Ninja custom field using Ninja-Property-Set-Piped.
#
# Prerequisites:
#   - Device custom field: patchCacheHtml
#   - Run on the CACHE SERVER device via scheduled automation
# =============================================================================

# --- Parameters ---
param(
    [Parameter(Mandatory = $false)]
    [ValidateRange(1,365)]
    [int]$LastNDays = 7,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1,1000)]
    [int]$MaxLogFiles = 50
)

# --- Configuration ---
$CacheFolderDefault = 'C:\ProgramData\cache'
$CustomFieldName = 'patchCacheHtml'

#region ================================================================
#  LOGGING
#================================================================

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warning','Error','Debug','Progress')][string]$Level = 'Info'
    )
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$ts] [$Level] $Message"
}

#endregion

#region ================================================================
#  CACHE LISTENER LOG PARSING
#================================================================

$ErrorActionPreference = 'Stop'
$startTime = Get-Date

Write-Log '=== Patch Cache Health Monitor v2.0 (Custom Field Output) ===' 'Info'

$cacheProc = Get-Process -Name 'CacheListener' -ErrorAction SilentlyContinue
$reportTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
$html = [System.Collections.ArrayList]::new()

$borderColor = '#336699'
$headerStyle = "border-bottom:2px solid $borderColor;padding:8px;"
$labelStyle = "padding:5px 8px;border-top:1px solid #cccccc;color:#555555;width:180px;"
$valueStyle = "padding:5px 8px;border-top:1px solid #cccccc;"
$sectionStyle = "padding:8px;border-top:2px solid $borderColor;color:$borderColor;"

if (-not $cacheProc) {
    Write-Log 'CacheListener.exe is NOT running' 'Warning'
    [void]$html.Add("<table style='width:100%;border-collapse:collapse;font-family:monospace;font-size:13px;'>")
    [void]$html.Add("<tr><td style='$headerStyle' colspan='2'><b style='color:$borderColor;'>PATCH CACHE REPORT</b> <span style='color:#999999;'>| $reportTime</span></td></tr>")
    [void]$html.Add("<tr><td style='$labelStyle'>Hostname</td><td style='$valueStyle'>$($env:COMPUTERNAME)</td></tr>")
    [void]$html.Add("<tr><td style='$labelStyle'>Service</td><td style='$valueStyle'><b style='color:#cc0000;'>CacheListener.exe Not Running</b></td></tr>")
    [void]$html.Add("</table>")
    $skipParse = $true
} else {
    $skipParse = $false
    Write-Log "CacheListener.exe running - PID: $($cacheProc.Id)" 'Info'
}

if (-not $skipParse) {
    $searchPaths = [System.Collections.ArrayList]::new()
    $procPath = $cacheProc.Path
    if ($procPath) { [void]$searchPaths.Add((Split-Path $procPath -Parent)) }
    [void]$searchPaths.Add('C:\ProgramData\NinjaRMMAgent\logs\CacheListener')
    [void]$searchPaths.Add('C:\ProgramData\NinjaRMMAgent\logs')
    [void]$searchPaths.Add('C:\ProgramData\NinjaRMMAgent')
    [void]$searchPaths.Add($CacheFolderDefault)
    [void]$searchPaths.Add('C:\ProgramData\NinjaRMMAgent\cache')

    $logFiles = [System.Collections.ArrayList]::new()
    foreach ($sp in $searchPaths) {
        if (Test-Path $sp) {
            $found = Get-ChildItem -Path $sp -Filter 'CacheListener_*.log' -Recurse -ErrorAction SilentlyContinue
            foreach ($f in $found) { [void]$logFiles.Add($f) }
        }
    }

    $uniqueLogs = $logFiles | Sort-Object FullName -Unique
    $cutoffTime = (Get-Date).AddDays(-$LastNDays)

    function Get-CacheListenerLogTimestamp {
        param(
            [Parameter(Mandatory = $true)]
            [string]$LogFileName
        )

        # Expected name: CacheListener_YYYYMMDDTHH.MM.SS.MMM.log
        # Returns:
        # - StartOfDay: DateTime at 00:00:00 local time
        # - FileTimestamp: DateTime including time portion
        $startOfDay = $null
        $fileTs = $null

        if ($LogFileName -match '^CacheListener_(\d{4})(\d{2})(\d{2})T(\d{2})\.(\d{2})\.(\d{2})\.(\d{3})\.log$') {
            $year = [int]$Matches[1]
            $month = [int]$Matches[2]
            $day = [int]$Matches[3]
            $hour = [int]$Matches[4]
            $minute = [int]$Matches[5]
            $second = [int]$Matches[6]
            $ms = [int]$Matches[7]

            try {
                $startOfDay = Get-Date -Year $year -Month $month -Day $day -Hour 0 -Minute 0 -Second 0
                $fileTs = Get-Date -Year $year -Month $month -Day $day -Hour $hour -Minute $minute -Second $second -Millisecond $ms
            } catch { }
        }

        [pscustomobject]@{
            StartOfDay    = $startOfDay
            FileTimestamp = $fileTs
        }
    }

    $selectedLogs = @()
    foreach ($l in $uniqueLogs) {
        $parsed = Get-CacheListenerLogTimestamp -LogFileName $l.Name
        $fileTs = $parsed.FileTimestamp
        $startOfDay = $parsed.StartOfDay

        if (-not $fileTs) { $fileTs = $l.LastWriteTime }
        if (-not $startOfDay) { $startOfDay = $fileTs.Date }

        if ($fileTs -ge $cutoffTime) {
            $selectedLogs += [pscustomobject]@{
                FullName      = $l.FullName
                Name          = $l.Name
                FileTimestamp = $fileTs
                StartOfDay    = $startOfDay
            }
        }
    }

    $selectedLogs = $selectedLogs | Sort-Object FileTimestamp
    if ($MaxLogFiles -and $selectedLogs.Count -gt $MaxLogFiles) {
        $selectedLogs = $selectedLogs |
            Sort-Object FileTimestamp -Descending |
            Select-Object -First $MaxLogFiles |
            Sort-Object FileTimestamp
    }

    $latestLog = if ($selectedLogs -and $selectedLogs.Count -gt 0) { $selectedLogs[-1] } else { $null }

    if (-not $latestLog) {
        Write-Log 'No CacheListener log files found' 'Warning'
        [void]$html.Add("<table style='width:100%;border-collapse:collapse;font-family:monospace;font-size:13px;'>")
        [void]$html.Add("<tr><td style='$headerStyle' colspan='2'><b style='color:$borderColor;'>PATCH CACHE REPORT</b> <span style='color:#999999;'>| $reportTime</span></td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Hostname</td><td style='$valueStyle'>$($env:COMPUTERNAME)</td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Service</td><td style='$valueStyle'>Running (PID $($cacheProc.Id))</td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Log Files</td><td style='$valueStyle'>No CacheListener logs found</td></tr>")
        [void]$html.Add("</table>")
    } else {
        Write-Log "Parsing $($selectedLogs.Count) logs (last $LastNDays day(s))" 'Info'

        $transfers = [System.Collections.ArrayList]::new()
        $clientIPs = [System.Collections.Generic.HashSet[string]]::new()
        $errors = [System.Collections.ArrayList]::new()
        $heartbeatCount = 0
        $lastHeartbeat = $null
        $totalBytesServed = [long]0
        $activeDownloads = @{}

        foreach ($logMeta in $selectedLogs) {
            Write-Log "  Parsing log: $($logMeta.FullName)" 'Debug'
            $logContent = Get-Content $logMeta.FullName -ErrorAction Stop
            $logDate = $logMeta.StartOfDay

            foreach ($line in $logContent) {
                if ($line -match '^(\d{2}:\d{2}:\d{2}\.\d{3})') {
                    $timeStr = $Matches[1]
                    $lineTime = [datetime]::MinValue
                    try {
                        $lineTime = [datetime]::ParseExact($timeStr, 'HH:mm:ss.fff', $null)
                        $lineTime = $logDate.Add($lineTime.TimeOfDay)
                    } catch { continue }
                    if ($lineTime -lt $cutoffTime) { continue }
                } else { continue }

                if ($line -match 'Heartbeat request received') {
                    $heartbeatCount++
                    if (-not $lastHeartbeat -or $lineTime -gt $lastHeartbeat) { $lastHeartbeat = $lineTime }
                    continue
                }

                if ($line -match 'Download request received by\s+([\d\.]+)\.\s+Url=(.+?),\s+Download=') {
                    $clientIP = $Matches[1]
                    $requestedUrl = $Matches[2]
                    [void]$clientIPs.Add($clientIP)
                    if ($line -match '\[Session\s+(\d+)\]') {
                        $sessId = $Matches[1]
                        $activeDownloads[$sessId] = @{ IP = $clientIP; URL = $requestedUrl; StartTime = $lineTime; Type = 'Unknown' }
                    }
                    continue
                }

                if ($line -match '\[Session\s+(\d+)\]\s+Transferring existing file:') {
                    $sessId = $Matches[1]
                    if ($activeDownloads.ContainsKey($sessId)) { $activeDownloads[$sessId]['Type'] = 'CacheHit' }
                    continue
                }

                if ($line -match '\[Session\s+(\d+)\]\s+Downloading file:') {
                    $sessId = $Matches[1]
                    if ($activeDownloads.ContainsKey($sessId)) { $activeDownloads[$sessId]['Type'] = 'CacheMiss' }
                    continue
                }

                if ($line -match '\[Session\s+(\d+)\]\s+Transferring existing file while downloading:') {
                    $sessId = $Matches[1]
                    if ($activeDownloads.ContainsKey($sessId)) { $activeDownloads[$sessId]['Type'] = 'PartialCache' }
                    continue
                }

                if ($line -match '\[Session\s+(\d+)\]\s+Finish transfer\s+''(.+?)''\s+\[(\S+)\s+KB,\s+(\d+)\s+ms\s+\(([\d\.]+)\s+KB/s\)\]') {
                    $sessId = $Matches[1]
                    $fileUrl = $Matches[2]
                    $fileSizeKB = [double]$Matches[3]
                    $durationMs = [int]$Matches[4]
                    $speedKBs = [double]$Matches[5]
                    $fileName = $fileUrl.Split('/')[-1].Split('?')[0]
                    $clientIP = 'Unknown'
                    $transferType = 'CacheHit'
                    if ($activeDownloads.ContainsKey($sessId)) {
                        $clientIP = $activeDownloads[$sessId]['IP']
                        if ($activeDownloads[$sessId]['Type'] -ne 'Unknown') { $transferType = $activeDownloads[$sessId]['Type'] }
                    }
                    [void]$transfers.Add(@{
                        Session = $sessId; Time = $lineTime; ClientIP = $clientIP; FileName = $fileName
                        SizeKB = $fileSizeKB; DurationSec = [math]::Round($durationMs / 1000, 1)
                        SpeedKBs = [math]::Round($speedKBs, 1); Type = $transferType; FullURL = $fileUrl
                    })
                    $totalBytesServed += [long]($fileSizeKB * 1024)
                    continue
                }

                if ($line -match '\[Session\s+(\d+)\]\s+Finish Download') { continue }

                if ($line -match '^\d{2}:\d{2}:\d{2}\.\d{3}\s+[EW]\s+') {
                    if ($line -notmatch 'Ignoring SHA|timeout_in_secs|config\.h|listener_config') {
                        [void]$errors.Add("[$($lineTime.ToString('MM/dd HH:mm:ss'))] $($line.Trim())")
                    }
                    continue
                }
            }
        }

        # Cache folder size
        $cacheSizeGB = 'N/A'
        if (Test-Path $CacheFolderDefault) {
            try {
                $folderBytes = (Get-ChildItem $CacheFolderDefault -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                if ($folderBytes) { $cacheSizeGB = "$([math]::Round($folderBytes / 1GB, 2)) GB" }
                else { $cacheSizeGB = '0 GB' }
            } catch { $cacheSizeGB = 'Error' }
        }

        # Computed values
        $totalServedMB = [math]::Round($totalBytesServed / 1MB, 1)
        $transferCount = $transfers.Count
        $uniqueClientCount = $clientIPs.Count
        $errorCount = $errors.Count
        $cacheHits = @($transfers | Where-Object { $_.Type -eq 'CacheHit' }).Count
        $cacheMisses = @($transfers | Where-Object { $_.Type -eq 'CacheMiss' }).Count
        $partials = @($transfers | Where-Object { $_.Type -eq 'PartialCache' }).Count
        $hitRate = if ($transferCount -gt 0) { "$([math]::Round(($cacheHits / $transferCount) * 100, 0))%" } else { 'N/A' }
        $lastHB = if ($lastHeartbeat) { $lastHeartbeat.ToString('MM/dd HH:mm') } else { 'None' }

        # Build HTML dashboard
        [void]$html.Add("<table style='width:100%;border-collapse:collapse;font-family:monospace;font-size:13px;'>")
        [void]$html.Add("<tr><td style='$headerStyle' colspan='2'><b style='color:$borderColor;'>PATCH CACHE REPORT</b> <span style='color:#999999;'>| Last $LastNDays day(s) | $reportTime</span></td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Hostname</td><td style='$valueStyle'>$($env:COMPUTERNAME)</td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Service</td><td style='$valueStyle'>Running (PID $($cacheProc.Id))</td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Heartbeats</td><td style='$valueStyle'>$heartbeatCount (last: $lastHB)</td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Logs Parsed</td><td style='$valueStyle;font-size:11px;'>$($selectedLogs.Count) (latest: $($latestLog.FullName))</td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Cache Folder</td><td style='$valueStyle'>$cacheSizeGB</td></tr>")

        # Transfers section
        [void]$html.Add("<tr><td style='$sectionStyle' colspan='2'><b>TRANSFERS</b></td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Total Transfers</td><td style='$valueStyle'><b>$transferCount</b></td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Data Served</td><td style='$valueStyle'><b>$totalServedMB MB</b></td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Cache Hits</td><td style='$valueStyle'>$cacheHits</td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Cache Misses</td><td style='$valueStyle'>$cacheMisses</td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Partial (mid-DL)</td><td style='$valueStyle'>$partials</td></tr>")
        [void]$html.Add("<tr><td style='$labelStyle'>Hit Rate</td><td style='$valueStyle'>$hitRate</td></tr>")

        $clientList = if ($uniqueClientCount -gt 0) { ($clientIPs -join ', ') } else { 'None' }
        [void]$html.Add("<tr><td style='$labelStyle'>Unique Clients</td><td style='$valueStyle'>$uniqueClientCount -- $clientList</td></tr>")

        if ($transferCount -gt 0) {
            $lastXfer = $transfers | Sort-Object { $_.Time } | Select-Object -Last 1
            $sizeMB = [math]::Round($lastXfer.SizeKB / 1024, 1)
            $lastStr = "$($lastXfer.Time.ToString('MM/dd HH:mm')) | $($lastXfer.FileName) | ${sizeMB}MB @ $($lastXfer.SpeedKBs) KB/s | $($lastXfer.ClientIP)"
            [void]$html.Add("<tr><td style='$labelStyle'>Last Transfer</td><td style='$valueStyle;font-size:11px;'>$lastStr</td></tr>")
        } else {
            [void]$html.Add("<tr><td style='$labelStyle'>Last Transfer</td><td style='$valueStyle'>None in last $LastNDays day(s)</td></tr>")
        }

        # Recent transfers table
        if ($transferCount -gt 0) {
            [void]$html.Add("<tr><td style='$sectionStyle' colspan='2'><b>RECENT TRANSFERS</b> (last 25)</td></tr>")
            [void]$html.Add("<tr><td colspan='2' style='padding:0;'>")
            [void]$html.Add("<table style='width:100%;border-collapse:collapse;font-family:monospace;font-size:11px;'>")
            [void]$html.Add("<tr>")
            [void]$html.Add("<td style='padding:4px;border-bottom:1px solid $borderColor;color:$borderColor;'><b>Time</b></td>")
            [void]$html.Add("<td style='padding:4px;border-bottom:1px solid $borderColor;color:$borderColor;'><b>Type</b></td>")
            [void]$html.Add("<td style='padding:4px;border-bottom:1px solid $borderColor;color:$borderColor;'><b>Client</b></td>")
            [void]$html.Add("<td style='padding:4px;border-bottom:1px solid $borderColor;color:$borderColor;'><b>File</b></td>")
            [void]$html.Add("<td style='padding:4px;border-bottom:1px solid $borderColor;color:$borderColor;text-align:right;'><b>Size</b></td>")
            [void]$html.Add("<td style='padding:4px;border-bottom:1px solid $borderColor;color:$borderColor;text-align:right;'><b>Speed</b></td>")
            [void]$html.Add("</tr>")

            $recentTransfers = $transfers | Sort-Object { $_.Time } | Select-Object -Last 25
            foreach ($t in $recentTransfers) {
                $sizeMB = [math]::Round($t.SizeKB / 1024, 1)
                $typeColor = switch ($t.Type) {
                    'CacheHit' { '#228B22' }
                    'CacheMiss' { '#cc6600' }
                    'PartialCache' { '#336699' }
                    default { '#555555' }
                }
                $rowBorder = "border-top:1px solid #eeeeee;"
                [void]$html.Add("<tr>")
                [void]$html.Add("<td style='padding:3px 4px;$rowBorder'>$($t.Time.ToString('MM/dd HH:mm'))</td>")
                [void]$html.Add("<td style='padding:3px 4px;$rowBorder'><b style='color:$typeColor;'>$($t.Type)</b></td>")
                [void]$html.Add("<td style='padding:3px 4px;$rowBorder'>$($t.ClientIP)</td>")
                $shortName = $t.FileName
                if ($shortName.Length -gt 40) { $shortName = $shortName.Substring(0, 37) + '...' }
                [void]$html.Add("<td style='padding:3px 4px;$rowBorder;font-size:10px;'>$shortName</td>")
                [void]$html.Add("<td style='padding:3px 4px;$rowBorder;text-align:right;'>${sizeMB}MB</td>")
                [void]$html.Add("<td style='padding:3px 4px;$rowBorder;text-align:right;'>$($t.SpeedKBs) KB/s</td>")
                [void]$html.Add("</tr>")
            }
            [void]$html.Add("</table>")
            [void]$html.Add("</td></tr>")
        }

        # Errors section
        [void]$html.Add("<tr><td style='$sectionStyle' colspan='2'><b>ERRORS/WARNINGS</b> ($errorCount)</td></tr>")
        if ($errorCount -gt 0) {
            $displayErrors = $errors | Select-Object -Last 15
            foreach ($e in $displayErrors) {
                $safeText = $e.Replace('<', '&lt;').Replace('>', '&gt;')
                [void]$html.Add("<tr><td colspan='2' style='padding:3px 8px;border-top:1px solid #eeeeee;font-size:11px;color:#cc0000;'>$safeText</td></tr>")
            }
            if ($errorCount -gt 15) {
                [void]$html.Add("<tr><td colspan='2' style='padding:3px 8px;font-size:11px;color:#999999;'>...and $($errorCount - 15) more</td></tr>")
            }
        } else {
            [void]$html.Add("<tr><td colspan='2' style='padding:5px 8px;color:#228B22;'>None</td></tr>")
        }

        [void]$html.Add("</table>")

        # Console summary
        Write-Log '============================================' 'Info'
        Write-Log "PATCH CACHE REPORT (Last $LastNDays day(s))" 'Info'
        Write-Log '============================================' 'Info'
        Write-Log "Logs Parsed:     $($selectedLogs.Count)" 'Info'
        Write-Log "Service:          Running (PID $($cacheProc.Id))" 'Info'
        Write-Log "Heartbeats:       $heartbeatCount (last: $lastHB)" 'Info'
        Write-Log "Transfers:        $transferCount (Hits:$cacheHits Miss:$cacheMisses Partial:$partials)" 'Info'
        Write-Log "Data Served:      $totalServedMB MB" 'Info'
        Write-Log "Hit Rate:         $hitRate" 'Info'
        Write-Log "Unique Clients:   $uniqueClientCount -- $clientList" 'Info'
        Write-Log "Cache Folder:     $cacheSizeGB" 'Info'
        Write-Log "Errors:           $errorCount" 'Info'
        Write-Log '============================================' 'Info'
    }
}

#endregion

#region ================================================================
#  PUBLISH TO CUSTOM FIELD
#================================================================

$finalHtml = [string]($html -join '')

Write-Log "Writing HTML to custom field '$CustomFieldName' (HTML: $($finalHtml.Length) chars)" 'Progress'
try {
    $finalHtml | Ninja-Property-Set-Piped $CustomFieldName
    Write-Log "Custom field write complete: $CustomFieldName" 'Info'
} catch {
    Write-Log "Failed writing custom field '$CustomFieldName': $($_.Exception.Message)" 'Error'
    throw
}

$elapsed = (Get-Date) - $startTime
Write-Log "=== Complete in $([math]::Round($elapsed.TotalSeconds, 1))s ===" 'Info'
Write-Log "Custom Field: $CustomFieldName" 'Info'

#endregion
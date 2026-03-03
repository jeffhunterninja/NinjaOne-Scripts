<#
.SYNOPSIS
    Uploads HTML reports produced by Invoke-ScriptTracker.ps1 to NinjaOne as global Knowledge Base articles. Creates new articles or updates existing ones by name; skips updates when content is unchanged (hash-based).
.DESCRIPTION
    Reads all .html files under BaseOutputFolder (same output as Invoke-ScriptTracker.ps1), derives a unique article name from the time-frame folder and the HTML <title>, and creates or updates global KB articles via the NinjaOne API. If an article already exists and the content hash matches the last uploaded version, the update is skipped by default. Uses a sidecar .kb-article-hashes.json file to track content hashes. Credentials: Ninja-Property-Get -> env (NINJAONE_INSTANCE, NINJAONE_CLIENT_ID, NINJAONE_CLIENT_SECRET) -> parameters. Requires NinjaOneDocs module.
.PARAMETER BaseOutputFolder
    Folder containing the Script Tracker HTML output (e.g. Past 10 Days subfolder). Defaults to C:\RMM\Reports\Script Tracking.
.PARAMETER DestinationFolderPath
    NinjaOne KB folder path for newly created articles. Default: Script Tracking.
.PARAMETER NinjaOneInstance
    NinjaOne instance host (e.g. app.ninjaone.com). Resolved from custom properties or env if not provided.
.PARAMETER NinjaOneClientId
    NinjaOne API client ID. Resolved from custom properties or env if not provided.
.PARAMETER NinjaOneClientSecret
    NinjaOne API client secret. Resolved from custom properties or env if not provided.
.PARAMETER SkipUnchanged
    When an article exists, skip PATCH if content hash matches the last stored hash. Default: $true.
.PARAMETER ForceUpdate
    If set, perform PATCH even when hash matches (overrides SkipUnchanged for this run).
.PARAMETER PruneStaleArticles
    After create/update, remove articles from the destination KB folder that are not in the current report set (stale). Lists articles under DestinationFolderPath, then archives those whose names are not in the current run's article list. Default: $false (opt-in).
.PARAMETER WhatIf
    Common parameter (from SupportsShouldProcess). Lists what would be created, updated, or skipped without calling the API or writing the hash file. When -PruneStaleArticles is also set, lists which articles would be pruned (archived).
.EXAMPLE
    .\Publish-ScriptTrackerToKnowledgeBase.ps1
    Publishes all HTML under the default report folder; skips updates when content is unchanged.
.EXAMPLE
    .\Publish-ScriptTrackerToKnowledgeBase.ps1 -WhatIf
    Shows what would be created, updated, or skipped without making changes.
.EXAMPLE
    .\Publish-ScriptTrackerToKnowledgeBase.ps1 -ForceUpdate
    Forces all existing articles to be updated even when content is unchanged.
.EXAMPLE
    .\Publish-ScriptTrackerToKnowledgeBase.ps1 -PruneStaleArticles
    Publishes reports and archives any KB articles under the destination folder that are no longer in the current report set (e.g. scripts that dropped out of Past 10 Days).
.LINK
    Invoke-ScriptTracker.ps1
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter()]
    [string]$BaseOutputFolder = 'C:\RMM\Reports\Script Tracking',
    [Parameter()]
    [string]$DestinationFolderPath = 'Script Tracking',
    [Parameter()]
    [string]$NinjaOneInstance = '',
    [Parameter()]
    [string]$NinjaOneClientId = '',
    [Parameter()]
    [string]$NinjaOneClientSecret = '',
    [Parameter()]
    [switch]$SkipUnchanged = $true,
    [Parameter()]
    [switch]$ForceUpdate = $false,
    [Parameter()]
    [switch]$PruneStaleArticles = $false
)

$ErrorActionPreference = 'Stop'

# NinjaOne KB article content limit (20M); use 19.9M to leave margin
$MaxKbArticleContentLength = 19900000

# --- In-line: extract <title> from HTML ---
function Get-HtmlTitle {
    param([Parameter(Mandatory)][string]$Html)
    if ([string]::IsNullOrWhiteSpace($Html)) { return 'Untitled' }
    if ($Html -match '(?s)<title>\s*(.*?)\s*</title>') { return $Matches[1].Trim() }
    return 'Untitled'
}

# --- In-line: get first path segment (time frame) from path relative to base ---
function Get-TimeFrameFromRelativePath {
    param([Parameter(Mandatory)][string]$RelativePath)
    $segments = $RelativePath -split [regex]::Escape([System.IO.Path]::DirectorySeparatorChar) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($segments.Count -gt 0) { return $segments[0] }
    return 'Reports'
}

# --- In-line: get relative directory from file path (normalized to / for API) ---
function Get-RelativeDirectoryFromPath {
    param([Parameter(Mandatory)][string]$RelativePath)
    $dir = [System.IO.Path]::GetDirectoryName($RelativePath)
    if ([string]::IsNullOrWhiteSpace($dir)) { return '' }
    $dir = $dir -replace '\\', '/'
    return $dir.Trim('/')
}

# --- In-line: convert path to NinjaOne KB folder format (pipe-separated) ---
function Get-NinjaOneFolderPath {
    param([Parameter(Mandatory)][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $segments = $Path -split '[/\\]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    return ($segments -join '|').Trim('|')
}

# --- In-line: normalize device detail title: "Device Detail: X (id)" as-is; "X - device Y" -> "Device Detail: X (Y)" ---
function Get-NormalizedArticleTitle {
    param([Parameter(Mandatory)][string]$Title)
    if ([string]::IsNullOrWhiteSpace($Title)) { return $Title }
    $trimmed = $Title.Trim()
    if ($trimmed -match '^Device\s+Detail:\s+.+\s+\(.+\)\s*$') { return $trimmed }
    if ($trimmed -match '^(.+)\s+-\s+device\s+(.+)\s*$') { return "Device Detail: $($Matches[1].Trim()) ($($Matches[2].Trim()))" }
    return $trimmed
}

# --- In-line: normalize HTML for hashing (strip run-time "Generated on" timestamp so only material changes affect hash) ---
function Get-NormalizedContentForHash {
    param([Parameter(Mandatory)][string]$Content)
    if ([string]::IsNullOrEmpty($Content)) { return $Content }
    return [regex]::Replace($Content, 'Generated on \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', 'Generated on <report>')
}

# --- In-line: SHA256 hash of string (UTF-8) as hex ---
function Get-ContentHash {
    param([Parameter(Mandatory)][string]$Content)
    $utf8 = [System.Text.Encoding]::UTF8
    $bytes = $utf8.GetBytes($Content)
    $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
    return [BitConverter]::ToString($hash) -replace '-', ''
}

# --- In-line: load hash file from BaseOutputFolder ---
function Get-StoredHashes {
    param([Parameter(Mandatory)][string]$Folder)
    $path = Join-Path $Folder '.kb-article-hashes.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return @{} }
    try {
        $json = Get-Content -LiteralPath $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($json)) { return @{} }
        $obj = $json | ConvertFrom-Json
        $ht = @{}
        $obj.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
        return $ht
    } catch {
        Write-Warning "Could not read hash file '$path': $_. Using empty hash cache."
        return @{}
    }
}

# --- In-line: save hash file ---
function Set-StoredHashes {
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)]$Hashtable)
    $path = Join-Path $Folder '.kb-article-hashes.json'
    $obj = [PSCustomObject]$Hashtable
    $json = $obj | ConvertTo-Json
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($path, $json, $utf8NoBom)
}

# --- In-line: get global KB article by exact name; returns $null, single article, or throws if multiple ---
function Get-GlobalKBArticleByName {
    param([Parameter(Mandatory)][string]$Name)
    $trimName = if ([string]::IsNullOrWhiteSpace($Name)) { '' } else { $Name.Trim() }
    if ([string]::IsNullOrWhiteSpace($trimName)) { return $null }
    $qs = "articleName=$([uri]::EscapeDataString($trimName))&includeArchived=true"
    $list = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/global/articles' -QueryParams $qs
    $hits = @($list | Where-Object { $_.name -eq $trimName })
    if ($hits.Count -gt 1) { throw "Multiple KB articles matched name '$trimName' (ids: $(($hits | ForEach-Object { $_.id }) -join ','))." }
    $hit = $hits | Select-Object -First 1
    if ($hit -and $hit.PSObject.Properties['id'] -and ([long]$hit.id) -gt 0) { return $hit }
    return $null
}

# --- In-line: list global KB articles under a destination folder path (for prune) ---
function Get-KBArticlesUnderFolder {
    param(
        [Parameter(Mandatory)][string]$DestinationFolderPathFilter
    )
    $destPathNorm = (Get-NinjaOneFolderPath -Path $DestinationFolderPathFilter.Trim()).Trim('|')
    if ([string]::IsNullOrWhiteSpace($destPathNorm)) { return @() }
    try {
        $list = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/global/articles' -QueryParams 'includeArchived=true' -Paginate
    } catch {
        Write-Warning "Could not list global KB articles for prune: $_. Skipping prune."
        return @()
    }
    $result = [System.Collections.Generic.List[object]]::new()
    foreach ($a in @($list)) {
        if (-not $a -or [string]::IsNullOrWhiteSpace($a.name)) { continue }
        $id = $null
        if ($a.PSObject.Properties['id'] -and ([long]$a.id) -gt 0) { $id = [long]$a.id }
        if (-not $id) { continue }
        $path = if ($a.PSObject.Properties['path']) { [string]$a.path } else { '' }
        $pathNorm = if ([string]::IsNullOrWhiteSpace($path)) { '' } else { (Get-NinjaOneFolderPath -Path $path).Trim('|') }
        $underDest = -not [string]::IsNullOrWhiteSpace($pathNorm) -and ($pathNorm -eq $destPathNorm -or $pathNorm.StartsWith($destPathNorm + '|'))
        if (-not $underDest) { continue }
        $isArchived = $false
        if ($a.PSObject.Properties['archived']) { $isArchived = [bool]$a.archived }
        [void]$result.Add([PSCustomObject]@{ Name = $a.name.Trim(); Id = $id; Archived = $isArchived })
    }
    return @($result)
}

# --- NinjaOneDocs module ---
try {
    $moduleName = 'NinjaOneDocs'
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Install-Module -Name $moduleName -Force -AllowClobber
    }
    Import-Module $moduleName -ErrorAction Stop
} catch {
    Write-Error "Failed to import NinjaOneDocs module. Error: $_"
    exit 1
}

# --- Credentials: Ninja-Property-Get -> env -> parameters ---
$resolvedInstance = $null
$resolvedClientId = $null
$resolvedClientSecret = $null
try { $fromNinja = Ninja-Property-Get ninjaoneInstance; if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $resolvedInstance = $fromNinja } } catch { }
if ([string]::IsNullOrWhiteSpace($resolvedInstance)) { $resolvedInstance = $env:NINJAONE_INSTANCE }
if ([string]::IsNullOrWhiteSpace($resolvedInstance) -and $PSBoundParameters.ContainsKey('NinjaOneInstance')) { $resolvedInstance = $NinjaOneInstance }

try { $fromNinja = Ninja-Property-Get ninjaoneClientId; if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $resolvedClientId = $fromNinja } } catch { }
if ([string]::IsNullOrWhiteSpace($resolvedClientId)) { $resolvedClientId = $env:NINJAONE_CLIENT_ID }
if ([string]::IsNullOrWhiteSpace($resolvedClientId) -and $PSBoundParameters.ContainsKey('NinjaOneClientId')) { $resolvedClientId = $NinjaOneClientId }

try { $fromNinja = Ninja-Property-Get ninjaoneClientSecret; if (-not [string]::IsNullOrWhiteSpace($fromNinja)) { $resolvedClientSecret = $fromNinja } } catch { }
if ([string]::IsNullOrWhiteSpace($resolvedClientSecret)) { $resolvedClientSecret = $env:NINJAONE_CLIENT_SECRET }
if ([string]::IsNullOrWhiteSpace($resolvedClientSecret) -and $PSBoundParameters.ContainsKey('NinjaOneClientSecret')) { $resolvedClientSecret = $NinjaOneClientSecret }

if ([string]::IsNullOrWhiteSpace($resolvedInstance) -or [string]::IsNullOrWhiteSpace($resolvedClientId) -or [string]::IsNullOrWhiteSpace($resolvedClientSecret)) {
    Write-Error "Missing required API credentials. Set ninjaoneInstance, ninjaoneClientId, ninjaoneClientSecret in NinjaOne custom properties, or use env vars NINJAONE_INSTANCE, NINJAONE_CLIENT_ID, NINJAONE_CLIENT_SECRET, or script parameters."
    exit 1
}

# --- Validate BaseOutputFolder ---
$baseFolder = $BaseOutputFolder
if (-not (Test-Path -LiteralPath $baseFolder -PathType Container)) {
    Write-Error "BaseOutputFolder not found or not a directory: $baseFolder. Run Invoke-ScriptTracker.ps1 first to generate HTML reports."
    exit 1
}

# --- Enumerate HTML files and build article list ---
$baseFullPath = (Resolve-Path -LiteralPath $baseFolder).Path
$htmlFiles = Get-ChildItem -LiteralPath $baseFullPath -Filter '*.html' -Recurse -File -ErrorAction SilentlyContinue
if (-not $htmlFiles -or $htmlFiles.Count -eq 0) {
    Write-Warning "No .html files found under $baseFolder. Run Invoke-ScriptTracker.ps1 first to generate reports."
    exit 0
}

$destRoot = $DestinationFolderPath.Trim().Trim('/').Trim('\')
$articleList = [System.Collections.Generic.List[object]]::new()
$seenNames = @{}
foreach ($file in $htmlFiles) {
    $rawContent = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($rawContent)) { continue }
    $title = Get-HtmlTitle -Html $rawContent
    $normalizedTitle = Get-NormalizedArticleTitle -Title $title
    $relativePath = $file.FullName.Substring($baseFullPath.Length).TrimStart([System.IO.Path]::DirectorySeparatorChar)
    $relativeDir = Get-RelativeDirectoryFromPath -RelativePath $relativePath
    $timeFrame = Get-TimeFrameFromRelativePath -RelativePath $relativePath
    $articleName = "$timeFrame - $normalizedTitle"
    $articleDestPath = if ([string]::IsNullOrWhiteSpace($relativeDir)) { $destRoot } else { "$destRoot/$relativeDir" }
    $contentForHash = Get-NormalizedContentForHash -Content $rawContent
    $contentHash = Get-ContentHash -Content $contentForHash
    if ($seenNames.ContainsKey($articleName)) {
        Write-Warning "Duplicate article name '$articleName' (from $relativePath); keeping first occurrence."
        continue
    }
    $seenNames[$articleName] = $true
    [void]$articleList.Add([PSCustomObject]@{
        ArticleName             = $articleName
        HtmlContent             = $rawContent
        ContentHash             = $contentHash
        RelativePath            = $relativePath
        ArticleDestinationPath  = $articleDestPath
    })
}

# --- Replace content over API limit with placeholder so publish does not fail ---
$placeholderHtml = @'
<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>Size limit</title></head>
<body style="margin: 20px; font-family: 'Segoe UI', sans-serif;">
<p>This article exceeded the Knowledge Base size limit and was not published. Consider using a shorter time window or reducing detail in Script Tracker (e.g. -MaxDetailRows).</p>
</body>
</html>
'@
$placeholderHash = Get-ContentHash -Content $placeholderHtml
foreach ($item in $articleList) {
    if ($item.HtmlContent.Length -gt $MaxKbArticleContentLength) {
        Write-Warning "Article '$($item.ArticleName)' exceeds KB size limit ($($item.HtmlContent.Length) chars). Replacing with placeholder."
        $item.HtmlContent = $placeholderHtml
        $item.ContentHash = $placeholderHash
    }
}

if ($articleList.Count -eq 0) {
    Write-Warning "No valid HTML content to publish."
    exit 0
}

# --- Connect to NinjaOne (needed for classification even in WhatIf) ---
try {
    Connect-NinjaOne -NinjaOneInstance $resolvedInstance -NinjaOneClientID $resolvedClientId -NinjaOneClientSecret $resolvedClientSecret
} catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit 1
}

$toCreate = [System.Collections.Generic.List[object]]::new()
$toUpdate = [System.Collections.Generic.List[object]]::new()
$toSkip = [System.Collections.Generic.List[object]]::new()
$errors = [System.Collections.Generic.List[string]]::new()

$storedHashes = Get-StoredHashes -Folder $baseFullPath
$effectiveSkip = $SkipUnchanged -and -not $ForceUpdate

foreach ($item in $articleList) {
    try {
        $existing = Get-GlobalKBArticleByName -Name $item.ArticleName
    } catch {
        [void]$errors.Add($_.Exception.Message)
        continue
    }
    if (-not $existing) {
        [void]$toCreate.Add($item)
    } else {
        $storedHash = $storedHashes[$item.ArticleName]
        if ($effectiveSkip -and $null -ne $storedHash -and $storedHash -eq $item.ContentHash) {
            [void]$toSkip.Add($item)
        } else {
            $item | Add-Member -NotePropertyName 'ExistingId' -NotePropertyValue ([long]$existing.id) -Force
            $isArchived = $false
            if ($existing.PSObject.Properties.Name -contains 'archived') { $isArchived = [bool]$existing.archived }
            $item | Add-Member -NotePropertyName 'ExistingArchived' -NotePropertyValue $isArchived -Force
            [void]$toUpdate.Add($item)
        }
    }
}

# --- Expected article names (current report set) for prune ---
$expectedArticleNames = @{}
foreach ($item in $articleList) { $expectedArticleNames[$item.ArticleName] = $true }

if ($WhatIfPreference) {
    Write-Host "WhatIf: Would create $($toCreate.Count) article(s), update $($toUpdate.Count), skip $($toSkip.Count)."
    if ($toCreate.Count -gt 0) {
        $names = ($toCreate | ForEach-Object { $_.ArticleName }) -join '; '
        Write-Host "  Create: $names"
    }
    if ($toUpdate.Count -gt 0) {
        $names = ($toUpdate | ForEach-Object { $_.ArticleName }) -join '; '
        Write-Host "  Update: $names"
    }
    if ($toSkip.Count -gt 0) {
        $names = ($toSkip | ForEach-Object { $_.ArticleName }) -join '; '
        Write-Host "  Skip (unchanged): $names"
    }
    if ($PruneStaleArticles) {
        $allInFolder = Get-KBArticlesUnderFolder -DestinationFolderPathFilter $destRoot
        $wouldPrune = @($allInFolder | Where-Object { -not $_.Archived -and -not $expectedArticleNames.ContainsKey($_.Name) })
        if ($wouldPrune.Count -gt 0) {
            Write-Host "WhatIf: Would prune (archive) $($wouldPrune.Count) stale article(s): $(($wouldPrune | ForEach-Object { $_.Name }) -join '; ')"
        } else {
            Write-Host "WhatIf: No stale articles to prune."
        }
    }
    if ($errors.Count -gt 0) { foreach ($e in $errors) { Write-Warning $e } }
    exit 0
}

if ($errors.Count -gt 0) {
    foreach ($e in $errors) { Write-Error $e }
    exit 1
}

$created = 0
$updated = 0

foreach ($item in $toCreate) {
    try {
        $body = [PSCustomObject]@{
            name                    = $item.ArticleName
            destinationFolderPath   = Get-NinjaOneFolderPath -Path $item.ArticleDestinationPath
            content                 = @{ html = [string]$item.HtmlContent }
        }
        $resp = Invoke-NinjaOneRequest -Path 'knowledgebase/articles' -Method POST -InputObject @($body) -AsArray
        if ($resp) {
            $created++
            Write-Host "Created: $($item.ArticleName)"
        }
    } catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -match 'exceeds the maximum length|String length') {
            Write-Warning "Skipped create for '$($item.ArticleName)': content over size limit. $errMsg"
        } else {
            Write-Error "Failed to create '$($item.ArticleName)': $_"
            throw
        }
    }
}

foreach ($item in $toUpdate) {
    try {
        $body = [PSCustomObject]@{
            id                     = $item.ExistingId
            name                   = $item.ArticleName
            destinationFolderPath  = Get-NinjaOneFolderPath -Path $item.ArticleDestinationPath
            content                = @{ html = [string]$item.HtmlContent }
        }
        if ($item.ExistingArchived) { $body | Add-Member -NotePropertyName 'archived' -NotePropertyValue $false -Force }
        $resp = Invoke-NinjaOneRequest -Path 'knowledgebase/articles' -Method PATCH -InputObject @($body) -AsArray
        if ($resp) {
            $updated++
            #Write-Host "Updated: $($item.ArticleName)"
        }
    } catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -match 'exceeds the maximum length|String length') {
            Write-Warning "Skipped update for '$($item.ArticleName)': content over size limit. $errMsg"
        } else {
            Write-Error "Failed to update '$($item.ArticleName)': $_"
            throw
        }
    }
}

foreach ($item in $toSkip) {
    Write-Host "Skipped (unchanged): $($item.ArticleName)"
}

# --- Prune stale articles: archive those in KB under destination folder but not in current report set ---
$pruned = 0
if ($PruneStaleArticles) {
    $allInFolder = Get-KBArticlesUnderFolder -DestinationFolderPathFilter $destRoot
    $toPrune = @($allInFolder | Where-Object { -not $_.Archived -and -not $expectedArticleNames.ContainsKey($_.Name) })
    foreach ($article in $toPrune) {
        try {
            $body = [PSCustomObject]@{
                id       = $article.Id
                archived = $true
            }
            Invoke-NinjaOneRequest -Path 'knowledgebase/articles' -Method PATCH -InputObject @($body) -AsArray | Out-Null
            $pruned++
            Write-Host "Pruned (archived): $($article.Name)"
        } catch {
            Write-Warning "Failed to archive stale article '$($article.Name)' (id=$($article.Id)): $_"
        }
    }
}

if ($created -gt 0 -or $updated -gt 0) {
    $newHashes = @{}
    foreach ($item in $articleList) { $newHashes[$item.ArticleName] = $item.ContentHash }
    Set-StoredHashes -Folder $baseFullPath -Hashtable $newHashes
}

$pruneSummary = if ($PruneStaleArticles) { ", Pruned: $pruned" } else { '' }
Write-Host "Done. Created: $created, Updated: $updated, Skipped: $($toSkip.Count)$pruneSummary."

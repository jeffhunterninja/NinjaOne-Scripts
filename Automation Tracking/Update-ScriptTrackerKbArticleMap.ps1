<#
.SYNOPSIS
    Builds or updates the KB article ID mapping file used by Invoke-ScriptTracker.ps1 for deep links. Runs independently of the Tracker and Publish workflow.
.DESCRIPTION
    Connects to NinjaOne, lists global Knowledge Base articles via the knowledgebase/global/articles API (not the org-scoped folder API), optionally filters by destination folder path (e.g. Script Tracking), and writes a JSON file mapping article name to id and parentFolderId. Invoke-ScriptTracker.ps1 reads this file when generating HTML so links between KB articles use valid NinjaOne deep links. Run after Publish-ScriptTrackerToKnowledgeBase.ps1 to refresh the map with new article IDs, or on a schedule. Credentials: Ninja-Property-Get -> env (NINJAONE_INSTANCE, NINJAONE_CLIENT_ID, NINJAONE_CLIENT_SECRET) -> parameters. Requires NinjaOneDocs module.
.PARAMETER BaseOutputFolder
    Folder where the mapping file is written (same as Invoke-ScriptTracker.ps1 output). Defaults to C:\RMM\Reports\Script Tracking.
.PARAMETER DestinationFolderPath
    NinjaOne KB folder path to enumerate (e.g. Script Tracking). Only articles under this folder are included in the map. Default: Script Tracking.
.PARAMETER NinjaOneInstance
    NinjaOne instance host (e.g. app.ninjaone.com). Resolved from custom properties or env if not provided.
.PARAMETER NinjaOneClientId
    NinjaOne API client ID. Resolved from custom properties or env if not provided.
.PARAMETER NinjaOneClientSecret
    NinjaOne API client secret. Resolved from custom properties or env if not provided.
.PARAMETER MergeWithExisting
    If the mapping file already exists, merge new entries with existing ones so entries outside the destination folder are preserved. Default: $true.
.EXAMPLE
    .\Update-ScriptTrackerKbArticleMap.ps1
    Refreshes the article ID map from NinjaOne and writes .kb-article-ids.json to the default report folder.
.EXAMPLE
    .\Update-ScriptTrackerKbArticleMap.ps1 -BaseOutputFolder "C:\Reports\KB" -MergeWithExisting:$false
    Writes a new map file containing only articles under the default Script Tracking folder; does not merge with existing.
.LINK
    Invoke-ScriptTracker.ps1
.LINK
    Publish-ScriptTrackerToKnowledgeBase.ps1
#>

[CmdletBinding()]
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
    [switch]$MergeWithExisting = $true
)

$ErrorActionPreference = 'Stop'

# --- In-line: convert path to NinjaOne KB folder format (pipe-separated) ---
function Get-NinjaOneFolderPath {
    param([Parameter(Mandatory)][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $segments = $Path -split '[/\\|]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    return ($segments -join '|').Trim('|')
}

# --- In-line: build article map from global KB using knowledgebase/global/articles (not org-scoped folder API) ---
function Get-GlobalKBArticleMap {
    param(
        [Parameter(Mandatory)][hashtable]$ResultMap,
        [string]$DestinationFolderPathFilter = ''
    )
    try {
        $query = 'includeArchived=true'
        $list = Invoke-NinjaOneRequest -Method GET -Path 'knowledgebase/global/articles' -QueryParams $query -Paginate
    } catch {
        Write-Warning "Could not list global KB articles: $_. Skipping."
        return
    }
    $articles = @($list)
    $destPathNorm = $null
    if (-not [string]::IsNullOrWhiteSpace($DestinationFolderPathFilter)) {
        $destPathNorm = (Get-NinjaOneFolderPath -Path $DestinationFolderPathFilter.Trim()).Trim('|')
    }
    foreach ($a in $articles) {
        if (-not $a -or [string]::IsNullOrWhiteSpace($a.name)) { continue }
        $id = $null
        $parentId = $null
        if ($a.PSObject.Properties['id'] -and ([long]$a.id) -gt 0) { $id = [long]$a.id }
        if ($a.PSObject.Properties['parentFolderId'] -and ([long]$a.parentFolderId) -gt 0) { $parentId = [long]$a.parentFolderId }
        if ($id -le 0 -or $parentId -le 0) { continue }
        if ($null -ne $destPathNorm) {
            $path = if ($a.PSObject.Properties['path']) { [string]$a.path } else { '' }
            $pathNorm = if ([string]::IsNullOrWhiteSpace($path)) { '' } else { (Get-NinjaOneFolderPath -Path $path).Trim('|') }
            $underDest = -not [string]::IsNullOrWhiteSpace($pathNorm) -and ($pathNorm -eq $destPathNorm -or $pathNorm.StartsWith($destPathNorm + '|'))
            if (-not $underDest) { continue }
        }
        $ResultMap[$a.name.Trim()] = @{ id = $id; parentFolderId = $parentId }
    }
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
    Write-Error "Missing required API credentials. Set ninjaoneInstance, ninjaoneClientId, ninjaoneClientSecret in NinjaOne custom properties, or use env vars NINJAONE_INSTANCE, NINJAONE_CLIENT_ID, NINJAONE_CLIENT_SECRET, or parameters."
    exit 1
}

try {
    Connect-NinjaOne -NinjaOneInstance $resolvedInstance -NinjaOneClientID $resolvedClientId -NinjaOneClientSecret $resolvedClientSecret
} catch {
    Write-Error "Failed to connect to NinjaOne API. Error: $_"
    exit 1
}

$articleMap = @{}
Get-GlobalKBArticleMap -ResultMap $articleMap -DestinationFolderPathFilter $DestinationFolderPath

if ($MergeWithExisting) {
    $mapPath = Join-Path $BaseOutputFolder '.kb-article-ids.json'
    if (Test-Path -LiteralPath $mapPath -PathType Leaf) {
        try {
            $json = Get-Content -LiteralPath $mapPath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($json)) {
                $existing = $json | ConvertFrom-Json
                $existing.PSObject.Properties | ForEach-Object {
                    $key = $_.Name
                    $val = $_.Value
                    if ($val -is [PSCustomObject] -and $val.PSObject.Properties['id'] -and $val.PSObject.Properties['parentFolderId']) {
                        if (-not $articleMap.ContainsKey($key)) { $articleMap[$key] = @{ id = [long]$val.id; parentFolderId = [long]$val.parentFolderId } }
                    }
                }
            }
        } catch {
            Write-Warning "Could not read existing map file '$mapPath': $_. Writing new map only."
        }
    }
}

$outputDir = $BaseOutputFolder
if (-not (Test-Path -LiteralPath $outputDir -PathType Container)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
$mapPath = Join-Path $outputDir '.kb-article-ids.json'
$toSerialize = @{}
foreach ($k in $articleMap.Keys) {
    $toSerialize[$k] = [PSCustomObject]@{ id = $articleMap[$k].id; parentFolderId = $articleMap[$k].parentFolderId }
}
$obj = [PSCustomObject]$toSerialize
$json = $obj | ConvertTo-Json
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($mapPath, $json, $utf8NoBom)
Write-Host "Wrote $($articleMap.Count) article(s) to $mapPath"

#Requires -Version 5.1
<#
.SYNOPSIS
  Checks if specified tags are present on the current device using the NinjaOne API.

.DESCRIPTION
  Evaluates device tag presence against tags supplied via parameters or NinjaOne script
  variables. Supports two evaluation modes: ANY (at least one tag must be present) and
  ALL (all specified tags must be present). Intended for use with NinjaOne compound
  conditions and automation policies.

.EXIT CODES
  0 = Match found (tag condition satisfied)
  1 = No match (tag condition not satisfied)
  2 = Error (validation failure, Get-NinjaTag failure, etc.)
#>

[CmdletBinding()]
param(
    [string]$Tags,
    [string]$Mode = 'any'
)

$ErrorActionPreference = 'Stop'

# NinjaOne script variables populate env vars; use them when present
if ($null -ne $env:tagsToSearch -and -not [string]::IsNullOrWhiteSpace($env:tagsToSearch)) {
    $Tags = $env:tagsToSearch.Trim()
}
if ($null -ne $env:mode -and -not [string]::IsNullOrWhiteSpace($env:mode)) {
    $Mode = $env:mode.Trim().ToLowerInvariant()
} elseif ([string]::IsNullOrWhiteSpace($Mode)) {
    $Mode = 'any'
}

# Normalize Mode for switch comparison
$Mode = ($Mode -as [string]).Trim().ToLowerInvariant()

try {
    # Convert input to array, trimming whitespace and ensuring it's always an array
    $tagArray = @($Tags -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if (-not $tagArray -or $tagArray.Count -eq 0) {
        Write-Error "No tags specified."
        exit 2
    }

    # Get currently assigned tags from NinjaOne; ensure array form
    $currentTags = @(Get-NinjaTag)

    if (-not $currentTags) {
        Write-Error "Unable to retrieve current tags."
        exit 2
    }

    Write-Verbose "Raw Tags Input: '$Tags'"
    Write-Verbose "Parsed Tags: $($tagArray -join ', ')"
    Write-Verbose "Current Ninja Tags: $($currentTags -join ', ')"

    # If only one tag was specified, check for it directly
    if ($tagArray.Count -eq 1) {
        if ($currentTags -contains $tagArray[0]) {
            Write-Host "Tag '$($tagArray[0])' is present."
            exit 0
        } else {
            Write-Host "Tag '$($tagArray[0])' is NOT present."
            exit 1
        }
    }

    # Multi-tag logic based on $Mode
    switch ($Mode) {
        "all" {
            $missingTags = $tagArray | Where-Object { $currentTags -notcontains $_ }
            if ($missingTags.Count -eq 0) {
                Write-Host "All specified tags are present."
                exit 0
            } else {
                Write-Host "Not all specified tags are present. Missing: $($missingTags -join ', ')"
                exit 1
            }
        }
        "any" {
            $anyFound = $tagArray | Where-Object { $currentTags -contains $_ }
            if ($anyFound.Count -gt 0) {
                Write-Host "At least one specified tag is present: $($anyFound -join ', ')"
                exit 0
            } else {
                Write-Host "None of the specified tags are present."
                exit 1
            }
        }
        default {
            Write-Error "Invalid Mode: '$Mode'. Use 'any' or 'all'."
            exit 2
        }
    }
}
catch {
    Write-Error "An error occurred: $_"
    exit 2
}

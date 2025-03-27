param (
    [string]$Tags,
    
    [string]$Mode
)

# Assign environment variables
$Tags = $env:tagsToSearch
$Mode = $env:mode

Write-Output $Tags

try {
    # Convert input to array, trimming whitespace and ensuring it's always an array
    $tagArray = @($Tags -split ',' | ForEach-Object { $_.Trim() })

    if (-not $tagArray -or $tagArray.Count -eq 0) {
        Write-Error "No tags specified."
        exit 2
    }

    # Get currently assigned tags from NinjaOne
    $currentTags = Get-NinjaTag

    if (-not $currentTags) {
        Write-Error "Unable to retrieve current tags."
        exit 2
    }

    Write-Host "Raw Tags Input: '$Tags'"
    Write-Host "Parsed Tags: $($tagArray -join ', ')"
    Write-Host "Current Ninja Tags: $($currentTags -join ', ')"

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
            # All specified tags must be present
            $allFound = $tagArray | ForEach-Object { $currentTags -contains $_ } | Where-Object { $_ -eq $false } | Measure-Object
            if ($allFound.Count -eq 0) {
                Write-Host "All specified tags are present."
                exit 0
            } else {
                Write-Host "Not all specified tags are present."
                exit 1
            }
        }
        "any" {
            # At least one specified tag must be present
            $anyFound = $tagArray | Where-Object { $currentTags -contains $_ }
            if ($anyFound.Count -gt 0) {
                Write-Host "At least one specified tag is present: $($anyFound -join ', ')"
                exit 0
            } else {
                Write-Host "None of the specified tags are present."
                exit 1
            }
        }
    }
}
catch {
    Write-Error "An error occurred: $_"
    exit 2
}

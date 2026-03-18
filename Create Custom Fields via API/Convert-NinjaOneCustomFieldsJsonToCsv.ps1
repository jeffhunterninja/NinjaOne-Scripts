#Requires -Version 5.1
<#
.SYNOPSIS
    Converts a NinjaOne custom fields JSON export to the CSV format used by New-NinjaOneCustomField.ps1 -CsvPath.

.DESCRIPTION
    Reads a JSON file containing custom field definitions (either a root-level array or an object with a
    "customFields" array) and writes a CSV with columns: Label, Description, Type, DefinitionScope,
    TechnicianPermission, ScriptPermission, ApiPermission, DefaultValue, DropdownValues. The output is
    compatible with New-NinjaOneCustomField.ps1 -CsvPath. The JSON file can be retrieved using developer tools
    in a web browser, navigating to the network tab, then navigating in NinjaOne to the device custom fields page.

.PARAMETER JsonPath
    Path to the JSON file. Default: customfieldstemplate.json in the script directory.

.PARAMETER CsvPath
    Path for the output CSV. Default: NinjaOne-CustomFields-Example.csv in the script directory.

.EXAMPLE
    .\Convert-NinjaOneCustomFieldsJsonToCsv.ps1
    Converts .\customfieldstemplate.json to .\NinjaOne-CustomFields-Example.csv.

.EXAMPLE
    .\Convert-NinjaOneCustomFieldsJsonToCsv.ps1 -JsonPath 'C:\export.json' -CsvPath 'C:\fields.csv'
    Converts the specified JSON file to the specified CSV path.
#>

[CmdletBinding()]
param(
    [string]$JsonPath = (Join-Path (Split-Path -Parent $PSCommandPath) 'customfieldstemplate.json'),
    [string]$CsvPath  = (Join-Path (Split-Path -Parent $PSCommandPath) 'NinjaOne-CustomFields-Example.csv')
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $JsonPath -PathType Leaf)) {
    Write-Error "JsonPath does not exist or is not a file: $JsonPath"
}

$raw = Get-Content -LiteralPath $JsonPath -Raw
$parsed = $raw | ConvertFrom-Json

$items = $null
if ($parsed -is [System.Array]) {
    $items = $parsed
} elseif ($parsed.PSObject.Properties['customFields']) {
    $items = @($parsed.customFields)
} else {
    Write-Error "JSON must be a root-level array or an object with a 'customFields' array."
}

$csvRows = [System.Collections.ArrayList]::new()
foreach ($obj in $items) {
    $name = [string]$obj.name
    if ([string]::IsNullOrWhiteSpace($name)) { continue }

    $label = $name.Trim()
    $description = if ($null -eq $obj.description) { '' } else { [string]$obj.description }
    if (-not [string]::IsNullOrWhiteSpace($description)) { $description = $description.Trim() }

    $attrType = [string]$obj.attributeType
    if ([string]::IsNullOrWhiteSpace($attrType)) { $attrType = 'TEXT' }
    else { $attrType = $attrType.Trim().ToUpperInvariant() }

    $defScope = $obj.definitionScope
    $defScopeStr = 'NODE'
    if ($defScope -and $defScope -is [System.Array]) {
        $parts = @($defScope | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($parts.Count -gt 0) { $defScopeStr = $parts -join ';' }
    } elseif ($defScope -and -not ($defScope -is [System.Array])) {
        $defScopeStr = [string]$defScope
    }

    $techPerm = [string]$obj.technicianPermission
    if ([string]::IsNullOrWhiteSpace($techPerm)) { $techPerm = 'NONE' }
    else { $techPerm = $techPerm.Trim() }

    $scriptPerm = [string]$obj.scriptPermission
    if ([string]::IsNullOrWhiteSpace($scriptPerm)) { $scriptPerm = 'NONE' }
    else { $scriptPerm = $scriptPerm.Trim() }

    $apiPerm = [string]$obj.apiPermission
    if ([string]::IsNullOrWhiteSpace($apiPerm)) { $apiPerm = 'NONE' }
    else { $apiPerm = $apiPerm.Trim() }

    $defaultVal = ''
    $rawDefault = $obj.defaultValue
    if ($null -ne $rawDefault) { $rawDefault = [string]$rawDefault }
    if (-not [string]::IsNullOrWhiteSpace($rawDefault)) {
        $isUuid = $rawDefault -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        $content = $obj.content
        $values = $null
        if ($content -and $content.PSObject.Properties['values']) { $values = $content.values }
        if ($isUuid -and $values -and ($attrType -eq 'DROPDOWN' -or $attrType -eq 'MULTI_SELECT')) {
            $matchVal = $values | Where-Object { [string]$_.id -eq $rawDefault } | Select-Object -First 1
            if ($matchVal -and $matchVal.PSObject.Properties['name']) {
                $defaultVal = [string]$matchVal.name
            } else {
                $defaultVal = $rawDefault.Trim()
            }
        } else {
            $defaultVal = $rawDefault.Trim()
        }
    }

    $dropdownStr = ''
    $content = $obj.content
    if ($content -and $content.PSObject.Properties['values']) {
        $vals = @($content.values | ForEach-Object {
            if ($_.PSObject.Properties['name']) { [string]$_.name } else { $null }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($vals.Count -gt 0) { $dropdownStr = $vals -join ';' }
    }

    $row = [PSCustomObject]@{
        Label                = $label
        Description          = $description
        Type                 = $attrType
        DefinitionScope      = $defScopeStr
        TechnicianPermission = $techPerm
        ScriptPermission     = $scriptPerm
        ApiPermission        = $apiPerm
        DefaultValue         = $defaultVal
        DropdownValues       = $dropdownStr
    }
    $null = $csvRows.Add($row)
}

if ($csvRows.Count -eq 0) {
    Write-Warning "No custom field rows with a non-blank name were found in the JSON."
}

$csvRows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($csvRows.Count) rows to $CsvPath"

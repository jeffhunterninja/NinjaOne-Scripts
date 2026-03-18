#Requires -Version 5.1
<#
.SYNOPSIS
    Bulk-creates NinjaOne custom fields from a CSV file via the Public API using OAuth 2.0 Authorization Code flow.

.DESCRIPTION
    Creates NinjaOne custom fields (node attributes) in bulk from a CSV file using the
    bulkCreateNodeAttributes API endpoint. Uses Authorization Code flow: opens a browser for
    user sign-in, captures the redirect with the authorization code via a local HTTP listener,
    then exchanges the code for an access token.

    API reference:
    - bulkCreateNodeAttributes: https://ca.ninjarmm.com/apidocs-beta/core-resources/operations/bulkCreateNodeAttributes

    App must be registered in NinjaOne as a Regular Web Application with Authorization Code
    grant type and a redirect URI matching -RedirectUri (e.g. http://localhost:8888/).

.PARAMETER NinjaOneInstance
    NinjaOne hostname only (e.g. app.ninjarmm.com, eu.ninjarmm.com), no path or scheme.

.PARAMETER ClientId
    API Client ID from NinjaOne (Administration > Apps > API).

.PARAMETER ClientSecret
    API Client Secret. Can be omitted if $env:NinjaOneClientSecret is set.

.PARAMETER RedirectUri
    Redirect URI registered for the app. Must match exactly (e.g. http://localhost:8888/).
    The script starts an HTTP listener on this URL to capture the authorization code.

.PARAMETER Scope
    Space-separated OAuth scopes (e.g. monitoring management). Default: monitoring management.

.PARAMETER AccessToken
    Use this access token and skip the authorization flow. Useful when token is from a previous run or external source.

.PARAMETER TokenFile
    Path to a file containing a cached access token (plain text). If present and readable, skips the authorization flow.

.PARAMETER CsvPath
    Path to a CSV file for bulk create. CSV must have a Label column (required); FieldName is derived from Label using camelCase. Optional columns: Description, Type, DefinitionScope, TechnicianPermission, ScriptPermission, ApiPermission, DefaultValue, DropdownValues. Use semicolons to separate multiple values in DefinitionScope (e.g. NODE;END_USER) and DropdownValues (e.g. Low;Medium;High). Rows with blank Label are skipped.

.PARAMETER ApiBasePath
    API path segment. Default: v2 (full path https://{instance}/v2/custom-fields). Use api/v2 only if your tenant explicitly requires it.

.PARAMETER OAuthInstance
    Hostname to use for OAuth (authorize and token) only. If not set, NinjaOneInstance is used. Set to app.ninjarmm.com if your EU (or other region) browser authorize step returns 404; some tenants must authorize at app.ninjarmm.com.

.PARAMETER OAuthPathPrefix
    Path segment before oauth. Default: ws (URLs become /ws/oauth/authorize and /ws/oauth/token). If your instance returns 404 on authorize, try '' to use /oauth/authorize and /oauth/token.

.PARAMETER UseTestConfig
    Loads NinjaOneInstance, ClientId, and related settings from the $TestConfig hashtable in this script (edit placeholders; never commit secrets).

.EXAMPLE
    .\New-NinjaOneCustomField.ps1 -NinjaOneInstance app.ninjarmm.com -ClientId 'your-id' -RedirectUri 'http://localhost:8888/' -CsvPath '.\NinjaOne-CustomFields-Example.csv'
    Bulk creates custom fields from the CSV file (one row per field; Label required, Type and other columns optional).

.EXAMPLE
    .\New-NinjaOneCustomField.ps1 -UseTestConfig
    Runs using the $TestConfig block after you edit placeholders (instance, client id, CSV path). Prefer $env:NinjaOneClientSecret for the secret.

.NOTES
    Requires PowerShell 5.1+. RedirectUri must be registered in NinjaOne. Client secret can be passed as parameter or via $env:NinjaOneClientSecret.
#>

[CmdletBinding(DefaultParameterSetName = 'BulkCsv')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'BulkCsv')]
    [string]$NinjaOneInstance,
    [Parameter(Mandatory = $true, ParameterSetName = 'BulkCsv')]
    [string]$ClientId,
    [Parameter(Mandatory = $true, ParameterSetName = 'TestConfig')]
    [switch]$UseTestConfig,
    [string]$ClientSecret = $env:NinjaOneClientSecret,
    [string]$RedirectUri = 'http://localhost:8888/',
    [string]$Scope = 'monitoring management',
    [string]$AccessToken,
    [string]$TokenFile,
    [string]$CsvPath = (Join-Path (Split-Path -Parent $PSCommandPath) 'NinjaOne-CustomFields-Example.csv'),
    [string]$ApiBasePath = 'v2',
    [string]$OAuthInstance,
    [string]$OAuthPathPrefix = 'ws'
)

$ErrorActionPreference = 'Stop'

# Edit placeholders when using -UseTestConfig. Do not commit real Client Secret here; use $env:NinjaOneClientSecret.
$TestConfig = @{
    NinjaOneInstance  = 'app.ninjarmm.com'
    ClientId          = 'your-client-id'
    ClientSecret      = ''  # empty = use $ClientSecret / $env:NinjaOneClientSecret
    RedirectUri       = 'http://localhost:8888/'
    Scope             = 'monitoring management'
    ApiBasePath       = 'v2'
    OAuthInstance     = ''  # e.g. app.ninjarmm.com if authorize 404 on EU instance
    OAuthPathPrefix   = 'ws'  # or '' for /oauth/authorize
    CsvPath           = (Join-Path (Split-Path -Parent $PSCommandPath) 'NinjaOne-CustomFields-Example.csv')
}

if ($UseTestConfig) {
    $NinjaOneInstance  = $TestConfig.NinjaOneInstance
    $ClientId          = $TestConfig.ClientId
    if (-not [string]::IsNullOrWhiteSpace([string]$TestConfig.ClientSecret)) {
        $ClientSecret = $TestConfig.ClientSecret
    }
    $RedirectUri       = $TestConfig.RedirectUri
    $Scope             = $TestConfig.Scope
    $ApiBasePath       = $TestConfig.ApiBasePath
    if (-not [string]::IsNullOrWhiteSpace([string]$TestConfig.OAuthInstance)) {
        $OAuthInstance = $TestConfig.OAuthInstance
    }
    if ($null -ne $TestConfig.OAuthPathPrefix) { $OAuthPathPrefix = $TestConfig.OAuthPathPrefix }
    if (-not [string]::IsNullOrWhiteSpace([string]$TestConfig.CsvPath)) { $CsvPath = $TestConfig.CsvPath }
    if ($ClientId -eq 'your-client-id' -or [string]::IsNullOrWhiteSpace($ClientId)) {
        throw 'Edit $TestConfig in the script: set ClientId (and NinjaOneInstance) before using -UseTestConfig.'
    }
}

$effectiveParamSet = 'BulkCsv'

# --- In-line: parse redirect URI into prefix for HttpListener (must end with /) ---
function Get-RedirectUriParts {
    param([string]$Uri)
    if ([string]::IsNullOrWhiteSpace($Uri)) { throw 'RedirectUri is required.' }
    $u = [System.Uri]$Uri
    $path = $u.AbsolutePath
    if (-not $path.EndsWith('/')) { $path += '/' }
    $hostPart = $u.Host
    if ([string]::IsNullOrWhiteSpace($hostPart)) { $hostPart = 'localhost' }
    $prefix = "$($u.Scheme)://${hostPart}:$($u.Port)$path"
    return @{ Prefix = $prefix; Port = $u.Port; FullUri = $Uri }
}

# --- In-line: convert label string to camelCase for use as fieldName ---
function ConvertTo-CamelCaseFromLabel {
    param([string]$Label)
    if ([string]::IsNullOrWhiteSpace($Label)) { return '' }
    $trimmed = $Label.Trim()
    $parts = @([regex]::Replace($trimmed, '[^a-zA-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_ })
    if ($null -eq $parts -or $parts.Count -eq 0) {
        throw "Label must produce a non-empty field name. Label='$Label' yields no alphanumeric words."
    }
    $result = ''
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $word = $parts[$i]
        if ($i -eq 0) {
            $result += $word.Substring(0, 1).ToLowerInvariant()
            if ($word.Length -gt 1) { $result += $word.Substring(1).ToLowerInvariant() }
        } else {
            $result += $word.Substring(0, 1).ToUpperInvariant()
            if ($word.Length -gt 1) { $result += $word.Substring(1).ToLowerInvariant() }
        }
    }
    return $result
}

# --- In-line: start HTTP listener, wait for one request, return code and state from query ---
function Receive-AuthorizationCodeFromListener {
    param(
        [string]$ListenerPrefix,
        [string]$SuccessHtmlTitle = 'Success',
        [string]$SuccessHtmlBody = 'Authorization successful. You can close this window.'
    )
    $listener = $null
    try {
        $listener = New-Object System.Net.HttpListener
        $listener.Prefixes.Add($ListenerPrefix)
        $listener.Start()
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        $query = $request.Url.Query
        $code = $null
        $state = $null
        if (-not [string]::IsNullOrEmpty($query)) {
            $query = $query.TrimStart('?')
            foreach ($pair in $query -split '&') {
                $kv = $pair -split '=', 2
                if ($kv.Count -ge 2) {
                    $key = [System.Uri]::UnescapeDataString($kv[0])
                    $val = [System.Uri]::UnescapeDataString($kv[1].Replace('+', ' '))
                    if ($key -eq 'code') { $code = $val }
                    if ($key -eq 'state') { $state = $val }
                }
            }
        }
        $html = @"
<!DOCTYPE html><html><head><title>$SuccessHtmlTitle</title></head><body><p>$SuccessHtmlBody</p></body></html>
"@
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
        $response.ContentLength64 = $buffer.Length
        $response.ContentType = 'text/html; charset=utf-8'
        $response.StatusCode = 200
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.OutputStream.Close()
        return @{ Code = $code; State = $state }
    }
    finally {
        if ($null -ne $listener) {
            try { $listener.Stop() } catch { }
            $listener.Close()
        }
    }
}

# --- In-line: get access token via Authorization Code flow (listener + browser + token exchange) ---
function Get-NinjaOneAccessTokenFromAuthCode {
    param(
        [string]$Instance,
        [string]$ClientID,
        [string]$ClientSecret,
        [string]$RedirectUriFull,
        [string]$Scope,
        [string]$ListenerPrefix,
        [string]$OAuthInstanceOverride,
        [string]$OAuthPathPrefix = 'ws'
    )
    if ([string]::IsNullOrWhiteSpace($ClientSecret)) {
        throw 'ClientSecret is required for authorization code flow. Set -ClientSecret or $env:NinjaOneClientSecret.'
    }
    $base = $Instance -replace '/ws$', ''
    $oauthHost = if ([string]::IsNullOrWhiteSpace($OAuthInstanceOverride)) { $base } else { $OAuthInstanceOverride.Trim() -replace '/ws$', '' }
    $pathPrefix = if ([string]::IsNullOrWhiteSpace($OAuthPathPrefix)) { '' } else { $OAuthPathPrefix.Trim().TrimEnd('/') }
    $oauthAuthPath = if ($pathPrefix -eq '') { 'oauth/authorize' } else { "$pathPrefix/oauth/authorize" }
    $oauthTokenPath = if ($pathPrefix -eq '') { 'oauth/token' } else { "$pathPrefix/oauth/token" }
    $state = [Guid]::NewGuid().ToString('N')
    # Authorize URL must NOT include client_secret (only in token exchange). Match NinjaOne working example.
    $authUrl = "https://${oauthHost}/${oauthAuthPath}?" + "response_type=code&client_id=" + $ClientID + "&redirect_uri=" + $RedirectUriFull + "&state=" + $state + "&scope=" + ($Scope -replace ' ', '%20')
    Write-Verbose "OAuth: host=$oauthHost path=$oauthAuthPath redirect=$RedirectUriFull scope=$Scope"
    Write-Verbose "OAuth authorize URL: $authUrl"
    Write-Host 'Opening browser for NinjaOne sign-in. Approve the app to continue.'
    try {
        Start-Process $authUrl
    } catch {
        Write-Host "Could not start browser. Open this URL in a browser: $authUrl"
    }
    $result = Receive-AuthorizationCodeFromListener -ListenerPrefix $ListenerPrefix
    $code = $result.Code
    if ([string]::IsNullOrWhiteSpace($code)) {
        throw 'No authorization code received. The redirect may have contained an error or the user denied access.'
    }
    $body = @{
        grant_type    = 'authorization_code'
        client_id     = $ClientID
        client_secret = $ClientSecret
        redirect_uri  = $RedirectUriFull
        code          = $code
    }
    $tokenUrl = "https://$oauthHost/$oauthTokenPath"
    $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing
    return $response
}

# --- In-line: ensure we have an access token (cached, file, or new via auth code flow) ---
function Get-AccessToken {
    param(
        [string]$Instance,
        [string]$ClientID,
        [string]$ClientSecret,
        [string]$RedirectUriFull,
        [string]$Scope,
        [string]$ProvidedToken,
        [string]$TokenFilePath,
        [string]$OAuthInstanceOverride,
        [string]$OAuthPathPrefix = 'ws'
    )
    if (-not [string]::IsNullOrWhiteSpace($ProvidedToken)) {
        return $ProvidedToken
    }
    if (-not [string]::IsNullOrWhiteSpace($TokenFilePath) -and (Test-Path -LiteralPath $TokenFilePath -PathType Leaf)) {
        $token = (Get-Content -LiteralPath $TokenFilePath -Raw).Trim()
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            return $token
        }
    }
    $uriParts = Get-RedirectUriParts -Uri $RedirectUriFull
    $tokenResponse = Get-NinjaOneAccessTokenFromAuthCode -Instance $Instance -ClientID $ClientID -ClientSecret $ClientSecret -RedirectUriFull $RedirectUriFull -Scope $Scope -ListenerPrefix $uriParts.Prefix -OAuthInstanceOverride $OAuthInstanceOverride -OAuthPathPrefix $OAuthPathPrefix
    return $tokenResponse.access_token
}

# --- In-line: POST to NinjaOne API with Bearer token ---
function Invoke-NinjaOnePost {
    param(
        [string]$Instance,
        [string]$ApiBasePath,
        [string]$Path,
        [string]$AccessToken,
        [object]$Body
    )
    $base = $Instance -replace '/ws$', ''
    $uri = "https://$base/$ApiBasePath/$Path"
    $json = if ($null -eq $Body) { '{}' } else { $Body | ConvertTo-Json -Depth 15 -Compress:$false }
    # PowerShell ConvertTo-Json serializes single-element arrays as scalars; API expects definitionScope as array. Fix "definitionScope": "X" -> ["X"].
    if ($json -match '"definitionScope"\s*:\s*"') {
        $json = [regex]::Replace($json, '"definitionScope"\s*:\s*"([^"]*)"', '"definitionScope": ["$1"]')
    }
    $requestUri = [System.Uri]::new($uri)
    $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Post, $requestUri)
    $request.Headers.TryAddWithoutValidation('Authorization', "Bearer $AccessToken") | Out-Null
    $request.Headers.TryAddWithoutValidation('Accept', 'application/json') | Out-Null
    $request.Content = [System.Net.Http.StringContent]::new($json, [System.Text.Encoding]::UTF8, 'application/json')
    try {
        $handler = [System.Net.Http.HttpClientHandler]::new()
        $client = [System.Net.Http.HttpClient]::new($handler)
        try {
            $response = $client.SendAsync($request).GetAwaiter().GetResult()
            $responseBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            if (-not $response.IsSuccessStatusCode) {
                $statusCode = [int]$response.StatusCode
                if (-not [string]::IsNullOrWhiteSpace($responseBody)) {
                    try {
                        $err = $responseBody | ConvertFrom-Json
                        $detail = $err.errorMessage
                        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = $err.reason }
                        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = $err.message }
                        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = $err.error }
                        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = $responseBody }
                        $msg = "HTTP $statusCode - $detail"
                        if ($err.incidentId) { $msg += " (incidentId: $($err.incidentId))" }
                        throw $msg
                    } catch {
                        if ($_.Exception.Message -match '^HTTP \d+ - ') { throw }
                    }
                }
                throw "HTTP $statusCode calling $uri - $responseBody"
            }
            if ([string]::IsNullOrWhiteSpace($responseBody)) {
                return $null
            }
            return $responseBody | ConvertFrom-Json
        } finally {
            $handler.Dispose()
            $client.Dispose()
        }
    } finally {
        $request.Dispose()
    }
}

# --- Convert a hashtable field definition (bulk item) to API shape ---
function ConvertTo-CustomFieldApiBody {
    param([hashtable]$Def)
    $label = $Def['label']
    if ([string]::IsNullOrWhiteSpace($label)) {
        throw "Each field definition must have at least 'label'. fieldName is optional and derived from label when omitted."
    }
    $fieldName = $Def['fieldName']
    if ([string]::IsNullOrWhiteSpace($fieldName)) {
        $fieldName = ConvertTo-CamelCaseFromLabel -Label $label
    }
    $defScope = if ($Def['definitionScope']) { @($Def['definitionScope']) } else { @('NODE') }
    $body = @{
        label                 = $label
        fieldName             = $fieldName
        scope                 = 'NODE_ROLE'
        definitionScope       = $defScope
        type                  = if ($Def['type']) { $Def['type'] } else { 'TEXT' }
        technicianPermission  = if ($Def['technicianPermission']) { $Def['technicianPermission'] } else { 'NONE' }
        scriptPermission      = if ($Def['scriptPermission']) { $Def['scriptPermission'] } else { 'NONE' }
        apiPermission         = if ($Def['apiPermission']) { $Def['apiPermission'] } else { 'NONE' }
        addToDefaultTab       = $false
    }
    if ($Def['description']) { $body.description = $Def['description'] }
    if ($Def['defaultValue']) { $body.defaultValue = $Def['defaultValue'] }
    if ($Def['dropdownValues'] -and $Def['dropdownValues'].Count -gt 0) {
        $body.content = @{
            values   = @($Def['dropdownValues'] | ForEach-Object { @{ name = $_ } })
            required = $false
        }
    }
    return $body
}

# --- Main ---
Write-Host "Starting NinjaOne custom field creation..."
Write-Host "  Mode:     $effectiveParamSet"

$instance = $NinjaOneInstance.Trim()
if ([string]::IsNullOrWhiteSpace($instance)) {
    throw 'NinjaOneInstance is required.'
}
Write-Host "  Instance: $instance"

Write-Host "Acquiring access token..."
$token = Get-AccessToken -Instance $instance -ClientID $ClientId -ClientSecret $ClientSecret -RedirectUriFull $RedirectUri -Scope $Scope -ProvidedToken $AccessToken -TokenFilePath $TokenFile -OAuthInstanceOverride $OAuthInstance -OAuthPathPrefix $OAuthPathPrefix
Write-Host "Access token acquired."

if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    throw 'CsvPath is required. Provide a path to a CSV file containing custom field definitions.'
}
$validTypes = @('DROPDOWN','MULTI_SELECT','CHECKBOX','TEXT','TEXT_MULTILINE','TEXT_ENCRYPTED','NUMERIC','DECIMAL','DATE','DATE_TIME','TIME','ATTACHMENT','NODE_DROPDOWN','NODE_MULTI_SELECT','CLIENT_DROPDOWN','CLIENT_MULTI_SELECT','CLIENT_LOCATION_DROPDOWN','CLIENT_LOCATION_MULTI_SELECT','CLIENT_DOCUMENT_DROPDOWN','CLIENT_DOCUMENT_MULTI_SELECT','EMAIL','PHONE','IP_ADDRESS','WYSIWYG','URL','MONETARY','IDENTIFIER','TOTP')
$resolvedCsvPath = if ($PSCmdlet) { $PSCmdlet.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CsvPath) } else { if ([System.IO.Path]::IsPathRooted($CsvPath)) { $CsvPath } else { Join-Path (Split-Path -Parent $PSCommandPath) $CsvPath } }
if (-not (Test-Path -LiteralPath $resolvedCsvPath -PathType Leaf)) {
    throw "CsvPath must point to an existing file: $CsvPath"
}
$rows = Import-Csv -LiteralPath $resolvedCsvPath
$customFields = @()
foreach ($row in $rows) {
    $labelVal = [string]$row.Label
    if ([string]::IsNullOrWhiteSpace($labelVal)) {
        continue
    }
    $labelVal = $labelVal.Trim()
    $fieldNameVal = ConvertTo-CamelCaseFromLabel -Label $labelVal
    $typeVal = [string]$row.Type
    if ([string]::IsNullOrWhiteSpace($typeVal)) { $typeVal = 'TEXT' }
    else {
        $typeUpper = $typeVal.Trim().ToUpperInvariant()
        if ($typeUpper -notin $validTypes) { $typeVal = 'TEXT' }
        else { $typeVal = $typeUpper }
    }
    $scopeVal = [string]$row.DefinitionScope
    $definitionScopeArr = if ([string]::IsNullOrWhiteSpace($scopeVal)) { @('NODE') } else { @($scopeVal.Trim() -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
    if ($definitionScopeArr.Count -eq 0) { $definitionScopeArr = @('NODE') }
    $dropdownVal = [string]$row.DropdownValues
    $dropdownArr = if ([string]::IsNullOrWhiteSpace($dropdownVal)) { @() } else { @($dropdownVal.Trim() -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
    $def = @{
        label                 = $labelVal
        fieldName             = $fieldNameVal
        type                  = $typeVal
        definitionScope       = $definitionScopeArr
        technicianPermission  = if ([string]::IsNullOrWhiteSpace([string]$row.TechnicianPermission)) { 'NONE' } else { ([string]$row.TechnicianPermission).Trim() }
        scriptPermission      = if ([string]::IsNullOrWhiteSpace([string]$row.ScriptPermission)) { 'NONE' } else { ([string]$row.ScriptPermission).Trim() }
        apiPermission         = if ([string]::IsNullOrWhiteSpace([string]$row.ApiPermission)) { 'NONE' } else { ([string]$row.ApiPermission).Trim() }
    }
    $descProp = [string]$row.Description
    if (-not [string]::IsNullOrWhiteSpace($descProp)) { $def['description'] = $descProp.Trim() }
    $defaultProp = [string]$row.DefaultValue
    if (-not [string]::IsNullOrWhiteSpace($defaultProp)) { $def['defaultValue'] = $defaultProp.Trim() }
    if ($dropdownArr.Count -gt 0) { $def['dropdownValues'] = $dropdownArr }
    $customFields += ConvertTo-CustomFieldApiBody -Def $def
}
if ($customFields.Count -eq 0) {
    throw 'CSV file produced no valid rows. Each row must have a non-blank Label.'
}
$payload = @{ customFields = $customFields }
$fieldCount = $customFields.Count
Write-Host "Submitting $fieldCount custom field(s) from CSV..."
$result = Invoke-NinjaOnePost -Instance $instance -ApiBasePath $ApiBasePath -Path 'custom-fields/bulk' -AccessToken $token -Body $payload
if ($null -eq $result) { Write-Host "API returned success (empty response). $fieldCount field(s) submitted." }
else { Write-Host "API returned a response for $fieldCount field(s):"; $result }

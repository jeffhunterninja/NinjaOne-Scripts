#Requires -Version 5.1
<#
.SYNOPSIS
Syncs NinjaOne device assigned owners from the assignedUser custom field.

.DESCRIPTION
Authenticates with NinjaOne using OAuth Authorization Code (PKCE), loads all devices,
reads the assignedUser custom field for each device, and sets device owner when an
end-user account matches that email.

By default, creates missing end users in the device's organization before assigning
ownership. Pass -CreateMissingEndUsers:$false to skip create and only assign when
an end user already exists.

Devices whose current assigned owner email already matches assignedUser are skipped.

This script is standalone and does not dot-source shared helpers.

.PARAMETER NinjaOneInstance
NinjaOne hostname or base URL (for example: ca.ninjarmm.com or https://ca.ninjarmm.com).

.PARAMETER ClientId
OAuth application client ID (Authorization Code app with PKCE).

.PARAMETER RedirectUri
OAuth redirect URI. Must match EXACTLY one Redirect URL registered on the
Client App in Administration > Apps > API (including trailing slash).
Default: http://localhost:8888/

.PARAMETER Scope
OAuth scope string.
Default: monitoring management

.PARAMETER CustomFieldName
Device custom field API name containing the assigned-user email.
Default: assignedUser

.PARAMETER CreateMissingEndUsers
When set (default), creates an end user for emails that do not already match an
end-user account. Uses the device's organizationId. Pass -CreateMissingEndUsers:$false
to disable create.

.PARAMETER ResultsPath
Optional path to export per-device results CSV.
Default: script folder with timestamped filename.

.EXAMPLE
.\Sync-NinjaDeviceOwnerFromAssignedUser.ps1 `
  -NinjaOneInstance ca.ninjarmm.com `
  -ClientId "your-client-id"

.EXAMPLE
.\Sync-NinjaDeviceOwnerFromAssignedUser.ps1 `
  -NinjaOneInstance ca.ninjarmm.com `
  -ClientId "your-client-id" `
  -CreateMissingEndUsers:$false `
  -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $false)]
    [string]$NinjaOneInstance = 'ca.ninjarmm.com',

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [string]$RedirectUri = 'http://localhost:8888/',

    [string]$Scope = 'monitoring management',

    [string]$CustomFieldName = 'assignedUser',

    [switch]$CreateMissingEndUsers = $false,

    [string]$ResultsPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-BaseUrl {
    param([Parameter(Mandatory = $true)][string]$Instance)
    $trimmed = $Instance.Trim()
    if (-not $trimmed) { throw 'NinjaOneInstance cannot be empty.' }
    if ($trimmed -notmatch '^https?://') {
        $trimmed = "https://$trimmed"
    }
    $uri = [Uri]$trimmed
    if ($uri.Scheme -ne 'https' -and $uri.Host -ne 'localhost') {
        throw "NinjaOneInstance must use https (received: '$trimmed')."
    }
    return $trimmed.TrimEnd('/')
}

function Resolve-RedirectUri {
    param([Parameter(Mandatory = $true)][string]$UriValue)

    # Strip accidental quotes/whitespace from CLI or copy/paste (e.g. http://localhost:8888/')
    $trimmed = $UriValue.Trim().Trim("`"", "'")
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        throw 'RedirectUri cannot be empty.'
    }
    if ($trimmed -notmatch '^https?://') {
        throw "RedirectUri must be an absolute http(s) URL (received: '$UriValue')."
    }

    $uri = [Uri]$trimmed
    if ($uri.Host -ne 'localhost' -and $uri.Host -ne '127.0.0.1' -and $uri.Scheme -ne 'https') {
        throw "RedirectUri for non-localhost hosts must use https (received: '$trimmed')."
    }

    # HttpListener prefixes require a trailing slash on the authority root.
    if ([string]::IsNullOrEmpty($uri.AbsolutePath) -or $uri.AbsolutePath -eq '/') {
        return "$($uri.Scheme)://$($uri.Authority)/"
    }

    $path = $uri.AbsolutePath
    if (-not $path.EndsWith('/')) {
        $path = "$path/"
    }
    return "$($uri.Scheme)://$($uri.Authority)$path"
}

function New-PkceVerifier {
    $bytes = New-Object byte[] 64
    [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
    $base64 = [Convert]::ToBase64String($bytes)
    return ($base64.TrimEnd('=').Replace('+', '-').Replace('/', '_'))
}

function Get-PkceChallenge {
    param([Parameter(Mandatory = $true)][string]$Verifier)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::ASCII.GetBytes($Verifier)
        $hash = $sha.ComputeHash($bytes)
        return ([Convert]::ToBase64String($hash).TrimEnd('=').Replace('+', '-').Replace('/', '_'))
    } finally {
        $sha.Dispose()
    }
}

function New-RandomState {
    return ([Guid]::NewGuid().ToString('N') + [Guid]::NewGuid().ToString('N'))
}

function Get-TokenFromAuthorizationCode {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$ClientIdValue,
        [Parameter(Mandatory = $true)][string]$RedirectUriValue,
        [Parameter(Mandatory = $true)][string]$ScopeValue
    )

    $normalizedRedirect = Resolve-RedirectUri -UriValue $RedirectUriValue
    $redirect = [Uri]$normalizedRedirect
    $listener = [System.Net.HttpListener]::new()
    # Listener prefix is scheme://host:port/ ; callback path is still matched via the request URL.
    $listener.Prefixes.Add("$($redirect.Scheme)://$($redirect.Host):$($redirect.Port)/")
    $listener.Start()

    $verifier = New-PkceVerifier
    $challenge = Get-PkceChallenge -Verifier $verifier
    $state = New-RandomState

    $authorizeUri = '{0}/ws/oauth/authorize?response_type=code&client_id={1}&redirect_uri={2}&scope={3}&state={4}&code_challenge={5}&code_challenge_method=S256' -f `
        $BaseUrl, `
        [Uri]::EscapeDataString($ClientIdValue), `
        [Uri]::EscapeDataString($normalizedRedirect), `
        [Uri]::EscapeDataString($ScopeValue), `
        [Uri]::EscapeDataString($state), `
        [Uri]::EscapeDataString($challenge)

    Write-Host "OAuth redirect_uri (must match Client App exactly): $normalizedRedirect"
    Write-Host "OAuth client_id: $ClientIdValue"
    Write-Host 'Opening browser for NinjaOne sign-in...'
    Start-Process $authorizeUri | Out-Null

    Write-Host "Waiting for OAuth callback on $normalizedRedirect ..."
    $context = $listener.GetContext()

    try {
        $request = $context.Request
        $response = $context.Response
        $code = $request.QueryString['code']
        $receivedState = $request.QueryString['state']
        $errorValue = $request.QueryString['error']

        if ($errorValue) {
            $html = '<html><body><h3>Authentication failed.</h3><p>You can close this tab.</p></body></html>'
            $bytes = [Text.Encoding]::UTF8.GetBytes($html)
            $response.ContentType = 'text/html'
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
            $response.Close()
            throw "OAuth authorization failed: $errorValue"
        }

        if ([string]::IsNullOrWhiteSpace($code)) {
            throw 'OAuth callback did not include an authorization code.'
        }
        if ($receivedState -ne $state) {
            throw 'OAuth callback state mismatch. Request may be invalid.'
        }

        $html = '<html><body><h3>Authentication complete.</h3><p>You can close this tab and return to PowerShell.</p></body></html>'
        $bytes = [Text.Encoding]::UTF8.GetBytes($html)
        $response.ContentType = 'text/html'
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        $response.Close()

        $tokenBody = @{
            grant_type    = 'authorization_code'
            client_id     = $ClientIdValue
            redirect_uri  = $normalizedRedirect
            code          = $code
            code_verifier = $verifier
            scope         = $ScopeValue
        }

        return Invoke-RestMethod -Method Post -Uri "$BaseUrl/ws/oauth/token" `
            -Headers @{ accept = 'application/json'; 'Content-Type' = 'application/x-www-form-urlencoded' } `
            -Body $tokenBody
    }
    finally {
        if ($listener.IsListening) {
            $listener.Stop()
        }
        $listener.Close()
    }
}

function Invoke-NinjaRequest {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')][string]$Method,
        [Parameter(Mandatory = $true)][string]$Endpoint,
        [object]$Body
    )

    $uri = "$BaseUrl/ws/$($Endpoint.TrimStart('/'))"
    $headers = @{
        accept        = 'application/json'
        Authorization = "Bearer $AccessToken"
    }

    $params = @{
        Method  = $Method
        Uri     = $uri
        Headers = $headers
    }

    if ($null -ne $Body) {
        if ($Body -is [string]) {
            $params['Body'] = $Body
        }
        else {
            $params['Body'] = ($Body | ConvertTo-Json -Depth 10 -Compress)
        }
        $params['ContentType'] = 'application/json'
    }

    return Invoke-RestMethod @params
}

function ConvertTo-ApiArray {
    param($InputObject)

    if ($null -eq $InputObject) {
        return @()
    }

    # Bare JSON arrays from Invoke-RestMethod
    if ($InputObject -is [System.Array]) {
        return @($InputObject)
    }

    # Generic lists / ArrayLists from some hosts
    if ($InputObject -is [System.Collections.IEnumerable] -and
        -not ($InputObject -is [string]) -and
        -not ($InputObject -is [System.Collections.IDictionary]) -and
        -not ($InputObject -is [PSCustomObject])) {
        $tmp = New-Object System.Collections.ArrayList
        foreach ($el in $InputObject) {
            [void]$tmp.Add($el)
        }
        return @($tmp.ToArray())
    }

    $psObj = $InputObject.PSObject
    if ($psObj) {
        foreach ($propName in @('results', 'data', 'items', 'users', 'devices')) {
            if ($psObj.Properties[$propName]) {
                $value = $InputObject.$propName
                if ($null -eq $value) {
                    return @()
                }
                if ($value -is [System.Array]) {
                    return @($value)
                }
                if ($value -is [System.Collections.IEnumerable] -and
                    -not ($value -is [string]) -and
                    -not ($value -is [System.Collections.IDictionary]) -and
                    -not ($value -is [PSCustomObject])) {
                    $tmp = New-Object System.Collections.ArrayList
                    foreach ($el in $value) {
                        [void]$tmp.Add($el)
                    }
                    return @($tmp.ToArray())
                }
                return @($value)
            }
        }
    }

    return @($InputObject)
}

function ConvertTo-ScalarString {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string]) {
        $s = $Value.Trim()
        if ([string]::IsNullOrWhiteSpace($s)) { return $null }
        return $s
    }
    if ($Value -is [System.Array]) {
        foreach ($item in $Value) {
            $s = ConvertTo-ScalarString -Value $item
            if ($s) { return $s }
        }
        return $null
    }
    if ($Value -is [PSCustomObject] -or $Value -is [System.Collections.IDictionary]) {
        foreach ($name in @('email', 'Email', 'value', 'text', 'name', 'uid', 'id')) {
            $propValue = $null
            if ($Value -is [System.Collections.IDictionary]) {
                if ($Value.Contains($name)) { $propValue = $Value[$name] }
            }
            else {
                $prop = $Value.PSObject.Properties[$name]
                if ($prop) { $propValue = $prop.Value }
            }
            $s = ConvertTo-ScalarString -Value $propValue
            if ($s) { return $s }
        }
        return $null
    }
    $asString = ($Value -as [string])
    if ([string]::IsNullOrWhiteSpace($asString)) { return $null }
    return $asString.Trim()
}

function Get-PageInfo {
    param($Response)

    $hasMore = $false
    $nextCursor = $null

    # Bare arrays have no pageInfo
    if ($null -eq $Response -or $Response -is [System.Array]) {
        return [PSCustomObject]@{
            HasMore    = $false
            NextCursor = $null
        }
    }

    if ($Response.PSObject -and $Response.PSObject.Properties['pageInfo'] -and $null -ne $Response.pageInfo) {
        $pi = $Response.pageInfo
        if ($pi.PSObject.Properties['hasMore'] -and $pi.hasMore) {
            $hasMore = $true
        }
        if ($pi.PSObject.Properties['nextCursor'] -and $null -ne $pi.nextCursor) {
            $nextCursor = [string]$pi.nextCursor
            if ([string]::IsNullOrWhiteSpace($nextCursor)) {
                $nextCursor = $null
            }
        }
    }

    return [PSCustomObject]@{
        HasMore    = $hasMore
        NextCursor = $nextCursor
    }
}

function Get-NinjaPagedResults {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUrl,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [Parameter(Mandatory = $true)][string]$EndpointBase,
        [int]$PageSize = 500
    )

    $collected = New-Object System.Collections.ArrayList
    $after = $null
    $pageIndex = 0

    do {
        $pageIndex++
        $hasQuery = ($EndpointBase.IndexOf('?') -ge 0)
        $separator = if ($hasQuery) { '&' } else { '?' }
        $ep = "$EndpointBase$separator" + "pageSize=$PageSize"
        if (-not [string]::IsNullOrWhiteSpace($after)) {
            $ep += "&after=$([System.Uri]::EscapeDataString([string]$after))"
        }

        try {
            $resp = Invoke-NinjaRequest -BaseUrl $BaseUrl -AccessToken $AccessToken -Method GET -Endpoint $ep
        }
        catch {
            throw "GET /ws/$ep failed: $($_.Exception.Message)"
        }

        $items = @(ConvertTo-ApiArray -InputObject $resp)
        foreach ($item in $items) {
            # Cast to [object] so an accidental nested array is one Add() arg, not a splat.
            if ($item -is [System.Array]) {
                foreach ($sub in $item) {
                    [void]$collected.Add([object]$sub)
                }
            }
            else {
                [void]$collected.Add([object]$item)
            }
        }

        $page = Get-PageInfo -Response $resp
        if ($page.HasMore -and -not [string]::IsNullOrWhiteSpace($page.NextCursor)) {
            $after = [string]$page.NextCursor
        }
        else {
            $after = $null
        }

        # Safety: avoid infinite loops if API keeps returning the same cursor
        if ($pageIndex -gt 10000) {
            throw "Pagination exceeded 10000 pages for endpoint '$EndpointBase'."
        }
    } while (-not [string]::IsNullOrWhiteSpace($after))

    return @($collected.ToArray())
}

function Test-PlausibleEmail {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return $false }
    return ($Email -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function ConvertTo-NormalizedEmail {
    param($Value)
    $raw = ConvertTo-ScalarString -Value $Value
    if (-not $raw) { return $null }
    return $raw.Trim().ToLowerInvariant()
}

function Get-NamePartsFromEmail {
    param([Parameter(Mandatory = $true)][string]$Email)

    $local = ($Email -split '@', 2)[0]
    $tokens = @($local -split '[.\-_]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    function ConvertTo-TitleCaseToken {
        param([string]$Token)
        if ([string]::IsNullOrWhiteSpace($Token)) { return $Token }
        if ($Token.Length -eq 1) { return $Token.ToUpperInvariant() }
        return ($Token.Substring(0, 1).ToUpperInvariant() + $Token.Substring(1).ToLowerInvariant())
    }

    if ($tokens.Count -eq 0) {
        return [PSCustomObject]@{ FirstName = 'User'; LastName = 'User' }
    }
    if ($tokens.Count -eq 1) {
        return [PSCustomObject]@{
            FirstName = (ConvertTo-TitleCaseToken -Token $tokens[0])
            LastName  = 'User'
        }
    }

    $first = ConvertTo-TitleCaseToken -Token $tokens[0]
    $last = ConvertTo-TitleCaseToken -Token ($tokens[$tokens.Count - 1])
    return [PSCustomObject]@{ FirstName = $first; LastName = $last }
}

function Get-DeviceDisplayName {
    param($Device)
    foreach ($name in @('displayName', 'systemName', 'name', 'dnsName')) {
        if ($Device.PSObject.Properties[$name]) {
            $value = ConvertTo-ScalarString -Value $Device.$name
            if ($value) { return $value }
        }
    }
    $id = ConvertTo-ScalarString -Value $Device.id
    if ($id) { return "Device $id" }
    return 'Unknown device'
}

function Get-DeviceOrganizationId {
    param($Device)
    foreach ($name in @('organizationId', 'orgId', 'clientId')) {
        if ($Device.PSObject.Properties[$name] -and $null -ne $Device.$name) {
            $n = 0
            if ([int]::TryParse(($Device.$name -as [string]), [ref]$n) -and $n -gt 0) {
                return $n
            }
        }
    }
    if ($Device.PSObject.Properties['organization'] -and $Device.organization) {
        $org = $Device.organization
        if ($org.PSObject.Properties['id'] -and $null -ne $org.id) {
            $n = 0
            if ([int]::TryParse(($org.id -as [string]), [ref]$n) -and $n -gt 0) {
                return $n
            }
        }
    }
    return $null
}

function Get-OwnerRefValues {
    param($Value)

    $found = New-Object System.Collections.ArrayList

    if ($null -eq $Value) {
        return @()
    }

    if ($Value -is [System.Array]) {
        foreach ($item in $Value) {
            foreach ($nested in (Get-OwnerRefValues -Value $item)) {
                [void]$found.Add($nested)
            }
        }
        return @($found.ToArray())
    }

    if ($Value -is [PSCustomObject] -or $Value -is [System.Collections.IDictionary]) {
        foreach ($propName in @(
                'uid', 'userUid', 'assignedUserUid', 'ownerUid', 'assignedOwnerUid',
                'id', 'userId', 'assignedUserId', 'ownerId', 'assignedOwnerId',
                'email', 'Email', 'value'
            )) {
            $propValue = $null
            if ($Value -is [System.Collections.IDictionary]) {
                if ($Value.Contains($propName)) { $propValue = $Value[$propName] }
            }
            else {
                $prop = $Value.PSObject.Properties[$propName]
                if ($prop) { $propValue = $prop.Value }
            }
            if ($null -ne $propValue) {
                foreach ($nested in (Get-OwnerRefValues -Value $propValue)) {
                    [void]$found.Add($nested)
                }
            }
        }
        return @($found.ToArray())
    }

    $s = ConvertTo-ScalarString -Value $Value
    if ($s) {
        [void]$found.Add($s)
    }
    return @($found.ToArray())
}

function Get-DeviceOwnerRefs {
    param($Device)

    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $ordered = New-Object System.Collections.ArrayList

    $propertyNames = @(
        'ownerUid', 'assignedUserUid', 'assignedOwnerUid',
        'ownerId', 'assignedUserId', 'assignedOwnerId',
        'userUid', 'userId',
        'owner', 'assignedUser', 'assignedTo', 'primaryUser',
        'user', 'endUser', 'userInfo', 'ownerInfo', 'assignedOwner'
    )

    foreach ($propName in $propertyNames) {
        if (-not $Device.PSObject.Properties[$propName]) { continue }
        $propValue = $Device.$propName
        if ($null -eq $propValue) { continue }

        foreach ($ref in (Get-OwnerRefValues -Value $propValue)) {
            if ($seen.Add($ref)) {
                [void]$ordered.Add($ref)
            }
        }
    }

    return @($ordered.ToArray())
}

function Get-DeviceOwnerEmail {
    param(
        $Device,
        [hashtable]$EndUserByUid,
        [hashtable]$EndUserById,
        [hashtable]$EndUserByEmail
    )

    # Direct email on nested owner objects
    foreach ($objName in @('owner', 'assignedOwner', 'assignedUser', 'primaryUser', 'user', 'endUser', 'ownerInfo')) {
        if (-not $Device.PSObject.Properties[$objName] -or $null -eq $Device.$objName) { continue }
        $obj = $Device.$objName
        if ($obj.PSObject.Properties['email']) {
            $email = ConvertTo-NormalizedEmail -Value $obj.email
            if ($email) { return $email }
        }
        if ($obj.PSObject.Properties['Email']) {
            $email = ConvertTo-NormalizedEmail -Value $obj.Email
            if ($email) { return $email }
        }
    }

    $refs = @(Get-DeviceOwnerRefs -Device $Device)
    foreach ($ref in $refs) {
        if ([string]::IsNullOrWhiteSpace($ref)) { continue }
        if (Test-PlausibleEmail -Email $ref) {
            return $ref.Trim().ToLowerInvariant()
        }
        if ($EndUserByUid.ContainsKey($ref)) {
            $eu = $EndUserByUid[$ref]
            if ($eu.Email) { return $eu.Email }
        }
        if ($EndUserById.ContainsKey($ref)) {
            $eu = $EndUserById[$ref]
            if ($eu.Email) { return $eu.Email }
        }
        if ($EndUserByEmail.ContainsKey($ref.ToLowerInvariant())) {
            return $ref.ToLowerInvariant()
        }
    }
    return $null
}

function Get-CustomFieldValueFromRow {
    param(
        $Row,
        [Parameter(Mandatory = $true)][string]$FieldName
    )

    if ($null -eq $Row) { return $null }

    if ($Row.PSObject.Properties[$FieldName]) {
        return ConvertTo-ScalarString -Value $Row.$FieldName
    }

    foreach ($containerName in @('fields', 'customFields', 'values', 'attributes')) {
        if (-not $Row.PSObject.Properties[$containerName] -or $null -eq $Row.$containerName) { continue }
        $container = $Row.$containerName
        if ($container -is [System.Collections.IDictionary]) {
            if ($container.Contains($FieldName)) {
                return ConvertTo-ScalarString -Value $container[$FieldName]
            }
        }
        elseif ($container.PSObject.Properties[$FieldName]) {
            return ConvertTo-ScalarString -Value $container.$FieldName
        }
        elseif ($container -is [System.Array]) {
            foreach ($item in $container) {
                $itemName = $null
                if ($item.PSObject.Properties['name']) { $itemName = ConvertTo-ScalarString -Value $item.name }
                elseif ($item.PSObject.Properties['fieldName']) { $itemName = ConvertTo-ScalarString -Value $item.fieldName }
                elseif ($item.PSObject.Properties['key']) { $itemName = ConvertTo-ScalarString -Value $item.key }
                if ($itemName -and $itemName.Equals($FieldName, [StringComparison]::OrdinalIgnoreCase)) {
                    if ($item.PSObject.Properties['value']) {
                        return ConvertTo-ScalarString -Value $item.value
                    }
                }
            }
        }
    }
    return $null
}

function Get-RowDeviceId {
    param($Row)
    foreach ($name in @('deviceId', 'nodeId', 'id', 'DeviceId', 'NodeId')) {
        if ($Row.PSObject.Properties[$name] -and $null -ne $Row.$name) {
            $n = 0
            if ([int]::TryParse(($Row.$name -as [string]), [ref]$n) -and $n -gt 0) {
                return $n
            }
        }
    }
    return $null
}

function ConvertTo-EndUserRecord {
    param($UserObject)

    $email = ConvertTo-NormalizedEmail -Value $UserObject.email
    if (-not $email -and $UserObject.PSObject.Properties['Email']) {
        $email = ConvertTo-NormalizedEmail -Value $UserObject.Email
    }

    $uid = $null
    foreach ($name in @('uid', 'UID', 'userUid')) {
        if ($UserObject.PSObject.Properties[$name]) {
            $uid = ConvertTo-ScalarString -Value $UserObject.$name
            if ($uid) { break }
        }
    }

    $id = $null
    foreach ($name in @('id', 'Id', 'userId')) {
        if ($UserObject.PSObject.Properties[$name] -and $null -ne $UserObject.$name) {
            $id = ConvertTo-ScalarString -Value $UserObject.$name
            if ($id) { break }
        }
    }

    $orgId = $null
    foreach ($name in @('organizationId', 'orgId', 'clientId')) {
        if ($UserObject.PSObject.Properties[$name] -and $null -ne $UserObject.$name) {
            $n = 0
            if ([int]::TryParse(($UserObject.$name -as [string]), [ref]$n) -and $n -gt 0) {
                $orgId = $n
                break
            }
        }
    }

    $ownerKey = if ($uid) { $uid } elseif ($id) { $id } else { $null }

    return [PSCustomObject]@{
        Id             = $id
        Uid            = $uid
        OwnerKey       = $ownerKey
        Email          = $email
        OrganizationId = $orgId
        Raw            = $UserObject
    }
}

function Register-EndUserMaps {
    param(
        [Parameter(Mandatory = $true)]$EndUserRecord,
        [Parameter(Mandatory = $true)][hashtable]$ByEmail,
        [Parameter(Mandatory = $true)][hashtable]$ByUid,
        [Parameter(Mandatory = $true)][hashtable]$ById
    )

    if ($EndUserRecord.Email -and -not $ByEmail.ContainsKey($EndUserRecord.Email)) {
        $ByEmail[$EndUserRecord.Email] = $EndUserRecord
    }
    if ($EndUserRecord.Uid -and -not $ByUid.ContainsKey($EndUserRecord.Uid)) {
        $ByUid[$EndUserRecord.Uid] = $EndUserRecord
    }
    if ($EndUserRecord.Id -and -not $ById.ContainsKey($EndUserRecord.Id)) {
        $ById[$EndUserRecord.Id] = $EndUserRecord
    }
}

function New-ResultRow {
    param(
        [int]$DeviceId,
        [string]$DeviceName,
        $OrganizationId,
        [string]$AssignedUserEmail,
        [string]$Action,
        [string]$EndUserUid,
        [string]$Message
    )

    return [PSCustomObject]@{
        DeviceId          = $DeviceId
        DeviceName        = $DeviceName
        OrganizationId    = $OrganizationId
        AssignedUserEmail = $AssignedUserEmail
        Action            = $Action
        EndUserUid        = $EndUserUid
        Message           = $Message
    }
}

try {
    if ([string]::IsNullOrWhiteSpace($ClientId)) {
        throw 'ClientId is required.'
    }
    if ([string]::IsNullOrWhiteSpace($CustomFieldName)) {
        throw 'CustomFieldName cannot be empty.'
    }

    $baseUrl = Resolve-BaseUrl -Instance $NinjaOneInstance
    $RedirectUri = Resolve-RedirectUri -UriValue $RedirectUri
    $fieldName = $CustomFieldName.Trim()

    Write-Host 'Authenticating with NinjaOne via authorization code...'
    Write-Host "If you get 'Invalid redirect_uri', add this exact URL to the Client App Redirect URLs: $RedirectUri"
    $tokenResponse = Get-TokenFromAuthorizationCode `
        -BaseUrl $baseUrl `
        -ClientIdValue $ClientId `
        -RedirectUriValue $RedirectUri `
        -ScopeValue $Scope

    $accessToken = $tokenResponse.access_token
    if ([string]::IsNullOrWhiteSpace($accessToken)) {
        throw 'Access token was not returned from OAuth token response.'
    }

    Write-Host 'Loading devices...'
    $devices = @(Get-NinjaPagedResults -BaseUrl $baseUrl -AccessToken $accessToken -EndpointBase 'api/v2/devices-detailed')
    Write-Host "Loaded $($devices.Count) device(s)."

    Write-Host 'Loading end users...'
    $endUserRaw = @(Get-NinjaPagedResults -BaseUrl $baseUrl -AccessToken $accessToken -EndpointBase 'api/v2/user/end-users')
    if ($endUserRaw.Count -eq 0) {
        Write-Verbose 'end-users endpoint returned no rows; falling back to /api/v2/users filtered to END_USER.'
        $allUsers = @(Get-NinjaPagedResults -BaseUrl $baseUrl -AccessToken $accessToken -EndpointBase 'api/v2/users')
        $endUserRaw = @($allUsers | Where-Object {
                $_.PSObject.Properties['userType'] -and (($_.userType -as [string]) -eq 'END_USER')
            })
    }
    Write-Host "Loaded $($endUserRaw.Count) end user(s)."

    $endUserByEmail = @{}
    $endUserByUid = @{}
    $endUserById = @{}
    foreach ($eu in $endUserRaw) {
        $rec = ConvertTo-EndUserRecord -UserObject $eu
        Register-EndUserMaps -EndUserRecord $rec -ByEmail $endUserByEmail -ByUid $endUserByUid -ById $endUserById
    }

    Write-Host "Loading custom field '$fieldName'..."
    $cfEndpoint = "api/v2/queries/custom-fields?fields=$([uri]::EscapeDataString($fieldName))"
    $cfRows = @()
    try {
        $cfRows = @(Get-NinjaPagedResults -BaseUrl $baseUrl -AccessToken $accessToken -EndpointBase $cfEndpoint)
    }
    catch {
        Write-Warning "queries/custom-fields failed ($($_.Exception.Message)); trying custom-fields-detailed..."
        $cfEndpointDetailed = "api/v2/queries/custom-fields-detailed?fields=$([uri]::EscapeDataString($fieldName))"
        $cfRows = @(Get-NinjaPagedResults -BaseUrl $baseUrl -AccessToken $accessToken -EndpointBase $cfEndpointDetailed)
    }

    $assignedUserByDeviceId = @{}
    foreach ($row in $cfRows) {
        $deviceId = Get-RowDeviceId -Row $row
        if (-not $deviceId) { continue }
        $value = Get-CustomFieldValueFromRow -Row $row -FieldName $fieldName
        if ($null -eq $value) { continue }
        $assignedUserByDeviceId[$deviceId] = $value
    }
    Write-Host "Found $($assignedUserByDeviceId.Count) device(s) with a value for '$fieldName'."

    $results = New-Object System.Collections.ArrayList
    $createdEmails = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    foreach ($device in $devices) {
        $deviceId = 0
        if (-not [int]::TryParse((ConvertTo-ScalarString -Value $device.id), [ref]$deviceId) -or $deviceId -le 0) {
            continue
        }

        $deviceName = Get-DeviceDisplayName -Device $device
        $orgId = Get-DeviceOrganizationId -Device $device

        $rawCf = $null
        if ($assignedUserByDeviceId.ContainsKey($deviceId)) {
            $rawCf = $assignedUserByDeviceId[$deviceId]
        }

        $assignedEmail = ConvertTo-NormalizedEmail -Value $rawCf
        if ([string]::IsNullOrWhiteSpace($assignedEmail)) {
            $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                        -AssignedUserEmail '' -Action 'SkippedEmpty' -EndUserUid '' `
                        -Message "Custom field '$fieldName' is empty.")) | Out-Null
            continue
        }

        if (-not (Test-PlausibleEmail -Email $assignedEmail)) {
            $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                        -AssignedUserEmail $assignedEmail -Action 'Failed' -EndUserUid '' `
                        -Message 'assignedUser value is not a plausible email address.')) | Out-Null
            continue
        }

        $currentOwnerEmail = Get-DeviceOwnerEmail -Device $device `
            -EndUserByUid $endUserByUid -EndUserById $endUserById -EndUserByEmail $endUserByEmail

        if ($currentOwnerEmail -and $currentOwnerEmail -eq $assignedEmail) {
            $matched = $null
            if ($endUserByEmail.ContainsKey($assignedEmail)) { $matched = $endUserByEmail[$assignedEmail] }
            $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                        -AssignedUserEmail $assignedEmail -Action 'AlreadyMatched' `
                        -EndUserUid $(if ($matched) { $matched.OwnerKey } else { '' }) `
                        -Message 'Device assigned owner already matches assignedUser email.')) | Out-Null
            continue
        }

        $endUser = $null
        $createdThisEmail = $false
        if ($endUserByEmail.ContainsKey($assignedEmail)) {
            $endUser = $endUserByEmail[$assignedEmail]
        }
        else {
            if (-not $CreateMissingEndUsers) {
                $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                            -AssignedUserEmail $assignedEmail -Action 'EndUserMissing' -EndUserUid '' `
                            -Message 'No matching end user found. Re-run with -CreateMissingEndUsers to create.')) | Out-Null
                continue
            }

            if (-not $orgId) {
                $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId '' `
                            -AssignedUserEmail $assignedEmail -Action 'Failed' -EndUserUid '' `
                            -Message 'Cannot create end user: device has no organizationId.')) | Out-Null
                continue
            }

            if ($createdEmails.Contains($assignedEmail)) {
                # Another device already requested create; map should have been updated.
                if ($endUserByEmail.ContainsKey($assignedEmail)) {
                    $endUser = $endUserByEmail[$assignedEmail]
                }
                else {
                    $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                                -AssignedUserEmail $assignedEmail -Action 'Failed' -EndUserUid '' `
                                -Message 'End user create was attempted earlier for this email but is not available in cache.')) | Out-Null
                    continue
                }
            }
            else {
                $names = Get-NamePartsFromEmail -Email $assignedEmail
                $createBody = @{
                    firstName         = $names.FirstName
                    lastName          = $names.LastName
                    email             = $assignedEmail
                    organizationId    = $orgId
                    fullPortalAccess  = $false
                }

                $createTarget = "$assignedEmail (org $orgId)"
                if ($PSCmdlet.ShouldProcess($createTarget, 'Create end user')) {
                    try {
                        $created = Invoke-NinjaRequest -BaseUrl $baseUrl -AccessToken $accessToken `
                            -Method POST -Endpoint 'api/v2/user/end-users' -Body $createBody
                        $endUser = ConvertTo-EndUserRecord -UserObject $created
                        if (-not $endUser.Email) { $endUser.Email = $assignedEmail }
                        if (-not $endUser.OwnerKey) {
                            throw 'Created end user response did not include uid or id.'
                        }
                        Register-EndUserMaps -EndUserRecord $endUser -ByEmail $endUserByEmail -ByUid $endUserByUid -ById $endUserById
                        [void]$createdEmails.Add($assignedEmail)
                        $createdThisEmail = $true
                        Write-Host "Created end user $assignedEmail (owner key $($endUser.OwnerKey)) for org $orgId."
                    }
                    catch {
                        $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                                    -AssignedUserEmail $assignedEmail -Action 'Failed' -EndUserUid '' `
                                    -Message "Failed to create end user: $($_.Exception.Message)")) | Out-Null
                        continue
                    }
                }
                else {
                    $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                                -AssignedUserEmail $assignedEmail -Action 'EndUserMissing' -EndUserUid '' `
                                -Message 'End user create skipped by WhatIf/Confirm.')) | Out-Null
                    continue
                }
            }
        }

        if (-not $endUser -or [string]::IsNullOrWhiteSpace($endUser.OwnerKey)) {
            $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                        -AssignedUserEmail $assignedEmail -Action 'Failed' -EndUserUid '' `
                        -Message 'Resolved end user has no uid/id for owner assignment.')) | Out-Null
            continue
        }

        $ownerTarget = "device $deviceId ($deviceName) -> $($endUser.OwnerKey)"
        if ($PSCmdlet.ShouldProcess($ownerTarget, 'Set device assigned owner')) {
            try {
                Invoke-NinjaRequest -BaseUrl $baseUrl -AccessToken $accessToken `
                    -Method POST -Endpoint "api/v2/device/$deviceId/owner/$($endUser.OwnerKey)" `
                    -Body '{}' | Out-Null

                $action = if ($createdThisEmail) { 'EndUserCreatedAndOwnerSet' } else { 'OwnerSet' }
                $message = if ($createdThisEmail) {
                    'Created end user and set device owner.'
                }
                else {
                    'Device owner updated from assignedUser custom field.'
                }

                $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                            -AssignedUserEmail $assignedEmail -Action $action `
                            -EndUserUid $endUser.OwnerKey -Message $message)) | Out-Null
            }
            catch {
                $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                            -AssignedUserEmail $assignedEmail -Action 'Failed' `
                            -EndUserUid $endUser.OwnerKey `
                            -Message "Failed to set device owner: $($_.Exception.Message)")) | Out-Null
            }
        }
        else {
            $results.Add((New-ResultRow -DeviceId $deviceId -DeviceName $deviceName -OrganizationId $orgId `
                        -AssignedUserEmail $assignedEmail -Action 'SkippedWhatIf' `
                        -EndUserUid $endUser.OwnerKey `
                        -Message 'Owner set skipped by WhatIf/Confirm.')) | Out-Null
        }
    }

    if (-not $ResultsPath) {
        $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $PSCommandPath }
        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $ResultsPath = Join-Path $scriptDir "Sync-NinjaDeviceOwnerFromAssignedUser-Results-$timestamp.csv"
    }

    $results | Export-Csv -LiteralPath $ResultsPath -NoTypeInformation -Encoding UTF8

    $counts = @{}
    foreach ($row in $results) {
        $key = [string]$row.Action
        if (-not $counts.ContainsKey($key)) { $counts[$key] = 0 }
        $counts[$key] = [int]$counts[$key] + 1
    }

    Write-Host ''
    Write-Host 'Completed device owner sync from assignedUser.'
    Write-Host "Devices processed: $($results.Count)"
    foreach ($key in ($counts.Keys | Sort-Object)) {
        Write-Host ("{0,-28}: {1}" -f $key, $counts[$key])
    }
    Write-Host "Results : $ResultsPath"
}
catch {
    $msg = $_.Exception.Message
    $pos = $null
    if ($_.InvocationInfo) {
        $pos = $_.InvocationInfo.PositionMessage
    }
    $stack = $_.ScriptStackTrace
    Write-Error ("{0}`n{1}`n{2}" -f $msg, $pos, $stack)
    exit 1
}

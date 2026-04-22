<#
.SYNOPSIS
    Creates NinjaOne PSA tickets from a billing invoice export manifest and attaches each HTML invoice.

.DESCRIPTION
    Reads NinjaBillingInvoiceExport-manifest.json produced by Export-NinjaOneBillingInvoicePdfs.ps1,
    authenticates with OAuth2 refresh_token (delegated user / authorization-code app), then for each
    entry POSTs /v2/ticketing/ticket and attaches the HTML file via multipart
    POST /v2/ticketing/ticket/{id}/comment.

    Use a separate Ninja API OAuth application from your billing machine-to-machine client: the ticketing
    app must allow authorization code (and refresh token) with ticketing permissions and appropriate scopes.

    No third-party PowerShell modules required. Requires System.Net.Http for multipart upload.

.PARAMETER NinjaOneInstance
    NinjaOne instance hostname (e.g. app.ninjarmm.com). Defaults to NINJAONE_INSTANCE env var.

.PARAMETER NinjaOneClientId
    OAuth2 client ID for the refresh-token app. Defaults to NINJAONE_CLIENT_ID env var.

.PARAMETER NinjaOneClientSecret
    OAuth2 client secret. Defaults to NINJAONE_CLIENT_SECRET env var.

.PARAMETER NinjaOneRefreshToken
    Refresh token from the authorization code flow. Defaults to NINJAONE_REFRESH_TOKEN env var.

.PARAMETER ManifestPath
    Full path to NinjaBillingInvoiceExport-manifest.json. If omitted, use -OutputPath with the default manifest name.

.PARAMETER OutputPath
    Folder containing the manifest and HTML files (same as export -OutputPath). Ignored if -ManifestPath is set.

.PARAMETER TicketFormId
    Ticket form ID from Ninja (API: GET /v2/ticketing/ticket-form).

.PARAMETER TicketStatus
    Ticket status ID string for new tickets. Default is "1000" (per public API schema).

.PARAMETER TicketSubjectTemplate
    Subject line; placeholders {InvoiceNumber} and {ClientName} are replaced. Max 200 characters after expansion.

.PARAMETER TicketDescriptionBody
    Plain text for the initial ticket description (body).

.PARAMETER TicketType
    Ticket type enum value for the API (default TASK).

.PARAMETER TicketAttachmentComment
    Plain text body on the multipart comment that carries the HTML attachment.

.EXAMPLE
    .\New-NinjaOneBillingInvoiceTickets.ps1 -OutputPath .\Invoices -TicketFormId 5

.EXAMPLE
    .\New-NinjaOneBillingInvoiceTickets.ps1 -ManifestPath 'D:\Exports\NinjaBillingInvoiceExport-manifest.json' -TicketFormId 5

.EXAMPLE
    .\New-NinjaOneBillingInvoiceTickets.ps1 -OutputPath .\Invoices -TicketFormId 5 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$NinjaOneInstance      = $env:NINJAONE_INSTANCE,
    [string]$NinjaOneClientId      = $env:NINJAONE_CLIENT_ID,
    [string]$NinjaOneClientSecret  = $env:NINJAONE_CLIENT_SECRET,
    [string]$NinjaOneRefreshToken  = $env:NINJAONE_REFRESH_TOKEN,

    [string]$ManifestPath,
    [string]$OutputPath,

    [Parameter(Mandatory)]
    [int]$TicketFormId,

    [string]$TicketStatus = '1000',

    [string]$TicketSubjectTemplate = 'Billing invoice {InvoiceNumber} — {ClientName}',

    [string]$TicketDescriptionBody = 'Billing invoice export; details are in the attached HTML.',

    [ValidateSet('PROBLEM', 'QUESTION', 'INCIDENT', 'TASK', 'CHANGE_REQUEST', 'SERVICE_REQUEST', 'PROJECT', 'APPOINTMENT', 'MISCELLANEOUS')]
    [string]$TicketType = 'TASK',

    [string]$TicketAttachmentComment = 'Invoice HTML is attached.'
)

$ErrorActionPreference = 'Stop'
try { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop } catch {
    throw 'This script requires System.Net.Http (e.g. .NET Framework 4.5+ on Windows PowerShell).'
}

$DefaultManifestName = 'NinjaBillingInvoiceExport-manifest.json'

# ── Auth (refresh token) ─────────────────────────────────────────────────────

function Get-NinjaTokenFromRefreshToken {
    param(
        [Parameter(Mandatory)] [string]$ClientId,
        [Parameter(Mandatory)] [string]$ClientSecret,
        [Parameter(Mandatory)] [string]$RefreshToken,
        [Parameter(Mandatory)] [string]$Instance
    )
    $uri = "https://$($Instance.TrimEnd('/'))/ws/oauth/token"
    $body = @{
        grant_type    = 'refresh_token'
        client_id     = $ClientId
        client_secret = $ClientSecret
        refresh_token = $RefreshToken
    }
    $headers = @{
        'Accept'       = 'application/json'
        'Content-Type' = 'application/x-www-form-urlencoded'
    }
    return Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $body -TimeoutSec 30 -ErrorAction Stop
}

function Apply-NinjaRefreshTokenResponse {
    param(
        [Parameter(Mandatory)] [PSCustomObject]$Session,
        [Parameter(Mandatory)] $TokenResponse
    )
    if ([string]::IsNullOrWhiteSpace($TokenResponse.access_token)) {
        throw 'Token response did not include access_token.'
    }
    $Session.AuthHeader = "Bearer $($TokenResponse.access_token)"
    $Session.ExpiresAt  = if ($TokenResponse.expires_in) {
        (Get-Date).AddSeconds([int]$TokenResponse.expires_in - 60)
    } else {
        (Get-Date).AddMinutes(55)
    }
    if ($TokenResponse.refresh_token) {
        if ($TokenResponse.refresh_token -ne $Session.RefreshToken) {
            Write-Warning 'OAuth returned a new refresh_token; update NINJAONE_REFRESH_TOKEN (or your secret store) to avoid auth failures later.'
            $Session.RefreshToken = [string]$TokenResponse.refresh_token
        }
    }
}

# ── API wrapper with auto-refresh and exponential backoff ─────────────────────

function Invoke-NinjaApi {
    [CmdletBinding()]
    param(
        [ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',
        [Parameter(Mandatory)] [string]$Endpoint,
        [string]$Query,
        $Body,
        [int]$TimeoutSec = 60,
        [int]$MaxRetries = 4,
        [Parameter(Mandatory)] [PSCustomObject]$Session
    )

    function Get-HttpStatus($Err) {
        try { return [int]$Err.Exception.Response.StatusCode } catch { return $null }
    }
    function Get-RetryAfter($Err) {
        try {
            $raw = $Err.Exception.Response.Headers['Retry-After'] | Select-Object -First 1
            $sec = 0
            if ($raw -and [int]::TryParse($raw, [ref]$sec) -and $sec -gt 0) { return $sec }
        } catch { }
        return $null
    }
    function Refresh-Session([PSCustomObject]$S) {
        $t = Get-NinjaTokenFromRefreshToken -ClientId $S.ClientId -ClientSecret $S.ClientSecret `
            -RefreshToken $S.RefreshToken -Instance $S.Instance
        Apply-NinjaRefreshTokenResponse -Session $S -TokenResponse $t
    }

    if (-not $Session.ExpiresAt -or (Get-Date) -ge $Session.ExpiresAt) { Refresh-Session $Session }

    $base = "https://$($Session.Instance.TrimEnd('/'))/$($Endpoint.TrimStart('/'))"
    $uri  = if ($Query) { "${base}?${Query}" } else { $base }

    $attempt = 0
    while ($true) {
        $reqHeaders = @{ Authorization = $Session.AuthHeader; Accept = 'application/json' }
        $bodyJson   = $null
        if ($Body) {
            $bodyJson = $Body | ConvertTo-Json -Depth 20
            $reqHeaders['Content-Type'] = 'application/json'
        }
        try {
            return Invoke-RestMethod -Uri $uri -Method $Method -Headers $reqHeaders -Body $bodyJson -TimeoutSec $TimeoutSec -ErrorAction Stop
        } catch {
            $status = Get-HttpStatus $_
            $attempt++
            if ($status -eq 401 -and $attempt -le $MaxRetries) { Refresh-Session $Session; continue }
            $retryable = $status -in @(408, 429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -gt $MaxRetries) { throw }
            $wait  = Get-RetryAfter $_
            $sleep = if ($wait -and $wait -gt 0) { [Math]::Min($wait, 60) } else { [Math]::Min([Math]::Pow(2, $attempt), 30) }
            Write-Warning "HTTP $status — retrying in ${sleep}s (attempt $attempt/$MaxRetries)"
            Start-Sleep -Seconds $sleep
        }
    }
}

function Update-NinjaSessionToken {
    param([Parameter(Mandatory)] [PSCustomObject]$Session)
    $t = Get-NinjaTokenFromRefreshToken -ClientId $Session.ClientId -ClientSecret $Session.ClientSecret `
        -RefreshToken $Session.RefreshToken -Instance $Session.Instance
    Apply-NinjaRefreshTokenResponse -Session $Session -TokenResponse $t
}

function Expand-TicketSubject {
    param(
        [Parameter(Mandatory)] [string]$Template,
        [string]$InvoiceNumber,
        [string]$ClientName
    )
    $n = if ([string]::IsNullOrEmpty($ClientName)) { '' } else { $ClientName }
    $inv = if ([string]::IsNullOrEmpty($InvoiceNumber)) { '' } else { $InvoiceNumber }
    $s = $Template.Replace('{InvoiceNumber}', $inv).Replace('{ClientName}', $n)
    if ($s.Length -gt 200) { return $s.Substring(0, 200) }
    return $s
}

function Send-NinjaTicketCommentWithAttachment {
    param(
        [Parameter(Mandatory)] [PSCustomObject]$Session,
        [Parameter(Mandatory)] [int]$TicketId,
        [Parameter(Mandatory)] [string]$CommentBody,
        [Parameter(Mandatory)] [string]$HtmlFilePath
    )
    if (-not $Session.ExpiresAt -or (Get-Date) -ge $Session.ExpiresAt) { Update-NinjaSessionToken -Session $Session }

    $uri = "https://$($Session.Instance.TrimEnd('/'))/v2/ticketing/ticket/$TicketId/comment"
    $commentJson = (@{ public = $true; body = $CommentBody } | ConvertTo-Json -Compress -Depth 5)

    $client = [System.Net.Http.HttpClient]::new()
    try {
        $client.DefaultRequestHeaders.TryAddWithoutValidation('Authorization', $Session.AuthHeader) | Out-Null
        [void]$client.DefaultRequestHeaders.Accept.ParseAdd('application/json')

        $multipart = [System.Net.Http.MultipartFormDataContent]::new()
        try {
            $commentPart = [System.Net.Http.StringContent]::new(
                $commentJson,
                [System.Text.UTF8Encoding]::new($false),
                'application/json'
            )
            $multipart.Add($commentPart, 'comment')

            $fileStream = [System.IO.File]::OpenRead($HtmlFilePath)
            $fileContent = [System.Net.Http.StreamContent]::new($fileStream)
            $fileContent.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('text/html')
            $leaf = [System.IO.Path]::GetFileName($HtmlFilePath)
            $multipart.Add($fileContent, 'files', $leaf)

            $response = $client.PostAsync($uri, $multipart).GetAwaiter().GetResult()
            $respBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            if (-not $response.IsSuccessStatusCode) {
                throw "Attach HTML failed: HTTP $([int]$response.StatusCode) $respBody"
            }
        } finally {
            $multipart.Dispose()
        }
    } finally {
        $client.Dispose()
    }
}

# ── Validation ────────────────────────────────────────────────────────────────

if ([string]::IsNullOrWhiteSpace($NinjaOneInstance)) {
    throw "NinjaOne instance is required. Set NINJAONE_INSTANCE or pass -NinjaOneInstance."
}
if ([string]::IsNullOrWhiteSpace($NinjaOneClientId)) {
    throw "Client ID is required. Set NINJAONE_CLIENT_ID or pass -NinjaOneClientId."
}
if ([string]::IsNullOrWhiteSpace($NinjaOneClientSecret)) {
    throw "Client secret is required. Set NINJAONE_CLIENT_SECRET or pass -NinjaOneClientSecret."
}
if ([string]::IsNullOrWhiteSpace($NinjaOneRefreshToken)) {
    throw "Refresh token is required. Set NINJAONE_REFRESH_TOKEN or pass -NinjaOneRefreshToken."
}
if ($TicketFormId -le 0) {
    throw '-TicketFormId must be a positive integer. Use GET /v2/ticketing/ticket-form to list forms.'
}

$resolvedManifest = $null
if (-not [string]::IsNullOrWhiteSpace($ManifestPath)) {
    $resolvedManifest = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ManifestPath)
} elseif (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $outDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
    $resolvedManifest = Join-Path $outDir $DefaultManifestName
} else {
    throw 'Specify -ManifestPath to the manifest JSON file, or -OutputPath to the folder that contains NinjaBillingInvoiceExport-manifest.json.'
}

if (-not (Test-Path -LiteralPath $resolvedManifest -PathType Leaf)) {
    throw "Manifest not found: $resolvedManifest"
}

$manifestDir = [System.IO.Path]::GetDirectoryName($resolvedManifest)

# ── Load manifest ─────────────────────────────────────────────────────────────

$rawJson = [System.IO.File]::ReadAllText($resolvedManifest, [System.Text.Encoding]::UTF8)
$parsed = $rawJson | ConvertFrom-Json
$rows = @()
if ($parsed -is [System.Array]) {
    $rows = @($parsed)
} elseif ($null -ne $parsed) {
    $rows = @($parsed)
}

if ($rows.Count -eq 0) {
    Write-Host 'Manifest contains no invoice rows. Exiting.'
    exit 0
}

# ── Main ──────────────────────────────────────────────────────────────────────

Write-Host "Authenticating to $NinjaOneInstance (refresh token)..."
$firstToken = Get-NinjaTokenFromRefreshToken -ClientId $NinjaOneClientId -ClientSecret $NinjaOneClientSecret `
    -RefreshToken $NinjaOneRefreshToken -Instance $NinjaOneInstance
$session = [PSCustomObject]@{
    Instance     = $NinjaOneInstance
    ClientId     = $NinjaOneClientId
    ClientSecret = $NinjaOneClientSecret
    RefreshToken = [string]$NinjaOneRefreshToken
    AuthHeader   = ''
    ExpiresAt    = $null
}
Apply-NinjaRefreshTokenResponse -Session $session -TokenResponse $firstToken

$ticketsCreated = 0
$failed = 0
$skipped = 0
$errors = [System.Collections.Generic.List[string]]::new()

foreach ($row in $rows) {
    $invNum = if ($row.invoiceNumber) { [string]$row.invoiceNumber } else { "INV-$($row.invoiceId)" }
    $displayClient = if ($row.clientName) { [string]$row.clientName } else { 'Unknown' }
    $htmlName = if ($row.htmlFileName) { [string]$row.htmlFileName } else { '' }
    $orgId = $row.clientOrganizationId

    if ([string]::IsNullOrWhiteSpace($htmlName)) {
        Write-Warning "[$invNum] Manifest row missing htmlFileName; skipping."
        $failed++
        [void]$errors.Add("[$invNum] Manifest row missing htmlFileName.")
        continue
    }
    if ($null -eq $orgId) {
        Write-Warning "[$invNum] Manifest row missing clientOrganizationId; skipping."
        $failed++
        [void]$errors.Add("[$invNum] Manifest row missing clientOrganizationId.")
        continue
    }
    $orgIdInt = 0
    if (-not [int]::TryParse("$orgId", [ref]$orgIdInt)) {
        Write-Warning "[$invNum] Invalid clientOrganizationId; skipping."
        $failed++
        [void]$errors.Add("[$invNum] Invalid clientOrganizationId.")
        continue
    }

    $htmlPath = Join-Path $manifestDir $htmlName
    if (-not (Test-Path -LiteralPath $htmlPath -PathType Leaf)) {
        Write-Warning "[$invNum] HTML file not found: $htmlPath"
        $failed++
        [void]$errors.Add("[$invNum] HTML file not found: $htmlPath")
        continue
    }

    $shouldDesc = "Create PSA ticket with HTML attachment for $invNum ($displayClient)"
    if (-not $PSCmdlet.ShouldProcess($htmlPath, $shouldDesc)) {
        $skipped++
        continue
    }

    try {
        Write-Host "  [$invNum] $displayClient..." -NoNewline
        $subject = Expand-TicketSubject -Template $TicketSubjectTemplate -InvoiceNumber $invNum -ClientName $displayClient
        $newTicketBody = @{
            clientId     = $orgIdInt
            ticketFormId = $TicketFormId
            status       = $TicketStatus
            subject      = $subject
            type         = $TicketType
            description  = @{ public = $true; body = $TicketDescriptionBody }
        }
        $created = Invoke-NinjaApi -Method POST -Endpoint '/v2/ticketing/ticket' -Body $newTicketBody -Session $session
        if ($null -eq $created.id) { throw 'Create ticket response did not include id.' }
        $tid = [int]$created.id
        Send-NinjaTicketCommentWithAttachment -Session $session -TicketId $tid -CommentBody $TicketAttachmentComment -HtmlFilePath $htmlPath
        Write-Host " ticket #$tid"
        $ticketsCreated++
    } catch {
        Write-Host " FAILED"
        $msg = "[$invNum] $($_.Exception.Message)"
        Write-Warning $msg
        [void]$errors.Add($msg)
        $failed++
    }
}

Write-Host ''
Write-Host ('─' * 50)
Write-Host "  Tickets created : $ticketsCreated"
if ($failed  -gt 0) { Write-Host "  Failed          : $failed" }
if ($skipped -gt 0) { Write-Host "  Skipped         : $skipped (WhatIf)" }
Write-Host ('─' * 50)

if ($failed -gt 0) {
    Write-Host ''
    Write-Host 'Failures:'
    $errors | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

<#
.SYNOPSIS
    Uploads PDF files to the NinjaOne Knowledge Base.

.DESCRIPTION
    Authenticates to NinjaOne with OAuth2 client credentials, then uploads each PDF
    via POST /v2/knowledgebase/articles/upload (multipart form). Omit -OrganizationId
    to target the global knowledge base.

    NOTE: Endpoint and form-field names are taken from third-party docs
    (homotechsual NinjaOne module), not the official reference. Test with one file first.

.EXAMPLE
    .\Upload-PdfToNinjaKB.ps1 -Path .\docs -Region app -FolderPath "Guides/HR" -WhatIf

.EXAMPLE
    $env:NINJA_CLIENT_ID = '...'; $env:NINJA_CLIENT_SECRET = '...'
    .\Upload-PdfToNinjaKB.ps1 -Path .\manual.pdf -OrganizationId 12 -Region eu
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    # A PDF file or a folder containing PDFs.
    [Parameter(Mandatory)][string]$Path,

    # Regional NinjaOne host.
    [ValidateSet('app', 'eu', 'ca', 'oc', 'us2')][string]$Region = 'ca',

    # Omit for the global knowledge base.
    [int]$OrganizationId,

    [int]$FolderId,
    [string]$FolderPath,

    [switch]$Recurse,

    [string]$ClientId     = $env:NINJA_CLIENT_ID,
    [string]$ClientSecret = $env:NINJA_CLIENT_SECRET
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http

if (-not $ClientId -or -not $ClientSecret) {
    throw 'Provide -ClientId/-ClientSecret or set NINJA_CLIENT_ID / NINJA_CLIENT_SECRET.'
}

$baseUrl = "https://$Region.ninjarmm.com"
$maxBytes = 20MB   # NinjaOne attachment limit (verify; may change)

# Collect PDFs
if (-not (Test-Path -LiteralPath $Path)) { throw "Path not found: $Path" }
$item = Get-Item -LiteralPath $Path
$files = if ($item.PSIsContainer) {
    Get-ChildItem -LiteralPath $item.FullName -Filter *.pdf -File -Recurse:$Recurse
} else {
    if ($item.Extension -ne '.pdf') { throw "Not a PDF: $($item.FullName)" }
    @($item)
}
if (-not $files) { Write-Warning 'No PDF files found.'; return }

# Authenticate
$tokenResponse = Invoke-RestMethod -Method Post -Uri "$baseUrl/ws/oauth/token" -Body @{
    grant_type    = 'client_credentials'
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = 'management'
}
$headers = @{ Authorization = "Bearer $($tokenResponse.access_token)" }

$client = [System.Net.Http.HttpClient]::new()
$client.DefaultRequestHeaders.Authorization =
    [System.Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $tokenResponse.access_token)

function Add-JsonField($content, $name, $value) {
    $part = [System.Net.Http.StringContent]::new(($value | ConvertTo-Json -Compress), [Text.Encoding]::UTF8, 'application/json')
    $content.Add($part, $name)
}

$results = foreach ($file in $files) {
    $status = 'Uploaded'
    $detail = $null
    try {
        if ($file.Length -gt $maxBytes) { throw "Exceeds $($maxBytes / 1MB) MB limit." }

        if ($PSCmdlet.ShouldProcess($file.FullName, 'Upload to NinjaOne KB')) {
            $form = [System.Net.Http.MultipartFormDataContent]::new()
            if ($PSBoundParameters.ContainsKey('OrganizationId')) { Add-JsonField $form 'organizationId' $OrganizationId }
            if ($PSBoundParameters.ContainsKey('FolderId'))       { Add-JsonField $form 'folderId' $FolderId }
            if ($FolderPath)                                      { Add-JsonField $form 'folderPath' $FolderPath }

            $stream = [System.IO.File]::OpenRead($file.FullName)
            try {
                $filePart = [System.Net.Http.StreamContent]::new($stream)
                $filePart.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse('application/octet-stream')
                $form.Add($filePart, 'files', $file.Name)

                $response = $client.PostAsync("$baseUrl/v2/knowledgebase/articles/upload", $form).GetAwaiter().GetResult()
                $body = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                if (-not $response.IsSuccessStatusCode) { throw "HTTP $([int]$response.StatusCode): $body" }
                $detail = $body
            } finally {
                $stream.Dispose(); $form.Dispose()
            }
        } else {
            $status = 'Skipped (WhatIf)'
        }
    } catch {
        $status = 'Failed'
        $detail = $_.Exception.Message
    }
    [pscustomobject]@{ File = $file.Name; Status = $status; Detail = $detail }
}

$client.Dispose()
$results | Format-Table -AutoSize
if ($results.Status -contains 'Failed') { exit 1 }

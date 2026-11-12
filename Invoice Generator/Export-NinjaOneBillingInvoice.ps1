<#
.SYNOPSIS
    Exports NinjaOne billing invoices for a given month as HTML (and optionally PDF) files.

.DESCRIPTION
    Authenticates to the NinjaOne API via OAuth2 client credentials, retrieves billing
    invoices for the specified month, fetches full invoice details including all line items,
    and renders each as a styled UTF-8 HTML document under -OutputPath. With -ConvertToPdf,
    each HTML file is also printed to PDF using headless Microsoft Edge or Google Chrome.

    After a successful run, writes NinjaBillingInvoiceExport-manifest.json in -OutputPath
    listing each exported invoice (invoice id, organization client id, paths) for use with
    New-NinjaOneBillingInvoiceTickets.ps1, which uses a refresh token for PSA ticketing.
    The manifest describes only the invoices exported by THIS run and replaces any
    existing manifest in -OutputPath (use a separate -OutputPath per run/period if you
    need to keep earlier manifests).

    Testing: run with -WhatIf to skip disk and manifest; zero invoices exits 0.

    No third-party PowerShell modules required.

.PARAMETER NinjaOneInstance
    NinjaOne instance hostname (e.g. app.ninjarmm.com). Defaults to NINJAONE_INSTANCE env var.

.PARAMETER NinjaOneClientId
    OAuth2 client ID. Defaults to NINJAONE_CLIENT_ID env var.

.PARAMETER NinjaOneClientSecret
    OAuth2 client secret. Defaults to NINJAONE_CLIENT_SECRET env var.

.PARAMETER Month
    Month number (1-12) for the billing period. Defaults to previous month.

.PARAMETER Year
    Year for the billing period. Defaults to the year of the previous month
    (so running in January defaults to December of last year).

.PARAMETER OutputPath
    Folder where invoices will be saved. Created if it does not exist.
    Relative paths are resolved against the current PowerShell location.

.PARAMETER OrganizationId
    Optional. Filter invoices to a specific NinjaOne organization/client ID.
    (Alias: -ClientId, kept for backward compatibility. Not the OAuth client id.)

.PARAMETER IssuerName
.PARAMETER IssuerAddress
.PARAMETER IssuerPhone
.PARAMETER IssuerWebPage
    Optional "From" block. Used when the invoice has no billContent from the API
    (observed for PENDING/APPROVED invoices). API values win when present.

.PARAMETER ConvertToPdf
    Also render each invoice to PDF using headless Edge or Chrome (must be installed).

.PARAMETER ExportCsv
    Also write two CSV files to -OutputPath, laid out like the NinjaOne UI exports:
      NinjaInvoices-Summary_<yyyy>-<mm>.csv     one row per invoice
      NinjaInvoices-LineItems_<yyyy>-<mm>.csv   one row per invoice line item
    Created/Last Updated are shown in this computer's local time zone. Text cells that
    start with = + - @ get a leading apostrophe so spreadsheets don't treat them as formulas.

.PARAMETER CsvPerOrganization
    Implies -ExportCsv. Instead of one combined pair of CSVs, writes a pair per organization:
      NinjaInvoices-Summary_<Organization>_<yyyy>-<mm>.csv
      NinjaInvoices-LineItems_<Organization>_<yyyy>-<mm>.csv
    Organizations that share a name get their organization id appended to the file name.

.PARAMETER CsvOnly
    Write only the CSV files (implies -ExportCsv). No HTML/PDF and no ticket manifest is
    produced. Cannot be combined with -PdfOnly / -ConvertToPdf.

.PARAMETER PdfOnly
    Output PDF files only. Implies -ConvertToPdf; the intermediate HTML is created in a
    temp folder and deleted. The manifest then lists pdfFileName (no htmlFileName), and
    New-NinjaOneBillingInvoiceTickets.ps1 attaches the PDF.

.PARAMETER Status
    One or more invoice statuses to export. Defaults to COMPLETE.
    Valid values: PENDING, APPROVED, COMPLETE, FAILED, ARCHIVED, ANY (no status filter).
    Example: -Status PENDING,APPROVED

.EXAMPLE
    .\Export-NinjaOneBillingInvoices.ps1 -Month 3 -Year 2026 -OutputPath .\Invoices

.EXAMPLE
    .\Export-NinjaOneBillingInvoices.ps1 -Status APPROVED -OrganizationId 42 -ConvertToPdf

.EXAMPLE
    .\Export-NinjaOneBillingInvoices.ps1 -Status PENDING,APPROVED -PdfOnly

.EXAMPLE
    .\Export-NinjaOneBillingInvoices.ps1 -Status PENDING,APPROVED -CsvOnly

.EXAMPLE
    .\Export-NinjaOneBillingInvoices.ps1 -Status PENDING -PdfOnly -ExportCsv

.EXAMPLE
    .\Export-NinjaOneBillingInvoices.ps1 -Status PENDING,APPROVED -CsvOnly -CsvPerOrganization

.EXAMPLE
    .\Export-NinjaOneBillingInvoices.ps1 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$NinjaOneInstance     = $env:NINJAONE_INSTANCE,
    [string]$NinjaOneClientId     = $env:NINJAONE_CLIENT_ID,
    [string]$NinjaOneClientSecret = $env:NINJAONE_CLIENT_SECRET,

    [ValidateRange(1, 12)]
    [int]$Month = (Get-Date).AddMonths(-1).Month,

    [ValidateRange(2000, 2100)]
    [int]$Year = (Get-Date).AddMonths(-1).Year,

    [string]$OutputPath = '.\Invoices',

    [Alias('ClientId')]
    [int]$OrganizationId,

    [ValidateSet('PENDING', 'APPROVED', 'COMPLETE', 'FAILED', 'ARCHIVED', 'ANY')]
    [string[]]$Status = 'COMPLETE',

    [string]$IssuerName,
    [string]$IssuerAddress,
    [string]$IssuerPhone,
    [string]$IssuerWebPage,

    [switch]$ConvertToPdf,

    [switch]$PdfOnly,

    [switch]$ExportCsv,

    [switch]$CsvOnly,

    [switch]$CsvPerOrganization
)

$ErrorActionPreference = 'Stop'
if ($PdfOnly) { $ConvertToPdf = [switch]::new($true) }
if ($CsvOnly) {
    if ($ConvertToPdf) { throw '-CsvOnly cannot be combined with -PdfOnly or -ConvertToPdf.' }
    $ExportCsv = [switch]::new($true)
}
if ($CsvPerOrganization) { $ExportCsv = [switch]::new($true) }
Add-Type -AssemblyName System.Web

# Windows PowerShell 5.1 may default to older TLS versions
if ($PSVersionTable.PSVersion.Major -lt 6) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
}

$ManifestFileName = 'NinjaBillingInvoiceExport-manifest.json'

# -- Auth ----------------------------------------------------------------------

function Get-NinjaToken {
    param(
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$Instance
    )
    $uri  = "https://$($Instance.TrimEnd('/'))/ws/oauth/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        grant_type    = 'client_credentials'
        scope         = 'monitoring management'
    }
    $headers = @{
        'Accept'       = 'application/json'
        'Content-Type' = 'application/x-www-form-urlencoded'
    }
    return Invoke-RestMethod -Uri $uri -Method POST -Headers $headers -Body $body -TimeoutSec 30 -ErrorAction Stop
}

# -- API wrapper with auto-refresh and exponential backoff ---------------------

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
            $h = $Err.Exception.Response.Headers
            # PowerShell 7 (HttpResponseHeaders): typed RetryAfter property
            if ($h.PSObject.Properties['RetryAfter'] -and $h.RetryAfter -and $h.RetryAfter.Delta) {
                return [int][Math]::Ceiling($h.RetryAfter.Delta.TotalSeconds)
            }
        } catch { }
        try {
            # Windows PowerShell 5.1 (WebHeaderCollection): string indexer
            $raw = $Err.Exception.Response.Headers['Retry-After'] | Select-Object -First 1
            $sec = 0
            if ($raw -and [int]::TryParse([string]$raw, [ref]$sec) -and $sec -gt 0) { return $sec }
        } catch { }
        return $null
    }
    function Get-ApiErrorDetail($Err) {
        try {
            if ($Err.ErrorDetails -and $Err.ErrorDetails.Message) { return [string]$Err.ErrorDetails.Message }
        } catch { }
        try {
            $stream = $Err.Exception.Response.GetResponseStream()
            if ($stream) {
                $stream.Position = 0
                return (New-Object System.IO.StreamReader($stream)).ReadToEnd()
            }
        } catch { }
        return $null
    }
    function Update-Session([PSCustomObject]$S) {
        $t = Get-NinjaToken -ClientId $S.ClientId -ClientSecret $S.ClientSecret -Instance $S.Instance
        $S.AuthHeader = "Bearer $($t.access_token)"
        $S.ExpiresAt  = if ($t.expires_in) { (Get-Date).AddSeconds([int]$t.expires_in - 60) } else { (Get-Date).AddMinutes(55) }
    }

    if (-not $Session.ExpiresAt -or (Get-Date) -ge $Session.ExpiresAt) { Update-Session $Session }

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
            if ($status -eq 401 -and $attempt -le $MaxRetries) { Update-Session $Session; continue }
            $retryable = $status -in @(408, 429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -gt $MaxRetries) {
                $detail = Get-ApiErrorDetail $_
                if ($detail) { throw "$($_.Exception.Message) -- $detail" }
                throw
            }
            $wait  = Get-RetryAfter $_
            $sleep = if ($wait -and $wait -gt 0) { [Math]::Min($wait, 60) } else { [Math]::Min([Math]::Pow(2, $attempt), 30) }
            Write-Warning "HTTP $status - retrying in ${sleep}s (attempt $attempt/$MaxRetries)"
            Start-Sleep -Seconds $sleep
        }
    }
}

# -- Formatting helpers --------------------------------------------------------

function Format-Money {
    param([object]$Amount, [string]$Currency = 'USD')
    if ($null -eq $Amount) { return '-' }
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    return "$(([double]$Amount).ToString('N2', $inv)) $(HtmlEnc $Currency)"
}

function Format-InvoiceDate {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    try { return ([datetime]$Value).ToString('MMM d, yyyy', [System.Globalization.CultureInfo]::InvariantCulture) }
    catch { return (HtmlEnc $Value) }
}

function HtmlEnc([string]$s) {
    return [System.Web.HttpUtility]::HtmlEncode($s)
}

# HTML-encode and keep line breaks (addresses are often multi-line)
function HtmlEncMultiline([string]$s) {
    return ((HtmlEnc $s) -replace "\r?\n", '<br>')
}

# Only allow http(s) URLs in href attributes
function Test-HttpUrl([string]$s) {
    return ($s -match '^https?://[^\s''"<>]+$')
}

# -- File helpers --------------------------------------------------------------

# Make a string safe to use as part of a Windows file name
function Get-SafeFileName {
    param([string]$Name, [int]$MaxLength = 80)
    $s = $Name -replace '[\x00-\x1F\\/:*?"<>|]', '_'
    $s = $s.Trim().TrimEnd('.', ' ')
    if ($s.Length -gt $MaxLength) { $s = $s.Substring(0, $MaxLength).TrimEnd('.', ' ') }
    if ([string]::IsNullOrWhiteSpace($s)) { $s = '_' }
    if ($s -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$') { $s = "_$s" }
    return $s
}

# [System.IO.File] resolves relative paths against the .NET process directory, not the
# PowerShell location, so always hand it absolute paths.
function Resolve-OutputFolder {
    param([string]$Path)
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
}

function Find-PdfBrowser {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )
    foreach ($c in $candidates) { if ($c -and (Test-Path -LiteralPath $c)) { return $c } }
    return $null
}

function Convert-HtmlToPdf {
    param([string]$Browser, [string]$HtmlPath, [string]$PdfPath)
    $profileDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ninja-pdf-" + [guid]::NewGuid().ToString('N'))
    try {
        $fileUri = ([System.Uri]$HtmlPath).AbsoluteUri
        $browserArgs = @(
            '--headless', '--disable-gpu', '--no-pdf-header-footer',
            "--user-data-dir=`"$profileDir`"",
            "--print-to-pdf=`"$PdfPath`"",
            "`"$fileUri`""
        )
        $p = Start-Process -FilePath $Browser -ArgumentList $browserArgs -Wait -PassThru -WindowStyle Hidden
        if (-not (Test-Path -LiteralPath $PdfPath)) {
            throw "PDF was not created (browser exit code $($p.ExitCode))."
        }
    } finally {
        if (Test-Path -LiteralPath $profileDir) { Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# -- CSV helpers (column layout mirrors the NinjaOne UI exports) ---------------

$CsvInvariant = [System.Globalization.CultureInfo]::InvariantCulture

$SummaryCsvHeader = @('Invoice ID', 'Status', 'Number', 'Agreement Name', 'Organization', 'Invoice generation type',
                      'Currency', 'Subtotal', 'Total Tax', 'Total', 'Billing Period Start Date', 'Billing Period End Date',
                      'Due Date', 'Created', 'Last Updated')
$LineCsvHeader    = @('Invoice ID', 'Number', 'Organization', 'Product Name', 'Description', 'Type', 'Taxable', 'Currency',
                      'Quantity', 'Billing Period Start Date', 'Billing Period End Date', 'Unit Price', 'Subtotal')

# Quote only when needed (like the UI export). -Text fields that start with = + - @ are
# prefixed with an apostrophe so spreadsheet apps don't evaluate them as formulas.
function ConvertTo-CsvField {
    param($Value, [switch]$Text)
    if ($null -eq $Value) { return '' }
    $s = [string]$Value
    if ($Text -and $s -match '^[=+\-@\t\r]') { $s = "'" + $s }
    if ($s -match '[",\r\n]') { $s = '"' + $s.Replace('"', '""') + '"' }
    return $s
}

function ConvertTo-CsvLine {
    param([object[]]$Fields, [bool[]]$IsText)
    $out = for ($i = 0; $i -lt $Fields.Count; $i++) {
        if ($IsText[$i]) { ConvertTo-CsvField $Fields[$i] -Text } else { ConvertTo-CsvField $Fields[$i] }
    }
    return ($out -join ',')
}

function Format-CsvMoney($v) {
    if ($null -eq $v -or "$v" -eq '') { return '' }
    return ([double]$v).ToString('0.00####', $CsvInvariant)
}
function Format-CsvQuantity($v) {
    if ($null -eq $v -or "$v" -eq '') { return '' }
    return ([double]$v).ToString('0.########', $CsvInvariant)
}
function Format-CsvDate($v) {
    if ($null -eq $v -or "$v" -eq '') { return '' }
    try { return ([datetime]::Parse([string]$v, $CsvInvariant)).ToString('M/d/yyyy', $CsvInvariant) } catch { return [string]$v }
}
# Accepts epoch seconds (list endpoint) or an ISO timestamp (detail endpoint); renders local time
function Format-CsvDateTime($v) {
    if ($null -eq $v -or "$v" -eq '') { return '' }
    try {
        $dt = if ($v -is [string]) { [DateTimeOffset]::Parse($v, $CsvInvariant) }
              else { [DateTimeOffset]::FromUnixTimeMilliseconds([long]([double]$v * 1000)) }
        return $dt.LocalDateTime.ToString('M/d/yyyy, h:mm tt', $CsvInvariant)
    } catch { return [string]$v }
}

function New-SummaryCsvLine {
    param($ListInvoice, [string]$OrganizationName)
    $i = $ListInvoice
    return ConvertTo-CsvLine -Fields @(
        $i.id, $i.status, $i.invoiceNumber, $i.agreementName, $OrganizationName, $i.invoiceGenerationType,
        $i.currency, (Format-CsvMoney $i.subTotal), (Format-CsvMoney $i.totalTax), (Format-CsvMoney $i.total),
        (Format-CsvDate $i.billingPeriodStartDate), (Format-CsvDate $i.billingPeriodEndDate), (Format-CsvDate $i.dueDate),
        (Format-CsvDateTime $i.createTime), (Format-CsvDateTime $i.updateTime)
    ) -IsText @($false, $false, $false, $true, $true, $false, $false, $false, $false, $false, $false, $false, $false, $false, $false)
}

function New-LineCsvLines {
    param($ListInvoice, $Detail, [string]$OrganizationName)
    $products = @()
    foreach ($n in 'products', 'agreementProducts', 'ticketProducts', 'oneOffProducts') {
        if ($Detail.PSObject.Properties[$n] -and $Detail.$n) { $products += @($Detail.$n) }
    }
    foreach ($p in $products) {
        $sub = if ($null -ne $p.subTotal) { $p.subTotal } else { $p.subTotalWithDiscount }
        $taxable = if ($null -ne $p.taxable) { ([bool]$p.taxable).ToString().ToLowerInvariant() } else { '' }
        $start = if ($p.billingPeriodStartDate) { $p.billingPeriodStartDate } else { $Detail.billingPeriodStartDate }
        $end   = if ($p.billingPeriodEndDate)   { $p.billingPeriodEndDate }   else { $Detail.billingPeriodEndDate }
        ConvertTo-CsvLine -Fields @(
            $ListInvoice.id, $ListInvoice.invoiceNumber, $OrganizationName, $p.name, $p.description, $p.quantityType,
            $taxable, $Detail.currency, (Format-CsvQuantity $p.quantity), (Format-CsvDate $start), (Format-CsvDate $end),
            (Format-CsvMoney $p.price), (Format-CsvMoney $sub)
        ) -IsText @($false, $false, $true, $true, $true, $false, $false, $false, $false, $false, $false, $false, $false)
    }
}

function Write-CsvFile {
    param([string]$Path, [string[]]$Header, [string[]]$Lines, $Encoding)
    $text = (@($Header -join ',') + @($Lines)) -join "`r`n"
    [System.IO.File]::WriteAllText($Path, $text + "`r`n", $Encoding)
}

# -- HTML invoice builder ------------------------------------------------------

function Build-InvoiceHtml {
    param(
        [PSCustomObject]$Invoice,
        [string]$ClientName,          # fallback "Bill To" name (list/detail client.name is often null)
        [hashtable]$IssuerFallback    # fallback "From" block: Name/Address/Phone/WebPage
    )

    $cur        = if ($Invoice.currency) { $Invoice.currency } else { 'USD' }
    $invoiceNum = HtmlEnc $(if ($Invoice.invoiceNumber) { $Invoice.invoiceNumber } else { "INV-$($Invoice.id)" })

    # Pre-compute conditional blocks so the here-string stays readable
    $dueDateHtml = if ($Invoice.dueDate) { "<p>Due: $(Format-InvoiceDate $Invoice.dueDate)</p>" } else { '' }

    $issuer = $Invoice.billContent
    $fromLines = [System.Collections.Generic.List[string]]::new()
    $toLines   = [System.Collections.Generic.List[string]]::new()
    if ($issuer) {
        if ($issuer.issuerName)    { $fromLines.Add("<strong>$(HtmlEnc $issuer.issuerName)</strong>") }
        if ($issuer.issuerAddress) { $fromLines.Add((HtmlEncMultiline $issuer.issuerAddress)) }
        if ($issuer.issuerPhone)   { $fromLines.Add((HtmlEnc $issuer.issuerPhone)) }
        if ($issuer.issuerWebPage) {
            if (Test-HttpUrl $issuer.issuerWebPage) {
                $fromLines.Add("<a href='$(HtmlEnc $issuer.issuerWebPage)'>$(HtmlEnc $issuer.issuerWebPage)</a>")
            } else {
                $fromLines.Add((HtmlEnc $issuer.issuerWebPage))
            }
        }

        if ($issuer.billName)    { $toLines.Add("<strong>$(HtmlEnc $issuer.billName)</strong>") }
        if ($issuer.billAddress) { $toLines.Add((HtmlEncMultiline $issuer.billAddress)) }
        if ($issuer.billEmail)   { $toLines.Add((HtmlEnc $issuer.billEmail)) }
        if ($issuer.billPhone)   { $toLines.Add((HtmlEnc $issuer.billPhone)) }
    }
    if ($fromLines.Count -eq 0 -and $IssuerFallback) {
        if ($IssuerFallback.Name)    { $fromLines.Add("<strong>$(HtmlEnc $IssuerFallback.Name)</strong>") }
        if ($IssuerFallback.Address) { $fromLines.Add((HtmlEncMultiline $IssuerFallback.Address)) }
        if ($IssuerFallback.Phone)   { $fromLines.Add((HtmlEnc $IssuerFallback.Phone)) }
        if ($IssuerFallback.WebPage) {
            if (Test-HttpUrl $IssuerFallback.WebPage) {
                $fromLines.Add("<a href='$(HtmlEnc $IssuerFallback.WebPage)'>$(HtmlEnc $IssuerFallback.WebPage)</a>")
            } else {
                $fromLines.Add((HtmlEnc $IssuerFallback.WebPage))
            }
        }
    }
    if ($toLines.Count -eq 0) {
        $nm = if ($Invoice.client -and $Invoice.client.name) { $Invoice.client.name } else { $ClientName }
        if ($nm) { $toLines.Add("<strong>$(HtmlEnc $nm)</strong>") }
    }
    $fromHtml = $fromLines -join '<br>'
    $toHtml   = $toLines   -join '<br>'

    $agreementHtml = ''
    if ($Invoice.agreement -and $Invoice.agreement.name) {
        $agreementHtml = "<div class='meta-item'><span class='meta-label'>Agreement</span>$(HtmlEnc $Invoice.agreement.name)</div>"
    }
    $intervalHtml = ''
    if ($Invoice.interval) {
        $intervalHtml = "<div class='meta-item'><span class='meta-label'>Interval</span>$(HtmlEnc $Invoice.interval)</div>"
    }

    # Merge all product arrays
    $allProducts = @()
    if ($Invoice.products)          { $allProducts += @($Invoice.products) }
    if ($Invoice.agreementProducts) { $allProducts += @($Invoice.agreementProducts) }
    if ($Invoice.ticketProducts)    { $allProducts += @($Invoice.ticketProducts) }
    if ($Invoice.oneOffProducts)    { $allProducts += @($Invoice.oneOffProducts) }

    $lineRowsHtml = [System.Text.StringBuilder]::new()
    foreach ($p in $allProducts) {
        $name      = if ($p.name)        { HtmlEnc $p.name }        else { '(unnamed)' }
        $desc      = if ($p.description) { HtmlEnc $p.description } else { '' }
        $nameCell  = if ($desc) { "<strong>$name</strong><br><span class='subdesc'>$desc</span>" } else { "<strong>$name</strong>" }
        $qty       = if ($null -ne $p.quantity) { [double]$p.quantity } else { 0 }
        $unitPrice = if ($null -ne $p.price)    { [double]$p.price }    else { 0 }
        $lineAmt   = if ($null -ne $p.subTotalWithDiscount) { [double]$p.subTotalWithDiscount }
                     elseif ($null -ne $p.subTotal)         { [double]$p.subTotal }
                     else                                   { $unitPrice * $qty }
        $discCell  = if ($p.discount -and [double]$p.discount -ne 0) {
                         "<br><span class='subdesc negative'>(- $(Format-Money $p.discount $cur))</span>"
                     } else { '' }

        $qtyCell   = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.####}', $qty)
        if ($p.quantityType) { $qtyCell += "<br><span class='subdesc'>$(HtmlEnc $p.quantityType)</span>" }

        [void]$lineRowsHtml.AppendLine(
            "<tr><td class='td-name'>$nameCell</td><td class='td-num'>$qtyCell</td>" +
            "<td class='td-num'>$(Format-Money $unitPrice $cur)</td>" +
            "<td class='td-num'><span class='line-total'>$(Format-Money $lineAmt $cur)</span>$discCell</td></tr>"
        )
    }

    $discountRowHtml = ''
    if ($Invoice.discount -and [double]$Invoice.discount -ne 0) {
        $discountRowHtml = "<tr><td class='td-label'>Discount</td><td class='td-num negative'>- $(Format-Money $Invoice.discount $cur)</td></tr>"
    }
    # taxRate is normally a fraction (0.0825); treat values > 1 as already-percent (8.25)
    $taxPct = ''
    if ($Invoice.taxRate) {
        $rate = [double]$Invoice.taxRate
        if ($rate -gt 1) { $rate = $rate / 100 }
        $taxPct = ($rate * 100).ToString('0.##', [System.Globalization.CultureInfo]::InvariantCulture) + '%'
    }
    $taxLabel    = if ($taxPct) { "Tax ($taxPct)" } else { 'Tax' }

    $notesHtml = ''
    if (-not [string]::IsNullOrWhiteSpace($Invoice.invoiceNote)) {
        $notesHtml = "<div class='notes-section'><h3>Notes</h3><p>$(HtmlEnc $Invoice.invoiceNote)</p></div>"
    }

    $generatedDate = (Get-Date).ToString('MMMM d, yyyy')
    $periodStart   = Format-InvoiceDate $Invoice.billingPeriodStartDate
    $periodEnd     = Format-InvoiceDate $Invoice.billingPeriodEndDate
    $periodText    = "$periodStart &ndash; $periodEnd"
    $invDate       = Format-InvoiceDate $Invoice.invoiceDate
    $statusVal     = if ($Invoice.status) { HtmlEnc $Invoice.status } else { '' }
    $paymentHtml   = ''
    if ($Invoice.paymentStatus) {
        $paymentHtml = "<p>Payment: $(HtmlEnc $Invoice.paymentStatus)"
        if ($null -ne $Invoice.outstandingBalance) { $paymentHtml += " (outstanding $(Format-Money $Invoice.outstandingBalance $cur))" }
        $paymentHtml += '</p>'
    }
    $subtotal      = Format-Money $Invoice.subTotal $cur
    $totalTax      = Format-Money $Invoice.totalTax $cur
    $grandTotal    = Format-Money $Invoice.total $cur
    $lineRows      = $lineRowsHtml.ToString()

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Invoice $invoiceNum</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; color: #222; background: #fff; padding: 32px 40px; }
  .header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 28px; border-bottom: 3px solid #1a56db; padding-bottom: 18px; }
  .invoice-title { font-size: 36px; font-weight: 700; color: #1a56db; letter-spacing: 2px; }
  .invoice-meta { text-align: right; }
  .inv-num { font-size: 18px; font-weight: 600; color: #1a56db; }
  .invoice-meta p { margin-top: 4px; color: #555; }
  .parties { display: flex; gap: 48px; margin-bottom: 24px; }
  .party-block { flex: 1; }
  .party-block h3 { font-size: 11px; text-transform: uppercase; letter-spacing: 1px; color: #888; margin-bottom: 6px; }
  .party-block p { line-height: 1.7; }
  .meta-section { background: #f7f8fc; border-radius: 6px; padding: 14px 18px; margin-bottom: 24px; display: flex; flex-wrap: wrap; gap: 24px; }
  .meta-item { flex: 1; min-width: 150px; }
  .meta-label { font-weight: 600; color: #555; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; display: block; margin-bottom: 2px; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
  thead tr { background: #1a56db; color: #fff; }
  thead th { padding: 10px 14px; text-align: left; font-size: 12px; font-weight: 600; }
  thead th.th-num { text-align: right; }
  tbody tr:nth-child(even) { background: #f7f8fc; }
  td { padding: 10px 14px; vertical-align: top; border-bottom: 1px solid #e5e7eb; }
  .td-num { text-align: right; }
  .td-name { width: 45%; }
  .subdesc { color: #777; font-size: 11px; }
  .line-total { font-weight: 600; }
  .negative { color: #b91c1c; }
  .totals-wrap { display: flex; justify-content: flex-end; margin-bottom: 20px; }
  .totals-table { width: 320px; border-collapse: collapse; }
  .totals-table td { padding: 6px 12px; border: none; }
  .td-label { text-align: right; color: #555; }
  .totals-divider td { border-top: 1px solid #d1d5db; }
  .grand-total td { font-size: 15px; font-weight: 700; color: #1a56db; padding-top: 10px; }
  .notes-section { margin-top: 20px; padding: 14px 16px; background: #fffbeb; border-left: 3px solid #f59e0b; border-radius: 4px; }
  .notes-section h3 { font-size: 11px; text-transform: uppercase; color: #92400e; margin-bottom: 6px; }
  .notes-section p { line-height: 1.6; }
  .footer { margin-top: 36px; border-top: 1px solid #e5e7eb; padding-top: 10px; text-align: center; font-size: 11px; color: #aaa; }
  @page { size: Letter; margin: 14mm; }
  @media print {
    body { padding: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    tr { page-break-inside: avoid; }
  }
</style>
</head>
<body>

<div class="header">
  <div class="invoice-title">INVOICE</div>
  <div class="invoice-meta">
    <div class="inv-num">$invoiceNum</div>
    <p>Date: $invDate</p>
    $dueDateHtml
    <p>Status: $statusVal</p>
    $paymentHtml
  </div>
</div>

<div class="parties">
  <div class="party-block">
    <h3>From</h3>
    <p>$fromHtml</p>
  </div>
  <div class="party-block">
    <h3>Bill To</h3>
    <p>$toHtml</p>
  </div>
</div>

<div class="meta-section">
  <div class="meta-item">
    <span class="meta-label">Billing Period</span>
    $periodText
  </div>
  $agreementHtml
  $intervalHtml
</div>

<table>
  <thead>
    <tr>
      <th>Description</th>
      <th class="th-num">Qty</th>
      <th class="th-num">Unit Price</th>
      <th class="th-num">Amount</th>
    </tr>
  </thead>
  <tbody>
    $lineRows
  </tbody>
</table>

<div class="totals-wrap">
  <table class="totals-table">
    <tr>
      <td class="td-label">Subtotal</td>
      <td class="td-num">$subtotal</td>
    </tr>
    $discountRowHtml
    <tr>
      <td class="td-label">$taxLabel</td>
      <td class="td-num">$totalTax</td>
    </tr>
    <tr class="totals-divider"><td colspan="2"></td></tr>
    <tr class="grand-total">
      <td class="td-label">Total</td>
      <td class="td-num">$grandTotal</td>
    </tr>
  </table>
</div>

$notesHtml

<div class="footer">Generated $generatedDate &middot; NinjaOne Billing</div>

</body>
</html>
"@
    return $html
}

# -- Validation ----------------------------------------------------------------

if ([string]::IsNullOrWhiteSpace($NinjaOneInstance))     { throw "NinjaOne instance is required. Set NINJAONE_INSTANCE or pass -NinjaOneInstance." }
if ([string]::IsNullOrWhiteSpace($NinjaOneClientId))     { throw "Client ID is required. Set NINJAONE_CLIENT_ID or pass -NinjaOneClientId." }
if ([string]::IsNullOrWhiteSpace($NinjaOneClientSecret)) { throw "Client secret is required. Set NINJAONE_CLIENT_SECRET or pass -NinjaOneClientSecret." }

# -- Main ----------------------------------------------------------------------

Write-Host "Authenticating to $NinjaOneInstance..."
$tokenResp = Get-NinjaToken -ClientId $NinjaOneClientId -ClientSecret $NinjaOneClientSecret -Instance $NinjaOneInstance
$session = [PSCustomObject]@{
    Instance     = $NinjaOneInstance
    ClientId     = $NinjaOneClientId
    ClientSecret = $NinjaOneClientSecret
    AuthHeader   = "Bearer $($tokenResp.access_token)"
    ExpiresAt    = if ($tokenResp.expires_in) { (Get-Date).AddSeconds([int]$tokenResp.expires_in - 60) } else { (Get-Date).AddMinutes(55) }
}

# Compute billing period date range
$periodFrom = "$Year-$($Month.ToString('D2'))-01"
$lastDay    = [DateTime]::DaysInMonth($Year, $Month)
$periodTo   = "$Year-$($Month.ToString('D2'))-$($lastDay.ToString('D2'))"
Write-Host "Fetching invoices for period $periodFrom to $periodTo (status filter: $($Status -join ', '))..."

$query = "periodFrom=$periodFrom&periodTo=$periodTo"
if ($PSBoundParameters.ContainsKey('OrganizationId')) { $query += "&clientId=$OrganizationId" }

$rawList = Invoke-NinjaApi -Endpoint '/v2/billing/invoices' -Query $query -Session $session
$invoiceList = if ($rawList -is [Array])                                 { @($rawList) }
               elseif ($rawList.PSObject.Properties['results'])          { @($rawList.results) }
               elseif ($rawList.PSObject.Properties['data'])             { @($rawList.data) }
               else                                                      { @($rawList) }

# This script does not follow pages. If the response looks paged, warn instead of
# silently exporting a partial set.
if ($rawList -isnot [Array] -and $rawList) {
    $pagingProps = @('cursor', 'nextCursor', 'next', 'nextPage', 'nextPageToken', 'after', 'hasMore', 'totalPages', 'pageCount')
    $found = @($pagingProps | Where-Object { $rawList.PSObject.Properties[$_] -and $rawList.($_) })
    if ($found.Count -gt 0) {
        Write-Warning "Invoice list response contains paging fields ($($found -join ', ')); results may be incomplete. Pagination is not implemented."
    }
}

# Status is not a supported query param on this endpoint - filter client-side
$statusLabel = $Status -join ', '
if ($Status -notcontains 'ANY') {
    $invoiceList = @($invoiceList | Where-Object { $Status -contains $_.status })
}
Write-Host "Found $($invoiceList.Count) invoice(s) with status '$statusLabel'."

if ($invoiceList.Count -eq 0) {
    Write-Host "No invoices to process. Exiting."
    exit 0
}

$OutputPath = Resolve-OutputFolder $OutputPath
if (-not $WhatIfPreference -and -not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$pdfBrowser = $null
if ($ConvertToPdf) {
    $pdfBrowser = Find-PdfBrowser
    if (-not $pdfBrowser) { throw "-ConvertToPdf requires Microsoft Edge or Google Chrome, but neither was found." }
}

# With -PdfOnly the intermediate HTML lives in a temp folder and is deleted after each PDF
$htmlDir = $OutputPath
if ($PdfOnly -and -not $WhatIfPreference) {
    $htmlDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ninja-invoice-html-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $htmlDir -Force | Out-Null
}

$utf8Bom   = New-Object System.Text.UTF8Encoding($true)
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# Invoice client.name is often null, so resolve organization names once: id -> name
$orgNames = @{}
try {
    $orgRaw = Invoke-NinjaApi -Endpoint '/v2/organizations' -Session $session
    $orgList = if ($orgRaw -is [Array]) { @($orgRaw) }
               elseif ($orgRaw.PSObject.Properties['results']) { @($orgRaw.results) }
               else { @($orgRaw) }
    foreach ($o in $orgList) { if ($null -ne $o.id -and $o.name) { $orgNames[[int]$o.id] = [string]$o.name } }
    Write-Host "Resolved $($orgNames.Count) organization name(s)."
} catch {
    Write-Warning "Could not load organization names ($($_.Exception.Message)); falling back to 'Customer #<id>'."
}

$issuerFallback = $null
if ($IssuerName -or $IssuerAddress -or $IssuerPhone -or $IssuerWebPage) {
    $issuerFallback = @{ Name = $IssuerName; Address = $IssuerAddress; Phone = $IssuerPhone; WebPage = $IssuerWebPage }
}

$csvSuffix      = '{0}-{1:D2}' -f $Year, $Month
$csvSummaryPath = Join-Path $OutputPath "NinjaInvoices-Summary_$csvSuffix.csv"
$csvLinesPath   = Join-Path $OutputPath "NinjaInvoices-LineItems_$csvSuffix.csv"
$csvSummary     = @{}   # invoice id -> summary CSV line
$csvLines       = @{}   # invoice id -> line-item CSV lines
$csvOrg         = @{}   # invoice id -> @{ Key; Name } (for -CsvPerOrganization)
$csvFilesWritten = [System.Collections.Generic.List[string]]::new()

$succeeded = 0
$failed    = 0
$skipped   = 0
$errors    = [System.Collections.Generic.List[string]]::new()
$manifestRows = [System.Collections.Generic.List[hashtable]]::new()

foreach ($inv in $invoiceList) {
    $invNum = if ($inv.invoiceNumber) { $inv.invoiceNumber } else { "INV-$($inv.id)" }
    $listOrgId     = if ($inv.client -and $null -ne $inv.client.id) { [int]$inv.client.id } else { $null }
    $displayClient = if ($inv.client -and $inv.client.name) { $inv.client.name }
                     elseif ($null -ne $listOrgId -and $orgNames.ContainsKey($listOrgId)) { $orgNames[$listOrgId] }
                     elseif ($null -ne $listOrgId) { "Customer #$listOrgId" }
                     else { 'Unknown' }
    $clientName    = Get-SafeFileName $displayClient 60
    $safeNum       = Get-SafeFileName $invNum 40
    $baseName      = "${safeNum}_${clientName}"
    $htmlName      = "$baseName.html"
    $pdfName       = "$baseName.pdf"
    $htmlPath      = Join-Path $htmlDir $htmlName      # temp folder when -PdfOnly
    $pdfPath       = Join-Path $OutputPath $pdfName

    $targetPath = if ($CsvOnly) { $csvSummaryPath } elseif ($PdfOnly) { $pdfPath } else { $htmlPath }
    $shouldDesc = if ($CsvOnly) { "Add invoice $invNum ($displayClient) to CSV export" }
                  else { "Generate invoice $(if ($PdfOnly) { 'PDF' } else { 'HTML' }) for $invNum ($displayClient)" }

    if ($PSCmdlet.ShouldProcess($targetPath, $shouldDesc)) {
        try {
            Write-Host "  [$invNum] $displayClient..." -NoNewline

            $detail = Invoke-NinjaApi -Endpoint "/v2/billing/invoices/$($inv.id)" -Session $session

            if (-not $CsvOnly) {
                $htmlContent = Build-InvoiceHtml -Invoice $detail -ClientName $displayClient -IssuerFallback $issuerFallback
                [System.IO.File]::WriteAllText($htmlPath, $htmlContent, $utf8Bom)

                try {
                    if ($pdfBrowser) { Convert-HtmlToPdf -Browser $pdfBrowser -HtmlPath $htmlPath -PdfPath $pdfPath }
                } finally {
                    if ($PdfOnly -and (Test-Path -LiteralPath $htmlPath)) { Remove-Item -LiteralPath $htmlPath -Force -ErrorAction SilentlyContinue }
                }

                $orgId = if ($detail.client -and $null -ne $detail.client.id) { [int]$detail.client.id } else { $null }
                if ($null -ne $orgId) {
                    $row = @{
                        invoiceId              = [int]$inv.id
                        clientOrganizationId   = $orgId
                        invoiceNumber          = $invNum
                        clientName             = $displayClient
                    }
                    if (-not $PdfOnly) { $row.htmlFileName = $htmlName }
                    if ($pdfBrowser)   { $row.pdfFileName  = $pdfName }
                    [void]$manifestRows.Add($row)
                } else {
                    Write-Warning "[$invNum] Invoice has no client.id (organization); omitted from manifest (file still saved)."
                }
            }

            if ($ExportCsv) {
                $csvSummary[[int]$inv.id] = New-SummaryCsvLine -ListInvoice $inv -OrganizationName $displayClient
                $csvLines[[int]$inv.id]   = @(New-LineCsvLines -ListInvoice $inv -Detail $detail -OrganizationName $displayClient)
                $csvOrg[[int]$inv.id]     = [pscustomobject]@{ Key = $(if ($null -ne $listOrgId) { "$listOrgId" } else { 'unknown' }); Name = $displayClient }
            }

            Write-Host " $(if ($CsvOnly) { 'csv' } elseif ($PdfOnly) { $pdfName } else { $htmlName })"
            $succeeded++
        } catch {
            Write-Host " FAILED"
            $msg = "[$invNum] $($_.Exception.Message)"
            Write-Warning $msg
            [void]$errors.Add($msg)
            $failed++
        }
    } else {
        $skipped++
    }
}

if ($PdfOnly -and $htmlDir -ne $OutputPath -and (Test-Path -LiteralPath $htmlDir)) {
    Remove-Item -LiteralPath $htmlDir -Recurse -Force -ErrorAction SilentlyContinue
}

if ($ExportCsv -and -not $WhatIfPreference -and $succeeded -gt 0) {
    $ids = @($csvSummary.Keys | Sort-Object -Descending)   # newest first, like the UI
    if ($CsvPerOrganization) {
        # group invoice ids by organization (ids are already newest-first)
        $byOrg = [ordered]@{}
        foreach ($id in $ids) {
            $k = $csvOrg[$id].Key
            if (-not $byOrg.Contains($k)) { $byOrg[$k] = [System.Collections.Generic.List[int]]::new() }
            $byOrg[$k].Add([int]$id)
        }
        $usedNames = @{}
        foreach ($k in $byOrg.Keys) {
            $orgIds = @($byOrg[$k])
            $safe = Get-SafeFileName $csvOrg[$orgIds[0]].Name 60
            if ($usedNames.ContainsKey($safe.ToLowerInvariant())) { $safe = "$safe-$k" }
            $usedNames[$safe.ToLowerInvariant()] = $true
            $sp = Join-Path $OutputPath "NinjaInvoices-Summary_${safe}_$csvSuffix.csv"
            $lp = Join-Path $OutputPath "NinjaInvoices-LineItems_${safe}_$csvSuffix.csv"
            Write-CsvFile -Path $sp -Header $SummaryCsvHeader -Lines @($orgIds | ForEach-Object { $csvSummary[$_] }) -Encoding $utf8Bom
            Write-CsvFile -Path $lp -Header $LineCsvHeader    -Lines @($orgIds | ForEach-Object { $csvLines[$_] })   -Encoding $utf8Bom
            [void]$csvFilesWritten.Add($sp); [void]$csvFilesWritten.Add($lp)
        }
    } else {
        Write-CsvFile -Path $csvSummaryPath -Header $SummaryCsvHeader -Lines @($ids | ForEach-Object { $csvSummary[$_] }) -Encoding $utf8Bom
        Write-CsvFile -Path $csvLinesPath   -Header $LineCsvHeader    -Lines @($ids | ForEach-Object { $csvLines[$_] })   -Encoding $utf8Bom
        [void]$csvFilesWritten.Add($csvSummaryPath); [void]$csvFilesWritten.Add($csvLinesPath)
    }
}

if (-not $CsvOnly -and -not $WhatIfPreference -and $succeeded -gt 0) {
    $manifestPath = Join-Path $OutputPath $ManifestFileName
    $rows = @($manifestRows | ForEach-Object { [PSCustomObject]$_ })
    if ($rows.Count -eq 0) {
        $json = '[]'
    } elseif ($rows.Count -eq 1) {
        $json = '[' + (ConvertTo-Json -InputObject $rows[0] -Depth 5 -Compress) + ']'
    } else {
        $json = ConvertTo-Json -InputObject $rows -Depth 5
    }
    [System.IO.File]::WriteAllText($manifestPath, $json, $utf8NoBom)
}

Write-Host ''
Write-Host ('-' * 50)
Write-Host "  Succeeded : $succeeded ($(if ($CsvOnly) { 'invoices in CSV' } elseif ($PdfOnly) { 'PDF files' } elseif ($pdfBrowser) { 'HTML + PDF files' } else { 'HTML files' }))"
if ($failed  -gt 0) { Write-Host "  Failed    : $failed" }
if ($skipped -gt 0) { Write-Host "  Skipped   : $skipped (WhatIf)" }
if ($succeeded -gt 0) {
    $resolved = Resolve-Path $OutputPath -ErrorAction SilentlyContinue
    Write-Host "  Output    : $resolved"
    if (-not $WhatIfPreference) {
        if (-not $CsvOnly) { Write-Host "  Manifest  : $(Join-Path $resolved $ManifestFileName)" }
        if ($ExportCsv) {
            if ($CsvPerOrganization) {
                Write-Host "  CSV       : $($csvFilesWritten.Count) files ($($csvFilesWritten.Count / 2) organization(s)) in $resolved"
            } else {
                Write-Host "  CSV       : $csvSummaryPath"
                Write-Host "              $csvLinesPath"
            }
        }
    }
}
Write-Host ('-' * 50)

if ($failed -gt 0) {
    Write-Host ''
    Write-Host 'Failures:'
    $errors | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

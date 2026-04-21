<#
.SYNOPSIS
    Exports NinjaOne billing invoices for a given month as PDF files.

.DESCRIPTION
    Authenticates to the NinjaOne API via OAuth2 client credentials, retrieves billing
    invoices for the specified month, fetches full invoice details including all line items,
    renders each as a styled HTML document, then converts to PDF using headless Chrome or Edge.
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
    Year for the billing period. Defaults to current year.

.PARAMETER OutputPath
    Folder where PDF invoices will be saved. Created if it does not exist.

.PARAMETER ClientId
    Optional. Filter invoices to a specific NinjaOne organization/client ID.

.PARAMETER Status
    Invoice status filter. Defaults to COMPLETE (invoices exported via the API).
    Valid values: PENDING, APPROVED, COMPLETE, FAILED, ARCHIVED.

.EXAMPLE
    .\Export-NinjaOneBillingInvoicePdfs.ps1 -Month 3 -Year 2026 -OutputPath .\Invoices

.EXAMPLE
    .\Export-NinjaOneBillingInvoicePdfs.ps1 -Status APPROVED -ClientId 42

.EXAMPLE
    .\Export-NinjaOneBillingInvoicePdfs.ps1 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$NinjaOneInstance     = $env:NINJAONE_INSTANCE,
    [string]$NinjaOneClientId     = $env:NINJAONE_CLIENT_ID,
    [string]$NinjaOneClientSecret = $env:NINJAONE_CLIENT_SECRET,

    [ValidateRange(1, 12)]
    [int]$Month = (Get-Date).AddMonths(-1).Month,

    [int]$Year = (Get-Date).AddMonths(-1).Year,

    [string]$OutputPath = '.\Invoices',

    [int]$ClientId,

    [ValidateSet('PENDING', 'APPROVED', 'COMPLETE', 'FAILED', 'ARCHIVED')]
    [string]$Status = 'COMPLETE'
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Web

# ── Auth ──────────────────────────────────────────────────────────────────────

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
        $t = Get-NinjaToken -ClientId $S.ClientId -ClientSecret $S.ClientSecret -Instance $S.Instance
        $S.AuthHeader = "Bearer $($t.access_token)"
        $S.ExpiresAt  = if ($t.expires_in) { (Get-Date).AddSeconds([int]$t.expires_in - 60) } else { (Get-Date).AddMinutes(55) }
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

# ── Formatting helpers ────────────────────────────────────────────────────────

function Format-Money {
    param([object]$Amount, [string]$Currency = 'USD')
    if ($null -eq $Amount) { return '-' }
    return "$([string]::Format('{0:N2}', [double]$Amount)) $Currency"
}

function Format-InvoiceDate {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    try { return ([datetime]$Value).ToString('MMM d, yyyy') } catch { return $Value }
}

function HtmlEnc([string]$s) {
    return [System.Web.HttpUtility]::HtmlEncode($s)
}

# ── HTML invoice builder ──────────────────────────────────────────────────────

function Build-InvoiceHtml {
    param([PSCustomObject]$Invoice)

    $cur        = if ($Invoice.currency) { $Invoice.currency } else { 'USD' }
    $invoiceNum = if ($Invoice.invoiceNumber) { $Invoice.invoiceNumber } else { "INV-$($Invoice.id)" }

    # Pre-compute conditional blocks so the here-string stays readable
    $dueDateHtml = if ($Invoice.dueDate) { "<p>Due: $(Format-InvoiceDate $Invoice.dueDate)</p>" } else { '' }

    $issuer = $Invoice.billContent
    $fromLines = [System.Collections.Generic.List[string]]::new()
    $toLines   = [System.Collections.Generic.List[string]]::new()
    if ($issuer) {
        if ($issuer.issuerName)    { $fromLines.Add("<strong>$(HtmlEnc $issuer.issuerName)</strong>") }
        if ($issuer.issuerAddress) { $fromLines.Add((HtmlEnc $issuer.issuerAddress)) }
        if ($issuer.issuerPhone)   { $fromLines.Add((HtmlEnc $issuer.issuerPhone)) }
        if ($issuer.issuerWebPage) { $fromLines.Add("<a href='$(HtmlEnc $issuer.issuerWebPage)'>$(HtmlEnc $issuer.issuerWebPage)</a>") }

        if ($issuer.billName)    { $toLines.Add("<strong>$(HtmlEnc $issuer.billName)</strong>") }
        if ($issuer.billAddress) { $toLines.Add((HtmlEnc $issuer.billAddress)) }
        if ($issuer.billEmail)   { $toLines.Add((HtmlEnc $issuer.billEmail)) }
        if ($issuer.billPhone)   { $toLines.Add((HtmlEnc $issuer.billPhone)) }
    }
    if ($toLines.Count -eq 0 -and $Invoice.client -and $Invoice.client.name) {
        $toLines.Add("<strong>$(HtmlEnc $Invoice.client.name)</strong>")
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

        [void]$lineRowsHtml.AppendLine(
            "<tr><td class='td-name'>$nameCell</td><td class='td-num'>$qty</td>" +
            "<td class='td-num'>$(Format-Money $unitPrice $cur)</td>" +
            "<td class='td-num'><span class='line-total'>$(Format-Money $lineAmt $cur)</span>$discCell</td></tr>"
        )
    }

    $discountRowHtml = ''
    if ($Invoice.discount -and [double]$Invoice.discount -ne 0) {
        $discountRowHtml = "<tr><td colspan='2' class='td-label'>Discount</td><td class='td-num negative'>- $(Format-Money $Invoice.discount $cur)</td></tr>"
    }
    $taxPct      = if ($Invoice.taxRate) { [string]::Format('{0:P1}', [double]$Invoice.taxRate) } else { '' }
    $taxLabel    = if ($taxPct) { "Tax ($taxPct)" } else { 'Tax' }

    $notesHtml = ''
    if (-not [string]::IsNullOrWhiteSpace($Invoice.invoiceNote)) {
        $notesHtml = "<div class='notes-section'><h3>Notes</h3><p>$(HtmlEnc $Invoice.invoiceNote)</p></div>"
    }

    $generatedDate = (Get-Date).ToString('MMMM d, yyyy')
    $periodStart   = Format-InvoiceDate $Invoice.billingPeriodStartDate
    $periodEnd     = Format-InvoiceDate $Invoice.billingPeriodEndDate
    $invDate       = Format-InvoiceDate $Invoice.invoiceDate
    $statusVal     = if ($Invoice.status) { $Invoice.status } else { '' }
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
    $periodStart – $periodEnd
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

<div class="footer">Generated $generatedDate · NinjaOne Billing</div>

</body>
</html>
"@
    return $html
}

# ── PDF conversion via headless Chrome/Edge ───────────────────────────────────

function ConvertTo-Pdf {
    param([string]$HtmlPath, [string]$PdfPath)

    $candidates = @(
        'chrome.exe',
        'msedge.exe',
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:LocalAppData\Google\Chrome\Application\chrome.exe"
    )

    $browser = $null
    foreach ($candidate in $candidates) {
        $found = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($found) { $browser = $found.Source; break }
        if (Test-Path $candidate -ErrorAction SilentlyContinue) { $browser = $candidate; break }
    }

    if (-not $browser) {
        throw "Chrome or Edge not found. Install either browser or add its executable to PATH, then retry."
    }

    $absPdfPath = [System.IO.Path]::GetFullPath($PdfPath)
    $fileUri = 'file:///' + $HtmlPath.Replace('\', '/').TrimStart('/')
    & $browser --headless=new "--print-to-pdf=$absPdfPath" --no-margins --disable-gpu --no-pdf-header-footer "$fileUri" 2>$null

    if ($LASTEXITCODE -and $LASTEXITCODE -ne 0) {
        throw "Browser process exited with code $LASTEXITCODE while converting '$HtmlPath'."
    }
    if (-not (Test-Path $absPdfPath)) {
        throw "PDF was not created at '$absPdfPath'. Verify the browser supports headless PDF printing."
    }
}

# ── Validation ────────────────────────────────────────────────────────────────

if ([string]::IsNullOrWhiteSpace($NinjaOneInstance))     { throw "NinjaOne instance is required. Set NINJAONE_INSTANCE or pass -NinjaOneInstance." }
if ([string]::IsNullOrWhiteSpace($NinjaOneClientId))     { throw "Client ID is required. Set NINJAONE_CLIENT_ID or pass -NinjaOneClientId." }
if ([string]::IsNullOrWhiteSpace($NinjaOneClientSecret)) { throw "Client secret is required. Set NINJAONE_CLIENT_SECRET or pass -NinjaOneClientSecret." }

# ── Main ──────────────────────────────────────────────────────────────────────

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
Write-Host "Fetching invoices for period $periodFrom to $periodTo (status filter: $Status)..."

$query = "periodFrom=$periodFrom&periodTo=$periodTo"
if ($PSBoundParameters.ContainsKey('ClientId')) { $query += "&clientId=$ClientId" }

$rawList = Invoke-NinjaApi -Endpoint '/v2/billing/invoices' -Query $query -Session $session
$invoiceList = if ($rawList -is [Array])                                 { @($rawList) }
               elseif ($rawList.PSObject.Properties['results'])          { @($rawList.results) }
               elseif ($rawList.PSObject.Properties['data'])             { @($rawList.data) }
               else                                                      { @($rawList) }

# Status is not a supported query param on this endpoint — filter client-side
$invoiceList = @($invoiceList | Where-Object { $_.status -eq $Status })
Write-Host "Found $($invoiceList.Count) invoice(s) with status '$Status'."

if ($invoiceList.Count -eq 0) {
    Write-Host "No invoices to process. Exiting."
    exit 0
}

if (-not $WhatIfPreference -and -not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$succeeded = 0
$failed    = 0
$skipped   = 0
$errors    = [System.Collections.Generic.List[string]]::new()

foreach ($inv in $invoiceList) {
    $invNum     = if ($inv.invoiceNumber) { $inv.invoiceNumber } else { "INV-$($inv.id)" }
    $clientName = if ($inv.client -and $inv.client.name) {
                      $inv.client.name -replace '[\\/:*?"<>|]', '_'
                  } else { 'Unknown' }
    $safeNum    = $invNum -replace '[\\/:*?"<>|]', '_'
    $pdfName    = "${safeNum}_${clientName}.pdf"
    $pdfPath    = Join-Path $OutputPath $pdfName

    if ($PSCmdlet.ShouldProcess($pdfPath, "Generate invoice PDF for $invNum ($clientName)")) {
        try {
            Write-Host "  [$invNum] $clientName..." -NoNewline

            $detail      = Invoke-NinjaApi -Endpoint "/v2/billing/invoices/$($inv.id)" -Session $session
            $htmlContent = Build-InvoiceHtml -Invoice $detail
            $tmpHtml     = Join-Path $env:TEMP "ninja_invoice_$($inv.id)_$([System.IO.Path]::GetRandomFileName()).html"

            [System.IO.File]::WriteAllText($tmpHtml, $htmlContent, [System.Text.Encoding]::UTF8)
            ConvertTo-Pdf -HtmlPath $tmpHtml -PdfPath $pdfPath
            Remove-Item $tmpHtml -Force -ErrorAction SilentlyContinue

            Write-Host " $pdfName"
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

Write-Host ''
Write-Host ('─' * 50)
Write-Host "  Succeeded : $succeeded"
if ($failed  -gt 0) { Write-Host "  Failed    : $failed" }
if ($skipped -gt 0) { Write-Host "  Skipped   : $skipped (WhatIf)" }
if ($succeeded -gt 0) {
    $resolved = Resolve-Path $OutputPath -ErrorAction SilentlyContinue
    Write-Host "  Output    : $resolved"
}
Write-Host ('─' * 50)

if ($failed -gt 0) {
    Write-Host ''
    Write-Host 'Failures:'
    $errors | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

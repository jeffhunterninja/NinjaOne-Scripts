# NinjaOne Billing Invoice Tools

PowerShell scripts that export NinjaOne billing invoices as HTML, PDF or CSV, and optionally
create PSA tickets with the invoice attached.

| Script | Purpose |
|---|---|
| `Export-NinjaOneBillingInvoices.ps1` | Export invoices for a month as HTML, PDF and/or CSV. Writes a manifest for ticketing. |
| `New-NinjaOneBillingInvoiceTickets.ps1` | Create a PSA ticket per exported invoice and attach the PDF (or HTML). |
| `Invoke-NinjaBillingInvoiceManager.ps1` | Separate WPF GUI for export + ticketing (its own code, not covered here). |

## Requirements

- Windows PowerShell 5.1 or PowerShell 7. No extra modules.
- PDF output needs Microsoft Edge or Google Chrome installed (used headless to print the PDF).
- A NinjaOne API application (client credentials) allowed to read billing invoices and organizations.

## Setup

The scripts read credentials from environment variables (or the matching parameters):

```powershell
$env:NINJAONE_INSTANCE      = 'app.ninjarmm.com'   # your instance host name, no https://
$env:NINJAONE_CLIENT_ID     = '<client id>'
$env:NINJAONE_CLIENT_SECRET = '<client secret>'
```

Keep real credentials out of git. If you store them in a helper script, add it to `.gitignore`.

## Export modes

Every example below uses `-Month 10 -Year 2026 -Status PENDING,APPROVED`. Adjust to taste.
Add `-WhatIf` to any command to preview without writing files or calling the detail endpoints.

### Defaults

With no parameters the script exports **last month's COMPLETE invoices** as HTML into `.\Invoices`.

```powershell
.\Export-NinjaOneBillingInvoices.ps1
```

Tenants that mostly have PENDING or APPROVED invoices need `-Status` (or `-Status ANY`), otherwise
nothing matches and the script exits with "No invoices to process".

### HTML (one file per invoice)

```powershell
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING,APPROVED
```

### PDF only

HTML is generated in a temp folder, printed to PDF, then deleted. Only `.pdf` files remain.

```powershell
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING,APPROVED -PdfOnly
```

Put your company in the "From" block (the API's invoice issuer details are often empty):

```powershell
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING,APPROVED -PdfOnly `
    -IssuerName "Your MSP Name" -IssuerAddress "123 Main St`nTampa, FL 33601" `
    -IssuerPhone "555-0100" -IssuerWebPage "https://example.com"
```

### HTML and PDF

```powershell
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING,APPROVED -ConvertToPdf
```

### CSV: combined

Two files containing every invoice, laid out like the NinjaOne UI exports:

- `NinjaInvoices-Summary_2026-10.csv` (one row per invoice)
- `NinjaInvoices-LineItems_2026-10.csv` (one row per line item)

```powershell
# CSV only
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING,APPROVED -CsvOnly

# CSV in addition to PDFs
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING,APPROVED -PdfOnly -ExportCsv
```

### CSV: one pair per organization

```powershell
# CSV only, per organization
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING,APPROVED -CsvOnly -CsvPerOrganization

# PDFs plus per-organization CSVs
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING,APPROVED -PdfOnly -CsvPerOrganization
```

Produces `NinjaInvoices-Summary_<Organization>_2026-10.csv` and
`NinjaInvoices-LineItems_<Organization>_2026-10.csv` for each organization. If two organizations
share a name, the second file name gets the organization id appended.

### Narrow the export

```powershell
# One organization only (NinjaOne organization id; -ClientId still works as an alias)
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING -OrganizationId 3

# Every status
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status ANY

# Custom output folder (relative paths resolve against your current PowerShell location)
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING -OutputPath C:\Billing\2026-10
```

## Parameter reference (export)

| Parameter | Default | Notes |
|---|---|---|
| `-Month` / `-Year` | previous month | Billing period to export. |
| `-Status` | `COMPLETE` | One or more of `PENDING, APPROVED, COMPLETE, FAILED, ARCHIVED`, or `ANY`. |
| `-OrganizationId` | all | Only invoices for this organization. Alias `-ClientId`. |
| `-OutputPath` | `.\Invoices` | Created if missing. |
| `-ConvertToPdf` | off | Write PDFs in addition to HTML. |
| `-PdfOnly` | off | PDFs only; implies `-ConvertToPdf`; intermediate HTML is deleted. |
| `-ExportCsv` | off | Also write the combined summary and line-item CSVs. |
| `-CsvPerOrganization` | off | Implies `-ExportCsv`; one CSV pair per organization instead of combined. |
| `-CsvOnly` | off | Only CSVs (implies `-ExportCsv`). No HTML/PDF, no ticket manifest. Cannot be combined with `-PdfOnly`/`-ConvertToPdf`. |
| `-IssuerName`, `-IssuerAddress`, `-IssuerPhone`, `-IssuerWebPage` | none | "From" block used when the API provides none. |
| `-WhatIf` | | Preview only. |

Instance and credentials can also be passed as `-NinjaOneInstance`, `-NinjaOneClientId`, `-NinjaOneClientSecret`.

## What gets written

| Mode | Files in `-OutputPath` |
|---|---|
| default | `<number>_<customer>.html`, manifest |
| `-ConvertToPdf` | `.html` and `.pdf`, manifest |
| `-PdfOnly` | `.pdf`, manifest |
| `-CsvOnly` | CSV files only |
| `-ExportCsv` / `-CsvPerOrganization` | CSV files in addition to the above |

The manifest, `NinjaBillingInvoiceExport-manifest.json`, lists the invoices exported by that run and is
**replaced on every run**. Use a separate `-OutputPath` per run if you need to keep earlier manifests.

CSV notes:

- Created and Last Updated are shown in the local time zone of the computer running the script.
- Text cells starting with `=`, `+`, `-` or `@` get a leading apostrophe so spreadsheets don't run them as formulas.
- Files are UTF-8 with a BOM so Excel reads accented characters correctly.
- The `Type` column comes from each line item's `quantityType` field.

## Creating PSA tickets

`New-NinjaOneBillingInvoiceTickets.ps1` reads the manifest and, for each invoice, creates a ticket and
attaches the PDF (if the export produced one) or otherwise the HTML. It uses a **separate** OAuth
application that supports authorization code with a refresh token and ticketing permissions.

```powershell
$env:NINJAONE_REFRESH_TOKEN = '<refresh token>'

# Preview
.\New-NinjaOneBillingInvoiceTickets.ps1 -OutputPath .\Invoices -TicketFormId 5 -WhatIf

# Create tickets for the invoices in .\Invoices
.\New-NinjaOneBillingInvoiceTickets.ps1 -OutputPath .\Invoices -TicketFormId 5

# Point at a manifest elsewhere, with a custom subject
.\New-NinjaOneBillingInvoiceTickets.ps1 -ManifestPath 'D:\Exports\NinjaBillingInvoiceExport-manifest.json' `
    -TicketFormId 5 -TicketSubjectTemplate 'Invoice {InvoiceNumber} - {ClientName}'
```

Typical end-to-end run:

```powershell
.\Export-NinjaOneBillingInvoices.ps1 -Month 10 -Year 2026 -Status PENDING,APPROVED -PdfOnly -IssuerName "Your MSP Name"
.\New-NinjaOneBillingInvoiceTickets.ps1 -OutputPath .\Invoices -TicketFormId 5 -WhatIf
.\New-NinjaOneBillingInvoiceTickets.ps1 -OutputPath .\Invoices -TicketFormId 5
```

Find your ticket form id with `GET /v2/ticketing/ticket-form` in the NinjaOne API. The ticket script
has no duplicate protection: running it twice on the same manifest creates two sets of tickets.

## Known limitations

- Pagination of the invoice list is not implemented. The list endpoint returned a plain array with no
  paging fields in testing; the script warns if a response ever looks paged.
- Tax-rate display assumes a fraction (0.0825) but also accepts a percent (8.25). Test invoices so far had 0% tax.
- PDF conversion needs Edge or Chrome installed on the machine running the export.

# Activity Discrepancy Comparison

Compares NinjaOne API activities with webhook.site requests to identify delivery gaps. The script uses a **webhook-first, date-based** flow: it fetches webhook.site requests in the given time window first, parses activity payloads, and uses the **date of the oldest activity** in that dataset as the starting point. It then fetches from NinjaOne using the **after** parameter (afterUnixEpoch) so the NinjaOne set is aligned by time with what webhook received. If there are no webhook activities in the window, it uses the -After parameter for the NinjaOne date range. Use this to detect activities in NinjaOne but not at webhook.site, or webhook requests that don't match NinjaOne records.

## Prerequisites

- **PowerShell** 5.1+ (PowerShell 7 recommended)
- **NinjaOne** OAuth app with Client Credentials flow and scopes `monitoring` and `management`
- **NinjaOne webhook** configured to POST activity payloads to your webhook.site URL
- **webhook.site** token (free at [webhook.site](https://webhook.site))

## Parameters

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `-WebhookTokenId` | string | Yes | webhook.site token UUID (from your webhook URL) |
| `-After` | datetime | No | Start of time window (default: 7 days ago). End is always now. Used for webhook.site request list; when falling back (no webhook activities), also used as NinjaOne date range start. |
| `-WebhookApiKey` | string | No | webhook.site API key (if token requires auth) |
| `-NinjaInstance` | string | No* | NinjaOne instance URL (e.g. https://app.ninjarmm.com) |
| `-NinjaClientId` | string | No* | NinjaOne OAuth client ID |
| `-NinjaClientSecret` | string | No* | NinjaOne OAuth client secret |
| `-OutputPath` | string | No | Directory for CSV/JSON exports (default: current directory) |
| `-IncludeMatched` | switch | No | Include matched activities in CSV export |
| `-SortOrder` | string | No | Sort order by activityTime: `ascending` (oldest first) or `descending` (newest first). Default: ascending |

\* When run in NinjaOne context (e.g. via API server or script deployment), `NinjaInstance`, `NinjaClientId`, and `NinjaClientSecret` can be read from NinjaOne custom properties via `Get-NinjaProperty`. Otherwise pass them explicitly.

## Usage Examples

```powershell
# Basic run (requires NinjaOne context or env vars for credentials)
.\Compare-NinjaActivityDiscrepancies.ps1 -WebhookTokenId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

# With explicit NinjaOne credentials
.\Compare-NinjaActivityDiscrepancies.ps1 -WebhookTokenId "xxx" `
  -NinjaInstance "https://app.ninjarmm.com" `
  -NinjaClientId "your-client-id" `
  -NinjaClientSecret "your-secret"

# Last 24 hours, custom output path
.\Compare-NinjaActivityDiscrepancies.ps1 -WebhookTokenId "xxx" `
  -After (Get-Date).AddDays(-1) `
  -OutputPath C:\Reports

# Include matched activities in CSV export
.\Compare-NinjaActivityDiscrepancies.ps1 -WebhookTokenId "xxx" -IncludeMatched
```

## Output

### Console

Output appears in this order:

1. **Progress**: `Fetching webhook.site requests ($After to $Before)...` → `Parsed X NinjaOne activity payload(s) from Y webhook request(s).` → optional `Oldest webhook activity date: ... UTC` or warnings (e.g. no webhook activities; using -After for NinjaOne) → `Fetching NinjaOne activities (after ... UTC to ... UTC)...` → `Fetched X activities from NinjaOne.`
2. **Summary**: NinjaOne window (after oldest webhook activity date to Before), NinjaOne total, Webhook total, Matched, In NinjaOnly, In WebhookOnly
3. **Sample: In NinjaOne Only (first 10)** and **Sample: In Webhook Only (first 10)** (when non-empty): each line shows id, activityTime, deviceId, type, status (sorted by activityTime per `-SortOrder`)
4. **Export messages**: For each CSV/JSON, either `Exported ... to <path>` or `No ... to export.`
5. **Breakdown: In Ninja Only by type** — table columns Type, Count (or `(none)`)
6. **Breakdown: All Webhook activities by type** — table columns Type, Count (or `(none)`)
7. **Discrepancy by type** — table columns Type, AllNinja, AllWebhook, Matched, InNinjaOnly, InWebhookOnly (or `(none)`), then `Exported TypeBreakdown (N rows) to <path>`

### CSV Files

| File | Description |
|------|-------------|
| `InNinjaOnly.csv` | Activities found in NinjaOne API but not in webhook.site (sorted by activityTime) |
| `InWebhookOnly.csv` | Activities found in webhook.site but not in NinjaOne API (sorted by activityTime) |
| `AllWebhookActivities.csv` | All activities parsed from webhook.site in the time window (sorted by activityTime) |
| `AllNinjaOneActivities.csv` | All activities fetched from NinjaOne in the window (sorted by activityTime) |
| `TypeBreakdown.csv` | Per-activity-type counts: Type, AllNinja, AllWebhook, Matched, InNinjaOnly, InWebhookOnly |
| `Matched.csv` | Activities present in both, sorted by activityTime (only when `-IncludeMatched` is used) |

### JSON

- **discrepancy-report.json**: Full report containing **GeneratedAt** (ISO timestamp), **TimeWindow** (After/Before), **NinjaOneWindow** (Description, After, Before), **Summary** (NinjaOneTotal, WebhookTotal, Matched, InNinjaOnly, InWebhookOnly), **InNinjaOnly**, **InWebhookOnly** (raw activity objects), **AllWebhookActivities**, **AllNinjaOneActivities** (full raw sets), and **TypeBreakdown** (per-type counts)

### Script variants

The folder contains three scripts with the same behavior but different output filenames (for separate webhook tokens or tenants):

| Script | CSV prefix | JSON report |
|--------|------------|-------------|
| `Compare-NinjaActivityDiscrepancies.ps1` | (none) | `discrepancy-report.json` |
| `Compare-MSPActivities.ps1` | MSP | `mspdiscrepancy-report.json` |
| `Compare-IITActivities.ps1` | IIT | `IITdiscrepancy-report.json` |

Example: the base script writes `InNinjaOnly.csv` and `AllNinjaOneActivities.csv`; the MSP variant writes `MSPInNinjaOnly.csv` and `MSPAllNinjaOneActivities.csv`. `Matched.csv` is unchanged across variants.

## NinjaOne Webhook Payload Format

The script supports these NinjaOne webhook payload shapes:

1. **Flat activity**: `{ "id": 123, "activityTime": 1234567890, "deviceId": 1, ... }`
2. **Wrapper with `activity`**: `{ "activity": { "id": 123, ... } }`
3. **Wrapper with `activities` array**: `{ "activities": [ { "id": 123, ... }, ... ] }`

Configure your NinjaOne webhook to send activities; the script will parse the request body and extract activity objects.

The included `webhook.json` configures the webhook to request **all 38 activity types** from NinjaOne documentation (ACTIONSET, ACTION, CONDITION, CONDITION_ACTIONSET, CONDITION_ACTION, ANTIVIRUS, PATCH_MANAGEMENT, TEAMVIEWER, MONITOR, SYSTEM, COMMENT, SHADOWPROTECT, IMAGEMANAGER, HELP_REQUEST, SOFTWARE_PATCH_MANAGEMENT, SPLASHTOP, CLOUDBERRY, CLOUDBERRY_BACKUP, SCHEDULED_TASK, RDP, SCRIPTING, SECURITY, REMOTE_TOOLS, VIRTUALIZATION, PSA, MDM, NINJA_REMOTE, NINJA_QUICK_CONNECT, NINJA_NETWORK_DISCOVERY, NINJA_BACKUP, NINJA_TICKETING, KNOWLEDGE_BASE, RELATED_ITEM, CLIENT_CHECKLIST, CHECKLIST_TEMPLATE, DOCUMENTATION, MICROSOFT_INTUNE, DYNAMIC_POLICY) so the payload captures every activity.

## Troubleshooting

| Issue | Likely Cause | Solution |
|-------|--------------|----------|
| "NinjaOne credentials required" | Not in NinjaOne context, params not passed | Pass `-NinjaInstance`, `-NinjaClientId`, `-NinjaClientSecret` or run from NinjaOne API server |
| Webhook.site 401/403 | Token requires API key | Pass `-WebhookApiKey` |
| Webhook.site rate limit | Too many requests | Script uses per_page=100 and ~550ms delay between pages; reduce time window if needed |
| Empty webhook results | Wrong token ID or date window | Verify token UUID; ensure `-After` to now covers when webhooks were sent. Script will use `-After` for NinjaOne range and warn. |
| "Could not parse X webhook request(s)" | Non-JSON or unexpected payload | Check NinjaOne webhook configuration; payload should be JSON with activity structure |
| Wrong discrepancy counts | Timezone mismatch | Script uses UTC for both NinjaOne (`activityTime` epoch) and webhook.site date filters |

## Notes

- **Webhook-first flow**: Webhook.site is queried first; the date of the oldest activity (activityTime) in that dataset is used as the NinjaOne **after** parameter. This aligns the comparison by time with what webhook received.
- **Fallback**: If the webhook window has no parseable activities with activityTime, the script uses the -After parameter for the NinjaOne date range and warns.
- **Webhook.site rate limit**: 120 requests/minute. The script throttles pagination.
- **Deduplication**: Multiple webhook requests for the same activity id (e.g. retries) are deduplicated; last occurrence wins.
- **Run context**: Intended for API server, scheduled task, or manual run—not a device script.

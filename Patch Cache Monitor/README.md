# Patch Cache Health Monitor v2 (NinjaOne)

Builds a patch cache utilization report from Ninja `CacheListener` logs and writes the HTML directly to a device custom field using `Ninja-Property-Set-Piped`.

## Attribution

All credit to Ryan Snyder, the creator of the original script.

## What it does

- Detects whether `CacheListener.exe` is running.
- Locates `CacheListener_*.log` files by searching:
  - The directory containing the running `CacheListener` process
  - `C:\ProgramData\NinjaRMMAgent\logs\CacheListener`
  - `C:\ProgramData\NinjaRMMAgent\logs`
  - `C:\ProgramData\NinjaRMMAgent`
  - `C:\ProgramData\cache`
  - `C:\ProgramData\NinjaRMMAgent\cache`
- Filters logs to the last **N** days (`LastNDays`) using the timestamp embedded in the log filename (with fallbacks), and limits how many log files are parsed (`MaxLogFiles`).
- Parses log lines for heartbeats, download requests (client IP, URL, session), transfer type (cache hit, cache miss, partial), completed transfers (size, duration, speed), and error/warning lines (with known noise filtered).
- Computes total size of the cache folder at `C:\ProgramData\cache`.
- Produces a styled HTML report: summary metrics, up to 25 recent transfers, and up to 15 recent errors (with a count if more exist).
- Pipes the final HTML to the device custom field **`patchCacheHtml`**.

Recursive scans of log locations and the cache folder can take noticeable time on very large directory trees.

The HTML includes **client IP addresses and file names** from logs; consider whether that meets your org’s privacy expectations before exposing the custom field broadly.

## Prerequisites

- A **WYSIWYG** device custom field named **`patchCacheHtml`**. Use a field type that supports the HTML length your reports may reach.
- Run the script on the **patch cache server** (the device where `CacheListener` runs), typically via **scheduled automation** in NinjaOne.

## Parameters

| Parameter     | Range   | Default | Role                                                                 |
| ------------- | ------- | ------- | -------------------------------------------------------------------- |
| `LastNDays`   | 1–365   | 7       | Window for log file selection and per-line time filtering.           |
| `MaxLogFiles` | 1–1000  | 50      | Maximum log files parsed after the date filter (newest by timestamp). |

## NinjaOne usage

1. Create the `patchCacheHtml` custom field.
2. Assign the custom field to devices with the cache server device role.
3. Add a scheduled automation or task targeting the cache server featuring the script **`Monitor-PatchCacheServer.ps1`**.
4. View results on the device record under the **`patchCacheHtml`** custom field (HTML renders where NinjaOne supports it for that field type).

## Console output

The script also emits timestamped lines to the host (`Write-Host`) and ends with a text summary (logs parsed, service PID, heartbeats, transfers, data served, hit rate, clients, cache folder size, error count) for run history and troubleshooting in Ninja script output.

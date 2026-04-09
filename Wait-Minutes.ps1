# ============================================================
# NinjaOne Script Variable:
#   Name : waitMinutes  (Integer)
#   Type : Integer
# ============================================================

# Read the script variable injected by NinjaOne
$waitMinutes = $env:waitMinutes

# --- Validation ---
if ([string]::IsNullOrWhiteSpace($waitMinutes)) {
    Write-Error "Script variable 'waitMinutes' is not set. Please define it in NinjaOne."
    exit 1
}

if (-not [int]::TryParse($waitMinutes, [ref]$null)) {
    Write-Error "Script variable 'waitMinutes' must be an integer. Received: '$waitMinutes'"
    exit 1
}

$minutes = [int]$waitMinutes

if ($minutes -lt 0) {
    Write-Error "Script variable 'waitMinutes' must be a non-negative integer. Received: $minutes"
    exit 1
}

# --- Wait ---
if ($minutes -eq 0) {
    Write-Host "waitMinutes is 0 — skipping wait and proceeding immediately."
} else {
    $totalSeconds = $minutes * 60
    Write-Host "Waiting $minutes minute(s) ($totalSeconds seconds) before proceeding..."

    for ($i = 1; $i -le $totalSeconds; $i++) {
        $elapsed   = $i
        $remaining = $totalSeconds - $i
        $pct       = [math]::Round(($elapsed / $totalSeconds) * 100)

        Write-Progress `
            -Activity "Waiting $minutes minute(s)" `
            -Status   "$([math]::Floor($remaining / 60))m $($remaining % 60)s remaining" `
            -PercentComplete $pct

        Start-Sleep -Seconds 1
    }

    Write-Progress -Activity "Waiting $minutes minute(s)" -Completed
    Write-Host "Wait complete. Proceeding..."
}

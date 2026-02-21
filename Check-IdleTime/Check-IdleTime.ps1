#Requires -Version 5.1
<#
.SYNOPSIS
  Per-user idle time while running as SYSTEM by launching a helper in each user session
  that calls GetLastInputInfo and exits with idle seconds. Writes idle time to NinjaOne custom fields.

.EXIT CODES
  0 = OK (no threshold or idle < threshold)
  1 = ALERT (idle >= threshold)
  2 = Not elevated

.EXAMPLE
  .\Check-IdleTime.ps1
  Run with default settings (no threshold). Measures all active sessions and writes idle time to NinjaOne custom fields.

.EXAMPLE
  .\Check-IdleTime.ps1 -ThresholdMinutes 60 -PerProcessTimeoutSeconds 15
  Alert if idle >= 60 minutes; allow up to 15 seconds per session for the helper to complete.
#>

[CmdletBinding()]
param(
  [string]$UserName,
  [ValidateRange(0, [int]::MaxValue)]
  [int]$ThresholdMinutes = 0,
  [ValidateRange(1, 300)]
  [int]$PerProcessTimeoutSeconds = 10
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# NinjaOne script variable "Threshold Minutes" populates $env:thresholdminutes; use it when present and valid
if ($null -ne $env:thresholdminutes -and -not [string]::IsNullOrWhiteSpace($env:thresholdminutes)) {
  $parsed = 0
  if ([int]::TryParse($env:thresholdminutes.Trim(), [ref]$parsed) -and $parsed -ge 0) {
    $ThresholdMinutes = $parsed
  }
}

function Test-IsElevated {
  $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = [System.Security.Principal.WindowsPrincipal]::new($id)
  return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Format-Minutes([int]$Minutes) {
  $ts = [TimeSpan]::FromMinutes([double]$Minutes)
  $parts = @()
  if ($ts.Days)    { $parts += "$($ts.Days) $(if ($ts.Days -eq 1) { 'day' } else { 'days' })" }
  if ($ts.Hours)   { $parts += "$($ts.Hours) $(if ($ts.Hours -eq 1) { 'hour' } else { 'hours' })" }
  if ($ts.Minutes) { $parts += "$($ts.Minutes) $(if ($ts.Minutes -eq 1) { 'minute' } else { 'minutes' })" }
  if (-not $parts) { $parts = @('0 minutes') }
  $parts -join ', '
}

# Optional: 'query user' for display correlation
function Get-QueryUser {
  try {
    $raw = @(query.exe user 2>$null)
    if ($raw.Count -le 1) { return $null }
    $lines  = $raw.ForEach({ $_ -replace '\s{2,}', ',' })
    $header = $lines[0].Split(',').Trim()
    $result = @()
    for ($i = 1; $i -lt $lines.Count; $i++) {
      $cols = $lines[$i].Split(',').ForEach({ $_.Trim().Trim('>') })
      if ($cols.Count -eq 5) { $cols = @($cols[0], $null, $cols[1], $cols[2], $cols[3], $cols[4]) }
      $obj = "" | Select-Object $header
      for ($j = 0; $j -lt [Math]::Min($cols.Count, $header.Count); $j++) { $obj.$($header[$j]) = $cols[$j] }
      $result += $obj
    }
    return $result
  } catch { return $null }
}

# ---------- C# interop: enumerate sessions, get user tokens, spawn helper in session, capture exit code ----------
if (-not ("UserIdleHelper" -as [type])) {
Add-Type -Language CSharp @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class UserIdleHelper
{
    // --- WTS API ---
    public enum WTS_CONNECTSTATE_CLASS {
        WTSActive, WTSConnected, WTSConnectQuery, WTSShadow, WTSDisconnected,
        WTSIdle, WTSListen, WTSReset, WTSDown, WTSInit
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WTS_SESSION_INFO {
        public int SessionId;
        [MarshalAs(UnmanagedType.LPStr)] public string pWinStationName;
        public WTS_CONNECTSTATE_CLASS State;
    }

    [DllImport("Wtsapi32.dll", SetLastError=true)]
    public static extern bool WTSEnumerateSessions(IntPtr hServer, int Reserved, int Version,
        out IntPtr ppSessionInfo, out int pCount);

    [DllImport("Wtsapi32.dll")] public static extern void WTSFreeMemory(IntPtr pMemory);

    [DllImport("Wtsapi32.dll", SetLastError=true)]
    public static extern bool WTSQueryUserToken(int sessionId, out IntPtr Token);

    // --- Advapi32 for tokens/process ---
    [Flags] public enum TOKEN_ACCESS : uint {
        TOKEN_ASSIGN_PRIMARY = 0x0001,
        TOKEN_DUPLICATE      = 0x0002,
        TOKEN_QUERY          = 0x0008,
        TOKEN_ADJUST_DEFAULT = 0x0080,
        TOKEN_ADJUST_SESSIONID = 0x0100
    }

    public enum SECURITY_IMPERSONATION_LEVEL { Anonymous, Identification, Impersonation, Delegation }
    [Flags] public enum TOKEN_TYPE { TokenPrimary = 1, TokenImpersonation }

    [DllImport("Advapi32.dll", SetLastError=true)]
    public static extern bool DuplicateTokenEx(
        IntPtr hExistingToken,
        uint dwDesiredAccess,
        IntPtr lpTokenAttributes,
        SECURITY_IMPERSONATION_LEVEL ImpersonationLevel,
        TOKEN_TYPE TokenType,
        out IntPtr phNewToken);

    [StructLayout(LayoutKind.Sequential)]
    public struct STARTUPINFO {
        public int cb;
        public string lpReserved;
        public string lpDesktop;
        public string lpTitle;
        public int dwX;
        public int dwY;
        public int dwXSize;
        public int dwYSize;
        public int dwXCountChars;
        public int dwYCountChars;
        public int dwFillAttribute;
        public int dwFlags;
        public short wShowWindow;
        public short cbReserved2;
        public IntPtr lpReserved2;
        public IntPtr hStdInput;
        public IntPtr hStdOutput;
        public IntPtr hStdError;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION {
        public IntPtr hProcess;
        public IntPtr hThread;
        public int dwProcessId;
        public int dwThreadId;
    }

    [DllImport("Advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool CreateProcessAsUser(
        IntPtr hToken,
        string lpApplicationName,
        string lpCommandLine,
        IntPtr lpProcessAttributes,
        IntPtr lpThreadAttributes,
        bool bInheritHandles,
        uint dwCreationFlags,
        IntPtr lpEnvironment,
        string lpCurrentDirectory,
        ref STARTUPINFO lpStartupInfo,
        out PROCESS_INFORMATION lpProcessInformation);

    [DllImport("Kernel32.dll", SetLastError=true)]
    public static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);

    [DllImport("Kernel32.dll", SetLastError=true)]
    public static extern bool GetExitCodeProcess(IntPtr hProcess, out uint lpExitCode);

    [DllImport("Kernel32.dll", SetLastError=true)]
    public static extern bool CloseHandle(IntPtr hObject);

    // --- Userenv for environment block ---
    [DllImport("Userenv.dll", SetLastError=true)]
    public static extern bool CreateEnvironmentBlock(out IntPtr lpEnvironment, IntPtr hToken, bool bInherit);
    [DllImport("Userenv.dll", SetLastError=true)]
    public static extern bool DestroyEnvironmentBlock(IntPtr lpEnvironment);

    public const uint CREATE_UNICODE_ENVIRONMENT = 0x00000400;
    public const uint CREATE_NO_WINDOW = 0x08000000;
    public const uint STARTF_USESHOWWINDOW = 0x00000001;
    public const short SW_HIDE = 0;

    public class SessionInfo {
        public int SessionId;
        public string WinStation;
        public string State; // stringified
    }

    public static SessionInfo[] ListSessions() {
        System.Collections.Generic.List<SessionInfo> list = new System.Collections.Generic.List<SessionInfo>();
        IntPtr p = IntPtr.Zero; int count = 0;
        if (!WTSEnumerateSessions(IntPtr.Zero, 0, 1, out p, out count) || p == IntPtr.Zero || count <= 0) {
            return list.ToArray();
        }
        try {
            int size = Marshal.SizeOf(typeof(WTS_SESSION_INFO));
            for (int i = 0; i < count; i++) {
                var itemPtr = new IntPtr(p.ToInt64() + (i * size));
                var si = (WTS_SESSION_INFO)Marshal.PtrToStructure(itemPtr, typeof(WTS_SESSION_INFO));
                list.Add(new SessionInfo {
                    SessionId = si.SessionId,
                    WinStation = si.pWinStationName,
                    State      = si.State.ToString()
                });
            }
        } finally {
            WTSFreeMemory(p);
        }
        return list.ToArray();
    }

    // Launches a hidden powershell in the target session, waits, returns exit code
    public static bool RunInSessionAndGetExitCode(int sessionId, string commandLine, int timeoutMs, out int exit)
    {
        exit = -1;
        IntPtr hUserTok = IntPtr.Zero;
        IntPtr hPrimTok = IntPtr.Zero;
        IntPtr env = IntPtr.Zero;
        PROCESS_INFORMATION pi = new PROCESS_INFORMATION();
        try {
            if (!WTSQueryUserToken(sessionId, out hUserTok) || hUserTok == IntPtr.Zero) return false;

            // Need a primary token for CreateProcessAsUser
            uint access = (uint)(
                TOKEN_ACCESS.TOKEN_ASSIGN_PRIMARY |
                TOKEN_ACCESS.TOKEN_DUPLICATE |
                TOKEN_ACCESS.TOKEN_QUERY |
                TOKEN_ACCESS.TOKEN_ADJUST_DEFAULT |
                TOKEN_ACCESS.TOKEN_ADJUST_SESSIONID
            );

            if (!DuplicateTokenEx(hUserTok, access, IntPtr.Zero,
                                  SECURITY_IMPERSONATION_LEVEL.Impersonation,
                                  TOKEN_TYPE.TokenPrimary, out hPrimTok) || hPrimTok == IntPtr.Zero) {
                return false;
            }

            if (!CreateEnvironmentBlock(out env, hPrimTok, false)) {
                env = IntPtr.Zero; // still try without environment
            }

            STARTUPINFO si = new STARTUPINFO();
            si.cb = Marshal.SizeOf(typeof(STARTUPINFO));
            si.dwFlags = (int)STARTF_USESHOWWINDOW;
            si.wShowWindow = SW_HIDE;

            uint flags = CREATE_UNICODE_ENVIRONMENT | CREATE_NO_WINDOW;
            bool ok = CreateProcessAsUser(
                hPrimTok,
                null,
                commandLine,
                IntPtr.Zero, IntPtr.Zero,
                false,
                flags,
                env,
                null,
                ref si,
                out pi
            );
            if (!ok || pi.hProcess == IntPtr.Zero) return false;

            // Wait; only use exit code when process actually exited (not on timeout)
            const uint WAIT_OBJECT_0 = 0;
            const uint WAIT_TIMEOUT = 0x102;
            uint wait = WaitForSingleObject(pi.hProcess, (uint)timeoutMs);
            if (wait == WAIT_TIMEOUT) return false;
            if (wait != WAIT_OBJECT_0) return false;

            uint code;
            if (!GetExitCodeProcess(pi.hProcess, out code)) return false;

            exit = unchecked((int)code);
            return true;
        }
        finally {
            if (pi.hThread != IntPtr.Zero) CloseHandle(pi.hThread);
            if (pi.hProcess != IntPtr.Zero) CloseHandle(pi.hProcess);
            if (env != IntPtr.Zero) DestroyEnvironmentBlock(env);
            if (hPrimTok != IntPtr.Zero) CloseHandle(hPrimTok);
            if (hUserTok != IntPtr.Zero) CloseHandle(hUserTok);
        }
    }
}
"@
}

if (-not (Test-IsElevated)) {
  Write-Error "Access Denied. Please run with Administrator privileges."
  exit 2
}

if (-not (Get-Command Ninja-Property-Set -ErrorAction SilentlyContinue)) {
  Write-Warning "Ninja-Property-Set cmdlet not found; NinjaOne custom fields will not be updated."
}

$timeoutMs = [Math]::Max(3000, $PerProcessTimeoutSeconds * 1000)

# Build the tiny inline helper that runs INSIDE the user’s session:
# - Adds GetLastInputInfo P/Invoke
# - Computes idle milliseconds = Environment.TickCount - LastInputTime
# - Exits with idle seconds as process exit code
# - Note: Environment.TickCount wraps at ~49 days; idle may be wrong after long uptime (known limitation)
$helperCmd = @'
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
public static class GLI {
  [StructLayout(LayoutKind.Sequential)] struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
  [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LASTINPUTINFO lii);
  public static uint GetIdleMs() {
    var lii = new LASTINPUTINFO(); lii.cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(lii);
    if(!GetLastInputInfo(ref lii)) throw new Win32Exception(Marshal.GetLastWin32Error());
    return (uint)Environment.TickCount - lii.dwTime;
  }
}
"@
$ms = [GLI]::GetIdleMs()
$sec = [int][Math]::Floor($ms / 1000.0)
[Environment]::Exit($sec)
'@

# Command line for CreateProcessAsUser (hidden, no profile)
# NOTE: Use -EncodedCommand to avoid quoting pitfalls
$bytes   = [System.Text.Encoding]::Unicode.GetBytes($helperCmd)
$encCmd  = [Convert]::ToBase64String($bytes)
$psLine  = "powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -EncodedCommand $encCmd"

# 1) Enumerate sessions
$sessions = [UserIdleHelper]::ListSessions() | ForEach-Object {
  # normalize WinStation casing
  [PSCustomObject]@{
    SessionId  = $_.SessionId
    WinStation = $_.WinStation
    State      = $_.State
  }
}
Write-Verbose "Enumerated $($sessions.Count) session(s)."

# Optional filter for display (does not affect measurement)
$quserRows = Get-QueryUser
if ($UserName -and $quserRows) {
  $targetUsers = ($quserRows | Where-Object { $_.USERNAME -like $UserName }).USERNAME | Select-Object -Unique
} else {
  $targetUsers = $null
}

# 2) For every session in a connected/active-ish state, run helper and capture exit code (idle seconds)
$measured = [System.Collections.Generic.List[object]]::new()
foreach ($s in $sessions) {
  # We’ll try for Active/Connected/Idle states (you can loosen this if needed)
  if ($s.State -in @('WTSActive','WTSConnected','WTSIdle')) {
    # If user filter was provided, skip sessions that don't belong to filtered users (based on quser mapping)
    if ($targetUsers -and $quserRows) {
      $map = $quserRows | Where-Object { $_.SESSIONNAME -eq $s.WinStation -or $_.ID -eq "$($s.SessionId)" } | Select-Object -First 1
      if ($map -and ($targetUsers -notcontains $map.USERNAME)) { continue }
    }

    $exit = -1
    Write-Verbose "Measuring session $($s.SessionId) ($($s.WinStation), $($s.State))."
    $ok = [UserIdleHelper]::RunInSessionAndGetExitCode([int]$s.SessionId, $psLine, $timeoutMs, [ref]$exit)
    $idleSec = if ($ok -and $exit -ge 0) { [int]$exit } else { $null }
    $idleMin = if ($idleSec -ne $null) { [int][Math]::Floor($idleSec / 60.0) } else { $null }

    $measured.Add([PSCustomObject]@{
      SessionId     = $s.SessionId
      WinStation    = $s.WinStation
      State         = $s.State
      IdleSeconds   = $idleSec
      IdleMinutes   = $idleMin
      MeasuredVia   = if ($idleSec -ne $null) { 'CreateProcessAsUser:GetLastInputInfo' } else { 'Failed' }
    })
  }
}

# 3) Display results
Write-Host "=== Per-Session Idle (measured inside each user session) ==="
if ($measured.Count -gt 0) {
  $measured | Sort-Object SessionId |
    Select-Object SessionId, WinStation, State, IdleMinutes, IdleSeconds, MeasuredVia |
    Format-Table -AutoSize
} else {
  Write-Host "(No sessions measured or all failed)"
}
Write-Host ""

if ($quserRows) {
  $quserTable = foreach ($r in $quserRows) {
    [PSCustomObject]@{
      UserName    = $r.USERNAME
      SessionName = $r.SESSIONNAME
      Id          = $r.ID
      State       = $r.STATE
      LogonTime   = $r.'LOGON TIME'
      IdleTimeRaw = $r.'IDLE TIME'
    }
  }
  Write-Host "=== 'query user' (reference) ==="
  $quserTable | Format-Table -AutoSize
  Write-Host ""
}

# 4) Choose session for threshold evaluation (console first, then most-idle active, else any)
$eval = $null
$console = $measured | Where-Object { $_.WinStation -match '^(Console|console)$' -and $_.IdleMinutes -ne $null } | Select-Object -First 1
if ($console) {
  $eval = $console
  Write-Verbose "Evaluated session: Console (SessionId $($eval.SessionId), IdleMinutes $($eval.IdleMinutes))."
} else {
  $active = $measured | Where-Object { $_.State -eq 'WTSActive' -and $_.IdleMinutes -ne $null } | Sort-Object IdleMinutes -Descending | Select-Object -First 1
  if ($active) { $eval = $active; Write-Verbose "Evaluated session: most-idle Active (SessionId $($eval.SessionId), IdleMinutes $($eval.IdleMinutes))." } else {
    $any = $measured | Where-Object { $_.IdleMinutes -ne $null } | Sort-Object IdleMinutes -Descending | Select-Object -First 1
    if ($any) { $eval = $any; Write-Verbose "Evaluated session: fallback (SessionId $($eval.SessionId), IdleMinutes $($eval.IdleMinutes))." }
  }
}
if (-not $eval) { Write-Verbose "No session measured; using SYSTEM fallback (0 minutes)." }

# If still nothing, report SYSTEM fallback (rare; e.g., no active sessions)
$usedFallback = $false
if (-not $eval) {
  $usedFallback = $true
  $idleMin = 0
  $friendly = '0 minutes'
} else {
  $idleMin = [int]$eval.IdleMinutes
  $friendly = Format-Minutes $idleMin
}

# 5) Optional: write to NinjaOne CFs
try { Ninja-Property-Set idleTime $friendly } catch { Write-Verbose "Ninja-Property-Set failed: $_" }
try { Ninja-Property-Set idleTimeStatus $idleMin } catch { Write-Verbose "Ninja-Property-Set failed: $_" }
try { Ninja-Property-Set idleTimeMinutes $idleMin } catch { Write-Verbose "Ninja-Property-Set failed: $_" }

# 6) Summary
$summary = [PSCustomObject]@{
  ComputerName       = $env:COMPUTERNAME
  EvaluatedSessionId = if ($eval) { $eval.SessionId } else { -1 }
  EvaluatedStation   = if ($eval) { $eval.WinStation } else { 'None' }
  EvaluatedState     = if ($eval) { $eval.State } else { 'None' }
  IdleMinutes        = $idleMin
  IdleTime           = $friendly
  ThresholdMinutes   = $ThresholdMinutes
  ThresholdExceeded  = ($ThresholdMinutes -gt 0 -and $idleMin -ge $ThresholdMinutes)
  UsedFallback       = $usedFallback
}

Write-Host "=== Summary ==="
$null = ($summary | Format-List * | Out-String) | ForEach-Object { Write-Host $_ }

# 7) Exit
if ($ThresholdMinutes -gt 0 -and $idleMin -ge $ThresholdMinutes) {
  try { Ninja-Property-Set idleTimeStatus "ALERT: Idle $idleMin min (>= $ThresholdMinutes)" } catch { Write-Verbose "Ninja-Property-Set failed: $_" }
  Write-Error "Idle time threshold exceeded: $idleMin minutes (threshold: $ThresholdMinutes)."
  exit 1
}
exit 0

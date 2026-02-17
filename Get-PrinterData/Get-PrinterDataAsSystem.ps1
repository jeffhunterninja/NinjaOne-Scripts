#Requires -Version 5.1
<#
.SYNOPSIS
  Collects printer data from all logged-in user sessions via CreateProcessAsUser and populates NinjaOne custom fields.

.DESCRIPTION
  Runs as SYSTEM and uses CreateProcessAsUser to launch a helper in each active user session.
  Each helper runs Get-Printer in user context (capturing per-user HKCU and machine-wide HKLM printers),
  writes JSON to a per-session temp file. The main script merges results, deduplicates, and updates
  NinjaOne custom fields. When no users are logged in, falls back to Get-Printer in SYSTEM context
  (HKLM printers only).

.EXIT CODES
  0 = Success
  1 = Not elevated
  2 = Error (parse failure, Set-NinjaProperty failure)
#>

[CmdletBinding()]
param(
  [int]$PerProcessTimeoutSeconds = 20
)

$ErrorActionPreference = 'Stop'

function Test-IsElevated {
  $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
  $p  = [System.Security.Principal.WindowsPrincipal]::new($id)
  return $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ---------- C# interop: enumerate sessions, spawn helper in session ----------
if (-not ("SessionProcessHelper" -as [type])) {
Add-Type -Language CSharp @"
using System;
using System.Runtime.InteropServices;

public static class SessionProcessHelper
{
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
        public string State;
    }

    public static SessionInfo[] ListSessions() {
        var list = new System.Collections.Generic.List<SessionInfo>();
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

    public static bool RunInSession(int sessionId, string commandLine, int timeoutMs) {
        int exit;
        return RunInSessionAndGetExitCode(sessionId, commandLine, timeoutMs, out exit);
    }

    public static bool RunInSessionAndGetExitCode(int sessionId, string commandLine, int timeoutMs, out int exit) {
        exit = -1;
        IntPtr hUserTok = IntPtr.Zero;
        IntPtr hPrimTok = IntPtr.Zero;
        IntPtr env = IntPtr.Zero;
        PROCESS_INFORMATION pi = new PROCESS_INFORMATION();
        try {
            if (!WTSQueryUserToken(sessionId, out hUserTok) || hUserTok == IntPtr.Zero) return false;

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
                env = IntPtr.Zero;
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
  exit 1
}

if (-not (Get-Command Set-NinjaProperty -ErrorAction SilentlyContinue)) {
  Write-Warning "Set-NinjaProperty cmdlet not found; NinjaOne custom fields will not be updated."
}

$timeoutMs = [Math]::Max(5000, $PerProcessTimeoutSeconds * 1000)
$tempDir = Join-Path $env:SystemRoot "Temp"
$filePattern = "printer_info_*.json"

# Helper template: runs in user session, writes JSON to session-specific file
function Get-HelperCommand {
  param([int]$SessionId)
  @"
`$path = Join-Path `$env:SystemRoot "Temp\printer_info_$SessionId.json"
try {
  `$printers = Get-Printer | Select-Object Name, DriverName
  `$json = `$printers | ConvertTo-Json -Depth 2 -Compress
  [System.IO.File]::WriteAllText(`$path, `$json, [System.Text.UTF8Encoding]::new(`$false))
} catch {
  # Silently fail; main script will skip this session
}
"@
}

# 1) Enumerate sessions
$sessions = [SessionProcessHelper]::ListSessions() | ForEach-Object {
  [PSCustomObject]@{
    SessionId  = $_.SessionId
    WinStation = $_.WinStation
    State      = $_.State
  }
}
Write-Verbose "Enumerated $($sessions.Count) session(s)."

# 2) Run helper in each active/connected/idle session
$activeStates = @('WTSActive', 'WTSConnected', 'WTSIdle')
foreach ($s in $sessions) {
  if ($s.State -in $activeStates) {
    $helperCmd = Get-HelperCommand -SessionId $s.SessionId
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($helperCmd)
    $encCmd = [Convert]::ToBase64String($bytes)
    $psLine = "powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -EncodedCommand $encCmd"

    Write-Verbose "Collecting printers from session $($s.SessionId) ($($s.WinStation))."
    $null = [SessionProcessHelper]::RunInSession([int]$s.SessionId, $psLine, $timeoutMs)
  }
}

# 3) Collect and merge JSON files
$jsonFiles = @(Get-ChildItem -Path $tempDir -Filter $filePattern -ErrorAction SilentlyContinue)
$allPrinters = [System.Collections.Generic.List[PSCustomObject]]::new()

if ($jsonFiles.Count -gt 0) {
  foreach ($f in $jsonFiles) {
    try {
      $raw = Get-Content -Raw -Path $f.FullName
      $printers = $raw | ConvertFrom-Json
      $printers = @($printers)
      foreach ($p in $printers) {
        if ($p -and ($p.PSObject.Properties['Name'] -or $null -ne $p.Name)) {
          $allPrinters.Add($p)
        }
      }
    } catch {
      Write-Verbose "Could not parse $($f.Name): $_"
    }
  }
}

# 4) No-session fallback: run Get-Printer in SYSTEM context (HKLM only)
if ($allPrinters.Count -eq 0) {
  Write-Verbose "No session data collected; falling back to SYSTEM-context Get-Printer (HKLM printers only)."
  try {
    $printers = Get-Printer | Select-Object Name, DriverName
    foreach ($p in $printers) {
      $allPrinters.Add($p)
    }
  } catch {
    Write-Verbose "SYSTEM fallback Get-Printer failed: $_"
  }
}

# 5) Deduplicate and extract
$printerNames = $allPrinters | ForEach-Object { $_.Name } | Where-Object { $_ } | Sort-Object -Unique
$driverNames  = $allPrinters | ForEach-Object { $_.DriverName } | Where-Object { $_ } | Sort-Object -Unique

$printerValues = $printerNames -join "`r`n"
$driverValues  = $driverNames  -join "`r`n"

# 6) Update NinjaOne custom fields
try {
  if (Get-Command Set-NinjaProperty -ErrorAction SilentlyContinue) {
    Set-NinjaProperty "printers"       $printerValues
    Set-NinjaProperty "printerDrivers" $driverValues
  }
} catch {
  Write-Error "Could not write to NinjaOne custom fields: $_"
  exit 2
}

# 7) Cleanup temp files
foreach ($f in $jsonFiles) {
  try {
    Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue
  } catch {
    Write-Verbose "Could not remove temp file $($f.Name): $_"
  }
}

Write-Host "Collected $($printerNames.Count) printer(s), $($driverNames.Count) driver(s)."
exit 0

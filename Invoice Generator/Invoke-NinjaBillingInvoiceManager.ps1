#Requires -Version 5.1
<#
.SYNOPSIS
    WPF tool: export NinjaOne billing invoices to HTML and create PSA tickets using Authorization Code + PKCE only.

.DESCRIPTION
    Standalone script (no dot-sourcing). Billing export and PSA ticketing both use the same OAuth session:
    Authorization Code + PKCE, refresh_token without client_secret, scopes monitoring management offline_access.
    Calls https://{host}/v2/billing/... and /v2/ticketing/... on the signed-in instance. Optional encrypted refresh
    storage under %APPDATA%\NinjaBillingInvoiceManager.

.PARAMETER AllowInsecureHttp
    Allow http:// instance URLs for testing only.
#>
[CmdletBinding()]
param(
    [switch]$AllowInsecureHttp
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
try { Add-Type -AssemblyName System.Net.Http -ErrorAction Stop } catch {
    throw 'System.Net.Http is required for ticket attachments.'
}

#region Script state
$script:AllowInsecureHttp = $AllowInsecureHttp.IsPresent
$script:AppConfigDir  = Join-Path $env:APPDATA 'NinjaBillingInvoiceManager'
$script:AppConfigFile = Join-Path $script:AppConfigDir 'config.json'
$script:ManifestFileName = 'NinjaBillingInvoiceExport-manifest.json'

$script:MasterPassword = $null
$script:MasterPasswordVerifier = $null

$script:PsaAccessToken    = $null
$script:PsaRefreshToken   = $null
$script:PsaTokenExpiresAt = [datetime]::MinValue
$script:PsaBaseUrl       = ''
$script:PsaClientId      = ''

$script:AuthPS          = $null
$script:AuthHandle      = $null
$script:AuthVerifier    = $null
$script:AuthState       = $null
$script:AuthRedirectUri = $null
$script:AuthListener    = $null
$script:AuthTimeoutAt   = [datetime]::MinValue

$script:ManifestRowsTable = $null
$script:ManifestPathOnDisk = $null
$script:LastExportOutputPath = ''
#endregion

#region Normalize host
function Resolve-BaseUrl {
    param([string]$Instance)
    $u = if ($null -eq $Instance) { '' } else { $Instance.Trim() }
    if ([string]::IsNullOrWhiteSpace($u)) { $u = 'https://app.ninjarmm.com' }
    if ($u -notmatch '^[a-zA-Z][a-zA-Z0-9+\-.]*://') { $u = "https://$u" }
    $uri = $null
    if (-not [System.Uri]::TryCreate($u, [System.UriKind]::Absolute, [ref]$uri)) {
        throw "Invalid NinjaOne instance URL: '$Instance'."
    }
    if ($uri.Scheme -eq 'http' -and -not $script:AllowInsecureHttp) {
        throw 'Insecure HTTP is not allowed. Use HTTPS or -AllowInsecureHttp for local testing only.'
    }
    if ($uri.Scheme -ne 'https' -and $uri.Scheme -ne 'http') {
        throw "Unsupported URL scheme '$($uri.Scheme)'."
    }
    return $uri.AbsoluteUri.TrimEnd('/')
}
#endregion

#region Crypto (AES-256-CBC + PBKDF2)
function Protect-String {
    param([string]$PlainText, [string]$MasterPwd)
    $salt = [byte[]]::new(32)
    $iv   = [byte[]]::new(16)
    $rng  = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($salt); $rng.GetBytes($iv); $rng.Dispose()
    $kdf = [System.Security.Cryptography.Rfc2898DeriveBytes]::new(
        $MasterPwd, $salt, 100000, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $key = $kdf.GetBytes(32); $kdf.Dispose()
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Key = $key; $aes.IV = $iv
    $aes.Mode    = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $enc   = $aes.CreateEncryptor()
    $plain = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
    $cipher = $enc.TransformFinalBlock($plain, 0, $plain.Length)
    $enc.Dispose(); $aes.Dispose()
    $combined = [byte[]]::new(32 + 16 + $cipher.Length)
    [Array]::Copy($salt,   0, $combined, 0,  32)
    [Array]::Copy($iv,     0, $combined, 32, 16)
    [Array]::Copy($cipher, 0, $combined, 48, $cipher.Length)
    [Array]::Clear($key,   0, $key.Length)
    [Array]::Clear($plain, 0, $plain.Length)
    return [Convert]::ToBase64String($combined)
}

function Unprotect-String {
    param([string]$CipherText, [string]$MasterPwd)
    $combined = [Convert]::FromBase64String($CipherText)
    $salt   = [byte[]]::new(32)
    $iv     = [byte[]]::new(16)
    $cipher = [byte[]]::new($combined.Length - 48)
    [Array]::Copy($combined, 0,  $salt,   0, 32)
    [Array]::Copy($combined, 32, $iv,     0, 16)
    [Array]::Copy($combined, 48, $cipher, 0, $cipher.Length)
    $kdf = [System.Security.Cryptography.Rfc2898DeriveBytes]::new(
        $MasterPwd, $salt, 100000, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $key = $kdf.GetBytes(32); $kdf.Dispose()
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Key = $key; $aes.IV = $iv
    $aes.Mode    = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $dec   = $aes.CreateDecryptor()
    $plain = $dec.TransformFinalBlock($cipher, 0, $cipher.Length)
    $dec.Dispose(); $aes.Dispose()
    $result = [System.Text.Encoding]::UTF8.GetString($plain)
    [Array]::Clear($key,   0, $key.Length)
    [Array]::Clear($plain, 0, $plain.Length)
    return $result
}

function New-MasterPasswordVerifier {
    param([string]$MasterPwd)
    $salt = [byte[]]::new(32)
    $rng  = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($salt); $rng.Dispose()
    $kdf = [System.Security.Cryptography.Rfc2898DeriveBytes]::new(
        $MasterPwd, $salt, 100000, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $hash = $kdf.GetBytes(32); $kdf.Dispose()
    $combined = [byte[]]::new(64)
    [Array]::Copy($salt, 0, $combined, 0,  32)
    [Array]::Copy($hash, 0, $combined, 32, 32)
    [Array]::Clear($hash, 0, $hash.Length)
    return [Convert]::ToBase64String($combined)
}

function Test-MasterPasswordValid {
    param([string]$MasterPwd, [string]$Verifier)
    $combined   = [Convert]::FromBase64String($Verifier)
    $salt       = [byte[]]::new(32)
    $storedHash = [byte[]]::new(32)
    [Array]::Copy($combined, 0,  $salt,       0, 32)
    [Array]::Copy($combined, 32, $storedHash, 0, 32)
    $kdf = [System.Security.Cryptography.Rfc2898DeriveBytes]::new(
        $MasterPwd, $salt, 100000, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $computed = $kdf.GetBytes(32); $kdf.Dispose()
    $diff = 0
    for ($i = 0; $i -lt 32; $i++) { $diff = $diff -bor ($storedHash[$i] -bxor $computed[$i]) }
    [Array]::Clear($computed, 0, $computed.Length)
    return ($diff -eq 0)
}
#endregion

#region Config
function Get-AppConfig {
    $defaults = [PSCustomObject]@{
        PsaBaseUrl               = ''
        PsaClientId              = ''
        EncryptedPsaRefreshToken = ''
        MasterPasswordVerifier   = ''
    }
    if (Test-Path $script:AppConfigFile) {
        try {
            $raw = Get-Content $script:AppConfigFile -Raw | ConvertFrom-Json
            foreach ($prop in $raw.PSObject.Properties) {
                if ($defaults.PSObject.Properties[$prop.Name]) {
                    $defaults.$($prop.Name) = [string]$prop.Value
                }
            }
        } catch {
            Write-Verbose "Config load failed: $($_.Exception.Message)"
        }
    }
    return $defaults
}

function Save-AppConfig {
    param(
        [string]$PsaBaseUrl,
        [string]$PsaClientId,
        [string]$EncryptedPsaRefresh,
        [string]$Verifier
    )
    if (-not (Test-Path $script:AppConfigDir)) {
        New-Item -ItemType Directory -Path $script:AppConfigDir -Force | Out-Null
    }
    $disk = [ordered]@{
        PsaBaseUrl               = $PsaBaseUrl
        PsaClientId              = $PsaClientId
        EncryptedPsaRefreshToken = $EncryptedPsaRefresh
        MasterPasswordVerifier   = $Verifier
    }
    [PSCustomObject]$disk | ConvertTo-Json -Depth 5 |
        Set-Content $script:AppConfigFile -Encoding UTF8
}

function Save-CurrentSessionToDisk {
    if (-not $script:MasterPassword) { return }
    $verifier = if ($script:MasterPasswordVerifier) {
        $script:MasterPasswordVerifier
    } else {
        $v = New-MasterPasswordVerifier -MasterPwd $script:MasterPassword
        $script:MasterPasswordVerifier = $v
        $v
    }
    $cfgPrev = Get-AppConfig
    $encPsa = ''
    if ($script:PsaRefreshToken) {
        $encPsa = Protect-String -PlainText $script:PsaRefreshToken -MasterPwd $script:MasterPassword
    } else {
        $encPsa = $cfgPrev.EncryptedPsaRefreshToken
    }
    $psaUrl = $script:PsaBaseUrl
    if ([string]::IsNullOrWhiteSpace($psaUrl)) { $psaUrl = $cfgPrev.PsaBaseUrl }
    $psaCid = $script:PsaClientId
    if ([string]::IsNullOrWhiteSpace($psaCid)) { $psaCid = $cfgPrev.PsaClientId }
    Save-AppConfig -PsaBaseUrl $psaUrl `
        -PsaClientId $psaCid `
        -EncryptedPsaRefresh $encPsa `
        -Verifier $verifier
}

function Clear-SavedSession {
    if (Test-Path $script:AppConfigFile) {
        Remove-Item $script:AppConfigFile -Force -ErrorAction SilentlyContinue
    }
    $script:MasterPassword = $null
    $script:MasterPasswordVerifier = $null
    $script:PsaRefreshToken = $null
    $script:PsaAccessToken = $null
}
#endregion

#region Master password dialogs
function Show-MasterPasswordPrompt {
    param(
        [string]$Title = 'Master Password',
        [string]$Message = 'Enter your master password:',
        [switch]$IsNewPassword
    )
    $dlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" SizeToContent="WidthAndHeight" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" MinWidth="380" MaxWidth="440"
        Background="#FAFAFA">
  <StackPanel Margin="20">
    <TextBlock Text="$Message" TextWrapping="Wrap" Margin="0,0,0,12" FontSize="13"/>
    <TextBlock Text="Password" FontSize="11" Foreground="#555" Margin="0,0,0,4"/>
    <PasswordBox x:Name="pbPassword" Height="28" Margin="0,0,0,4"/>
    $(if ($IsNewPassword) {
    '<TextBlock Text="Confirm Password" FontSize="11" Foreground="#555" Margin="0,8,0,4"/>' +
    '<PasswordBox x:Name="pbConfirm" Height="28" Margin="0,0,0,4"/>' +
    '<TextBlock FontSize="11" Foreground="#888" Margin="0,2,0,0" Text="Minimum 8 characters."/>'
    } else { '' })
    <TextBlock x:Name="lblError" Foreground="Red" FontSize="11" Margin="0,6,0,0"
               TextWrapping="Wrap" Visibility="Collapsed"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
      <Button x:Name="btnOK" Content="OK" Width="80" IsDefault="True" Margin="0,0,8,0"/>
      <Button x:Name="btnCancel" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
</Window>
"@
    $dlgReader = New-Object System.Xml.XmlNodeReader ([xml]$dlgXaml)
    $dlg = [Windows.Markup.XamlReader]::Load($dlgReader)
    $dlg.Owner = $window
    $pbPwd      = $dlg.FindName('pbPassword')
    $pbConfirm  = if ($IsNewPassword) { $dlg.FindName('pbConfirm') } else { $null }
    $lblErr     = $dlg.FindName('lblError')
    $btnDlgOK   = $dlg.FindName('btnOK')
    $btnDlgCanc = $dlg.FindName('btnCancel')
    $btnDlgOK.Add_Click({
        $enteredPwd = $pbPwd.Password
        if ([string]::IsNullOrWhiteSpace($enteredPwd)) {
            $lblErr.Text = 'Password cannot be empty.'
            $lblErr.Visibility = 'Visible'
            return
        }
        if ($IsNewPassword) {
            if ($enteredPwd.Length -lt 8) {
                $lblErr.Text = 'Password must be at least 8 characters.'
                $lblErr.Visibility = 'Visible'
                return
            }
            if ($enteredPwd -ne $pbConfirm.Password) {
                $lblErr.Text = 'Passwords do not match.'
                $lblErr.Visibility = 'Visible'
                return
            }
        }
        $dlg.Tag = $enteredPwd
        $dlg.DialogResult = $true
        $dlg.Close()
    }.GetNewClosure())
    $btnDlgCanc.Add_Click({
        $dlg.DialogResult = $false
        $dlg.Close()
    }.GetNewClosure())
    $dlg.Add_ContentRendered({ $pbPwd.Focus() }.GetNewClosure())
    if ($dlg.ShowDialog()) { return $dlg.Tag }
    return $null
}

function Show-ChangeMasterPasswordPrompt {
    $dlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Change Master Password" SizeToContent="WidthAndHeight" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" MinWidth="380" MaxWidth="440"
        Background="#FAFAFA">
  <StackPanel Margin="20">
    <TextBlock Text="Enter your current password, then choose a new one."
               TextWrapping="Wrap" Margin="0,0,0,12" FontSize="13"/>
    <TextBlock Text="Current Password" FontSize="11" Foreground="#555" Margin="0,0,0,4"/>
    <PasswordBox x:Name="pbCurrent" Height="28" Margin="0,0,0,8"/>
    <TextBlock Text="New Password" FontSize="11" Foreground="#555" Margin="0,0,0,4"/>
    <PasswordBox x:Name="pbNew" Height="28" Margin="0,0,0,4"/>
    <TextBlock Text="Confirm New Password" FontSize="11" Foreground="#555" Margin="0,8,0,4"/>
    <PasswordBox x:Name="pbConfirmNew" Height="28" Margin="0,0,0,4"/>
    <TextBlock x:Name="lblError" Foreground="Red" FontSize="11" Margin="0,6,0,0"
               TextWrapping="Wrap" Visibility="Collapsed"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
      <Button x:Name="btnOK" Content="Change" Width="80" IsDefault="True" Margin="0,0,8,0"/>
      <Button x:Name="btnCancel" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
</Window>
"@
    $dlgReader = New-Object System.Xml.XmlNodeReader ([xml]$dlgXaml)
    $dlg = [Windows.Markup.XamlReader]::Load($dlgReader)
    $dlg.Owner = $window
    $pbCur      = $dlg.FindName('pbCurrent')
    $pbNew      = $dlg.FindName('pbNew')
    $pbConf     = $dlg.FindName('pbConfirmNew')
    $lblErr     = $dlg.FindName('lblError')
    $btnDlgOK   = $dlg.FindName('btnOK')
    $btnDlgCanc = $dlg.FindName('btnCancel')
    $btnDlgOK.Add_Click({
        $cur = $pbCur.Password
        if (-not $script:MasterPasswordVerifier -or
            -not (Test-MasterPasswordValid -MasterPwd $cur -Verifier $script:MasterPasswordVerifier)) {
            $lblErr.Text = 'Current password is incorrect.'
            $lblErr.Visibility = 'Visible'
            return
        }
        $npwd = $pbNew.Password
        if ($npwd.Length -lt 8) {
            $lblErr.Text = 'New password must be at least 8 characters.'
            $lblErr.Visibility = 'Visible'
            return
        }
        if ($npwd -ne $pbConf.Password) {
            $lblErr.Text = 'New passwords do not match.'
            $lblErr.Visibility = 'Visible'
            return
        }
        $dlg.Tag = $npwd
        $dlg.DialogResult = $true
        $dlg.Close()
    }.GetNewClosure())
    $btnDlgCanc.Add_Click({
        $dlg.DialogResult = $false
        $dlg.Close()
    }.GetNewClosure())
    $dlg.Add_ContentRendered({ $pbCur.Focus() }.GetNewClosure())
    if ($dlg.ShowDialog()) { return $dlg.Tag }
    return $null
}
#endregion

#region UI helper
function Push-UIUpdate {
    $frame = [System.Windows.Threading.DispatcherFrame]::new()
    $cb = [System.Windows.Threading.DispatcherOperationCallback]{
        param([object]$state)
        ([System.Windows.Threading.DispatcherFrame]$state).Continue = $false
        return $null
    }
    [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        $cb,
        $frame
    )
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Set-StatusBar {
    param([string]$Text)
    $lbl = $window.FindName('lblStatus')
    if ($lbl) {
        $lbl.Text = $Text
        Push-UIUpdate
    }
}
#endregion

#region PKCE + PSA tokens (public client, no secret on refresh)
function New-PkceVerifier {
    $buf = [byte[]]::new(48)
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($buf)
    return [Convert]::ToBase64String($buf) -replace '\+','-' -replace '/','_' -replace '=',''
}

function Get-PkceChallenge {
    param([string]$Verifier)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $hash = $sha.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($Verifier))
    return [Convert]::ToBase64String($hash) -replace '\+','-' -replace '/','_' -replace '=',''
}

function New-OAuthState { return New-PkceVerifier }

function Test-PsaTokenValid {
    return ($null -ne $script:PsaAccessToken -and
        -not [string]::IsNullOrWhiteSpace($script:PsaAccessToken) -and
        [datetime]::UtcNow -lt $script:PsaTokenExpiresAt)
}

function Test-PsaRefreshPresent {
    return -not [string]::IsNullOrWhiteSpace($script:PsaRefreshToken)
}

function Update-PsaTokensFromResponse {
    param($Response)
    if ([string]::IsNullOrWhiteSpace($Response.access_token)) {
        throw 'Token response did not include access_token.'
    }
    $script:PsaAccessToken = [string]$Response.access_token
    if ($Response.refresh_token) {
        $script:PsaRefreshToken = [string]$Response.refresh_token
    }
    $exp = if ($Response.expires_in) { [int]$Response.expires_in } else { 3600 }
    if ($exp -le 0) { $exp = 3600 }
    $script:PsaTokenExpiresAt = [datetime]::UtcNow.AddSeconds($exp - 60)
    if ($script:MasterPassword -and $script:PsaRefreshToken) {
        try { Save-CurrentSessionToDisk } catch { Write-Verbose $_.Exception.Message }
    }
}

function Invoke-PsaTokenRefresh {
    if (-not (Test-PsaRefreshPresent)) { throw 'No refresh token. Sign in again.' }
    $resp = Invoke-RestMethod -Uri "$($script:PsaBaseUrl)/ws/oauth/token" `
        -Method POST -UseBasicParsing `
        -Headers @{
            'Accept'       = 'application/json'
            'Content-Type' = 'application/x-www-form-urlencoded'
        } `
        -Body @{
            grant_type    = 'refresh_token'
            refresh_token = $script:PsaRefreshToken
            client_id     = $script:PsaClientId
        }
    Update-PsaTokensFromResponse $resp
}
#endregion

#region PSA session (/v2 with refresh, public client)
function Get-NinjaPsaTokenFromRefresh {
    param([string]$ClientId, [string]$RefreshToken, [string]$HostOnly)
    $uri = "https://$HostOnly/ws/oauth/token"
    $body = @{
        grant_type    = 'refresh_token'
        client_id     = $ClientId
        refresh_token = $RefreshToken
    }
    return Invoke-RestMethod -Uri $uri -Method POST -Headers @{
        'Accept' = 'application/json'; 'Content-Type' = 'application/x-www-form-urlencoded'
    } -Body $body -TimeoutSec 30 -ErrorAction Stop
}

function Apply-PsaRefreshToSession {
    param([PSCustomObject]$Session, $TokenResponse)
    if ([string]::IsNullOrWhiteSpace($TokenResponse.access_token)) { throw 'No access_token in refresh response.' }
    $Session.AuthHeader = "Bearer $($TokenResponse.access_token)"
    $Session.ExpiresAt  = if ($TokenResponse.expires_in) {
        [datetime]::UtcNow.AddSeconds([int]$TokenResponse.expires_in - 60)
    } else {
        [datetime]::UtcNow.AddMinutes(55)
    }
    if ($TokenResponse.refresh_token) {
        $Session.RefreshToken = [string]$TokenResponse.refresh_token
        $script:PsaRefreshToken = $Session.RefreshToken
    }
}

#region Billing: /v2 API (same PKCE refresh session as PSA)
function Invoke-NinjaBillingApi {
    [CmdletBinding()]
    param(
        [ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Endpoint,
        [string]$Query,
        $Body,
        [int]$TimeoutSec = 120,
        [int]$MaxRetries = 4,
        [Parameter(Mandatory)][PSCustomObject]$Session
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
    function Refresh-Billing([PSCustomObject]$S) {
        $t = Get-NinjaPsaTokenFromRefresh -ClientId $S.ClientId -RefreshToken $S.RefreshToken -HostOnly $S.HostOnly
        Apply-PsaRefreshToSession -Session $S -TokenResponse $t
        $script:PsaAccessToken = ($S.AuthHeader -replace '^Bearer ', '')
        $script:PsaTokenExpiresAt = $S.ExpiresAt
        if ($script:MasterPassword) { try { Save-CurrentSessionToDisk } catch { } }
    }
    if (-not $Session.ExpiresAt -or [datetime]::UtcNow -ge $Session.ExpiresAt) { Refresh-Billing $Session }
    $base = "https://$($Session.HostOnly)/$($Endpoint.TrimStart('/'))"
    $uri  = if ($Query) { "${base}?${Query}" } else { $base }
    $attempt = 0
    while ($true) {
        $reqHeaders = @{ Authorization = $Session.AuthHeader; Accept = 'application/json' }
        $bodyJson = $null
        if ($Body) {
            $bodyJson = $Body | ConvertTo-Json -Depth 20
            $reqHeaders['Content-Type'] = 'application/json'
        }
        try {
            return Invoke-RestMethod -Uri $uri -Method $Method -Headers $reqHeaders -Body $bodyJson -TimeoutSec $TimeoutSec -ErrorAction Stop
        } catch {
            $status = Get-HttpStatus $_
            $attempt++
            if ($status -eq 401 -and $attempt -le $MaxRetries) { Refresh-Billing $Session; continue }
            $retryable = $status -in @(408, 429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -gt $MaxRetries) { throw }
            $wait  = Get-RetryAfter $_
            $sleep = if ($wait -and $wait -gt 0) { [Math]::Min($wait, 60) } else { [Math]::Min([Math]::Pow(2, $attempt), 30) }
            Start-Sleep -Seconds $sleep
        }
    }
}
#endregion

function Invoke-NinjaPsaApi {
    [CmdletBinding()]
    param(
        [ValidateSet('GET', 'POST', 'PUT', 'PATCH', 'DELETE')]
        [string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Endpoint,
        [string]$Query,
        $Body,
        [int]$TimeoutSec = 120,
        [int]$MaxRetries = 4,
        [Parameter(Mandatory)][PSCustomObject]$Session
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
    function Refresh-Psa([PSCustomObject]$S) {
        $t = Get-NinjaPsaTokenFromRefresh -ClientId $S.ClientId -RefreshToken $S.RefreshToken -HostOnly $S.HostOnly
        Apply-PsaRefreshToSession -Session $S -TokenResponse $t
        $script:PsaAccessToken = ($S.AuthHeader -replace '^Bearer ', '')
        $script:PsaTokenExpiresAt = $S.ExpiresAt
        if ($script:MasterPassword) { try { Save-CurrentSessionToDisk } catch { } }
    }
    if (-not $Session.ExpiresAt -or [datetime]::UtcNow -ge $Session.ExpiresAt) { Refresh-Psa $Session }
    $base = "https://$($Session.HostOnly)/$($Endpoint.TrimStart('/'))"
    $uri  = if ($Query) { "${base}?${Query}" } else { $base }
    $attempt = 0
    while ($true) {
        $reqHeaders = @{ Authorization = $Session.AuthHeader; Accept = 'application/json' }
        $bodyJson = $null
        if ($Body) {
            $bodyJson = $Body | ConvertTo-Json -Depth 20
            $reqHeaders['Content-Type'] = 'application/json'
        }
        try {
            return Invoke-RestMethod -Uri $uri -Method $Method -Headers $reqHeaders -Body $bodyJson -TimeoutSec $TimeoutSec -ErrorAction Stop
        } catch {
            $status = Get-HttpStatus $_
            $attempt++
            if ($status -eq 401 -and $attempt -le $MaxRetries) { Refresh-Psa $Session; continue }
            $retryable = $status -in @(408, 429, 500, 502, 503, 504)
            if (-not $retryable -or $attempt -gt $MaxRetries) { throw }
            $wait  = Get-RetryAfter $_
            $sleep = if ($wait -and $wait -gt 0) { [Math]::Min($wait, 60) } else { [Math]::Min([Math]::Pow(2, $attempt), 30) }
            Start-Sleep -Seconds $sleep
        }
    }
}

function Update-PsaSessionFromScriptToken {
    param([Parameter(Mandatory)][PSCustomObject]$Session)
    if (-not (Test-PsaTokenValid)) { Invoke-PsaTokenRefresh }
    $Session.AuthHeader = "Bearer $($script:PsaAccessToken)"
    $Session.ExpiresAt  = $script:PsaTokenExpiresAt
    $Session.RefreshToken = $script:PsaRefreshToken
}

function Send-NinjaTicketCommentWithAttachment {
    param(
        [Parameter(Mandatory)][PSCustomObject]$Session,
        [Parameter(Mandatory)][int]$TicketId,
        [Parameter(Mandatory)][string]$CommentBody,
        [Parameter(Mandatory)][string]$HtmlFilePath
    )
    Update-PsaSessionFromScriptToken -Session $Session
    $uri = "https://$($Session.HostOnly)/v2/ticketing/ticket/$TicketId/comment"
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
#endregion

#region Invoice HTML + manifest
function Format-MoneyInv {
    param([object]$Amount, [string]$Currency = 'USD')
    if ($null -eq $Amount) { return '-' }
    return "$([string]::Format('{0:N2}', [double]$Amount)) $Currency"
}

function Format-InvoiceDateInv {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    try { return ([datetime]$Value).ToString('MMM d, yyyy') } catch { return $Value }
}

function HtmlEncInv([string]$s) {
    return [System.Web.HttpUtility]::HtmlEncode($s)
}

function Build-InvoiceHtml {
    param([PSCustomObject]$Invoice)
    $cur        = if ($Invoice.currency) { $Invoice.currency } else { 'USD' }
    $invoiceNum = if ($Invoice.invoiceNumber) { $Invoice.invoiceNumber } else { "INV-$($Invoice.id)" }
    $dueDateHtml = if ($Invoice.dueDate) { "<p>Due: $(Format-InvoiceDateInv $Invoice.dueDate)</p>" } else { '' }
    $issuer = $Invoice.billContent
    $fromLines = [System.Collections.Generic.List[string]]::new()
    $toLines   = [System.Collections.Generic.List[string]]::new()
    if ($issuer) {
        if ($issuer.issuerName)    { $fromLines.Add("<strong>$(HtmlEncInv $issuer.issuerName)</strong>") }
        if ($issuer.issuerAddress) { $fromLines.Add((HtmlEncInv $issuer.issuerAddress)) }
        if ($issuer.issuerPhone)   { $fromLines.Add((HtmlEncInv $issuer.issuerPhone)) }
        if ($issuer.issuerWebPage) { $fromLines.Add("<a href='$(HtmlEncInv $issuer.issuerWebPage)'>$(HtmlEncInv $issuer.issuerWebPage)</a>") }
        if ($issuer.billName)    { $toLines.Add("<strong>$(HtmlEncInv $issuer.billName)</strong>") }
        if ($issuer.billAddress) { $toLines.Add((HtmlEncInv $issuer.billAddress)) }
        if ($issuer.billEmail)   { $toLines.Add((HtmlEncInv $issuer.billEmail)) }
        if ($issuer.billPhone)   { $toLines.Add((HtmlEncInv $issuer.billPhone)) }
    }
    if ($toLines.Count -eq 0 -and $Invoice.client -and $Invoice.client.name) {
        $toLines.Add("<strong>$(HtmlEncInv $Invoice.client.name)</strong>")
    }
    $fromHtml = $fromLines -join '<br>'
    $toHtml   = $toLines   -join '<br>'
    $agreementHtml = ''
    if ($Invoice.agreement -and $Invoice.agreement.name) {
        $agreementHtml = "<div class='meta-item'><span class='meta-label'>Agreement</span>$(HtmlEncInv $Invoice.agreement.name)</div>"
    }
    $intervalHtml = ''
    if ($Invoice.interval) {
        $intervalHtml = "<div class='meta-item'><span class='meta-label'>Interval</span>$(HtmlEncInv $Invoice.interval)</div>"
    }
    $allProducts = @()
    if ($Invoice.products)          { $allProducts += @($Invoice.products) }
    if ($Invoice.agreementProducts) { $allProducts += @($Invoice.agreementProducts) }
    if ($Invoice.ticketProducts)    { $allProducts += @($Invoice.ticketProducts) }
    $lineRowsHtml = [System.Text.StringBuilder]::new()
    foreach ($p in $allProducts) {
        $name      = if ($p.name)        { HtmlEncInv $p.name }        else { '(unnamed)' }
        $desc      = if ($p.description) { HtmlEncInv $p.description } else { '' }
        $nameCell  = if ($desc) { "<strong>$name</strong><br><span class='subdesc'>$desc</span>" } else { "<strong>$name</strong>" }
        $qty       = if ($null -ne $p.quantity) { [double]$p.quantity } else { 0 }
        $unitPrice = if ($null -ne $p.price)    { [double]$p.price }    else { 0 }
        $lineAmt   = if ($null -ne $p.subTotalWithDiscount) { [double]$p.subTotalWithDiscount }
                     elseif ($null -ne $p.subTotal)         { [double]$p.subTotal }
                     else                                   { $unitPrice * $qty }
        $discCell  = if ($p.discount -and [double]$p.discount -ne 0) {
                         "<br><span class='subdesc negative'>(- $(Format-MoneyInv $p.discount $cur))</span>"
                     } else { '' }
        [void]$lineRowsHtml.AppendLine(
            "<tr><td class='td-name'>$nameCell</td><td class='td-num'>$qty</td>" +
            "<td class='td-num'>$(Format-MoneyInv $unitPrice $cur)</td>" +
            "<td class='td-num'><span class='line-total'>$(Format-MoneyInv $lineAmt $cur)</span>$discCell</td></tr>"
        )
    }
    $discountRowHtml = ''
    if ($Invoice.discount -and [double]$Invoice.discount -ne 0) {
        $discountRowHtml = "<tr><td colspan='2' class='td-label'>Discount</td><td class='td-num negative'>- $(Format-MoneyInv $Invoice.discount $cur)</td></tr>"
    }
    $taxPct   = if ($Invoice.taxRate) { [string]::Format('{0:P1}', [double]$Invoice.taxRate) } else { '' }
    $taxLabel = if ($taxPct) { "Tax ($taxPct)" } else { 'Tax' }
    $notesHtml = ''
    if (-not [string]::IsNullOrWhiteSpace($Invoice.invoiceNote)) {
        $notesHtml = "<div class='notes-section'><h3>Notes</h3><p>$(HtmlEncInv $Invoice.invoiceNote)</p></div>"
    }
    $generatedDate = (Get-Date).ToString('MMMM d, yyyy')
    $periodStart   = Format-InvoiceDateInv $Invoice.billingPeriodStartDate
    $periodEnd     = Format-InvoiceDateInv $Invoice.billingPeriodEndDate
    $invDate       = Format-InvoiceDateInv $Invoice.invoiceDate
    $statusVal     = if ($Invoice.status) { $Invoice.status } else { '' }
    $subtotal      = Format-MoneyInv $Invoice.subTotal $cur
    $totalTax      = Format-MoneyInv $Invoice.totalTax $cur
    $grandTotal    = Format-MoneyInv $Invoice.total $cur
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
  .ninja-invoice-ticket-banner { margin-bottom: 16px; padding: 10px 14px; background: #ecfdf5; border: 1px solid #10b981; border-radius: 6px; font-size: 13px; }
  .ninja-invoice-ticket-banner a { color: #047857; font-weight: 600; }
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
<div id="ninja-invoice-psa-ticket"></div>
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

function Expand-TicketSubject {
    param([string]$Template, [string]$InvoiceNumber, [string]$ClientName)
    $n = if ([string]::IsNullOrEmpty($ClientName)) { '' } else { $ClientName }
    $inv = if ([string]::IsNullOrEmpty($InvoiceNumber)) { '' } else { $InvoiceNumber }
    $s = $Template.Replace('{InvoiceNumber}', $inv).Replace('{ClientName}', $n)
    if ($s.Length -gt 200) { return $s.Substring(0, 200) }
    return $s
}

function Get-PsaTicketWebUrl {
    param([string]$BaseUrl, [int]$TicketId)
    $u = $BaseUrl.TrimEnd('/')
    return "$u/app/#/ticketing/ticket/$TicketId"
}

function Update-InvoiceHtmlWithTicket {
    param(
        [string]$HtmlFilePath,
        [string]$TicketUrl,
        [int]$TicketId
    )
    $html = [System.IO.File]::ReadAllText($HtmlFilePath, [System.Text.Encoding]::UTF8)
    $encUrl = [System.Web.HttpUtility]::HtmlEncode($TicketUrl)
    $banner = "<div class=`"ninja-invoice-ticket-banner`">PSA ticket: <a href=`"$encUrl`" target=`"_blank`" rel=`"noopener`">#$TicketId</a></div>"
    if ($html -match 'id="ninja-invoice-psa-ticket"') {
        $html = [regex]::Replace($html,
            '(?s)<div\s+id="ninja-invoice-psa-ticket"\s*>\s*</div>',
            "<div id=`"ninja-invoice-psa-ticket`">$banner</div>",
            1)
    } elseif ($html -match 'ninja-invoice-ticket-banner') {
        return
    } else {
        $html = $html -replace '(?i)</body>', "$banner`r`n</body>", 1
    }
    [System.IO.File]::WriteAllText($HtmlFilePath, $html, [System.Text.Encoding]::UTF8)
}

function Write-ManifestJson {
    param([System.Collections.IEnumerable]$Rows, [string]$ManifestPath)
    $list = @($Rows)
    if ($list.Count -eq 0) { $json = '[]' }
    elseif ($list.Count -eq 1) {
        $json = '[' + (ConvertTo-Json -InputObject $list[0] -Depth 6 -Compress) + ']'
    } else {
        $json = ConvertTo-Json -InputObject $list -Depth 6
    }
    [System.IO.File]::WriteAllText($ManifestPath, $json, [System.Text.Encoding]::UTF8)
}
#endregion

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="NinjaOne Billing Invoice Manager" Height="720" Width="820"
  WindowStartupLocation="CenterScreen" MinHeight="600" MinWidth="720">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="10,5"/>
      <Setter Property="Margin" Value="4,2"/>
    </Style>
  </Window.Resources>
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock Grid.Row="0" FontSize="18" FontWeight="Bold" Margin="0,0,0,8"
               Text="NinjaOne Billing Invoice Manager"/>
    <Expander Grid.Row="1" x:Name="expConnection" Header="Connection" IsExpanded="True" Margin="0,0,0,8">
      <StackPanel Margin="0,6,0,0">
        <GroupBox Header="NinjaOne (Authorization Code + PKCE)" Margin="0,0,0,8">
          <StackPanel Margin="8">
            <TextBlock TextWrapping="Wrap" FontSize="11" Foreground="#666" Margin="0,0,0,10"
              Text="HTML export and ticketing use this sign-in. OAuth app: public client, PKCE, offline_access."/>
            <Grid Margin="0,0,0,4">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="120"/><ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBlock VerticalAlignment="Center" Text="Instance"/>
              <TextBox Grid.Column="1" x:Name="tbPsaInstance" Height="26"/>
            </Grid>
            <Grid Margin="0,0,0,4">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="120"/><ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBlock VerticalAlignment="Center" Text="Client ID"/>
              <TextBox Grid.Column="1" x:Name="tbPsaClientId" Height="26"/>
            </Grid>
            <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
              <Button x:Name="btnPsaSignIn" Content="Sign in to NinjaOne" Width="170"/>
              <TextBlock x:Name="lblPsaAuthStatus" Margin="12,6,0,0" FontWeight="SemiBold" Foreground="Gray" Text="Not connected"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
              <Button x:Name="btnChangeMasterPwd" Content="Change master password" Visibility="Collapsed"/>
              <Button x:Name="btnClearSession" Content="Clear saved session" Margin="8,0,0,0" Visibility="Collapsed"/>
            </StackPanel>
          </StackPanel>
        </GroupBox>
      </StackPanel>
    </Expander>
    <TabControl Grid.Row="2" x:Name="tabMain" Margin="0,0,0,8">
      <TabItem Header="Export HTML">
        <Grid Margin="8">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock VerticalAlignment="Center" Text="Month" Margin="0,0,6,0"/>
            <ComboBox x:Name="cbMonth" Width="60" VerticalAlignment="Center"/>
            <TextBlock VerticalAlignment="Center" Text="Year" Margin="16,0,6,0"/>
            <TextBox x:Name="tbYear" Width="70" Height="26" VerticalAlignment="Center"/>
            <TextBlock VerticalAlignment="Center" Text="Status" Margin="16,0,6,0"/>
            <ComboBox x:Name="cbStatus" Width="120" VerticalAlignment="Center"/>
          </StackPanel>
          <Grid Grid.Row="1" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock VerticalAlignment="Center" Text="Org filter (optional)" Margin="0,0,8,0"/>
            <TextBox Grid.Column="1" x:Name="tbExportClientId" Height="26"/>
          </Grid>
          <Grid Grid.Row="2" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="tbOutputPath" Height="28" VerticalContentAlignment="Center"/>
            <Button Grid.Column="1" x:Name="btnBrowseOutput" Content="Browse..." Padding="12,4"/>
          </Grid>
          <Button Grid.Row="3" x:Name="btnGenerate" Content="Generate HTML + manifest" HorizontalAlignment="Left" Width="220"/>
          <TextBox Grid.Row="4" x:Name="tbExportLog" Margin="0,10,0,0"
                   IsReadOnly="True" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"
                   FontFamily="Consolas" FontSize="11"/>
        </Grid>
      </TabItem>
      <TabItem Header="Create tickets">
        <Grid Margin="8">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <Grid Grid.Row="0" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="tbManifestFolder" Height="28" VerticalContentAlignment="Center"/>
            <Button Grid.Column="1" x:Name="btnBrowseManifest" Content="Browse..."/>
            <Button Grid.Column="2" x:Name="btnLoadManifest" Content="Load manifest"/>
          </Grid>
          <Grid Grid.Row="1" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="80"/>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Ticket form ID" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBox Grid.Column="1" x:Name="tbTicketFormId" Height="26"/>
            <Button Grid.Column="2" x:Name="btnLoadForms" Content="Load forms" Margin="8,0,8,0"/>
            <ComboBox Grid.Column="3" x:Name="cbTicketForms" DisplayMemberPath="Display" SelectedValuePath="Id"/>
          </Grid>
          <Grid Grid.Row="2" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="80"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Subject template" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBox Grid.Column="1" x:Name="tbSubjectTemplate" Height="26"
                     Text="Billing invoice {InvoiceNumber} — {ClientName}"/>
            <TextBlock Grid.Column="2" Text="Status ID" Margin="12,0,8,0" VerticalAlignment="Center"/>
            <TextBox Grid.Column="3" x:Name="tbTicketStatus" Height="26" Text="1000"/>
          </Grid>
          <Grid Grid.Row="3" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Description" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <TextBox Grid.Column="1" x:Name="tbTicketDescription" Height="26"
                     Text="Billing invoice export; details are in the attached HTML."/>
          </Grid>
          <Grid Grid.Row="4" Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="140"/>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Type" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <ComboBox Grid.Column="1" x:Name="cbTicketType"/>
            <TextBlock Grid.Column="2" Text="Attachment comment" Margin="12,0,8,0" VerticalAlignment="Center"/>
            <TextBox Grid.Column="3" x:Name="tbAttachComment" Height="26" Text="Invoice HTML is attached."/>
          </Grid>
          <DataGrid Grid.Row="5" x:Name="dgManifest" AutoGenerateColumns="True"
                    CanUserAddRows="False" IsReadOnly="False" SelectionMode="Extended"
                    Margin="0,0,0,8"/>
          <StackPanel Grid.Row="6" Orientation="Horizontal">
            <Button x:Name="btnCreateTickets" Content="Create tickets for selected rows"/>
            <TextBlock x:Name="lblTicketHint" Margin="12,6,0,0" FontSize="11" Foreground="#666"
                       TextWrapping="Wrap"
                       Text="Select rows (Ctrl+click). Rows that already have a ticket ID are skipped."/>
          </StackPanel>
        </Grid>
      </TabItem>
    </TabControl>
    <TextBlock Grid.Row="3" x:Name="lblStatus" Text="" TextWrapping="Wrap" FontSize="12"/>
  </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Populate combos
$cbMonth = $window.FindName('cbMonth')
for ($m = 1; $m -le 12; $m++) { [void]$cbMonth.Items.Add($m) }
$cbMonth.SelectedItem = (Get-Date).AddMonths(-1).Month
$window.FindName('tbYear').Text = [string]((Get-Date).AddMonths(-1).Year)
foreach ($st in @('PENDING', 'APPROVED', 'COMPLETE', 'FAILED', 'ARCHIVED')) {
    [void]$window.FindName('cbStatus').Items.Add($st)
}
$window.FindName('cbStatus').SelectedItem = 'COMPLETE'
foreach ($tt in @('PROBLEM', 'QUESTION', 'INCIDENT', 'TASK', 'CHANGE_REQUEST', 'SERVICE_REQUEST', 'PROJECT', 'APPOINTMENT', 'MISCELLANEOUS')) {
    [void]$window.FindName('cbTicketType').Items.Add($tt)
}
$window.FindName('cbTicketType').SelectedItem = 'TASK'

$cfgLoad = Get-AppConfig
if ($cfgLoad.PsaBaseUrl) {
    try {
        $u = [uri]$cfgLoad.PsaBaseUrl
        $window.FindName('tbPsaInstance').Text = $u.Host
    } catch { $window.FindName('tbPsaInstance').Text = $cfgLoad.PsaBaseUrl }
}
if ($cfgLoad.PsaClientId) { $window.FindName('tbPsaClientId').Text = $cfgLoad.PsaClientId }
if ($cfgLoad.PsaBaseUrl) { $script:PsaBaseUrl = $cfgLoad.PsaBaseUrl }

$window.FindName('btnBrowseOutput').Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $window.FindName('tbOutputPath').Text = $fb.SelectedPath
    }
})

$window.FindName('btnBrowseManifest').Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $window.FindName('tbManifestFolder').Text = $fb.SelectedPath
    }
})

function Test-PsaSignedIn {
    return ((Test-PsaTokenValid) -or (Test-PsaRefreshPresent))
}

$window.FindName('btnGenerate').Add_Click({
    $log = $window.FindName('tbExportLog')
    $btn = $window.FindName('btnGenerate')
    if (-not (Test-PsaSignedIn)) {
        [System.Windows.MessageBox]::Show(
            'Sign in to NinjaOne first (Connection). HTML export uses the same session as ticketing.',
            'Export', 'OK', 'Warning') | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($script:PsaClientId)) {
        [System.Windows.MessageBox]::Show(
            'Client ID is required. Enter it in Connection, then sign in.',
            'Export', 'OK', 'Warning') | Out-Null
        return
    }
    $hostOnly = $null
    try {
        $hostOnly = ([uri]$script:PsaBaseUrl).Host
    } catch {
        [System.Windows.MessageBox]::Show('Invalid or missing instance. Sign in from Connection.', 'Export', 'OK', 'Warning') | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($hostOnly)) {
        [System.Windows.MessageBox]::Show('Invalid or missing instance. Sign in from Connection.', 'Export', 'OK', 'Warning') | Out-Null
        return
    }
    $outPath = $window.FindName('tbOutputPath').Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outPath)) {
        [System.Windows.MessageBox]::Show('Choose an output folder.', 'Export', 'OK', 'Warning') | Out-Null
        return
    }
    $outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outPath)
    $month = [int]$window.FindName('cbMonth').SelectedItem
    $year = [int]$window.FindName('tbYear').Text
    $status = [string]$window.FindName('cbStatus').SelectedItem
    $fcRaw = $window.FindName('tbExportClientId').Text.Trim()
    $filterClient = $null
    if (-not [string]::IsNullOrWhiteSpace($fcRaw)) {
        $fc = 0
        if (-not [int]::TryParse($fcRaw, [ref]$fc)) {
            [System.Windows.MessageBox]::Show('Org filter must be an integer client ID.', 'Export', 'OK', 'Warning') | Out-Null
            return
        }
        $filterClient = [Nullable[int]]$fc
    }
    $btn.IsEnabled = $false
    $log.Clear()
    $log.AppendText("Fetching invoice list (background)...`r`n")
    Push-UIUpdate
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        param($HostOnly, $ClientId, $RefreshToken, $Month, $Year, $Status, $FilterClientId)
        $ErrorActionPreference = 'Stop'
        function Get-RefTok {
            param($cid, $rt, $h)
            Invoke-RestMethod -Uri "https://$h/ws/oauth/token" -Method POST -Headers @{
                Accept = 'application/json'; 'Content-Type' = 'application/x-www-form-urlencoded'
            } -Body @{
                grant_type    = 'refresh_token'
                client_id     = $cid
                refresh_token = $rt
            } -TimeoutSec 30
        }
        function Apply-TokToSession {
            param($S, $t)
            if ([string]::IsNullOrWhiteSpace($t.access_token)) { throw 'Token response did not include access_token.' }
            $S.AuthHeader = "Bearer $($t.access_token)"
            if ($t.refresh_token) { $S.RefreshToken = [string]$t.refresh_token }
            $exp = if ($t.expires_in) { [int]$t.expires_in } else { 3600 }
            if ($exp -le 0) { $exp = 3600 }
            $S.ExpiresAt = [datetime]::UtcNow.AddSeconds($exp - 60)
        }
        function Invoke-Api {
            param($Method, $Endpoint, $Query, $Body, $Sess)
            function Ref {
                param($S)
                $t = Get-RefTok -cid $S.ClientId -rt $S.RefreshToken -h $S.HostOnly
                Apply-TokToSession -S $S -t $t
            }
            if (-not $Sess.ExpiresAt -or [datetime]::UtcNow -ge $Sess.ExpiresAt) { Ref $Sess }
            $base = "https://$($Sess.HostOnly)/$($Endpoint.TrimStart('/'))"
            $uri = if ($Query) { "$base`?$Query" } else { $base }
            $h = @{ Authorization = $Sess.AuthHeader; Accept = 'application/json' }
            $bj = $null
            if ($Body) { $bj = $Body | ConvertTo-Json -Depth 20; $h['Content-Type'] = 'application/json' }
            Invoke-RestMethod -Uri $uri -Method $Method -Headers $h -Body $bj -TimeoutSec 120
        }
        $tr = Get-RefTok -cid $ClientId -rt $RefreshToken -h $HostOnly
        $session = [PSCustomObject]@{
            HostOnly     = $HostOnly
            ClientId     = $ClientId
            RefreshToken = $RefreshToken
            AuthHeader   = ''
            ExpiresAt    = $null
        }
        Apply-TokToSession -S $session -t $tr
        $periodFrom = "$Year-$('{0:D2}' -f $Month)-01"
        $lastDay = [DateTime]::DaysInMonth($Year, $Month)
        $periodTo = "$Year-$('{0:D2}' -f $Month)-$lastDay"
        $query = "periodFrom=$periodFrom&periodTo=$periodTo"
        if ($null -ne $FilterClientId) { $query += "&clientId=$FilterClientId" }
        $rawList = Invoke-Api -Method GET -Endpoint '/v2/billing/invoices' -Query $query -Sess $session
        $invoiceList = if ($rawList -is [Array]) { @($rawList) }
        elseif ($rawList.PSObject.Properties['results']) { @($rawList.results) }
        elseif ($rawList.PSObject.Properties['data']) { @($rawList.data) }
        else { @($rawList) }
        $invoiceList = @($invoiceList | Where-Object { $_.status -eq $Status })
        return [PSCustomObject]@{ Items = $invoiceList; BillingSession = $session }
    }).AddArgument($hostOnly).AddArgument($script:PsaClientId).AddArgument($script:PsaRefreshToken).AddArgument($month).AddArgument($year).AddArgument($status).AddArgument($filterClient)
    $script:ExportJobPs = $ps
    $script:ExportJobRs = $rs
    $script:ExportJobHandle = $ps.BeginInvoke()
    $script:ExportJobOutPath = $outPath
    $script:ExportJobBillingSession = [PSCustomObject]@{
        HostOnly     = $hostOnly
        ClientId     = $script:PsaClientId
        RefreshToken = $script:PsaRefreshToken
        AuthHeader   = if (Test-PsaTokenValid) { "Bearer $($script:PsaAccessToken)" } else { '' }
        ExpiresAt    = if (Test-PsaTokenValid) { $script:PsaTokenExpiresAt } else { $null }
    }
    $script:ExportJobLog = $log
    $script:ExportJobBtn = $btn
    if (-not $script:ExportJobTimer) {
        $script:ExportJobTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:ExportJobTimer.Interval = [TimeSpan]::FromMilliseconds(300)
        $script:ExportJobTimer.Add_Tick({
            if (-not $script:ExportJobHandle.IsCompleted) { return }
            $script:ExportJobTimer.Stop()
            $invList = @()
            try {
                $col = $script:ExportJobPs.EndInvoke($script:ExportJobHandle)
                if ($col -and $col.Count -gt 0) {
                    $wrap = $col[0]
                    if ($wrap.Items) { $invList = @($wrap.Items) }
                    if ($wrap.BillingSession) {
                        $bs = $wrap.BillingSession
                        $script:ExportJobBillingSession = $bs
                        if ($bs.RefreshToken) { $script:PsaRefreshToken = $bs.RefreshToken }
                        if ($bs.AuthHeader) {
                            $script:PsaAccessToken = $bs.AuthHeader -replace '^Bearer ', ''
                        }
                        if ($bs.ExpiresAt) { $script:PsaTokenExpiresAt = $bs.ExpiresAt }
                        if ($script:MasterPassword) { try { Save-CurrentSessionToDisk } catch { } }
                    }
                }
            } catch {
                $script:ExportJobLog.AppendText("ERROR: $($_.Exception.Message)`r`n")
                $script:ExportJobBtn.IsEnabled = $true
                try { $script:ExportJobPs.Dispose() } catch { }
                try { $script:ExportJobRs.Close() } catch { }
                Set-StatusBar 'Export failed.'
                return
            } finally {
                try { $script:ExportJobPs.Dispose() } catch { }
                try { $script:ExportJobRs.Close() } catch { }
            }
            $script:ExportJobLog.AppendText("Found $($invList.Count) invoice(s). Generating HTML...`r`n")
            Push-UIUpdate
            if ($invList.Count -eq 0) {
                $script:ExportJobBtn.IsEnabled = $true
                Set-StatusBar 'No invoices to export.'
                return
            }
            if (-not (Test-Path $script:ExportJobOutPath)) {
                New-Item -ItemType Directory -Path $script:ExportJobOutPath -Force | Out-Null
            }
            $manifestPath = Join-Path $script:ExportJobOutPath $script:ManifestFileName
            $ticketByInvoiceId = @{}
            if (Test-Path -LiteralPath $manifestPath) {
                try {
                    $old = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
                    $oldRows = @()
                    if ($old -is [System.Array]) { $oldRows = @($old) } elseif ($null -ne $old) { $oldRows = @($old) }
                    foreach ($r in $oldRows) {
                        if ($null -ne $r.invoiceId -and $null -ne $r.ticketId) {
                            $ticketByInvoiceId[[int]$r.invoiceId] = [int]$r.ticketId
                        }
                    }
                } catch { }
            }
            $manifestRows = [System.Collections.Generic.List[hashtable]]::new()
            $sess = $script:ExportJobBillingSession
            $succeeded = 0
            $failed = 0
            foreach ($inv in $invList) {
                $invNum = if ($inv.invoiceNumber) { $inv.invoiceNumber } else { "INV-$($inv.id)" }
                $displayClient = if ($inv.client -and $inv.client.name) { $inv.client.name } else { 'Unknown' }
                $clientName = $displayClient -replace '[\\/:*?"<>|]', '_'
                $safeNum = $invNum -replace '[\\/:*?"<>|]', '_'
                $htmlName = "${safeNum}_${clientName}.html"
                $htmlPath = Join-Path $script:ExportJobOutPath $htmlName
                try {
                    $detail = Invoke-NinjaBillingApi -Endpoint "/v2/billing/invoices/$($inv.id)" -Session $sess
                    $htmlContent = Build-InvoiceHtml -Invoice $detail
                    [System.IO.File]::WriteAllText($htmlPath, $htmlContent, [System.Text.Encoding]::UTF8)
                    $orgId = if ($detail.client -and $null -ne $detail.client.id) { [int]$detail.client.id } else { $null }
                    if ($null -ne $orgId) {
                        $tid = $null
                        if ($ticketByInvoiceId.ContainsKey([int]$inv.id)) {
                            $tid = $ticketByInvoiceId[[int]$inv.id]
                        }
                        $row = @{
                            invoiceId            = [int]$inv.id
                            clientOrganizationId = $orgId
                            invoiceNumber        = $invNum
                            clientName           = $displayClient
                            htmlFileName         = $htmlName
                        }
                        if ($null -ne $tid) { $row['ticketId'] = $tid }
                        [void]$manifestRows.Add($row)
                    } else {
                        $script:ExportJobLog.AppendText("WARN [$invNum] No client id — HTML saved, not in manifest.`r`n")
                    }
                    $script:ExportJobLog.AppendText("  OK $htmlName`r`n")
                    $succeeded++
                } catch {
                    $script:ExportJobLog.AppendText("  FAIL [$invNum] $($_.Exception.Message)`r`n")
                    $failed++
                }
                Push-UIUpdate
            }
            if ($succeeded -gt 0) {
                $rows = @($manifestRows | ForEach-Object { [PSCustomObject]$_ })
                Write-ManifestJson -Rows $rows -ManifestPath $manifestPath
                $script:ExportJobLog.AppendText("`r`nManifest: $manifestPath`r`n")
                $script:LastExportOutputPath = $script:ExportJobOutPath
                $window.FindName('tbManifestFolder').Text = $script:ExportJobOutPath
            }
            $script:ExportJobLog.AppendText("Done. Succeeded: $succeeded  Failed: $failed`r`n")
            $script:ExportJobBtn.IsEnabled = $true
            Set-StatusBar "Export finished. Succeeded: $succeeded"
        })
    }
    $script:ExportJobTimer.Start()
})

$window.FindName('btnPsaSignIn').Add_Click({
    $btn = $window.FindName('btnPsaSignIn')
    $lbl = $window.FindName('lblPsaAuthStatus')
    try {
        $script:PsaBaseUrl = Resolve-BaseUrl -Instance $window.FindName('tbPsaInstance').Text
    } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Invalid instance', 'OK', 'Warning') | Out-Null
        return
    }
    $script:PsaClientId = $window.FindName('tbPsaClientId').Text.Trim()
    if ([string]::IsNullOrWhiteSpace($script:PsaClientId)) {
        [System.Windows.MessageBox]::Show('PSA Client ID is required.', 'Sign in', 'OK', 'Warning') | Out-Null
        return
    }
    $cfg = Get-AppConfig
    $hasSaved = -not [string]::IsNullOrWhiteSpace($cfg.EncryptedPsaRefreshToken) -and
        -not [string]::IsNullOrWhiteSpace($cfg.MasterPasswordVerifier)
    if ($hasSaved) {
        $btn.IsEnabled = $false
        $lbl.Text = 'Unlocking...'
        $lbl.Foreground = [System.Windows.Media.Brushes]::DarkOrange
        Push-UIUpdate
        $mp = Show-MasterPasswordPrompt -Title 'Unlock' -Message 'Enter master password to use saved PSA session:'
        if (-not $mp) {
            $lbl.Text = 'Not connected'
            $lbl.Foreground = [System.Windows.Media.Brushes]::Gray
            $btn.IsEnabled = $true
            return
        }
        if (-not (Test-MasterPasswordValid -MasterPwd $mp -Verifier $cfg.MasterPasswordVerifier)) {
            [System.Windows.MessageBox]::Show('Incorrect master password.', 'Unlock', 'OK', 'Warning') | Out-Null
            $lbl.Text = 'Not connected'
            $lbl.Foreground = [System.Windows.Media.Brushes]::Gray
            $btn.IsEnabled = $true
            return
        }
        $script:MasterPassword = $mp
        $script:MasterPasswordVerifier = $cfg.MasterPasswordVerifier
        try {
            $script:PsaRefreshToken = Unprotect-String -CipherText $cfg.EncryptedPsaRefreshToken -MasterPwd $mp
            if ($cfg.PsaBaseUrl) { $script:PsaBaseUrl = $cfg.PsaBaseUrl }
            Invoke-PsaTokenRefresh
            $lbl.Text = 'Connected'
            $lbl.Foreground = [System.Windows.Media.Brushes]::Green
            $window.FindName('btnChangeMasterPwd').Visibility = 'Visible'
            $window.FindName('btnClearSession').Visibility = 'Visible'
            Set-StatusBar 'PSA session restored from saved config.'
            $btn.IsEnabled = $true
            return
        } catch {
            $lbl.Text = 'Session expired'
            $lbl.Foreground = [System.Windows.Media.Brushes]::DarkOrange
            Set-StatusBar "Saved PSA session failed: $($_.Exception.Message)"
        }
        $btn.IsEnabled = $true
    }
    $btn.IsEnabled = $false
    $lbl.Text = 'Waiting for browser...'
    $lbl.Foreground = [System.Windows.Media.Brushes]::DarkOrange
    Set-StatusBar 'Complete sign-in in the browser.'
    Push-UIUpdate
    $script:AuthVerifier = New-PkceVerifier
    $script:AuthState = New-OAuthState
    $challenge = Get-PkceChallenge -Verifier $script:AuthVerifier
    $script:AuthTimeoutAt = [datetime]::UtcNow.AddMinutes(3)
    $script:AuthRedirectUri = 'http://localhost:8888/'
    $script:AuthListener = [System.Net.HttpListener]::new()
    $script:AuthListener.Prefixes.Add($script:AuthRedirectUri)
    try {
        $script:AuthListener.Start()
    } catch {
        $lbl.Text = 'Listener failed'
        $lbl.Foreground = [System.Windows.Media.Brushes]::Red
        Set-StatusBar $_.Exception.Message
        $btn.IsEnabled = $true
        return
    }
    $scopes = 'monitoring management offline_access'
    $scopeEncoded = [uri]::EscapeDataString($scopes)
    $authorizeUrl = "$($script:PsaBaseUrl)/ws/oauth/authorize?" +
        "response_type=code" +
        "&client_id=$([uri]::EscapeDataString($script:PsaClientId))" +
        "&redirect_uri=$([uri]::EscapeDataString($script:AuthRedirectUri))" +
        "&scope=$scopeEncoded" +
        "&state=$([uri]::EscapeDataString($script:AuthState))" +
        "&code_challenge=$challenge" +
        "&code_challenge_method=S256"
    Start-Process $authorizeUrl
    $lst = $script:AuthListener
    $script:AuthPS = [powershell]::Create()
    [void]$script:AuthPS.AddScript({
        param($listener)
        try {
            $ctx = $listener.GetContext()
            $q = $ctx.Request.Url.Query
            $html = '<html><body style="font-family:system-ui,sans-serif;text-align:center;padding-top:60px">' +
                '<h2>Authentication complete</h2><p>You may close this tab.</p></body></html>'
            $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
            $ctx.Response.ContentType = 'text/html'
            $ctx.Response.ContentLength64 = $buf.Length
            $ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
            $ctx.Response.Close()
            $listener.Stop()
            return $q
        } catch {
            try { $listener.Stop() } catch { }
            return "error=$($_.Exception.Message)"
        }
    }).AddArgument($lst)
    $script:AuthHandle = $script:AuthPS.BeginInvoke()
    $authTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $authTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $authTimer.Add_Tick({
        if (-not $script:AuthHandle.IsCompleted) {
            if ($script:AuthTimeoutAt -ne [datetime]::MinValue -and [datetime]::UtcNow -lt $script:AuthTimeoutAt) { return }
            try { $script:AuthListener.Stop() } catch { }
            try { $script:AuthListener.Close() } catch { }
            try { if ($script:AuthPS) { $script:AuthPS.Stop() } } catch { }
            try { if ($script:AuthPS) { $script:AuthPS.Dispose() } } catch { }
            $script:AuthPS = $null
            $script:AuthHandle = $null
            $lbl.Text = 'Timed out'
            $lbl.Foreground = [System.Windows.Media.Brushes]::Red
            Set-StatusBar 'OAuth timed out.'
            $btn.IsEnabled = $true
            $this.Stop()
            return
        }
        $this.Stop()
        $queryString = $null
        try {
            $queryString = ($script:AuthPS.EndInvoke($script:AuthHandle) | Select-Object -First 1) -as [string]
        } finally {
            try { $script:AuthPS.Dispose() } catch { }
            $script:AuthPS = $null
        }
        $returnedState = $null
        if ($queryString -match '[?&]state=([^&]+)') {
            $returnedState = [uri]::UnescapeDataString($Matches[1])
        }
        try {
            if ($queryString -match '[?&]code=([^&]+)') {
                if ([string]::IsNullOrWhiteSpace($script:AuthState) -or
                    [string]::IsNullOrWhiteSpace($returnedState) -or
                    $returnedState -ne $script:AuthState) {
                    throw 'OAuth state validation failed.'
                }
                $code = [uri]::UnescapeDataString($Matches[1])
                $resp = Invoke-RestMethod -Uri "$($script:PsaBaseUrl)/ws/oauth/token" -Method POST -UseBasicParsing `
                    -Headers @{
                        'Accept' = 'application/json'; 'Content-Type' = 'application/x-www-form-urlencoded'
                    } -Body @{
                        grant_type    = 'authorization_code'
                        code          = $code
                        redirect_uri  = $script:AuthRedirectUri
                        client_id     = $script:PsaClientId
                        code_verifier = $script:AuthVerifier
                    }
                if (-not $resp.refresh_token) {
                    throw 'No refresh_token in response. Ensure offline_access scope and refresh token grant on the OAuth app.'
                }
                $script:PsaRefreshToken = [string]$resp.refresh_token
                Update-PsaTokensFromResponse $resp
                $lbl.Text = 'Connected'
                $lbl.Foreground = [System.Windows.Media.Brushes]::Green
                $window.FindName('btnChangeMasterPwd').Visibility = 'Visible'
                $window.FindName('btnClearSession').Visibility = 'Visible'
                if (-not $script:MasterPassword) {
                    $newPwd = Show-MasterPasswordPrompt -Title 'Save session' `
                        -Message 'Set a master password to encrypt your PSA refresh token on disk.' -IsNewPassword
                    if ($newPwd) {
                        $script:MasterPassword = $newPwd
                        $script:MasterPasswordVerifier = New-MasterPasswordVerifier -MasterPwd $newPwd
                        Save-CurrentSessionToDisk
                        Set-StatusBar 'Signed in; session saved.'
                    } else {
                        Set-StatusBar 'Signed in; session not saved to disk.'
                    }
                } else {
                    Save-CurrentSessionToDisk
                    Set-StatusBar 'Signed in; session saved.'
                }
            } elseif ($queryString -match '[?&]error=([^&]+)') {
                $errMsg = [uri]::UnescapeDataString($Matches[1])
                $lbl.Text = 'Failed'
                $lbl.Foreground = [System.Windows.Media.Brushes]::Red
                Set-StatusBar "OAuth error: $errMsg"
            } else {
                $lbl.Text = 'Failed'
                $lbl.Foreground = [System.Windows.Media.Brushes]::Red
                Set-StatusBar 'No authorization code received.'
            }
        } catch {
            $lbl.Text = 'Error'
            $lbl.Foreground = [System.Windows.Media.Brushes]::Red
            Set-StatusBar $_.Exception.Message
        }
        $script:AuthState = $null
        $btn.IsEnabled = $true
    })
    $authTimer.Start()
})

$window.FindName('btnChangeMasterPwd').Add_Click({
    $np = Show-ChangeMasterPasswordPrompt
    if (-not $np) { return }
    $script:MasterPassword = $np
    $script:MasterPasswordVerifier = New-MasterPasswordVerifier -MasterPwd $np
    Save-CurrentSessionToDisk
    Set-StatusBar 'Master password updated; credentials re-encrypted.'
})

$window.FindName('btnClearSession').Add_Click({
    if ([System.Windows.MessageBox]::Show('Delete saved encrypted session on disk?', 'Clear', 'YesNo', 'Question') -ne 'Yes') { return }
    Clear-SavedSession
    $window.FindName('btnChangeMasterPwd').Visibility = 'Collapsed'
    $window.FindName('btnClearSession').Visibility = 'Collapsed'
    $window.FindName('lblPsaAuthStatus').Text = 'Not connected'
    $window.FindName('lblPsaAuthStatus').Foreground = [System.Windows.Media.Brushes]::Gray
    Set-StatusBar 'Saved session cleared.'
})

function Get-PsaApiHost {
    return ([uri]$script:PsaBaseUrl).Host
}

$window.FindName('btnLoadManifest').Add_Click({
    $dir = $window.FindName('tbManifestFolder').Text.Trim()
    if ([string]::IsNullOrWhiteSpace($dir)) {
        [System.Windows.MessageBox]::Show('Choose a folder that contains the manifest JSON.', 'Manifest', 'OK', 'Warning') | Out-Null
        return
    }
    $dir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($dir)
    $mp = Join-Path $dir $script:ManifestFileName
    if (-not (Test-Path -LiteralPath $mp)) {
        [System.Windows.MessageBox]::Show("Manifest not found:`n$mp", 'Manifest', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        $raw = Get-Content -LiteralPath $mp -Raw | ConvertFrom-Json
        $rows = @()
        if ($raw -is [System.Array]) { $rows = @($raw) } elseif ($null -ne $raw) { $rows = @($raw) }
        $dt = New-Object System.Data.DataTable
        [void]$dt.Columns.Add('invoiceId', [int])
        [void]$dt.Columns.Add('clientOrganizationId', [int])
        [void]$dt.Columns.Add('invoiceNumber', [string])
        [void]$dt.Columns.Add('clientName', [string])
        [void]$dt.Columns.Add('htmlFileName', [string])
        [void]$dt.Columns.Add('ticketId', [string])
        foreach ($r in $rows) {
            $tid = if ($null -ne $r.ticketId) { [string]$r.ticketId } else { '' }
            [void]$dt.Rows.Add(
                [int]$r.invoiceId,
                [int]$r.clientOrganizationId,
                [string]$r.invoiceNumber,
                [string]$r.clientName,
                [string]$r.htmlFileName,
                $tid
            )
        }
        $script:ManifestRowsTable = $dt
        $script:ManifestPathOnDisk = $mp
        $window.FindName('dgManifest').ItemsSource = $dt.DefaultView
        Set-StatusBar "Loaded $($dt.Rows.Count) manifest row(s)."
    } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Manifest', 'OK', 'Error') | Out-Null
    }
})

$window.FindName('btnLoadForms').Add_Click({
    if (-not (Test-PsaSignedIn)) {
        [System.Windows.MessageBox]::Show('Sign in to PSA first.', 'Forms', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        $hostOnly = Get-PsaApiHost
        $sess = [PSCustomObject]@{
            HostOnly     = $hostOnly
            ClientId     = $script:PsaClientId
            RefreshToken = $script:PsaRefreshToken
            AuthHeader   = ''
            ExpiresAt    = $null
        }
        Update-PsaSessionFromScriptToken -Session $sess
        $forms = Invoke-NinjaPsaApi -Method GET -Endpoint '/v2/ticketing/ticket-form' -Session $sess
        $list = @()
        $arr = @()
        if ($forms -is [Array]) { $arr = @($forms) }
        elseif ($forms -and $forms.PSObject.Properties['results']) { $arr = @($forms.results) }
        elseif ($forms -and $forms.PSObject.Properties['ticketForms']) { $arr = @($forms.ticketForms) }
        elseif ($forms -and $forms.PSObject.Properties['data']) { $arr = @($forms.data) }
        elseif ($null -ne $forms) { $arr = @($forms) }
        foreach ($f in $arr) {
            if ($null -eq $f.id) { continue }
            $nm = if ($f.name) { [string]$f.name } else { "Form $($f.id)" }
            $list += [PSCustomObject]@{ Id = [int]$f.id; Display = "$($f.id) — $nm" }
        }
        $cb = $window.FindName('cbTicketForms')
        $cb.ItemsSource = $list
        if ($list.Count -gt 0) { $cb.SelectedIndex = 0 }
        Set-StatusBar "Loaded $($list.Count) ticket form(s)."
    } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Forms', 'OK', 'Error') | Out-Null
    }
})

$window.FindName('cbTicketForms').Add_SelectionChanged({
    $cb = $window.FindName('cbTicketForms')
    if ($null -eq $cb.SelectedItem) { return }
    $id = $cb.SelectedItem.Id
    if ($null -ne $id) { $window.FindName('tbTicketFormId').Text = [string]$id }
})

$window.FindName('btnCreateTickets').Add_Click({
    if (-not (Test-PsaSignedIn)) {
        [System.Windows.MessageBox]::Show('Sign in to PSA first.', 'Tickets', 'OK', 'Warning') | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($script:ManifestPathOnDisk)) {
        [System.Windows.MessageBox]::Show('Load a manifest first.', 'Tickets', 'OK', 'Warning') | Out-Null
        return
    }
    $dt = $script:ManifestRowsTable
    if (-not $dt -or $dt.Rows.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Load a manifest first.', 'Tickets', 'OK', 'Warning') | Out-Null
        return
    }
    $tf = 0
    if (-not [int]::TryParse($window.FindName('tbTicketFormId').Text.Trim(), [ref]$tf) -or $tf -le 0) {
        [System.Windows.MessageBox]::Show('Enter a valid Ticket form ID.', 'Tickets', 'OK', 'Warning') | Out-Null
        return
    }
    $dg = $window.FindName('dgManifest')
    $selected = @($dg.SelectedItems)
    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show('Select one or more rows in the grid.', 'Tickets', 'OK', 'Warning') | Out-Null
        return
    }
    $manifestDir = [System.IO.Path]::GetDirectoryName($script:ManifestPathOnDisk)
    $hostOnly = Get-PsaApiHost
    $sess = [PSCustomObject]@{
        HostOnly     = $hostOnly
        ClientId     = $script:PsaClientId
        RefreshToken = $script:PsaRefreshToken
        AuthHeader   = ''
        ExpiresAt    = $null
    }
    Update-PsaSessionFromScriptToken -Session $sess
    $subjT = $window.FindName('tbSubjectTemplate').Text
    $descB = $window.FindName('tbTicketDescription').Text
    $tstat = $window.FindName('tbTicketStatus').Text.Trim()
    $tType = [string]$window.FindName('cbTicketType').SelectedItem
    $attachC = $window.FindName('tbAttachComment').Text
    $btn = $window.FindName('btnCreateTickets')
    $btn.IsEnabled = $false
    try {
        foreach ($rowView in $selected) {
            $row = $rowView.Row
            $tidExisting = [string]$row['ticketId']
            if (-not [string]::IsNullOrWhiteSpace($tidExisting)) { continue }
            $invNum = [string]$row['invoiceNumber']
            $clientName = [string]$row['clientName']
            $htmlName = [string]$row['htmlFileName']
            $orgId = [int]$row['clientOrganizationId']
            $invId = [int]$row['invoiceId']
            $htmlPath = Join-Path $manifestDir $htmlName
            if (-not (Test-Path -LiteralPath $htmlPath)) {
                [System.Windows.MessageBox]::Show("HTML not found: $htmlPath", 'Tickets', 'OK', 'Warning') | Out-Null
                continue
            }
            $subject = Expand-TicketSubject -Template $subjT -InvoiceNumber $invNum -ClientName $clientName
            $newTicketBody = @{
                clientId     = $orgId
                ticketFormId = $tf
                status       = $tstat
                subject      = $subject
                type         = $tType
                description  = @{ public = $true; body = $descB }
            }
            $created = Invoke-NinjaPsaApi -Method POST -Endpoint '/v2/ticketing/ticket' -Body $newTicketBody -Session $sess
            if ($null -eq $created.id) { throw 'Create ticket response missing id.' }
            $newTid = [int]$created.id
            Send-NinjaTicketCommentWithAttachment -Session $sess -TicketId $newTid -CommentBody $attachC -HtmlFilePath $htmlPath
            $ticketUrl = Get-PsaTicketWebUrl -BaseUrl $script:PsaBaseUrl -TicketId $newTid
            Update-InvoiceHtmlWithTicket -HtmlFilePath $htmlPath -TicketUrl $ticketUrl -TicketId $newTid
            $row['ticketId'] = [string]$newTid
        }
        $objects = foreach ($r in $dt.Rows) {
            $o = [ordered]@{
                invoiceId            = [int]$r['invoiceId']
                clientOrganizationId = [int]$r['clientOrganizationId']
                invoiceNumber        = [string]$r['invoiceNumber']
                clientName           = [string]$r['clientName']
                htmlFileName         = [string]$r['htmlFileName']
            }
            $tv = $r['ticketId']
            if ($tv -isnot [System.DBNull] -and -not [string]::IsNullOrWhiteSpace([string]$tv)) {
                $o['ticketId'] = [int][string]$tv
            }
            [PSCustomObject]$o
        }
        Write-ManifestJson -Rows $objects -ManifestPath $script:ManifestPathOnDisk
        $dg.Items.Refresh()
        Set-StatusBar 'Ticket creation finished; manifest and HTML updated.'
    } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Tickets', 'OK', 'Error') | Out-Null
        Set-StatusBar $_.Exception.Message
    } finally {
        $btn.IsEnabled = $true
    }
})

if ((Get-AppConfig).EncryptedPsaRefreshToken -and (Get-AppConfig).MasterPasswordVerifier) {
    $window.FindName('btnChangeMasterPwd').Visibility = 'Visible'
    $window.FindName('btnClearSession').Visibility = 'Visible'
}

[void]$window.ShowDialog()

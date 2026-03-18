#Requires -Version 5.1
<#
.SYNOPSIS
    NinjaOne ITAM Manager - unified WPF tool for equipment import, QR code
    generation, QR upload, and device-to-user assignment.

.DESCRIPTION
    Standalone PowerShell WPF application (no dot-sourcing) combining four
    ITAM workflows:

        Tab 1 - Import Equipment:   Create unmanaged or staged devices from CSV or manual entry.
        Tab 2 - Generate QR Codes:  Create device dashboard QR code images.
        Tab 3 - Upload QR Codes:    Attach QR PNGs to devices as related items.
        Tab 4 - Scan & Assign:      Scan user/device QR codes to assign owners.

    Authentication uses OAuth Authorization Code + PKCE (browser sign-in,
    no client secret required). Session state flows between tabs: imported
    device IDs pre-populate QR generation; generated QR output directory
    pre-fills QR upload.

.PARAMETER NinjaOneInstance
    NinjaOne instance hostname or base URL.
    Default: env:NINJA_BASE_URL or ca.ninjarmm.com.
    For branded or partner portals, use the branded host here (for example,
    rcs-sales.rmmservice.ca or https://rcs-sales.rmmservice.ca) so that the
    entire OAuth flow (authorize, consent, and redirect back to localhost)
    stays on the same host. If the browser is redirected to a regional host
    such as https://ca.ninjarmm.com/ws/oauth/consent and shows
    "Missing or empty sessionKey.", the redirect behavior is coming from the
    NinjaOne web app; this script always uses the instance you provide for
    /ws/oauth/authorize and /ws/oauth/token.

.PARAMETER ClientId
    OAuth application Client ID (Native / Authorization Code type).
    Default: env:NinjaOneClientId. If neither the parameter nor the env var is set,
    the user is prompted in the UI before sign-in.

.PARAMETER AllowInsecureHttp
    Allow http:// base URLs for development/testing only. By default, HTTPS is required.

.EXAMPLE
    .\Invoke-NinjaITAMManager.ps1

.EXAMPLE
    .\Invoke-NinjaITAMManager.ps1 -NinjaOneInstance app.ninjarmm.com -ClientId "your-client-id"
#>

[CmdletBinding()]
param(
    [string] $NinjaOneInstance = $(if ($env:NINJA_BASE_URL) { $env:NINJA_BASE_URL } else { 'ca.ninjarmm.com' }),
    [string] $ClientId = $env:NinjaOneClientId,
    [switch] $AllowInsecureHttp
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

#region Script-Scope State
$script:AccessToken      = $null
$script:RefreshToken     = $null
$script:TokenExpiresAt   = [datetime]::MinValue
$script:NinjaBaseUrl     = ''
$script:NinjaClientId    = $ClientId

$script:AuthPS           = $null
$script:AuthHandle       = $null
$script:AuthVerifier     = $null
$script:AuthState        = $null
$script:AuthRedirectUri  = $null
$script:AuthListener     = $null
$script:AuthTimeoutAt    = [datetime]::MinValue

$script:OrgCache         = $null
$script:LocationCache    = $null
$script:RoleCache        = $null
$script:StagedRoleCache   = $null
$script:CsvData          = $null

$script:ImportedDevices  = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:QROutputDirectory = ''
$script:GeneratedQRFiles = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:UploadFileMap    = [System.Collections.Generic.List[PSCustomObject]]::new()

$script:ScanUserInfo     = $null
$script:ScanDevices      = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:AllowInsecureHttp = $AllowInsecureHttp.IsPresent

$script:MasterPassword   = $null
$script:MasterPasswordVerifier = $null
$script:ITAMConfigDir    = Join-Path $env:APPDATA 'NinjaITAMManager'
$script:ITAMConfigFile   = Join-Path $script:ITAMConfigDir 'config.json'
#endregion

#region Cryptography (AES-256-CBC + PBKDF2)
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

#region Config File Management
function Get-ITAMConfig {
    $defaults = [PSCustomObject]@{
        NinjaInstance          = ''
        ClientId               = ''
        EncryptedRefreshToken  = ''
        MasterPasswordVerifier = ''
    }
    if (Test-Path $script:ITAMConfigFile) {
        try {
            $raw = Get-Content $script:ITAMConfigFile -Raw | ConvertFrom-Json
            foreach ($prop in $raw.PSObject.Properties) {
                if ($defaults.PSObject.Properties[$prop.Name]) {
                    $defaults.$($prop.Name) = $prop.Value
                }
            }
        } catch {
            Write-Verbose "Failed to load ITAM config: $($_.Exception.Message)"
        }
    }
    return $defaults
}

function Save-ITAMConfig {
    param(
        [string]$Instance,
        [string]$ClientIdValue,
        [string]$EncryptedRefreshToken,
        [string]$Verifier
    )
    if (-not (Test-Path $script:ITAMConfigDir)) {
        New-Item -ItemType Directory -Path $script:ITAMConfigDir -Force | Out-Null
    }
    $disk = [ordered]@{
        NinjaInstance          = $Instance
        ClientId               = $ClientIdValue
        EncryptedRefreshToken  = $EncryptedRefreshToken
        MasterPasswordVerifier = $Verifier
    }
    [PSCustomObject]$disk | ConvertTo-Json -Depth 5 |
        Set-Content $script:ITAMConfigFile -Encoding UTF8
}

function Save-CurrentSession {
    if (-not $script:MasterPassword) { return }
    if (-not (Test-RefreshTokenPresent)) { return }
    $plainRefresh = ConvertFrom-SecureToken $script:RefreshToken
    $encrypted = Protect-String -PlainText $plainRefresh -MasterPwd $script:MasterPassword
    $verifier  = if ($script:MasterPasswordVerifier) {
        $script:MasterPasswordVerifier
    } else {
        $v = New-MasterPasswordVerifier -MasterPwd $script:MasterPassword
        $script:MasterPasswordVerifier = $v
        $v
    }
    Save-ITAMConfig -Instance $script:NinjaBaseUrl `
                    -ClientIdValue $script:NinjaClientId `
                    -EncryptedRefreshToken $encrypted `
                    -Verifier $verifier
}

function Clear-SavedSession {
    if (Test-Path $script:ITAMConfigFile) {
        Remove-Item $script:ITAMConfigFile -Force -ErrorAction SilentlyContinue
    }
    $script:MasterPassword = $null
    $script:MasterPasswordVerifier = $null
}
#endregion

#region Master Password Dialogs
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
    '<TextBlock FontSize="11" Foreground="#888" Margin="0,2,0,0" Text="Minimum 8 characters. This password encrypts your saved session."/>'
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

    $ok = $dlg.ShowDialog()
    if ($ok) { return $dlg.Tag } else { return $null }
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
    <TextBlock FontSize="11" Foreground="#888" Margin="0,2,0,0"
               Text="Minimum 8 characters."/>
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

    $ok = $dlg.ShowDialog()
    if ($ok) { return $dlg.Tag } else { return $null }
}
#endregion

#region Helper Functions
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
        throw "Insecure HTTP is not allowed. Use an HTTPS NinjaOne URL, or run with -AllowInsecureHttp for local testing only."
    }
    if ($uri.Scheme -ne 'https' -and $uri.Scheme -ne 'http') {
        throw "Unsupported URL scheme '$($uri.Scheme)'. Use https:// (or http:// with -AllowInsecureHttp)."
    }

    return $uri.AbsoluteUri.TrimEnd('/')
}

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

function New-OAuthState {
    return New-PkceVerifier
}

function ConvertTo-SecureToken {
    param([string]$PlainToken)
    return ($PlainToken | ConvertTo-SecureString -AsPlainText -Force)
}

function ConvertFrom-SecureToken {
    param([securestring]$SecureToken)
    return [System.Net.NetworkCredential]::new('', $SecureToken).Password
}

function Test-RefreshTokenPresent {
    return ($null -ne $script:RefreshToken -and $script:RefreshToken.Length -gt 0)
}

function Test-TokenValid {
    return ($null -ne $script:AccessToken -and [datetime]::UtcNow -lt $script:TokenExpiresAt)
}

function Update-TokensFromResponse {
    param($Response)
    $accessProp = $Response.PSObject.Properties['access_token']
    if ($null -eq $accessProp -or [string]::IsNullOrWhiteSpace($accessProp.Value)) {
        throw "Token response did not include an access_token."
    }
    $script:AccessToken = ConvertTo-SecureToken $accessProp.Value

    $refreshRotated = $false
    $refreshProp = $Response.PSObject.Properties['refresh_token']
    if ($null -ne $refreshProp -and -not [string]::IsNullOrWhiteSpace($refreshProp.Value)) {
        $script:RefreshToken = ConvertTo-SecureToken $refreshProp.Value
        $refreshRotated = $true
    }
    $exp = if ($Response.expires_in) { [int]$Response.expires_in } else { 3600 }
    if ($exp -le 0) { $exp = 3600 }
    $script:TokenExpiresAt = [datetime]::UtcNow.AddSeconds($exp - 60)

    if ($refreshRotated -and $script:MasterPassword) {
        try { Save-CurrentSession } catch { Write-Verbose "Auto-save after token rotation failed: $($_.Exception.Message)" }
    }
}

function Invoke-TokenRefresh {
    if (-not (Test-RefreshTokenPresent)) {
        throw "No refresh token. Please sign in again."
    }
    $resp = Invoke-RestMethod -Uri "$($script:NinjaBaseUrl)/ws/oauth/token" `
        -Method POST -UseBasicParsing `
        -Headers @{
            'Accept'       = 'application/json'
            'Content-Type' = 'application/x-www-form-urlencoded'
        } `
        -Body @{
            grant_type    = 'refresh_token'
            refresh_token = (ConvertFrom-SecureToken $script:RefreshToken)
            client_id     = $script:NinjaClientId
        }
    Update-TokensFromResponse $resp
}

function Invoke-NinjaApi {
    param(
        [string]$Method = 'GET',
        [string]$Endpoint,
        [object]$Body
    )
    if (-not (Test-TokenValid)) { Invoke-TokenRefresh }
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $uri = "$($script:NinjaBaseUrl)/api/v2/$($Endpoint.TrimStart('/'))"
        $p = @{
            Uri             = $uri
            Method          = $Method
            UseBasicParsing = $true
            ErrorAction     = 'Stop'
            Headers         = @{
                'Authorization' = "Bearer $(ConvertFrom-SecureToken $script:AccessToken)"
                'Accept'        = 'application/json'
            }
        }
        if ($Method -ne 'GET') {
            # DELETE with no body: do not send body (NinjaOne rejects DELETE with body -> 400)
            if ($Method -eq 'DELETE' -and -not $Body) {
                # omit ContentType and Body
            } else {
                if ($Body) {
                    $p.ContentType = 'application/json'
                    $p.Body = ($Body | ConvertTo-Json -Depth 10)
                } else {
                    $p.ContentType = 'application/json'
                    $p.Body = '{}'
                }
            }
        }
        try {
            return Invoke-RestMethod @p
        } catch {
            $statusCode = 0
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            if ($attempt -eq 1 -and $statusCode -eq 401 -and (Test-RefreshTokenPresent)) {
                $script:TokenExpiresAt = [datetime]::MinValue
                Invoke-TokenRefresh
                continue
            }
            throw
        }
    }
}

function Get-ValidBearerToken {
    if (-not (Test-TokenValid)) { Invoke-TokenRefresh }
    $plain = $null
    if ($script:AccessToken) { $plain = ConvertFrom-SecureToken $script:AccessToken }
    if ([string]::IsNullOrWhiteSpace($plain) -and (Test-RefreshTokenPresent)) {
        Invoke-TokenRefresh
        $plain = ConvertFrom-SecureToken $script:AccessToken
    }
    if ([string]::IsNullOrWhiteSpace($plain)) {
        throw "Not authenticated. Please sign in again."
    }
    return $plain
}

function ConvertTo-ListItems {
    param(
        [Parameter(Mandatory)] $Response
    )
    if ($Response -is [Array]) {
        return @($Response)
    }

    $psObj = $Response.PSObject
    if ($psObj) {
        foreach ($propName in @('data', 'items', 'organizations', 'locations', 'roles', 'list', 'users', 'contacts')) {
            if ($psObj.Properties[$propName]) {
                $value = $Response.$propName
                if ($value -is [Array]) {
                    return @($value)
                }
            }
        }
    }

    return @($Response)
}

function Test-SignedIn {
    if (-not (Test-TokenValid) -and -not (Test-RefreshTokenPresent)) {
        $lblStatus.Text = 'Not signed in. Please sign in first.'
        [System.Media.SystemSounds]::Hand.Play()
        return $false
    }
    return $true
}

function Ensure-ApiCaches {
    if ($null -eq $script:OrgCache) {
        $lblStatus.Text = 'Loading organizations...'
        Push-UIUpdate
        $orgResp = Invoke-NinjaApi -Endpoint 'organizations'
        $script:OrgCache = ConvertTo-ListItems -Response $orgResp
    }
    if ($null -eq $script:LocationCache) {
        $lblStatus.Text = 'Loading locations...'
        Push-UIUpdate
        $locResp = Invoke-NinjaApi -Endpoint 'locations'
        $script:LocationCache = ConvertTo-ListItems -Response $locResp
    }
    if ($null -eq $script:RoleCache -or $null -eq $script:StagedRoleCache) {
        $lblStatus.Text = 'Loading device roles...'
        Push-UIUpdate
        $rolesResp = Invoke-NinjaApi -Endpoint 'noderole/list'
        $allRoles = ConvertTo-ListItems -Response $rolesResp
        if ($null -eq $script:RoleCache) {
            $script:RoleCache = @($allRoles | Where-Object { $_.nodeClass -eq 'UNMANAGED_DEVICE' })
        }
        if ($null -eq $script:StagedRoleCache) {
            $script:StagedRoleCache = @($allRoles | Where-Object { $_.nodeClass -ne 'UNMANAGED_DEVICE' })
        }
    }
}

function ConvertTo-ScalarString {
    param(
        [Parameter(Mandatory = $false)]
        $Value
    )

    if ($null -eq $Value) { return $null }

    $current = $Value
    while ($current -is [System.Array] -and $current.Count -gt 0) {
        $current = $current | Where-Object { $_ -ne $null -and -not [string]::IsNullOrWhiteSpace($_.ToString()) } | Select-Object -First 1
        if ($null -eq $current) { break }
    }

    if ($null -eq $current) { return $null }

    $s = $current -as [string]
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    return $s.Trim()
}

# Resolves a user/contact by numeric ID. When the same ID exists in both users and contacts,
# we prefer the contact when the user has userType TECHNICIAN so Scan & Assign assigns to the end user.
function Find-UserById {
    param([int]$UserId)
    try {
        $usersResp = Invoke-NinjaApi -Endpoint 'users'
        $allUsers = @(ConvertTo-ListItems -Response $usersResp)
        $matchedUsers = $allUsers | Where-Object { $_.id -eq $UserId }
        $m = $null
        if ($matchedUsers) {
            $m = $matchedUsers | Where-Object {
                $_.PSObject.Properties['userType'] -and $_.userType -eq 'END_USER'
            } | Select-Object -First 1
            if (-not $m) {
                $m = $matchedUsers | Select-Object -First 1
            }
        }
        if ($m) {
            $userType = if ($m.PSObject.Properties['userType']) { $m.userType } else { $null }
            if ($userType -eq 'END_USER') {
                $first = ConvertTo-ScalarString -Value $m.firstname
                $last  = ConvertTo-ScalarString -Value $m.lastname
                $email = ConvertTo-ScalarString -Value $m.email
                $uid   = ConvertTo-ScalarString -Value $m.uid
                $nameParts = @()
                if ($first) { $nameParts += $first }
                if ($last)  { $nameParts += $last }
                $n = $nameParts -join ' '
                if ([string]::IsNullOrWhiteSpace($n)) { $n = "User $UserId" }
                return [PSCustomObject]@{ Id = $m.id; Uid = if ($uid) { $uid } else { $m.id }; Name = $n; Email = $email }
            }
            # User is TECHNICIAN or other; when same ID exists in contacts, prefer contact so assignment goes to end user.
            try {
                $contactsResp = Invoke-NinjaApi -Endpoint 'contacts'
                $allContacts = @(ConvertTo-ListItems -Response $contactsResp)
                $c = $allContacts | Where-Object { $_.id -eq $UserId } | Select-Object -First 1
                if ($c) {
                    $first = ConvertTo-ScalarString -Value $c.firstname
                    $last  = ConvertTo-ScalarString -Value $c.lastname
                    $nameField = ConvertTo-ScalarString -Value $c.name
                    $email = ConvertTo-ScalarString -Value $c.email
                    $uid   = ConvertTo-ScalarString -Value $c.uid
                    $nameParts = @()
                    if ($first) { $nameParts += $first }
                    if ($last)  { $nameParts += $last }
                    $n = $nameParts -join ' '
                    if ([string]::IsNullOrWhiteSpace($n)) { $n = if ($nameField) { $nameField } else { "Contact $UserId" } }
                    return [PSCustomObject]@{ Id = $c.id; Uid = if ($uid) { $uid } else { $c.id }; Name = $n; Email = $email }
                }
            } catch {
                Write-Verbose "Contacts lookup for ID ${UserId} (technician fallback): $($_.Exception.Message)"
            }
            # No contact with this ID; return the user.
            $first = ConvertTo-ScalarString -Value $m.firstname
            $last  = ConvertTo-ScalarString -Value $m.lastname
            $email = ConvertTo-ScalarString -Value $m.email
            $uid   = ConvertTo-ScalarString -Value $m.uid
            $nameParts = @()
            if ($first) { $nameParts += $first }
            if ($last)  { $nameParts += $last }
            $n = $nameParts -join ' '
            if ([string]::IsNullOrWhiteSpace($n)) { $n = "User $UserId" }
            return [PSCustomObject]@{ Id = $m.id; Uid = if ($uid) { $uid } else { $m.id }; Name = $n; Email = $email }
        }
    } catch {
        Write-Verbose "Failed user lookup in users endpoint for ID ${UserId}: $($_.Exception.Message)"
    }
    try {
        $contactsResp = Invoke-NinjaApi -Endpoint 'contacts'
        $allContacts = @(ConvertTo-ListItems -Response $contactsResp)
        $m = $allContacts | Where-Object { $_.id -eq $UserId } | Select-Object -First 1
        if ($m) {
            $first = ConvertTo-ScalarString -Value $m.firstname
            $last  = ConvertTo-ScalarString -Value $m.lastname
            $nameField = ConvertTo-ScalarString -Value $m.name
            $email = ConvertTo-ScalarString -Value $m.email
            $uid   = ConvertTo-ScalarString -Value $m.uid

            $nameParts = @()
            if ($first) { $nameParts += $first }
            if ($last)  { $nameParts += $last }
            $n = $nameParts -join ' '
            if ([string]::IsNullOrWhiteSpace($n)) {
                $n = if ($nameField) { $nameField } else { "Contact $UserId" }
            }

            return [PSCustomObject]@{
                Id    = $m.id
                Uid   = if ($uid) { $uid } else { $m.id }
                Name  = $n
                Email = $email
            }
        }
    } catch {
        Write-Verbose "Failed user lookup in contacts endpoint for ID ${UserId}: $($_.Exception.Message)"
    }
    return $null
}

function Get-DeviceInfo {
    param([int]$DeviceId)
    $d = Invoke-NinjaApi -Endpoint "device/$DeviceId"
    if (-not $d) {
        throw "API returned no data for device $DeviceId"
    }

    # Use PSObject.Properties so missing properties don't throw under Set-StrictMode
    $displayNameProp = $d.PSObject.Properties['displayName']
    $systemNameProp  = $d.PSObject.Properties['systemName']
    $idProp          = $d.PSObject.Properties['id']

    $displayName = if ($displayNameProp) {
        ConvertTo-ScalarString -Value $displayNameProp.Value
    } else {
        $null
    }

    $systemName = if ($systemNameProp) {
        ConvertTo-ScalarString -Value $systemNameProp.Value
    } else {
        $null
    }

    $name = if ($displayName) {
        $displayName
    } elseif ($systemName) {
        $systemName
    } else {
        "Device $DeviceId"
    }

    $resolvedId = if ($idProp -and $null -ne $idProp.Value) { $idProp.Value } else { $DeviceId }

    return [PSCustomObject]@{
        Id   = $resolvedId
        Name = $name
    }
}

function Set-NinjaDeviceOwner {
    param([int]$DeviceId, $OwnerUid)
    Invoke-NinjaApi -Method POST -Endpoint "device/$DeviceId/owner/$OwnerUid"
}

function Get-QRData {
    param([string]$Text)
    $t = $Text.Trim()
    if ($t -match 'userDashboard/(\d+)') {
        $id = [int]$Matches[1]
        return @{ Type = 'user'; Id = $id }
    }
    if ($t -match 'deviceDashboard/(\d+)') {
        $id = [int]$Matches[1]
        return @{ Type = 'device'; Id = $id }
    }
    return $null
}

function Get-DateToUnixSeconds {
    param([datetime]$Date)
    return [int]($Date.ToUniversalTime() - [datetime]'1970-01-01').TotalSeconds
}

function ConvertTo-UnixMilliseconds {
    param([datetime]$Date)
    return [int64](($Date.ToUniversalTime() - [datetime]'1970-01-01').TotalSeconds * 1000)
}

function ConvertTo-OptionalDateParseResult {
    param(
        [string]$Value
    )
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return [PSCustomObject]@{ Success = $true; Date = $null; Message = $null }
    }
    try {
        return [PSCustomObject]@{ Success = $true; Date = [datetime]$Value; Message = $null }
    } catch {
        return [PSCustomObject]@{
            Success = $false
            Date    = $null
            Message = "Invalid date '$Value'. Expected a valid date such as YYYY-MM-DD."
        }
    }
}

function ConvertTo-OptionalIntAmountParseResult {
    param(
        [string]$Value
    )
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return [PSCustomObject]@{ Success = $true; Amount = $null; Message = $null }
    }
    if ($Value -match '^\d+(\.\d+)?$') {
        return [PSCustomObject]@{ Success = $true; Amount = [int][double]$Value; Message = $null }
    }
    return [PSCustomObject]@{
        Success = $false
        Amount  = $null
        Message = "Invalid amount '$Value'. Use numeric values only."
    }
}

function Build-UnmanagedDeviceBody {
    param(
        [string]$DisplayName,
        [int]$RoleId,
        [int]$OrgId,
        [int]$LocationId,
        [datetime]$WarrantyStart,
        [datetime]$WarrantyEnd,
        [string]$Serial
    )
    return @{
        name              = $DisplayName
        roleId            = $RoleId
        orgId             = $OrgId
        locationId        = $LocationId
        warrantyStartDate = (Get-DateToUnixSeconds -Date $WarrantyStart)
        warrantyEndDate   = (Get-DateToUnixSeconds -Date $WarrantyEnd)
        serialNumber      = $Serial
    }
}

function Build-StagedDeviceBody {
    param(
        [string]$Name,
        [int]$OrgId,
        [int]$LocationId,
        [int]$RoleId,
        [string]$AssignedUserUid,
        [datetime]$WarrantyStart,
        [datetime]$WarrantyEnd,
        [string]$ItamAssetId,
        [string]$ItamAssetStatus,
        [long]$ItamAssetPurchaseDate,
        [int]$ItamAssetPurchaseAmount,
        [string]$ItamAssetExpectedLifetime,
        [long]$ItamAssetEndOfLifeDate,
        [string]$ItamAssetSerialNumber
    )
    $body = @{
        name           = $Name
        orgId          = $OrgId
        locationId     = $LocationId
        roleId         = $RoleId
    }
    if (-not [string]::IsNullOrWhiteSpace($AssignedUserUid)) {
        $body['assignedUserUid'] = $AssignedUserUid.Trim()
    }
    if ($WarrantyStart) {
        $body['warrantyStartDate'] = (Get-DateToUnixSeconds -Date $WarrantyStart)
    }
    if ($WarrantyEnd) {
        $body['warrantyEndDate'] = (Get-DateToUnixSeconds -Date $WarrantyEnd)
    }
    if (-not [string]::IsNullOrWhiteSpace($ItamAssetId)) {
        $body['itamAssetId'] = $ItamAssetId.Trim()
    }
    if (-not [string]::IsNullOrWhiteSpace($ItamAssetStatus)) {
        $body['itamAssetStatus'] = $ItamAssetStatus.Trim()
    }
    if ($ItamAssetPurchaseDate -ne $null -and $ItamAssetPurchaseDate -ne 0) {
        $body['itamAssetPurchaseDate'] = $ItamAssetPurchaseDate
    }
    if ($ItamAssetPurchaseAmount -ne $null) {
        $body['itamAssetPurchaseAmount'] = $ItamAssetPurchaseAmount
    }
    if (-not [string]::IsNullOrWhiteSpace($ItamAssetExpectedLifetime)) {
        $body['itamAssetExpectedLifetime'] = $ItamAssetExpectedLifetime.Trim().ToLower()
    }
    if ($ItamAssetEndOfLifeDate -ne $null -and $ItamAssetEndOfLifeDate -ne 0) {
        $body['itamAssetEndOfLifeDate'] = $ItamAssetEndOfLifeDate
    }
    if (-not [string]::IsNullOrWhiteSpace($ItamAssetSerialNumber)) {
        $body['itamAssetSerialNumber'] = $ItamAssetSerialNumber.Trim()
    }
    return $body
}

function Build-AssetCustomFieldsBody {
    param(
        [string]$Make,
        [string]$Model,
        [string]$Serial,
        [string]$AssetStatus,
        [string]$ExpectedLifetime,
        [datetime]$PurchaseDate,
        [datetime]$EndOfLifeDate,
        [int]$PurchaseAmount
    )

    $cf = @{}
    if ($Make)   { $cf['manufacturer'] = $Make }
    if ($Model)  { $cf['model'] = $Model }
    if ($Serial) { $cf['itamAssetSerialNumber'] = $Serial }
    if ($PurchaseDate) {
        $cf['itamAssetPurchaseDate'] = (ConvertTo-UnixMilliseconds -Date $PurchaseDate)
    }
    if ($PurchaseAmount -ne $null) {
        $cf['itamAssetPurchaseAmount'] = $PurchaseAmount
    }
    if ($AssetStatus) {
        $cf['itamAssetStatus'] = $AssetStatus
    }
    if ($ExpectedLifetime) {
        $cf['itamAssetExpectedLifetime'] = $ExpectedLifetime.ToLower()
    }
    if ($EndOfLifeDate) {
        $cf['itamAssetEndOfLifeDate'] = (ConvertTo-UnixMilliseconds -Date $EndOfLifeDate)
    }
    return $cf
}

function Get-RowValue {
    param([PSCustomObject]$Row, [string]$ColumnName)
    $prop = $Row.PSObject.Properties | Where-Object { $_.Name -ieq $ColumnName } | Select-Object -First 1
    if (-not $prop) { return '' }
    $v = $prop.Value -as [string]
    if ($null -eq $v) { return '' }
    return $v.Trim()
}

function ConvertTo-DataTable {
    param([PSCustomObject[]]$Data)
    $dt = New-Object System.Data.DataTable
    if (-not $Data -or $Data.Count -eq 0) { return $dt }
    foreach ($prop in $Data[0].PSObject.Properties) {
        [void]$dt.Columns.Add($prop.Name, [string])
    }
    foreach ($row in $Data) {
        $dr = $dt.NewRow()
        foreach ($prop in $row.PSObject.Properties) {
            $dr[$prop.Name] = if ($null -eq $prop.Value) { '' } else { $prop.Value.ToString() }
        }
        $dt.Rows.Add($dr)
    }
    return $dt
}
#endregion

#region WPF XAML
$xaml = @"
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="NinjaOne ITAM Manager" Height="780" Width="720"
  WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip"
  MinHeight="620" MinWidth="600">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="12,6"/>
      <Setter Property="Margin" Value="4,2"/>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Margin" Value="0,0,0,8"/>
      <Setter Property="Padding" Value="4"/>
    </Style>
  </Window.Resources>
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" FontSize="20" FontWeight="Bold" Margin="0,0,0,8"
               Text="NinjaOne ITAM Manager"/>

    <Expander Grid.Row="1" x:Name="expSettings" Header="Connection Settings" Margin="0,0,0,8">
      <Border BorderBrush="#DDD" BorderThickness="0,1,0,0" Padding="8,10,8,4">
        <StackPanel>
          <Grid Margin="0,0,0,6">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="80"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock VerticalAlignment="Center" Text="Instance"/>
            <TextBox Grid.Column="1" x:Name="tbInstance" Height="26"
                     VerticalContentAlignment="Center"/>
          </Grid>
          <Grid Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="80"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock VerticalAlignment="Center" Text="Client ID"/>
            <TextBox Grid.Column="1" x:Name="tbClientId" Height="26"
                     VerticalContentAlignment="Center"/>
          </Grid>
          <StackPanel Orientation="Horizontal">
            <Button x:Name="btnSignIn" Content="Sign In to NinjaOne" Width="170"/>
            <TextBlock x:Name="lblAuthStatus" VerticalAlignment="Center"
                       Margin="12,0,0,0" FontWeight="SemiBold"
                       Text="Not connected" Foreground="Gray"/>
          </StackPanel>
          <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
            <Button x:Name="btnChangeMasterPwd" Content="Change Master Password"
                    Width="170" Visibility="Collapsed"/>
            <Button x:Name="btnClearSession" Content="Clear Saved Session"
                    Margin="8,0,0,0" Width="140" Visibility="Collapsed"/>
            <TextBlock x:Name="lblSessionHint" VerticalAlignment="Center"
                       Margin="12,0,0,0" FontSize="11" Foreground="#888"
                       Text="" TextWrapping="Wrap"/>
          </StackPanel>
        </StackPanel>
      </Border>
    </Expander>

    <TabControl Grid.Row="2" x:Name="tabControl" Margin="0,0,0,8">

      <!-- Tab 1: Import Equipment -->
      <TabItem Header="Import Equipment" x:Name="tabImport">
        <Grid Margin="8">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="140"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
            <RadioButton x:Name="rbCsv" Content="CSV Import" IsChecked="True"
                         Margin="0,0,16,0" VerticalAlignment="Center"/>
            <RadioButton x:Name="rbManual" Content="Manual Entry"
                         VerticalAlignment="Center"/>
          </StackPanel>

          <Grid Grid.Row="1">
            <!-- CSV Panel -->
            <DockPanel x:Name="pnlCsv">
              <Grid DockPanel.Dock="Top" Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="tbCsvPath" Height="28"
                         VerticalContentAlignment="Center"
                         ToolTip="Path to CSV file"/>
                <Button Grid.Column="1" x:Name="btnBrowseCsv" Content="Browse..."/>
                <Button Grid.Column="2" x:Name="btnImportCsv" Content="Import"
                        FontWeight="SemiBold"/>
              </Grid>
              <DataGrid x:Name="dgCsvPreview" IsReadOnly="True"
                        AutoGenerateColumns="True" CanUserAddRows="False"
                        CanUserDeleteRows="False" HeadersVisibility="Column"
                        VerticalScrollBarVisibility="Auto"/>
            </DockPanel>

            <!-- Manual Panel (initially hidden) -->
            <ScrollViewer x:Name="pnlManual" Visibility="Collapsed"
                          VerticalScrollBarVisibility="Auto">
              <StackPanel>
                <Grid Margin="0,0,0,4">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="110"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                    <RowDefinition Height="32"/>
                  </Grid.RowDefinitions>
                  <TextBlock Grid.Row="0" Text="Device type *" VerticalAlignment="Center"/>
                  <StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal">
                    <RadioButton x:Name="rbManualUnmanaged" Content="Unmanaged device" IsChecked="True"
                                Margin="0,0,16,0" VerticalAlignment="Center"/>
                    <RadioButton x:Name="rbManualStaged" Content="Staged device"
                                VerticalAlignment="Center"/>
                  </StackPanel>
                  <TextBlock Grid.Row="1" Text="Name *" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="1" Grid.Column="1" x:Name="tbManualName"
                           Height="26" VerticalContentAlignment="Center"/>
                  <TextBlock Grid.Row="2" Text="Role *" VerticalAlignment="Center"/>
                  <ComboBox Grid.Row="2" Grid.Column="1" x:Name="cbManualRole"
                            Height="26"/>
                  <TextBlock Grid.Row="3" Text="Organization *" VerticalAlignment="Center"/>
                  <ComboBox Grid.Row="3" Grid.Column="1" x:Name="cbManualOrg"
                            Height="26"/>
                  <TextBlock Grid.Row="4" Text="Location *" VerticalAlignment="Center"/>
                  <ComboBox Grid.Row="4" Grid.Column="1" x:Name="cbManualLoc"
                            Height="26"/>
                  <TextBlock Grid.Row="5" Text="Serial Number" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="5" Grid.Column="1" x:Name="tbManualSerial"
                           Height="26" VerticalContentAlignment="Center"/>
                  <TextBlock Grid.Row="6" Text="Make" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="6" Grid.Column="1" x:Name="tbManualMake"
                           Height="26" VerticalContentAlignment="Center"/>
                  <TextBlock Grid.Row="7" Text="Model" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="7" Grid.Column="1" x:Name="tbManualModel"
                           Height="26" VerticalContentAlignment="Center"/>
                  <TextBlock Grid.Row="8" Text="Purchase Date" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="8" Grid.Column="1" x:Name="tbManualPurchDate"
                           Height="26" VerticalContentAlignment="Center"
                           ToolTip="YYYY-MM-DD"/>
                  <TextBlock Grid.Row="9" Text="Purchase Amt" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="9" Grid.Column="1" x:Name="tbManualAmount"
                           Height="26" VerticalContentAlignment="Center"/>
                  <TextBlock Grid.Row="10" Text="Warranty Start" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="10" Grid.Column="1" x:Name="tbManualWarrantyStart"
                           Height="26" VerticalContentAlignment="Center"
                           ToolTip="YYYY-MM-DD"/>
                  <TextBlock Grid.Row="11" Text="Warranty End" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="11" Grid.Column="1" x:Name="tbManualWarrantyEnd"
                           Height="26" VerticalContentAlignment="Center"
                           ToolTip="YYYY-MM-DD"/>
                  <TextBlock Grid.Row="12" Text="Asset Status" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="12" Grid.Column="1" x:Name="tbManualAssetStatus"
                           Height="26" VerticalContentAlignment="Center"
                           ToolTip="e.g. In Use, Retired"/>
                  <TextBlock Grid.Row="13" Text="Expected Life" VerticalAlignment="Center"/>
                  <ComboBox Grid.Row="13" Grid.Column="1" x:Name="cbManualExpLifetime"
                            Height="26"/>
                  <TextBlock Grid.Row="14" Text="End of Life" VerticalAlignment="Center"/>
                  <TextBox Grid.Row="14" Grid.Column="1" x:Name="tbManualEolDate"
                           Height="26" VerticalContentAlignment="Center"
                           ToolTip="YYYY-MM-DD"/>
                </Grid>
                <Button x:Name="btnManualAdd" Content="Add Device"
                        HorizontalAlignment="Left" FontWeight="SemiBold"
                        Margin="0,4,0,0"/>
              </StackPanel>
            </ScrollViewer>
          </Grid>

          <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,4">
            <TextBlock Text="Imported Devices: "/>
            <TextBlock x:Name="lblImportCount" Text="0"/>
          </StackPanel>
          <ListBox Grid.Row="3" x:Name="lbImportResults"/>
        </Grid>
      </TabItem>

      <!-- Tab 2: Generate QR Codes -->
      <TabItem Header="Generate QR Codes" x:Name="tabQrGen">
        <Grid Margin="8">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="150"/>
          </Grid.RowDefinitions>

          <Grid Grid.Row="0" Margin="0,0,0,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox x:Name="tbQrDeviceId" Height="28"
                     VerticalContentAlignment="Center"
                     ToolTip="Enter a device ID to add"/>
            <Button Grid.Column="1" x:Name="btnQrAddDevice" Content="Add"/>
            <Button Grid.Column="2" x:Name="btnQrRefreshImport"
                    Content="From Import"
                    ToolTip="Load devices imported in the Import tab"/>
            <Button Grid.Column="3" x:Name="btnQrRemoveDevice" Content="Remove"/>
          </Grid>

          <ListBox Grid.Row="1" x:Name="lbQrDevices" Margin="0,0,0,8"/>

          <Grid Grid.Row="2" Margin="0,0,0,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="70"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Output Dir" VerticalAlignment="Center"/>
            <TextBox Grid.Column="1" x:Name="tbQrOutputDir" Height="28"
                     VerticalContentAlignment="Center"/>
            <Button Grid.Column="2" x:Name="btnQrBrowseDir" Content="Browse..."/>
          </Grid>

          <Grid Grid.Row="3" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="70"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="QR Size" VerticalAlignment="Center"/>
            <ComboBox Grid.Column="1" x:Name="cbQrSize" Width="80" Height="26"/>
            <Button Grid.Column="3" x:Name="btnQrGenerate"
                    Content="Generate QR Codes" FontWeight="SemiBold"/>
          </Grid>

          <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,4">
            <TextBlock Text="Generated: "/>
            <TextBlock x:Name="lblQrGenCount" Text="0"/>
          </StackPanel>
          <ListBox Grid.Row="5" x:Name="lbQrResults"/>
        </Grid>
      </TabItem>

      <!-- Tab 3: Upload QR Codes -->
      <TabItem Header="Upload QR Codes" x:Name="tabUpload">
        <Grid Margin="8">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Grid Grid.Row="0" Margin="0,0,0,4">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="80"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Image Dir" VerticalAlignment="Center"/>
            <TextBox Grid.Column="1" x:Name="tbUploadDir" Height="28"
                     VerticalContentAlignment="Center"/>
            <Button Grid.Column="2" x:Name="btnUploadBrowseDir"
                    Content="Browse..."/>
            <Button Grid.Column="3" x:Name="btnUploadScan"
                    Content="Scan Directory"/>
          </Grid>

          <Grid Grid.Row="1" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="80"/>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Description" VerticalAlignment="Center"/>
            <TextBox Grid.Column="1" x:Name="tbUploadDesc" Height="28"
                     VerticalContentAlignment="Center"
                     Text="Device dashboard QR code"/>
            <CheckBox Grid.Column="2" x:Name="chkUploadReplace"
                      Content="Replace existing" VerticalAlignment="Center"
                      Margin="12,0,0,0"/>
          </Grid>

          <ListBox Grid.Row="2" x:Name="lbUploadFiles" Margin="0,0,0,8"/>

          <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,0,0,4">
            <Button x:Name="btnUpload" Content="Upload All"
                    FontWeight="SemiBold"/>
            <TextBlock x:Name="lblUploadCount" VerticalAlignment="Center"
                       Margin="12,0,0,0" Foreground="Gray"
                       Text="0 files found"/>
          </StackPanel>
        </Grid>
      </TabItem>

      <!-- Tab 4: Scan & Assign -->
      <TabItem Header="Scan &amp; Assign" x:Name="tabScan">
        <Grid Margin="8">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <GroupBox Grid.Row="0" Header="Scanner Input">
            <TextBox x:Name="tbScanInput" Height="28" FontSize="12"
                     VerticalContentAlignment="Center" Margin="2"
                     ToolTip="Focus this field, then scan a QR code or paste a URL and press Enter."/>
          </GroupBox>

          <GroupBox Grid.Row="1" Header="Step 1 &#x2014; User">
            <Grid Margin="2">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock x:Name="lblScanUserInfo" TextWrapping="Wrap"
                         VerticalAlignment="Center" Foreground="Gray"
                         Text="No user scanned. Scan a user QR code to begin."/>
              <Button x:Name="btnScanClearUser" Grid.Column="1" Content="Clear"
                      Visibility="Collapsed" VerticalAlignment="Center"/>
            </Grid>
          </GroupBox>

          <GroupBox Grid.Row="2" Header="Step 2 &#x2014; Devices">
            <DockPanel Margin="2">
              <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal"
                          Margin="0,6,0,0">
                <TextBlock x:Name="lblScanDeviceCount" VerticalAlignment="Center"
                           Foreground="Gray" Text="0 devices scanned"/>
                <Button x:Name="btnScanRemoveDevice" Content="Remove Selected"
                        Margin="12,0,0,0" IsEnabled="False"/>
              </StackPanel>
              <ListBox x:Name="lbScanDevices" MinHeight="60"/>
            </DockPanel>
          </GroupBox>

          <StackPanel Grid.Row="3" Orientation="Horizontal"
                      HorizontalAlignment="Center" Margin="0,0,0,4">
            <Button x:Name="btnScanAssign" Content="Assign All to User"
                    Width="170" FontWeight="SemiBold" IsEnabled="False"/>
            <Button x:Name="btnScanReset" Content="Reset" Width="100"/>
          </StackPanel>
        </Grid>
      </TabItem>

    </TabControl>

    <Border Grid.Row="3" Background="#F5F5F5" CornerRadius="4" Padding="8,6">
      <TextBlock x:Name="lblStatus" TextWrapping="Wrap" FontSize="12"
                 Text="Enter connection settings and sign in to begin."/>
    </Border>
  </Grid>
</Window>
"@
#endregion

#region Create Window and Bind Elements
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$expSettings         = $window.FindName('expSettings')
$tbInstance          = $window.FindName('tbInstance')
$tbClientId          = $window.FindName('tbClientId')
$btnSignIn           = $window.FindName('btnSignIn')
$lblAuthStatus       = $window.FindName('lblAuthStatus')
$btnChangeMasterPwd  = $window.FindName('btnChangeMasterPwd')
$btnClearSession     = $window.FindName('btnClearSession')
$lblSessionHint      = $window.FindName('lblSessionHint')
$tabControl          = $window.FindName('tabControl')
$lblStatus           = $window.FindName('lblStatus')

$tabImport           = $window.FindName('tabImport')
$rbCsv               = $window.FindName('rbCsv')
$rbManual            = $window.FindName('rbManual')
$pnlCsv              = $window.FindName('pnlCsv')
$pnlManual           = $window.FindName('pnlManual')
$tbCsvPath           = $window.FindName('tbCsvPath')
$btnBrowseCsv        = $window.FindName('btnBrowseCsv')
$btnImportCsv        = $window.FindName('btnImportCsv')
$dgCsvPreview        = $window.FindName('dgCsvPreview')
$rbManualUnmanaged   = $window.FindName('rbManualUnmanaged')
$rbManualStaged      = $window.FindName('rbManualStaged')
$tbManualName        = $window.FindName('tbManualName')
$cbManualRole        = $window.FindName('cbManualRole')
$cbManualOrg         = $window.FindName('cbManualOrg')
$cbManualLoc         = $window.FindName('cbManualLoc')
$tbManualSerial      = $window.FindName('tbManualSerial')
$tbManualMake        = $window.FindName('tbManualMake')
$tbManualModel       = $window.FindName('tbManualModel')
$tbManualPurchDate   = $window.FindName('tbManualPurchDate')
$tbManualAmount      = $window.FindName('tbManualAmount')
$tbManualWarrantyStart = $window.FindName('tbManualWarrantyStart')
$tbManualWarrantyEnd = $window.FindName('tbManualWarrantyEnd')
$tbManualAssetStatus = $window.FindName('tbManualAssetStatus')
$cbManualExpLifetime = $window.FindName('cbManualExpLifetime')
$tbManualEolDate     = $window.FindName('tbManualEolDate')
$btnManualAdd        = $window.FindName('btnManualAdd')
$lbImportResults     = $window.FindName('lbImportResults')
$lblImportCount      = $window.FindName('lblImportCount')

$tabQrGen            = $window.FindName('tabQrGen')
$tbQrDeviceId        = $window.FindName('tbQrDeviceId')
$btnQrAddDevice      = $window.FindName('btnQrAddDevice')
$btnQrRefreshImport  = $window.FindName('btnQrRefreshImport')
$btnQrRemoveDevice   = $window.FindName('btnQrRemoveDevice')
$lbQrDevices         = $window.FindName('lbQrDevices')
$tbQrOutputDir       = $window.FindName('tbQrOutputDir')
$btnQrBrowseDir      = $window.FindName('btnQrBrowseDir')
$cbQrSize            = $window.FindName('cbQrSize')
$btnQrGenerate       = $window.FindName('btnQrGenerate')
$lbQrResults         = $window.FindName('lbQrResults')
$lblQrGenCount       = $window.FindName('lblQrGenCount')

$tabUpload           = $window.FindName('tabUpload')
$tbUploadDir         = $window.FindName('tbUploadDir')
$btnUploadBrowseDir  = $window.FindName('btnUploadBrowseDir')
$btnUploadScan       = $window.FindName('btnUploadScan')
$lbUploadFiles       = $window.FindName('lbUploadFiles')
$tbUploadDesc        = $window.FindName('tbUploadDesc')
$chkUploadReplace    = $window.FindName('chkUploadReplace')
$btnUpload           = $window.FindName('btnUpload')
$lblUploadCount      = $window.FindName('lblUploadCount')

$tabScan             = $window.FindName('tabScan')
$tbScanInput         = $window.FindName('tbScanInput')
$lblScanUserInfo     = $window.FindName('lblScanUserInfo')
$btnScanClearUser    = $window.FindName('btnScanClearUser')
$lbScanDevices       = $window.FindName('lbScanDevices')
$lblScanDeviceCount  = $window.FindName('lblScanDeviceCount')
$btnScanRemoveDevice = $window.FindName('btnScanRemoveDevice')
$btnScanAssign       = $window.FindName('btnScanAssign')
$btnScanReset        = $window.FindName('btnScanReset')
#endregion

#region Initialize Defaults
$script:ITAMConfig = Get-ITAMConfig
$hasSavedSession = -not [string]::IsNullOrWhiteSpace($script:ITAMConfig.EncryptedRefreshToken) -and
                   -not [string]::IsNullOrWhiteSpace($script:ITAMConfig.MasterPasswordVerifier)

if ($hasSavedSession) {
    $tbInstance.Text = if (-not [string]::IsNullOrWhiteSpace($script:ITAMConfig.NinjaInstance)) {
        $inst = $script:ITAMConfig.NinjaInstance -replace '^https?://', ''
        $inst.TrimEnd('/')
    } elseif ($NinjaOneInstance) { $NinjaOneInstance } else { '' }
    $tbClientId.Text = if (-not [string]::IsNullOrWhiteSpace($script:ITAMConfig.ClientId)) {
        $script:ITAMConfig.ClientId
    } elseif ($ClientId) { $ClientId } else { '' }
    $lblSessionHint.Text = 'Saved session found. Click Sign In to reconnect.'
    $btnClearSession.Visibility = 'Visible'
    $expSettings.IsExpanded = $true
} else {
    $tbInstance.Text = $NinjaOneInstance
    $tbClientId.Text = if ($ClientId) { $ClientId } else { '' }
}

$hasDefaults = -not [string]::IsNullOrWhiteSpace($tbClientId.Text) -and `
               -not [string]::IsNullOrWhiteSpace($tbInstance.Text)
if (-not $hasSavedSession) {
    $expSettings.IsExpanded = -not $hasDefaults
}

$tbQrOutputDir.Text = '.\DeviceQRCodes'
foreach ($size in @(100, 150, 200, 250, 300, 400, 500, 600)) {
    $cbQrSize.Items.Add("${size}px") | Out-Null
}
$cbQrSize.SelectedIndex = 2

foreach ($yr in @('1 years', '2 years', '3 years', '4 years', '5 years')) {
    $cbManualExpLifetime.Items.Add($yr) | Out-Null
}
#endregion

#region UI Helpers
function Push-UIUpdate {
    $frame = [System.Windows.Threading.DispatcherFrame]::new()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [Action]{ $frame.Continue = $false }
    )
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Update-ScanState {
    if ($null -eq $script:ScanUserInfo) {
        $lblStatus.Text = 'Scan a user QR code (Step 1).'
        $btnScanAssign.IsEnabled = $false
    } elseif ($script:ScanDevices.Count -eq 0) {
        $lblStatus.Text = "User: $($script:ScanUserInfo.Name). Scan device QR codes (Step 2)."
        $btnScanAssign.IsEnabled = $false
    } else {
        $lblStatus.Text = "Ready to assign $($script:ScanDevices.Count) device(s) to $($script:ScanUserInfo.Name)."
        $btnScanAssign.IsEnabled = $true
    }
    $lblScanDeviceCount.Text = "$($script:ScanDevices.Count) device(s) scanned"
    $btnScanRemoveDevice.IsEnabled = ($lbScanDevices.SelectedIndex -ge 0)
}

function Reset-ScanAll {
    $script:ScanUserInfo = $null
    $script:ScanDevices.Clear()
    $lbScanDevices.Items.Clear()
    $lblScanUserInfo.Text = 'No user scanned. Scan a user QR code to begin.'
    $lblScanUserInfo.Foreground = [System.Windows.Media.Brushes]::Gray
    $btnScanClearUser.Visibility = 'Collapsed'
    Update-ScanState
    $tbScanInput.Clear()
    $tbScanInput.Focus()
}
#endregion

#region Sign-In (Authorization Code + PKCE)
$btnSignIn.Add_Click({
    try {
        $script:NinjaBaseUrl = Resolve-BaseUrl -Instance $tbInstance.Text
    } catch {
        [System.Windows.MessageBox]::Show(
            $_.Exception.Message,
            'Invalid Instance URL', 'OK', 'Warning') | Out-Null
        return
    }
    $script:NinjaClientId = $tbClientId.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($script:NinjaClientId)) {
        [System.Windows.MessageBox]::Show(
            'Client ID is required. Enter the Client ID from your NinjaOne OAuth application.',
            'Missing Client ID', 'OK', 'Warning') | Out-Null
        return
    }

    $cfg = Get-ITAMConfig
    $hasSaved = -not [string]::IsNullOrWhiteSpace($cfg.EncryptedRefreshToken) -and
                -not [string]::IsNullOrWhiteSpace($cfg.MasterPasswordVerifier)
    if ($hasSaved) {
        $btnSignIn.IsEnabled = $false
        $lblAuthStatus.Text = 'Unlocking saved session...'
        $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::DarkOrange
        $lblStatus.Text = 'Enter your master password to reconnect using saved session.'
        Push-UIUpdate

        $masterPwd = Show-MasterPasswordPrompt -Title 'Unlock Saved Session' `
                        -Message 'Enter your master password to reconnect:'
        if ($null -eq $masterPwd) {
            $lblAuthStatus.Text = 'Not connected'
            $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::Gray
            $lblStatus.Text = 'Master password entry cancelled. Click Sign In to try again, or use browser sign-in.'
            $btnSignIn.IsEnabled = $true
            return
        }

        if (-not (Test-MasterPasswordValid -MasterPwd $masterPwd -Verifier $cfg.MasterPasswordVerifier)) {
            [System.Windows.MessageBox]::Show(
                'Incorrect master password. You can try again or clear the saved session to sign in via browser.',
                'Wrong Password', 'OK', 'Warning') | Out-Null
            $lblAuthStatus.Text = 'Not connected'
            $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::Gray
            $lblStatus.Text = 'Master password incorrect. Click Sign In to retry.'
            $btnSignIn.IsEnabled = $true
            return
        }

        try {
            $plainRefresh = Unprotect-String -CipherText $cfg.EncryptedRefreshToken -MasterPwd $masterPwd
            $script:RefreshToken = ConvertTo-SecureToken $plainRefresh
            $script:MasterPassword = $masterPwd
            $script:MasterPasswordVerifier = $cfg.MasterPasswordVerifier

            $lblAuthStatus.Text = 'Refreshing token...'
            Push-UIUpdate

            Invoke-TokenRefresh

            $lblAuthStatus.Text = 'Connected'
            $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::Green
            $expSettings.IsExpanded = $false
            $lblStatus.Text = 'Reconnected using saved session. Use any tab to begin.'
            $lblSessionHint.Text = ''
            $btnChangeMasterPwd.Visibility = 'Visible'
            $btnClearSession.Visibility = 'Visible'

            Save-CurrentSession

            $script:OrgCache      = $null
            $script:LocationCache = $null
            $script:RoleCache     = $null
            $script:StagedRoleCache = $null
            $btnSignIn.IsEnabled = $true
            return
        } catch {
            $script:RefreshToken = $null
            $script:MasterPassword = $masterPwd
            $script:MasterPasswordVerifier = $cfg.MasterPasswordVerifier
            $lblAuthStatus.Text = 'Session expired'
            $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::DarkOrange
            $lblStatus.Text = "Saved session expired or revoked. Falling back to browser sign-in..."
            Push-UIUpdate
        }
    }

    $btnSignIn.IsEnabled = $false
    $lblAuthStatus.Text = 'Waiting for browser sign-in...'
    $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::DarkOrange
    $lblStatus.Text = 'A browser window has been opened. Sign in to NinjaOne to continue.'
    Push-UIUpdate

    $script:AuthVerifier = New-PkceVerifier
    $script:AuthState = New-OAuthState
    $challenge = Get-PkceChallenge -Verifier $script:AuthVerifier
    $script:AuthTimeoutAt = [datetime]::UtcNow.AddMinutes(3)

    $script:AuthRedirectUri = "http://localhost:8888/"

    $script:AuthListener = [System.Net.HttpListener]::new()
    $script:AuthListener.Prefixes.Add($script:AuthRedirectUri)
    try {
        $script:AuthListener.Start()
    } catch {
        $lblAuthStatus.Text = 'Listener failed'
        $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::Red
        $lblStatus.Text = "Could not start local HTTP listener: $($_.Exception.Message)"
        $btnSignIn.IsEnabled = $true
        return
    }

    $scopes = 'monitoring management offline_access'
    $scopeEncoded = [uri]::EscapeDataString($scopes)

    $authorizeUrl = "$($script:NinjaBaseUrl)/ws/oauth/authorize?" +
        "response_type=code" +
        "&client_id=$([uri]::EscapeDataString($script:NinjaClientId))" +
        "&redirect_uri=$([uri]::EscapeDataString($script:AuthRedirectUri))" +
        "&scope=$scopeEncoded" +
        "&state=$([uri]::EscapeDataString($script:AuthState))" +
        "&code_challenge=$challenge" +
        "&code_challenge_method=S256"

    Start-Process $authorizeUrl

    $script:AuthPS = [powershell]::Create()
    [void]$script:AuthPS.AddScript({
        param($lst)
        try {
            $ctx = $lst.GetContext()
            $q   = $ctx.Request.Url.Query
            $html = '<html><body style="font-family:system-ui,sans-serif;text-align:center;padding-top:60px">' +
                    '<h2>Authentication complete</h2>' +
                    '<p>You may close this tab and return to the ITAM Manager.</p></body></html>'
            $buf = [System.Text.Encoding]::UTF8.GetBytes($html)
            $ctx.Response.ContentType    = 'text/html'
            $ctx.Response.ContentLength64 = $buf.Length
            $ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
            $ctx.Response.Close()
            $lst.Stop()
            return $q
        } catch {
            try { $lst.Stop() } catch { Write-Verbose "Auth listener stop failed: $($_.Exception.Message)" }
            return "error=$($_.Exception.Message)"
        }
    }).AddArgument($script:AuthListener)

    $script:AuthHandle = $script:AuthPS.BeginInvoke()

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(500)
    $timer.Add_Tick({
        if (-not $script:AuthHandle.IsCompleted) {
            if ($script:AuthTimeoutAt -ne [datetime]::MinValue -and
                [datetime]::UtcNow -lt $script:AuthTimeoutAt) {
                return
            }
            try { $script:AuthListener.Stop() } catch { }
            try { $script:AuthListener.Close() } catch { }
            try { if ($script:AuthPS) { $script:AuthPS.Stop() } } catch { }
            try { if ($script:AuthPS) { $script:AuthPS.Dispose() } } catch { }
            $script:AuthPS = $null
            $script:AuthHandle = $null
            $lblAuthStatus.Text = 'Timed out'
            $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::Red
            $lblStatus.Text = 'Sign-in timed out waiting for OAuth callback. Click Sign In to try again.'
            $btnSignIn.IsEnabled = $true
            $script:AuthState = $null
            $this.Stop()
            return
        }
        $this.Stop()

        try {
            $queryString = ($script:AuthPS.EndInvoke($script:AuthHandle) |
                Select-Object -First 1) -as [string]
            $script:AuthPS.Dispose()
            $script:AuthPS = $null

            $returnedState = $null
            if ($queryString -match '[?&]state=([^&]+)') {
                $returnedState = [uri]::UnescapeDataString($Matches[1])
            }

            if ($queryString -match '[?&]code=([^&]+)') {
                if ([string]::IsNullOrWhiteSpace($script:AuthState) -or
                    [string]::IsNullOrWhiteSpace($returnedState) -or
                    $returnedState -ne $script:AuthState) {
                    throw 'Authentication state validation failed. Please sign in again.'
                }
                $code = [uri]::UnescapeDataString($Matches[1])
                $baseUrl = Resolve-BaseUrl -Instance $tbInstance.Text
                $resp = Invoke-RestMethod `
                    -Uri "$baseUrl/ws/oauth/token" `
                    -Method POST -UseBasicParsing `
                    -Headers @{
                        'Accept'       = 'application/json'
                        'Content-Type' = 'application/x-www-form-urlencoded'
                    } `
                    -Body @{
                        grant_type    = 'authorization_code'
                        code          = $code
                        redirect_uri  = $script:AuthRedirectUri
                        client_id     = $script:NinjaClientId
                        code_verifier = $script:AuthVerifier
                    }

                Update-TokensFromResponse $resp
                $script:NinjaBaseUrl = $baseUrl
                $lblAuthStatus.Text = 'Connected'
                $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::Green
                $expSettings.IsExpanded = $false

                if (-not $script:MasterPassword) {
                    $newPwd = Show-MasterPasswordPrompt -Title 'Save Session' `
                                  -Message 'Set a master password to save your session across app restarts.' `
                                  -IsNewPassword
                    if ($newPwd) {
                        $script:MasterPassword = $newPwd
                        $script:MasterPasswordVerifier = New-MasterPasswordVerifier -MasterPwd $newPwd
                        Save-CurrentSession
                        $btnChangeMasterPwd.Visibility = 'Visible'
                        $btnClearSession.Visibility = 'Visible'
                        $lblSessionHint.Text = ''
                        $lblStatus.Text = 'Signed in successfully. Session saved.'
                    } else {
                        $lblStatus.Text = 'Signed in successfully. Session not saved (no master password set).'
                    }
                } else {
                    Save-CurrentSession
                    $btnChangeMasterPwd.Visibility = 'Visible'
                    $btnClearSession.Visibility = 'Visible'
                    $lblSessionHint.Text = ''
                    $lblStatus.Text = 'Signed in successfully. Session saved.'
                }

                $script:OrgCache      = $null
                $script:LocationCache = $null
                $script:RoleCache     = $null
            $script:StagedRoleCache = $null
            }
            elseif ($queryString -match '[?&]error=([^&]+)') {
                $errMsg = [uri]::UnescapeDataString($Matches[1])
                $lblAuthStatus.Text = 'Failed'
                $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::Red
                $lblStatus.Text = "Sign-in failed: $errMsg"
            }
            else {
                $lblAuthStatus.Text = 'Failed'
                $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::Red
                $lblStatus.Text = 'Sign-in failed: no authorization code received.'
            }
            $script:AuthState = $null
        } catch {
            $lblAuthStatus.Text = 'Error'
            $lblAuthStatus.Foreground = [System.Windows.Media.Brushes]::Red
            $errMsg = $_.Exception.Message
            if ($errMsg -match 'refresh_token') {
                $lblStatus.Text = 'Authentication error: refresh token missing. Verify your OAuth app allows Refresh Token grant and offline_access scope.'
            } else {
                $lblStatus.Text = "Authentication error: $errMsg"
            }
        }
        $btnSignIn.IsEnabled = $true
    })
    $timer.Start()
})

$btnChangeMasterPwd.Add_Click({
    $newPwd = Show-ChangeMasterPasswordPrompt
    if ($null -eq $newPwd) { return }
    $script:MasterPassword = $newPwd
    $script:MasterPasswordVerifier = New-MasterPasswordVerifier -MasterPwd $newPwd
    Save-CurrentSession
    $lblStatus.Text = 'Master password changed and session re-encrypted.'
})

$btnClearSession.Add_Click({
    $confirm = [System.Windows.MessageBox]::Show(
        'This will delete your saved session. You will need to sign in via browser on next launch. Continue?',
        'Clear Saved Session', 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }
    Clear-SavedSession
    $btnChangeMasterPwd.Visibility = 'Collapsed'
    $btnClearSession.Visibility = 'Collapsed'
    $lblSessionHint.Text = ''
    $lblStatus.Text = 'Saved session cleared.'
})
#endregion

#region Tab 1: Import Equipment

$rbCsv.Add_Checked({
    $pnlCsv.Visibility = 'Visible'
    $pnlManual.Visibility = 'Collapsed'
})

function Update-ManualRoleComboBox {
    $cbManualRole.Items.Clear()
    $roleSource = if ($rbManualStaged.IsChecked -eq $true) { $script:StagedRoleCache } else { $script:RoleCache }
    if (-not $roleSource) { return }
    foreach ($r in $roleSource) {
        $roleName = $r.name
        if ($roleName -is [Array]) {
            $roleName = $roleName | Select-Object -First 1
        }
        $roleName = $roleName -as [string]
        if (-not [string]::IsNullOrWhiteSpace($roleName)) {
            $cbManualRole.Items.Add($roleName) | Out-Null
        }
    }
}

$rbManual.Add_Checked({
    $pnlCsv.Visibility = 'Collapsed'
    $pnlManual.Visibility = 'Visible'
    if (-not (Test-SignedIn)) { return }
    try {
        Ensure-ApiCaches
        if ($cbManualOrg.Items.Count -eq 0) {
            foreach ($org in $script:OrgCache) {
                $orgName = $org.name
                if ($orgName -is [Array]) {
                    $orgName = $orgName | Select-Object -First 1
                }
                $orgName = $orgName -as [string]
                if (-not [string]::IsNullOrWhiteSpace($orgName)) {
                    $cbManualOrg.Items.Add($orgName) | Out-Null
                }
            }
        }
        Update-ManualRoleComboBox
        $lblStatus.Text = 'Manual entry mode. Fill in the fields and click Add Device.'
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match '401') { $msg = "$msg Check that your OAuth app has the correct scopes and that the instance URL matches the token issuer." }
        $lblStatus.Text = "Failed to load lookup data: $msg"
    }
})

$rbManualUnmanaged.Add_Checked({ Update-ManualRoleComboBox })
$rbManualStaged.Add_Checked({ Update-ManualRoleComboBox })

$cbManualOrg.Add_SelectionChanged({
    $cbManualLoc.Items.Clear()
    $selectedOrgName = $cbManualOrg.SelectedItem -as [string]
    if ([string]::IsNullOrWhiteSpace($selectedOrgName)) { return }
    if (-not $script:LocationCache) { return }
    $orgMatch = $script:OrgCache | Where-Object { $_.name -eq $selectedOrgName } |
        Select-Object -First 1
    if (-not $orgMatch) { return }
    $orgId = $orgMatch.id
    $filteredLocs = $script:LocationCache | Where-Object {
        ($_.organizationID -eq $orgId) -or ($_.organizationId -eq $orgId)
    }
    foreach ($loc in $filteredLocs) {
        $locName = $loc.name
        if ($locName -is [Array]) {
            $locName = $locName | Select-Object -First 1
        }
        $locName = $locName -as [string]
        if (-not [string]::IsNullOrWhiteSpace($locName)) {
            $cbManualLoc.Items.Add($locName) | Out-Null
        }
    }
    if ($cbManualLoc.Items.Count -gt 0) { $cbManualLoc.SelectedIndex = 0 }
})

$btnBrowseCsv.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dlg.Title = 'Select equipment CSV file'
    if ($dlg.ShowDialog()) {
        $tbCsvPath.Text = $dlg.FileName
        try {
            $script:CsvData = Import-Csv -LiteralPath $dlg.FileName -Encoding UTF8
            if ($script:CsvData -and $script:CsvData.Count -gt 0) {
                # Bind the CSV rows directly; DataGrid will auto-generate columns
                $dgCsvPreview.ItemsSource = $script:CsvData
                $lblStatus.Text = "Loaded $($script:CsvData.Count) row(s) from CSV. Click Import to create devices."
            } else {
                $lblStatus.Text = 'CSV is empty or has no data rows.'
            }
        } catch {
            $lblStatus.Text = "Failed to read CSV: $($_.Exception.Message)"
            $script:CsvData = $null
        }
    }
})

$btnImportCsv.Add_Click({
    if (-not (Test-SignedIn)) { return }
    if (-not $script:CsvData -or $script:CsvData.Count -eq 0) {
        $lblStatus.Text = 'No CSV data loaded. Browse for a CSV file first.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }

    try { Ensure-ApiCaches } catch {
        $msg = $_.Exception.Message
        if ($msg -match '401') { $msg = "$msg Check that your OAuth app has the correct scopes and that the instance URL matches the token issuer." }
        $lblStatus.Text = "Failed to load lookup data: $msg"
        return
    }

    $btnImportCsv.IsEnabled = $false
    $created = 0
    $failed  = 0
    $errors  = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $rowNum  = 0
    $total   = $script:CsvData.Count

    $headerNames = @($script:CsvData[0].PSObject.Properties.Name)
    $hasRoleHeader = $headerNames | Where-Object { $_ -ieq 'RoleName' } | Select-Object -First 1
    if (-not $hasRoleHeader) {
        $btnImportCsv.IsEnabled = $true
        $lblStatus.Text = "CSV is missing required header 'RoleName'."
        [System.Media.SystemSounds]::Hand.Play()
        return
    }

    foreach ($row in $script:CsvData) {
        $rowNum++
        $lblStatus.Text = "Importing row $rowNum of $total..."
        Push-UIUpdate

        try {
            $deviceTypeRaw = Get-RowValue -Row $row -ColumnName 'DeviceType'
            $isStaged = [string]::Equals($deviceTypeRaw.Trim(), 'Staged', [StringComparison]::OrdinalIgnoreCase)

            $name     = Get-RowValue -Row $row -ColumnName 'Name'
            $roleName = Get-RowValue -Row $row -ColumnName 'RoleName'
            if ([string]::IsNullOrWhiteSpace($roleName)) {
                throw "RoleName is required."
            }

            $orgId = $null
            $locId = $null
            $orgIdStr = Get-RowValue -Row $row -ColumnName 'OrganizationId'
            $locIdStr = Get-RowValue -Row $row -ColumnName 'LocationId'
            $orgName  = Get-RowValue -Row $row -ColumnName 'OrganizationName'
            $locName  = Get-RowValue -Row $row -ColumnName 'LocationName'

            if ($orgIdStr -and $locIdStr) {
                $orgId = [int]$orgIdStr
                $locId = [int]$locIdStr
            } elseif ($orgName -and $locName) {
                $orgMatch = $script:OrgCache |
                    Where-Object { $_.name -eq $orgName } |
                    Select-Object -First 1
                if (-not $orgMatch) {
                    throw "Organization not found: '$orgName'."
                }
                $orgId = $orgMatch.id
                $locMatch = $script:LocationCache | Where-Object {
                    ($_.name -eq $locName) -and
                    (($_.organizationID -eq $orgId) -or ($_.organizationId -eq $orgId))
                } | Select-Object -First 1
                if (-not $locMatch) {
                    throw "Location not found: '$locName' in org '$orgName'."
                }
                $locId = $locMatch.id
            } else {
                throw "Row must have OrganizationName/LocationName or OrganizationId/LocationId."
            }

            $roleSource = if ($isStaged) { $script:StagedRoleCache } else { $script:RoleCache }
            $roleMatch = $roleSource |
                Where-Object { $_.name -eq $roleName } |
                Select-Object -First 1
            if (-not $roleMatch) {
                $kind = if ($isStaged) { 'Staged' } else { 'Unmanaged' }
                throw "$kind device role not found: '$roleName'."
            }
            $roleId = $roleMatch.id

            $displayName = $name
            $make  = Get-RowValue -Row $row -ColumnName 'Make'
            $model = Get-RowValue -Row $row -ColumnName 'Model'
            if ([string]::IsNullOrWhiteSpace($displayName)) {
                $displayName = if ($make -and $model) { "$make $model" }
                               else { if ($isStaged) { "Staged $roleName $rowNum" } else { "Unmanaged $roleName $rowNum" } }
            }

            $serial = Get-RowValue -Row $row -ColumnName 'SerialNumber'
            if ([string]::IsNullOrWhiteSpace($serial)) { $serial = $null }

            $warrantyStart = Get-Date
            $warrantyEnd   = (Get-Date).AddYears(3)
            $ws = Get-RowValue -Row $row -ColumnName 'WarrantyStartDate'
            $we = Get-RowValue -Row $row -ColumnName 'WarrantyEndDate'
            $parsedWs = ConvertTo-OptionalDateParseResult -Value $ws
            if ($parsedWs.Success) {
                if ($parsedWs.Date) { $warrantyStart = $parsedWs.Date }
            } else {
                $warnings.Add("Row ${rowNum}: WarrantyStartDate - $($parsedWs.Message) Using current date.")
            }
            $parsedWe = ConvertTo-OptionalDateParseResult -Value $we
            if ($parsedWe.Success) {
                if ($parsedWe.Date) { $warrantyEnd = $parsedWe.Date }
            } else {
                $warnings.Add("Row ${rowNum}: WarrantyEndDate - $($parsedWe.Message) Using +3 years default.")
            }

            $assetStatus      = Get-RowValue -Row $row -ColumnName 'AssetStatus'
            $expectedLifetime = Get-RowValue -Row $row -ColumnName 'ExpectedLifetime'
            $eolStr           = Get-RowValue -Row $row -ColumnName 'EndOfLifeDate'
            $purch  = Get-RowValue -Row $row -ColumnName 'PurchaseDate'
            $amount = Get-RowValue -Row $row -ColumnName 'PurchaseAmount'

            $parsedPurch = ConvertTo-OptionalDateParseResult -Value $purch
            if (-not $parsedPurch.Success) {
                $warnings.Add("Row ${rowNum}: PurchaseDate - $($parsedPurch.Message) Value skipped.")
            }
            $parsedEol = ConvertTo-OptionalDateParseResult -Value $eolStr
            if (-not $parsedEol.Success) {
                $warnings.Add("Row ${rowNum}: EndOfLifeDate - $($parsedEol.Message) Value skipped.")
            }
            $parsedAmount = ConvertTo-OptionalIntAmountParseResult -Value $amount
            if (-not $parsedAmount.Success) {
                $warnings.Add("Row ${rowNum}: PurchaseAmount - $($parsedAmount.Message) Value skipped.")
            }

            if ($isStaged) {
                $purchDateMs = 0
                $eolDateMs   = 0
                if ($parsedPurch.Date) { $purchDateMs = (ConvertTo-UnixMilliseconds -Date $parsedPurch.Date) }
                if ($parsedEol.Date)    { $eolDateMs   = (ConvertTo-UnixMilliseconds -Date $parsedEol.Date) }
                $stagedBody = Build-StagedDeviceBody `
                    -Name $displayName `
                    -OrgId $orgId `
                    -LocationId $locId `
                    -RoleId $roleId `
                    -WarrantyStart $warrantyStart `
                    -WarrantyEnd $warrantyEnd `
                    -ItamAssetStatus $assetStatus `
                    -ItamAssetExpectedLifetime $expectedLifetime `
                    -ItamAssetSerialNumber $serial `
                    -ItamAssetPurchaseDate $purchDateMs `
                    -ItamAssetPurchaseAmount $parsedAmount.Amount `
                    -ItamAssetEndOfLifeDate $eolDateMs
                $result = Invoke-NinjaApi -Method POST -Endpoint 'staged-device' -Body $stagedBody
                $nodeId = $result.nodeId
                if (-not $nodeId -and $result.id) { $nodeId = $result.id }
                if (-not $nodeId) { throw "API did not return a nodeId or id." }
                $script:ImportedDevices.Add([PSCustomObject]@{ Id = $nodeId; Name = $displayName; Role = $roleName })
                $lbImportResults.Items.Add("ID: $nodeId | $displayName (Staged $roleName)") | Out-Null
                $created++
            } else {
                $body = Build-UnmanagedDeviceBody `
                    -DisplayName $displayName `
                    -RoleId $roleId `
                    -OrgId $orgId `
                    -LocationId $locId `
                    -WarrantyStart $warrantyStart `
                    -WarrantyEnd $warrantyEnd `
                    -Serial $serial

                $result = Invoke-NinjaApi -Method POST -Endpoint 'itam/unmanaged-device' -Body $body
                $nodeId = $result.nodeId
                if (-not $nodeId) { throw "API did not return a nodeId." }

                if ($make -or $model -or $purch -or $amount -or $serial `
                    -or $assetStatus -or $expectedLifetime -or $eolStr) {
                    $cf = Build-AssetCustomFieldsBody `
                        -Make $make `
                        -Model $model `
                        -Serial $serial `
                        -AssetStatus $assetStatus `
                        -ExpectedLifetime $expectedLifetime `
                        -PurchaseDate $parsedPurch.Date `
                        -EndOfLifeDate $parsedEol.Date `
                        -PurchaseAmount $parsedAmount.Amount

                    if ($cf.Count -gt 0) {
                        try {
                            Invoke-NinjaApi -Method PATCH `
                                -Endpoint "device/$nodeId/custom-fields" `
                                -Body $cf | Out-Null
                        } catch {
                            $warnings.Add("Row ${rowNum}: Created device $nodeId but custom-fields update failed: $($_.Exception.Message)")
                        }
                    }
                }

                $script:ImportedDevices.Add([PSCustomObject]@{
                    Id   = $nodeId
                    Name = $displayName
                    Role = $roleName
                })
                $lbImportResults.Items.Add(
                    "ID: $nodeId | $displayName ($roleName)") | Out-Null
                $created++
            }
        } catch {
            $failed++
            $errors.Add("Row ${rowNum}: $($_.Exception.Message)")
        }
    }

    $lblImportCount.Text = $script:ImportedDevices.Count.ToString()
    $summary = "Import complete. Created: $created, Failed: $failed."
    if ($warnings.Count -gt 0) {
        $summary += " Warnings: $($warnings.Count)."
    }
    if ($errors.Count -gt 0 -or $warnings.Count -gt 0) {
        $detail = $summary
        if ($errors.Count -gt 0) {
            $detail += "`r`n`r`nErrors:`r`n" + ($errors -join "`r`n")
        }
        if ($warnings.Count -gt 0) {
            $detail += "`r`n`r`nWarnings:`r`n" + ($warnings -join "`r`n")
        }
        [System.Windows.MessageBox]::Show(
            $detail, 'Import Results', 'OK', 'Warning') | Out-Null
    } else {
        [System.Media.SystemSounds]::Asterisk.Play()
    }
    $lblStatus.Text = $summary
    $btnImportCsv.IsEnabled = $true
})

$btnManualAdd.Add_Click({
    if (-not (Test-SignedIn)) { return }

    $name     = $tbManualName.Text.Trim()
    $roleName = $cbManualRole.SelectedItem -as [string]
    $orgName  = $cbManualOrg.SelectedItem -as [string]
    $locName  = $cbManualLoc.SelectedItem -as [string]

    if ([string]::IsNullOrWhiteSpace($roleName) -or
        [string]::IsNullOrWhiteSpace($orgName) -or
        [string]::IsNullOrWhiteSpace($locName)) {
        $lblStatus.Text = 'Role, Organization, and Location are required.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }

    $isStaged = $rbManualStaged.IsChecked -eq $true
    $roleSource = if ($isStaged) { $script:StagedRoleCache } else { $script:RoleCache }
    $make  = $tbManualMake.Text.Trim()
    $model = $tbManualModel.Text.Trim()
    $displayName = $name
    if ([string]::IsNullOrWhiteSpace($displayName)) {
        $displayName = if ($make -and $model) { "$make $model" }
                       else { if ($isStaged) { "Staged $roleName" } else { "Unmanaged $roleName" } }
    }

    $orgMatch = $script:OrgCache |
        Where-Object { $_.name -eq $orgName } | Select-Object -First 1
    $locMatch = $script:LocationCache | Where-Object {
        ($_.name -eq $locName) -and
        (($_.organizationID -eq $orgMatch.id) -or ($_.organizationId -eq $orgMatch.id))
    } | Select-Object -First 1
    $roleMatch = $roleSource |
        Where-Object { $_.name -eq $roleName } | Select-Object -First 1

    if (-not $orgMatch -or -not $locMatch -or -not $roleMatch) {
        $lblStatus.Text = 'Could not resolve organization, location, or role.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }

    $serial = $tbManualSerial.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($serial)) { $serial = $null }

    $warrantyStart = Get-Date
    $warrantyEnd   = (Get-Date).AddYears(3)
    $ws = $tbManualWarrantyStart.Text.Trim()
    $we = $tbManualWarrantyEnd.Text.Trim()
    $parsedWs = ConvertTo-OptionalDateParseResult -Value $ws
    if (-not $parsedWs.Success) {
        $lblStatus.Text = "Invalid Warranty Start date. $($parsedWs.Message)"
        [System.Media.SystemSounds]::Hand.Play()
        return
    }
    if ($parsedWs.Date) { $warrantyStart = $parsedWs.Date }
    $parsedWe = ConvertTo-OptionalDateParseResult -Value $we
    if (-not $parsedWe.Success) {
        $lblStatus.Text = "Invalid Warranty End date. $($parsedWe.Message)"
        [System.Media.SystemSounds]::Hand.Play()
        return
    }
    if ($parsedWe.Date) { $warrantyEnd = $parsedWe.Date }

    $assetStatus      = $tbManualAssetStatus.Text.Trim()
    $expectedLifetime = $cbManualExpLifetime.SelectedItem -as [string]
    $eolStr           = $tbManualEolDate.Text.Trim()
    $purch            = $tbManualPurchDate.Text.Trim()
    $amount           = $tbManualAmount.Text.Trim()

    $parsedPurch = ConvertTo-OptionalDateParseResult -Value $purch
    if (-not $parsedPurch.Success) {
        $lblStatus.Text = "Invalid Purchase Date. $($parsedPurch.Message)"
        [System.Media.SystemSounds]::Hand.Play()
        return
    }
    $parsedEol = ConvertTo-OptionalDateParseResult -Value $eolStr
    if (-not $parsedEol.Success) {
        $lblStatus.Text = "Invalid End of Life Date. $($parsedEol.Message)"
        [System.Media.SystemSounds]::Hand.Play()
        return
    }
    $parsedAmount = ConvertTo-OptionalIntAmountParseResult -Value $amount
    if (-not $parsedAmount.Success) {
        $lblStatus.Text = "Invalid Purchase Amount. $($parsedAmount.Message)"
        [System.Media.SystemSounds]::Hand.Play()
        return
    }

    $btnManualAdd.IsEnabled = $false
    $lblStatus.Text = "Creating device '$displayName'..."
    Push-UIUpdate

    try {
        if ($isStaged) {
            $purchDateMs = 0
            $eolDateMs   = 0
            if ($parsedPurch.Date) { $purchDateMs = (ConvertTo-UnixMilliseconds -Date $parsedPurch.Date) }
            if ($parsedEol.Date)    { $eolDateMs   = (ConvertTo-UnixMilliseconds -Date $parsedEol.Date) }
            $stagedBody = Build-StagedDeviceBody `
                -Name $displayName `
                -OrgId $orgMatch.id `
                -LocationId $locMatch.id `
                -RoleId $roleMatch.id `
                -WarrantyStart $warrantyStart `
                -WarrantyEnd $warrantyEnd `
                -ItamAssetStatus $assetStatus `
                -ItamAssetExpectedLifetime $expectedLifetime `
                -ItamAssetSerialNumber $serial `
                -ItamAssetPurchaseDate $purchDateMs `
                -ItamAssetPurchaseAmount $parsedAmount.Amount `
                -ItamAssetEndOfLifeDate $eolDateMs
            $result = Invoke-NinjaApi -Method POST -Endpoint 'staged-device' -Body $stagedBody
            $nodeId = $result.nodeId
            if (-not $nodeId -and $result.id) { $nodeId = $result.id }
            if (-not $nodeId) { throw "API did not return a nodeId or id." }
            $script:ImportedDevices.Add([PSCustomObject]@{ Id = $nodeId; Name = $displayName; Role = $roleName })
            $lbImportResults.Items.Add("ID: $nodeId | $displayName (Staged $roleName)") | Out-Null
            $lblImportCount.Text = $script:ImportedDevices.Count.ToString()
            $lblStatus.Text = "Created staged device '$displayName' (ID: $nodeId)."
        } else {
            $body = Build-UnmanagedDeviceBody `
                -DisplayName $displayName `
                -RoleId $roleMatch.id `
                -OrgId $orgMatch.id `
                -LocationId $locMatch.id `
                -WarrantyStart $warrantyStart `
                -WarrantyEnd $warrantyEnd `
                -Serial $serial
            $result = Invoke-NinjaApi -Method POST `
                -Endpoint 'itam/unmanaged-device' -Body $body
            $nodeId = $result.nodeId
            if (-not $nodeId) { throw "API did not return a nodeId." }

            $customFieldsWarning = $null
            if ($make -or $model -or $purch -or $amount -or $serial `
                -or $assetStatus -or $expectedLifetime -or $eolStr) {
                $cf = Build-AssetCustomFieldsBody `
                    -Make $make `
                    -Model $model `
                    -Serial $serial `
                    -AssetStatus $assetStatus `
                    -ExpectedLifetime $expectedLifetime `
                    -PurchaseDate $parsedPurch.Date `
                    -EndOfLifeDate $parsedEol.Date `
                    -PurchaseAmount $parsedAmount.Amount
                if ($cf.Count -gt 0) {
                    try {
                        Invoke-NinjaApi -Method PATCH `
                            -Endpoint "device/$nodeId/custom-fields" `
                            -Body $cf | Out-Null
                    } catch {
                        $customFieldsWarning = " Custom-fields update warning: $($_.Exception.Message)"
                    }
                }
            }

            $script:ImportedDevices.Add([PSCustomObject]@{
                Id   = $nodeId
                Name = $displayName
                Role = $roleName
            })
            $lbImportResults.Items.Add(
                "ID: $nodeId | $displayName ($roleName)") | Out-Null
            $lblImportCount.Text = $script:ImportedDevices.Count.ToString()
            $lblStatus.Text = "Created '$displayName' (ID: $nodeId).$customFieldsWarning"
        }
        [System.Media.SystemSounds]::Asterisk.Play()

        $tbManualName.Clear()
        $tbManualSerial.Clear()
        $tbManualMake.Clear()
        $tbManualModel.Clear()
        $tbManualPurchDate.Clear()
        $tbManualAmount.Clear()
        $tbManualWarrantyStart.Clear()
        $tbManualWarrantyEnd.Clear()
        $tbManualAssetStatus.Clear()
        $cbManualExpLifetime.SelectedIndex = -1
        $tbManualEolDate.Clear()
    } catch {
        $lblStatus.Text = "Failed to create device: $($_.Exception.Message)"
        [System.Media.SystemSounds]::Hand.Play()
    }
    $btnManualAdd.IsEnabled = $true
})
#endregion

#region Tab 2: Generate QR Codes

$btnQrAddDevice.Add_Click({
    $idText = $tbQrDeviceId.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($idText) -or $idText -notmatch '^\d+$') {
        $lblStatus.Text = 'Enter a valid numeric device ID.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }
    $devId = [int]$idText

    foreach ($existing in $lbQrDevices.Items) {
        if (($existing -as [string]) -match "^ID:\s*$devId\s") {
            $lblStatus.Text = "Device $devId is already in the list."
            return
        }
    }

    $devName = "Device $devId"
    if (Test-TokenValid -or (Test-RefreshTokenPresent)) {
        try {
            $info = Get-DeviceInfo -DeviceId $devId
            $devName = $info.Name
        } catch {
            $lblStatus.Text = "Device $devId added, but name lookup failed: $($_.Exception.Message)"
        }
    }

    $lbQrDevices.Items.Add("ID: $devId | $devName") | Out-Null
    $tbQrDeviceId.Clear()
    $lblStatus.Text = "Added device $devId to QR generation list."
})

$btnQrRefreshImport.Add_Click({
    if ($script:ImportedDevices.Count -eq 0) {
        $lblStatus.Text = 'No devices imported yet. Use the Import Equipment tab first.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }
    $added = 0
    foreach ($dev in $script:ImportedDevices) {
        $alreadyExists = $false
        foreach ($existing in $lbQrDevices.Items) {
            if (($existing -as [string]) -match "^ID:\s*$($dev.Id)\s") {
                $alreadyExists = $true
                break
            }
        }
        if (-not $alreadyExists) {
            $lbQrDevices.Items.Add("ID: $($dev.Id) | $($dev.Name)") | Out-Null
            $added++
        }
    }
    $lblStatus.Text = "Added $added device(s) from import. Total: $($lbQrDevices.Items.Count) in list."
})

$btnQrRemoveDevice.Add_Click({
    if ($lbQrDevices.SelectedIndex -ge 0) {
        $lbQrDevices.Items.RemoveAt($lbQrDevices.SelectedIndex)
    }
})

$btnQrBrowseDir.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = 'Select output directory for QR code images'
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $tbQrOutputDir.Text = $fbd.SelectedPath
    }
})

$btnQrGenerate.Add_Click({
    if ($lbQrDevices.Items.Count -eq 0) {
        $lblStatus.Text = 'No devices in the list. Add device IDs first.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }
    $outputDir = $tbQrOutputDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outputDir)) {
        $lblStatus.Text = 'Specify an output directory for QR code images.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }
    if (-not [System.IO.Path]::IsPathRooted($outputDir)) {
        $outputDir = Join-Path -Path (Get-Location).Path -ChildPath $outputDir
    }
    try {
        if (-not (Test-Path -LiteralPath $outputDir)) {
            New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        }
        if (-not (Test-Path -LiteralPath $outputDir -PathType Container)) {
            throw "Output directory is not accessible: $outputDir"
        }
    } catch {
        $lblStatus.Text = "Failed to prepare output directory: $($_.Exception.Message)"
        [System.Media.SystemSounds]::Hand.Play()
        return
    }

    $sizeText = $cbQrSize.SelectedItem -as [string]
    $size = if ($sizeText -match '^(\d+)') { [int]$Matches[1] } else { 200 }
    $baseUrl = if ($script:NinjaBaseUrl) { $script:NinjaBaseUrl }
               else { Resolve-BaseUrl -Instance $tbInstance.Text }

    $btnQrGenerate.IsEnabled = $false
    $lbQrResults.Items.Clear()
    $script:GeneratedQRFiles.Clear()
    $generated = 0
    $total = $lbQrDevices.Items.Count

    for ($i = 0; $i -lt $total; $i++) {
        $item = $lbQrDevices.Items[$i] -as [string]
        if ($item -match '^ID:\s*(\d+)') {
            $devId = [int]$Matches[1]
        } else { continue }

        $lblStatus.Text = "Generating QR code $($i + 1) of $total (Device $devId)..."
        Push-UIUpdate

        $dashUrl = "$baseUrl/#/deviceDashboard/$devId/overview"
        $encodedUrl = [uri]::EscapeDataString($dashUrl)
        $qrApiUrl = "https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=$encodedUrl&format=png"
        $outPath = Join-Path $outputDir "Device_$devId.png"

        try {
            Invoke-WebRequest -Uri $qrApiUrl -OutFile $outPath -UseBasicParsing -ErrorAction Stop
            $fileInfo = Get-Item -LiteralPath $outPath -ErrorAction Stop
            if ($fileInfo.Length -le 0) {
                throw "Downloaded file is empty for Device $devId."
            }
            $fileBytes = [System.IO.File]::ReadAllBytes($outPath)
            if ($fileBytes.Length -lt 8 -or
                $fileBytes[0] -ne 137 -or $fileBytes[1] -ne 80 -or
                $fileBytes[2] -ne 78 -or $fileBytes[3] -ne 71 -or
                $fileBytes[4] -ne 13 -or $fileBytes[5] -ne 10 -or
                $fileBytes[6] -ne 26 -or $fileBytes[7] -ne 10) {
                throw "Downloaded file is not a valid PNG for Device $devId."
            }
            $script:GeneratedQRFiles.Add([PSCustomObject]@{
                DeviceId = $devId; Path = $outPath
            })
            $lbQrResults.Items.Add("OK: Device_$devId.png") | Out-Null
            $generated++
        } catch {
            $lbQrResults.Items.Add(
                "FAILED: Device $devId - $($_.Exception.Message)") | Out-Null
        }
    }

    $script:QROutputDirectory = $outputDir
    $lblQrGenCount.Text = $generated.ToString()
    $lblStatus.Text = "Generated $generated of $total QR code(s) in: $outputDir"
    if ($generated -gt 0) { [System.Media.SystemSounds]::Asterisk.Play() }
    $btnQrGenerate.IsEnabled = $true
})
#endregion

#region Tab 3: Upload QR Codes

$btnUploadBrowseDir.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = 'Select directory containing Device_*.png files'
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $tbUploadDir.Text = $fbd.SelectedPath
    }
})

$btnUploadScan.Add_Click({
    $dir = $tbUploadDir.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($dir)) {
        $lblStatus.Text = 'Enter a directory path to scan for Device_*.png files.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }
    if (-not [System.IO.Path]::IsPathRooted($dir)) {
        $dir = Join-Path -Path (Get-Location).Path -ChildPath $dir
    }
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
        $lblStatus.Text = "Directory not found: $dir"
        [System.Media.SystemSounds]::Hand.Play()
        return
    }

    $lbUploadFiles.Items.Clear()
    $script:UploadFileMap.Clear()
    $files = Get-ChildItem -Path $dir -Filter 'Device_*.png' `
        -File -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        if ($f.Name -match '^Device_(\d+)\.png$') {
            $devId = [int]$Matches[1]
            $script:UploadFileMap.Add([PSCustomObject]@{
                File = $f; DeviceId = $devId
            })
            $lbUploadFiles.Items.Add(
                "Device $devId <- $($f.Name)") | Out-Null
        }
    }
    $lblUploadCount.Text = "$($script:UploadFileMap.Count) file(s) found"
    if ($script:UploadFileMap.Count -eq 0) {
        $lblStatus.Text = "No Device_*.png files found in: $dir"
    } else {
        $lblStatus.Text = "Found $($script:UploadFileMap.Count) QR image(s). Click Upload All to upload."
    }
})

$btnUpload.Add_Click({
    if (-not (Test-SignedIn)) { return }
    if ($script:UploadFileMap.Count -eq 0) {
        $lblStatus.Text = 'No files to upload. Scan a directory first.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }

    if (-not (Test-TokenValid)) {
        try { Invoke-TokenRefresh } catch {
            $lblStatus.Text = "Token refresh failed: $($_.Exception.Message)"
            return
        }
    }

    $btnUpload.IsEnabled = $false
    $description = $tbUploadDesc.Text.Trim()
    $replace = $chkUploadReplace.IsChecked
    $uploaded = 0
    $skipped  = 0
    $failedCount = 0
    $warnings = [System.Collections.Generic.List[string]]::new()
    $total = $script:UploadFileMap.Count

    for ($i = 0; $i -lt $total; $i++) {
        $item = $script:UploadFileMap[$i]
        $deviceId = $item.DeviceId
        $fileInfo = $item.File
        $lblStatus.Text = "Uploading $($i + 1) of $total (Device $deviceId)..."
        Push-UIUpdate

        $existingIds = [System.Collections.Generic.List[int]]::new()
        try {
            $listResp = Invoke-NinjaApi `
                -Endpoint "related-items/with-entity/NODE/$deviceId"
            $riItems = if ($listResp -is [Array]) { @($listResp) }
                       elseif ($listResp.PSObject.Properties['data']) { @($listResp.data) }
                       elseif ($listResp.PSObject.Properties['items']) { @($listResp.items) }
                       else { @($listResp) }
            $targetName = [System.IO.Path]::GetFileNameWithoutExtension($fileInfo.Name)
            foreach ($ri in $riItems) {
                if ($null -eq $ri.id) { continue }
                if ($ri.relEntityType -ne 'ATTACHMENT') { continue }
                $meta = $null
                if ($ri.value -and $ri.value.PSObject.Properties['metadata']) { $meta = $ri.value.metadata }
                if (-not $meta) { continue }
                $metaName = if ($meta.PSObject.Properties['name']) { $meta.name } else { $null }
                if ($metaName -and [string]::Equals($metaName, $targetName, [StringComparison]::OrdinalIgnoreCase)) {
                    $existingIds.Add([int]$ri.id)
                }
            }
        } catch {
            $statusCode = 0
            $resp = $null
            $ex = $_.Exception
            if ($ex -ne $null) {
                $respProp = $ex.PSObject.Properties['Response']
                if ($respProp -and $respProp.Value) {
                    $resp = $respProp.Value
                }
            }
            if ($resp -ne $null) {
                try {
                    $statusCode = [int]$resp.StatusCode
                } catch {
                    $statusCode = 0
                }
            }
            if ($statusCode -ne 404) {
                $warnings.Add("Device ${deviceId}: Failed to check existing attachments. $($_.Exception.Message)")
            }
        }

        if ($existingIds.Count -gt 0 -and -not $replace) {
            $lbUploadFiles.Items.RemoveAt($i)
            $lbUploadFiles.Items.Insert($i,
                "SKIPPED: Device $deviceId (already exists)")
            $skipped++
            continue
        }

        if ($existingIds.Count -gt 0 -and $replace) {
            foreach ($rid in $existingIds) {
                try {
                    Invoke-NinjaApi -Method DELETE `
                        -Endpoint "related-items/$rid" | Out-Null
                } catch {
                    $warnings.Add("Device ${deviceId}: Failed to remove existing related-item $rid. $($_.Exception.Message)")
                }
            }
        }

        try {
            $bearerToken = Get-ValidBearerToken
            $boundary = [System.Guid]::NewGuid().ToString()
            $fileBytes = [System.IO.File]::ReadAllBytes($fileInfo.FullName)
            $enc = [System.Text.Encoding]::UTF8
            $bodyParts = [System.Collections.Generic.List[byte]]::new()

            $preamble = "--$boundary`r`nContent-Disposition: form-data; name=`"description`"`r`n`r`n$description`r`n"
            $bodyParts.AddRange([byte[]]$enc.GetBytes($preamble))
            $filePartHeaders = "--$boundary`r`nContent-Disposition: form-data; name=`"file`"; filename=`"$($fileInfo.Name)`"`r`nContent-Type: image/png`r`n`r`n"
            $bodyParts.AddRange([byte[]]$enc.GetBytes($filePartHeaders))
            $bodyParts.AddRange([byte[]]$fileBytes)
            $closing = "`r`n--$boundary--`r`n"
            $bodyParts.AddRange([byte[]]$enc.GetBytes($closing))

            $bodyBytes = $bodyParts.ToArray()
            $uploadUri = "$($script:NinjaBaseUrl)/api/v2/related-items/entity/NODE/$deviceId/attachment"
            $contentType = "multipart/form-data; boundary=`"$boundary`""

            $uploadDone = $false
            for ($uploadAttempt = 1; $uploadAttempt -le 2 -and -not $uploadDone; $uploadAttempt++) {
                try {
                    Invoke-RestMethod -Uri $uploadUri -Method POST `
                        -Headers @{ 'Authorization' = "Bearer $bearerToken" } `
                        -ContentType $contentType -Body $bodyBytes `
                        -UseBasicParsing -ErrorAction Stop | Out-Null
                    $uploadDone = $true
                } catch {
                    $uploadStatus = 0
                    if ($_.Exception.Response) { $uploadStatus = [int]$_.Exception.Response.StatusCode }
                    if ($uploadAttempt -eq 1 -and $uploadStatus -eq 401 -and (Test-RefreshTokenPresent)) {
                        $script:TokenExpiresAt = [datetime]::MinValue
                        Invoke-TokenRefresh
                        $bearerToken = ConvertFrom-SecureToken $script:AccessToken
                    } else { throw }
                }
            }

            $lbUploadFiles.Items.RemoveAt($i)
            $lbUploadFiles.Items.Insert($i,
                "OK: Device $deviceId <- $($fileInfo.Name)")
            $uploaded++
        } catch {
            $lbUploadFiles.Items.RemoveAt($i)
            $lbUploadFiles.Items.Insert($i,
                "FAILED: Device $deviceId - $($_.Exception.Message)")
            $failedCount++
        }
    }

    $summary = "Upload complete. Uploaded: $uploaded, Skipped: $skipped, Failed: $failedCount."
    if ($warnings.Count -gt 0) {
        $summary += " Warnings: $($warnings.Count)."
    }
    $lblStatus.Text = $summary
    if ($warnings.Count -gt 0) {
        $detail = $summary + "`r`n`r`nWarnings:`r`n" + ($warnings -join "`r`n")
        [System.Windows.MessageBox]::Show(
            $detail, 'Upload Warnings', 'OK', 'Warning') | Out-Null
    }
    if ($uploaded -gt 0) { [System.Media.SystemSounds]::Asterisk.Play() }
    $btnUpload.IsEnabled = $true
})
#endregion

#region Tab 4: Scan & Assign

$tbScanInput.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -ne 'Return') { return }
    $e.Handled = $true
    $raw = $tbScanInput.Text
    $tbScanInput.Clear()

    if ([string]::IsNullOrWhiteSpace($raw)) { return }
    if (-not (Test-SignedIn)) { return }

    $qr = Get-QRData -Text $raw
    if (-not $qr) {
        $lblStatus.Text = 'Unrecognized QR code. Expected a NinjaOne userDashboard or deviceDashboard URL.'
        [System.Media.SystemSounds]::Hand.Play()
        return
    }

    if ($qr.Type -eq 'user') {
        $lblStatus.Text = "Scanned user ID: $($qr.Id). Looking up..."
        Push-UIUpdate
        try {
            $userResult = Find-UserById -UserId $qr.Id
        } catch {
            $lblStatus.Text = "API error looking up user: $($_.Exception.Message)"
            [System.Media.SystemSounds]::Hand.Play()
            return
        }
        if (-not $userResult) {
            $lblStatus.Text = "User ID $($qr.Id) not found in NinjaOne users or contacts."
            [System.Media.SystemSounds]::Hand.Play()
            return
        }
        $script:ScanUserInfo = $userResult
        $display = $userResult.Name
        if ($userResult.Email) { $display += "  ($($userResult.Email))" }
        $display += "  |  UID: $($userResult.Uid)"
        $lblScanUserInfo.Text = $display
        $lblScanUserInfo.Foreground = [System.Windows.Media.Brushes]::Black
        $btnScanClearUser.Visibility = 'Visible'
        [System.Media.SystemSounds]::Asterisk.Play()
        Update-ScanState
    }
    elseif ($qr.Type -eq 'device') {
        if ($null -eq $script:ScanUserInfo) {
            $lblStatus.Text = 'Scan a user QR code first (Step 1) before scanning devices.'
            [System.Media.SystemSounds]::Hand.Play()
            return
        }
        if ($script:ScanDevices | Where-Object { $_.Id -eq $qr.Id }) {
            $lblStatus.Text = "Device $($qr.Id) is already in the list."
            [System.Media.SystemSounds]::Exclamation.Play()
            return
        }
        $lblStatus.Text = "Scanned device ID: $($qr.Id). Looking up..."
        Push-UIUpdate
        try {
            $deviceResult = Get-DeviceInfo -DeviceId $qr.Id
        } catch {
            $requestUrl = "$($script:NinjaBaseUrl)/api/v2/device/$($qr.Id)"
            $errMsg = $_.Exception.Message
            $innerMsg = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { '' }
            $statusCode = ''
            if ($_.Exception -is [System.Net.WebException] -and $_.Exception.Response) {
                try { $statusCode = " HTTP status: $([int]$_.Exception.Response.StatusCode)" } catch { Write-Verbose "Unable to parse HTTP status code: $($_.Exception.Message)" }
            }
            $detail = "Device $($qr.Id) lookup failed.`r`n`r`nRequest URL:`r`n$requestUrl`r`n`r`nError:$statusCode`r`n$errMsg"
            if ($innerMsg) { $detail += "`r`n`r`nInner: $innerMsg" }
            $detail += "`r`n`r`nIf this is a newly imported or staged device, it may need to be approved first or take a moment to appear in the device list. Assignment requires the device to exist in NinjaOne."
            $lblStatus.Text = "Device $($qr.Id) not found or API error. See details."
            [System.Windows.MessageBox]::Show($detail, 'Device lookup failed', 'OK', 'Warning') | Out-Null
            [System.Media.SystemSounds]::Hand.Play()
            return
        }
        $script:ScanDevices.Add($deviceResult)
        $lbScanDevices.Items.Add(
            "ID: $($deviceResult.Id)  |  $($deviceResult.Name)") | Out-Null
        [System.Media.SystemSounds]::Asterisk.Play()
        Update-ScanState
    }

    $tbScanInput.Focus()
})

$btnScanClearUser.Add_Click({
    $script:ScanUserInfo = $null
    $lblScanUserInfo.Text = 'No user scanned. Scan a user QR code to begin.'
    $lblScanUserInfo.Foreground = [System.Windows.Media.Brushes]::Gray
    $btnScanClearUser.Visibility = 'Collapsed'
    Update-ScanState
    $tbScanInput.Focus()
})

$lbScanDevices.Add_SelectionChanged({
    $btnScanRemoveDevice.IsEnabled = ($lbScanDevices.SelectedIndex -ge 0)
})

$btnScanRemoveDevice.Add_Click({
    $idx = $lbScanDevices.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:ScanDevices.Count) {
        $script:ScanDevices.RemoveAt($idx)
        $lbScanDevices.Items.RemoveAt($idx)
        Update-ScanState
    }
    $tbScanInput.Focus()
})

$btnScanAssign.Add_Click({
    if ($null -eq $script:ScanUserInfo -or $script:ScanDevices.Count -eq 0) {
        return
    }

    $ownerUid = $script:ScanUserInfo.Uid
    $total    = $script:ScanDevices.Count
    $success  = 0
    $failed   = 0
    $errors   = [System.Collections.Generic.List[string]]::new()

    $btnScanAssign.IsEnabled = $false
    $btnScanReset.IsEnabled  = $false

    for ($i = 0; $i -lt $total; $i++) {
        $dev = $script:ScanDevices[$i]
        $lblStatus.Text = "Assigning $($dev.Name) ($($i + 1)/$total)..."
        Push-UIUpdate

        try {
            Set-NinjaDeviceOwner -DeviceId $dev.Id -OwnerUid $ownerUid
            $success++
        } catch {
            $failed++
            $errors.Add("$($dev.Name) (ID $($dev.Id)): $($_.Exception.Message)")
        }
    }

    $summary = "Done. $success of $total device(s) assigned to $($script:ScanUserInfo.Name)."
    if ($failed -gt 0) {
        $summary += " $failed failed."
        $detail = $summary + "`r`n`r`nErrors:`r`n" + ($errors -join "`r`n")
        [System.Windows.MessageBox]::Show(
            $detail, 'Assignment Results', 'OK', 'Warning') | Out-Null
    } else {
        [System.Media.SystemSounds]::Asterisk.Play()
    }

    $lblStatus.Text = $summary
    $btnScanAssign.IsEnabled = $true
    $btnScanReset.IsEnabled  = $true
    $tbScanInput.Focus()
})

$btnScanReset.Add_Click({
    Reset-ScanAll
})
#endregion

#region Cross-Tab Wiring
$tabControl.Add_SelectionChanged({
    $tab = $tabControl.SelectedItem
    if ($tab -eq $tabQrGen -and $script:ImportedDevices.Count -gt 0 `
        -and $lbQrDevices.Items.Count -eq 0) {
        $lblStatus.Text = "Imported devices available. Click 'From Import' to load them."
    }
    elseif ($tab -eq $tabUpload -and $script:QROutputDirectory `
        -and -not $tbUploadDir.Text) {
        $tbUploadDir.Text = $script:QROutputDirectory
        $lblStatus.Text = 'QR output directory pre-filled from Generate tab. Click Scan Directory to find files.'
    }
    elseif ($tab -eq $tabScan) {
        if (Test-TokenValid -or (Test-RefreshTokenPresent)) {
            $tbScanInput.Focus()
        }
    }
})
#endregion

#region Window Events
$window.Add_ContentRendered({
    if (Test-TokenValid -or (Test-RefreshTokenPresent)) {
        $tbScanInput.Focus()
    }
})

$window.Add_Activated({
    if ((Test-TokenValid -or (Test-RefreshTokenPresent)) `
        -and -not $expSettings.IsExpanded) {
        $tab = $tabControl.SelectedItem
        if ($tab -eq $tabScan) { $tbScanInput.Focus() }
    }
})

$window.Add_Closing({
    if ($script:AuthListener -and $script:AuthListener.IsListening) {
        try { $script:AuthListener.Stop() } catch { Write-Verbose "Auth listener stop during closing failed: $($_.Exception.Message)" }
    }
    if ($script:AuthPS) {
        try { $script:AuthPS.Stop() } catch { Write-Verbose "Auth runspace stop during closing failed: $($_.Exception.Message)" }
        try { $script:AuthPS.Dispose() } catch { Write-Verbose "Auth runspace dispose during closing failed: $($_.Exception.Message)" }
    }

    try { Save-CurrentSession } catch { Write-Verbose "Session save on close failed: $($_.Exception.Message)" }

    if ($script:AccessToken) {
        $script:AccessToken.Dispose()
        $script:AccessToken = $null
    }
    if ($script:RefreshToken) {
        $script:RefreshToken.Dispose()
        $script:RefreshToken = $null
    }
    $script:TokenExpiresAt = [datetime]::MinValue
    $script:AuthVerifier   = $null
    $script:AuthState      = $null
    $script:MasterPassword = $null
    [System.GC]::Collect()
})
#endregion

#region Show Window (STA check)
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $t = New-Object Threading.Thread({ $window.ShowDialog() | Out-Null })
    $t.SetApartmentState([Threading.ApartmentState]::STA)
    $t.Start()
    $t.Join()
} else {
    $window.ShowDialog() | Out-Null
}
#endregion

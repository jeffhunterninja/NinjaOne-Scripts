#Requires -Version 5.1
[CmdletBinding()]
param()

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Web

$ErrorActionPreference = 'Stop'

$script:ValidTypes = @(
    'DROPDOWN','MULTI_SELECT','CHECKBOX','TEXT','TEXT_MULTILINE','TEXT_ENCRYPTED','NUMERIC','DECIMAL',
    'DATE','DATE_TIME','TIME','ATTACHMENT','NODE_DROPDOWN','NODE_MULTI_SELECT','CLIENT_DROPDOWN',
    'CLIENT_MULTI_SELECT','CLIENT_LOCATION_DROPDOWN','CLIENT_LOCATION_MULTI_SELECT','CLIENT_DOCUMENT_DROPDOWN',
    'CLIENT_DOCUMENT_MULTI_SELECT','EMAIL','PHONE','IP_ADDRESS','WYSIWYG','URL','MONETARY','IDENTIFIER','TOTP'
)
$script:DefinitionScopes = @('NODE','END_USER','LOCATION','DOCUMENT','ORGANIZATION')
$script:PermissionLevels = @('NONE','READ_ONLY','READ_WRITE')
$script:DefaultScope = 'NODE'
$script:DefaultType = 'TEXT'
$script:AuthRedirectUri = 'http://localhost:8888/'
$script:TokenExpiresAt = [datetime]::MinValue
$script:AccessToken = $null
$script:RefreshToken = $null
$script:MasterPassword = $null
$script:MasterPasswordVerifier = $null
$script:ConfigDir = Join-Path $env:APPDATA 'NinjaCustomFieldManager'
$script:ConfigFile = Join-Path $script:ConfigDir 'config.json'
$script:NinjaBaseUrl = ''
$script:NinjaClientId = ''
$script:ApiBasePath = 'v2'

$script:BulkRows = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$script:CsvRows  = [System.Collections.ObjectModel.ObservableCollection[object]]::new()

function Resolve-BaseUrl {
    param([Parameter(Mandatory = $true)][string]$Instance)
    $u = $Instance.Trim()
    if ($u -notmatch '^[a-zA-Z][a-zA-Z0-9+\-.]*://') { $u = "https://$u" }
    $uri = [System.Uri]$u
    if ($uri.Scheme -ne 'https') { throw "Use HTTPS instance URL. Received: $($uri.Scheme)" }
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

function Protect-String {
    param([string]$PlainText, [string]$MasterPwd)
    $salt = [byte[]]::new(32)
    $iv = [byte[]]::new(16)
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($salt)
    $rng.GetBytes($iv)
    $rng.Dispose()
    $kdf = [System.Security.Cryptography.Rfc2898DeriveBytes]::new($MasterPwd, $salt, 100000, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $key = $kdf.GetBytes(32)
    $kdf.Dispose()
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Key = $key
    $aes.IV = $iv
    $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $enc = $aes.CreateEncryptor()
    $plain = [System.Text.Encoding]::UTF8.GetBytes($PlainText)
    $cipher = $enc.TransformFinalBlock($plain, 0, $plain.Length)
    $enc.Dispose()
    $aes.Dispose()
    $combined = [byte[]]::new(48 + $cipher.Length)
    [Array]::Copy($salt, 0, $combined, 0, 32)
    [Array]::Copy($iv, 0, $combined, 32, 16)
    [Array]::Copy($cipher, 0, $combined, 48, $cipher.Length)
    [Array]::Clear($key, 0, $key.Length)
    [Array]::Clear($plain, 0, $plain.Length)
    return [Convert]::ToBase64String($combined)
}

function Unprotect-String {
    param([string]$CipherText, [string]$MasterPwd)
    $combined = [Convert]::FromBase64String($CipherText)
    $salt = [byte[]]::new(32)
    $iv = [byte[]]::new(16)
    $cipher = [byte[]]::new($combined.Length - 48)
    [Array]::Copy($combined, 0, $salt, 0, 32)
    [Array]::Copy($combined, 32, $iv, 0, 16)
    [Array]::Copy($combined, 48, $cipher, 0, $cipher.Length)
    $kdf = [System.Security.Cryptography.Rfc2898DeriveBytes]::new($MasterPwd, $salt, 100000, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $key = $kdf.GetBytes(32)
    $kdf.Dispose()
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.Key = $key
    $aes.IV = $iv
    $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $dec = $aes.CreateDecryptor()
    $plain = $dec.TransformFinalBlock($cipher, 0, $cipher.Length)
    $dec.Dispose()
    $aes.Dispose()
    $result = [System.Text.Encoding]::UTF8.GetString($plain)
    [Array]::Clear($key, 0, $key.Length)
    [Array]::Clear($plain, 0, $plain.Length)
    return $result
}

function New-MasterPasswordVerifier {
    param([string]$MasterPwd)
    $salt = [byte[]]::new(32)
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($salt)
    $kdf = [System.Security.Cryptography.Rfc2898DeriveBytes]::new($MasterPwd, $salt, 100000, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $hash = $kdf.GetBytes(32)
    $kdf.Dispose()
    $combined = [byte[]]::new(64)
    [Array]::Copy($salt, 0, $combined, 0, 32)
    [Array]::Copy($hash, 0, $combined, 32, 32)
    return [Convert]::ToBase64String($combined)
}

function Test-MasterPasswordValid {
    param([string]$MasterPwd, [string]$Verifier)
    $combined = [Convert]::FromBase64String($Verifier)
    $salt = [byte[]]::new(32)
    $storedHash = [byte[]]::new(32)
    [Array]::Copy($combined, 0, $salt, 0, 32)
    [Array]::Copy($combined, 32, $storedHash, 0, 32)
    $kdf = [System.Security.Cryptography.Rfc2898DeriveBytes]::new($MasterPwd, $salt, 100000, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
    $computed = $kdf.GetBytes(32)
    $kdf.Dispose()
    $diff = 0
    for ($i = 0; $i -lt 32; $i++) { $diff = $diff -bor ($storedHash[$i] -bxor $computed[$i]) }
    return ($diff -eq 0)
}

function Get-AppConfig {
    $defaults = [PSCustomObject]@{ NinjaInstance=''; ClientId=''; EncryptedRefreshToken=''; MasterPasswordVerifier='' }
    if (Test-Path -LiteralPath $script:ConfigFile) {
        try {
            $raw = Get-Content -LiteralPath $script:ConfigFile -Raw | ConvertFrom-Json
            foreach ($prop in $raw.PSObject.Properties) {
                if ($defaults.PSObject.Properties[$prop.Name]) { $defaults.$($prop.Name) = $prop.Value }
            }
        } catch { }
    }
    return $defaults
}

function Save-AppConfig {
    param([string]$Instance,[string]$ClientIdValue,[string]$EncryptedRefreshToken,[string]$Verifier)
    if (-not (Test-Path -LiteralPath $script:ConfigDir)) { New-Item -ItemType Directory -Path $script:ConfigDir -Force | Out-Null }
    [PSCustomObject]@{
        NinjaInstance = $Instance
        ClientId = $ClientIdValue
        EncryptedRefreshToken = $EncryptedRefreshToken
        MasterPasswordVerifier = $Verifier
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:ConfigFile -Encoding UTF8
}

function Save-CurrentSession {
    if (-not $script:MasterPassword -or -not $script:RefreshToken) { return }
    $encrypted = Protect-String -PlainText (ConvertFrom-SecureToken $script:RefreshToken) -MasterPwd $script:MasterPassword
    if (-not $script:MasterPasswordVerifier) { $script:MasterPasswordVerifier = New-MasterPasswordVerifier -MasterPwd $script:MasterPassword }
    Save-AppConfig -Instance $script:NinjaBaseUrl -ClientIdValue $script:NinjaClientId -EncryptedRefreshToken $encrypted -Verifier $script:MasterPasswordVerifier
}

function Clear-SavedSession {
    if (Test-Path -LiteralPath $script:ConfigFile) { Remove-Item -LiteralPath $script:ConfigFile -Force -ErrorAction SilentlyContinue }
    $script:MasterPassword = $null
    $script:MasterPasswordVerifier = $null
}

function ConvertTo-CamelCaseFromLabel {
    param([string]$Label)
    if ([string]::IsNullOrWhiteSpace($Label)) { return '' }
    $parts = @([regex]::Replace($Label.Trim(), '[^a-zA-Z0-9]+', ' ').Trim() -split '\s+' | Where-Object { $_ })
    if ($parts.Count -eq 0) { throw "Label cannot be converted to fieldName." }
    $result = ''
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $word = $parts[$i]
        if ($i -eq 0) { $result += $word.Substring(0,1).ToLowerInvariant() + $word.Substring(1).ToLowerInvariant() }
        else { $result += $word.Substring(0,1).ToUpperInvariant() + $word.Substring(1).ToLowerInvariant() }
    }
    return $result
}

function ConvertTo-StringArray {
    param([string]$Value,[string]$DefaultValue = '')
    $vals = @($Value -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($vals.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($DefaultValue)) { return @($DefaultValue) }
    return @($vals)
}

function New-RowObject {
    return [PSCustomObject]@{
        Include = $true
        Label = ''
        FieldName = ''
        Type = $script:DefaultType
        DefinitionScope = $script:DefaultScope
        Description = ''
        DefaultValue = ''
        TechnicianPermission = 'NONE'
        ScriptPermission = 'NONE'
        ApiPermission = 'NONE'
        DropdownValues = ''
        Validation = ''
    }
}

function Set-Status {
    param([string]$Message)
    $txtStatus.Text = $Message
}

function Show-MasterPasswordPrompt {
    param([string]$Title,[string]$Message,[switch]$IsNewPassword)
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
Title="$Title" SizeToContent="WidthAndHeight" ResizeMode="NoResize" WindowStartupLocation="CenterOwner" MinWidth="380">
  <StackPanel Margin="20">
    <TextBlock Text="$Message" Margin="0,0,0,10" TextWrapping="Wrap"/>
    <TextBlock Text="Password"/>
    <PasswordBox x:Name="pbPassword" Height="28"/>
    $(if ($IsNewPassword) { '<TextBlock Margin="0,8,0,0" Text="Confirm Password"/><PasswordBox x:Name="pbConfirm" Height="28"/>' } else { '' })
    <TextBlock x:Name="lblError" Foreground="Red" Margin="0,8,0,0" Visibility="Collapsed"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button x:Name="btnOK" Content="OK" Width="80" IsDefault="True" Margin="0,0,8,0"/>
      <Button x:Name="btnCancel" Content="Cancel" Width="80" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
</Window>
"@
    $dialog = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$xaml)))
    $dialog.Owner = $window
    $pbPassword = $dialog.FindName('pbPassword')
    $pbConfirm = $dialog.FindName('pbConfirm')
    $lblError = $dialog.FindName('lblError')
    $dialog.FindName('btnOK').Add_Click({
        $pwd = $pbPassword.Password
        if ([string]::IsNullOrWhiteSpace($pwd)) { $lblError.Text='Password required.'; $lblError.Visibility='Visible'; return }
        if ($IsNewPassword) {
            if ($pwd.Length -lt 8) { $lblError.Text='Minimum 8 characters.'; $lblError.Visibility='Visible'; return }
            if ($pwd -ne $pbConfirm.Password) { $lblError.Text='Passwords do not match.'; $lblError.Visibility='Visible'; return }
        }
        $dialog.Tag = $pwd
        $dialog.DialogResult = $true
        $dialog.Close()
    })
    if ($dialog.ShowDialog() -eq $true) { return [string]$dialog.Tag }
    return $null
}

function Update-TokensFromResponse {
    param($Response)
    if (-not $Response.access_token) { throw 'Token response missing access_token.' }
    $script:AccessToken = ConvertTo-SecureToken $Response.access_token
    if ($Response.refresh_token) { $script:RefreshToken = ConvertTo-SecureToken $Response.refresh_token }
    $exp = if ($Response.expires_in) { [int]$Response.expires_in } else { 3600 }
    if ($exp -le 0) { $exp = 3600 }
    $script:TokenExpiresAt = [datetime]::UtcNow.AddSeconds($exp - 60)
}

function Invoke-TokenRefresh {
    if (-not $script:RefreshToken) { throw 'No refresh token. Sign in required.' }
    $resp = Invoke-RestMethod -Uri "$($script:NinjaBaseUrl)/ws/oauth/token" -Method Post -UseBasicParsing -ContentType 'application/x-www-form-urlencoded' -Body @{
        grant_type    = 'refresh_token'
        refresh_token = (ConvertFrom-SecureToken $script:RefreshToken)
        client_id     = $script:NinjaClientId
    }
    Update-TokensFromResponse -Response $resp
    Save-CurrentSession
}

function Get-ValidBearerToken {
    if (-not $script:AccessToken -or [datetime]::UtcNow -ge $script:TokenExpiresAt) { Invoke-TokenRefresh }
    return ConvertFrom-SecureToken $script:AccessToken
}

function Invoke-NinjaApi {
    param([string]$Method,[string]$Endpoint,[object]$Body)
    $token = Get-ValidBearerToken
    $uri = "$($script:NinjaBaseUrl)/$($script:ApiBasePath)/$($Endpoint.TrimStart('/'))"
    $headers = @{ Authorization = "Bearer $token"; Accept = 'application/json' }
    try {
        if ($Method -eq 'GET') { return Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -UseBasicParsing }
        return Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers -Body ($Body | ConvertTo-Json -Depth 15) -ContentType 'application/json' -UseBasicParsing
    } catch {
        if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 401) {
            Invoke-TokenRefresh
            $headers.Authorization = "Bearer $(Get-ValidBearerToken)"
            if ($Method -eq 'GET') { return Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -UseBasicParsing }
            return Invoke-RestMethod -Uri $uri -Method $Method -Headers $headers -Body ($Body | ConvertTo-Json -Depth 15) -ContentType 'application/json' -UseBasicParsing
        }
        throw
    }
}

function Get-NormalizedCustomFieldLabelKey {
    param([string]$Label)
    if ([string]::IsNullOrWhiteSpace($Label)) { return '' }
    return ([regex]::Replace($Label.Trim(), '\s+', ' ')).ToLowerInvariant()
}

function Get-NinjaOneCustomFieldDefinitions {
    $pageSize = 500
    $cursor = $null
    $all = [System.Collections.ArrayList]::new()
    $safety = 0
    do {
        $safety++
        if ($safety -gt 1000) { throw 'Pagination safety limit exceeded while reading custom fields.' }
        $endpoint = "custom-fields?pageSize=$pageSize"
        if (-not [string]::IsNullOrWhiteSpace($cursor)) {
            $endpoint += '&cursorName=' + [System.Uri]::EscapeDataString($cursor)
        }
        $resp = Invoke-NinjaApi -Method 'GET' -Endpoint $endpoint -Body $null
        if ($null -eq $resp -or $null -eq $resp.results) { break }
        $results = @($resp.results)
        foreach ($item in $results) { [void]$all.Add($item) }
        $cursor = $null
        foreach ($name in @('cursorName','nextCursorName','nextCursor','cursor')) {
            if ($resp.PSObject.Properties.Name -contains $name) {
                $v = $resp.$name
                if (-not [string]::IsNullOrWhiteSpace([string]$v)) { $cursor = [string]$v; break }
            }
        }
        if ($results.Count -lt $pageSize -or [string]::IsNullOrWhiteSpace($cursor)) { break }
    } while ($true)
    return ,@($all.ToArray())
}

function New-ExistingCustomFieldLookup {
    param([object[]]$DefinitionObjects)
    $names = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $labels = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $DefinitionObjects) {
        if ($item.fieldName) { [void]$names.Add([string]$item.fieldName) }
        if ($item.label) { [void]$labels.Add((Get-NormalizedCustomFieldLabelKey -Label ([string]$item.label))) }
    }
    return @{ FieldNames = $names; LabelKeys = $labels }
}

function ConvertTo-CustomFieldApiBody {
    param([psobject]$Row)
    $label = [string]$Row.Label
    if ([string]::IsNullOrWhiteSpace($label)) { throw 'Label is required.' }
    $fieldName = if ([string]::IsNullOrWhiteSpace([string]$Row.FieldName)) { ConvertTo-CamelCaseFromLabel -Label $label } else { [string]$Row.FieldName }
    $type = ([string]$Row.Type).ToUpperInvariant()
    if ($type -notin $script:ValidTypes) { $type = 'TEXT' }
    $defScope = ConvertTo-StringArray -Value ([string]$Row.DefinitionScope) -DefaultValue 'NODE'
    $ddVals = ConvertTo-StringArray -Value ([string]$Row.DropdownValues)
    $body = @{
        label                = $label.Trim()
        fieldName            = $fieldName.Trim()
        scope                = 'NODE_ROLE'
        definitionScope      = @($defScope)
        type                 = $type
        technicianPermission = if ($Row.TechnicianPermission) { [string]$Row.TechnicianPermission } else { 'NONE' }
        scriptPermission     = if ($Row.ScriptPermission) { [string]$Row.ScriptPermission } else { 'NONE' }
        apiPermission        = if ($Row.ApiPermission) { [string]$Row.ApiPermission } else { 'NONE' }
        addToDefaultTab      = $false
    }
    if ($Row.Description) { $body.description = [string]$Row.Description }
    if ($Row.DefaultValue) { $body.defaultValue = [string]$Row.DefaultValue }
    if ($ddVals.Count -gt 0) { $body.content = @{ values = @($ddVals | ForEach-Object { @{ name = $_ } }); required = $false } }
    return $body
}

function Validate-Row {
    param([psobject]$Row)
    if ([string]::IsNullOrWhiteSpace([string]$Row.Label)) { throw 'Label is required.' }
    $derived = ConvertTo-CamelCaseFromLabel -Label ([string]$Row.Label)
    if ([string]::IsNullOrWhiteSpace([string]$Row.FieldName)) { $Row.FieldName = $derived }
    $type = ([string]$Row.Type).ToUpperInvariant()
    if ($type -notin $script:ValidTypes) { $type = 'TEXT'; $Row.Type = 'TEXT' }
    $tech = if ($Row.TechnicianPermission) { ([string]$Row.TechnicianPermission).ToUpperInvariant() } else { 'NONE' }
    $scriptPerm = if ($Row.ScriptPermission) { ([string]$Row.ScriptPermission).ToUpperInvariant() } else { 'NONE' }
    $api = if ($Row.ApiPermission) { ([string]$Row.ApiPermission).ToUpperInvariant() } else { 'NONE' }
    if ($tech -notin $script:PermissionLevels) { $tech = 'NONE' }
    if ($scriptPerm -notin $script:PermissionLevels) { $scriptPerm = 'NONE' }
    if ($api -notin $script:PermissionLevels) { $api = 'NONE' }
    $Row.TechnicianPermission = $tech
    $Row.ScriptPermission = $scriptPerm
    $Row.ApiPermission = $api
    if ([string]::IsNullOrWhiteSpace([string]$Row.DefinitionScope)) { $Row.DefinitionScope = 'NODE' }
}

function Submit-RowCollection {
    param([System.Collections.ObjectModel.ObservableCollection[object]]$Rows,[string]$SourceName)
    $selected = @($Rows | Where-Object { $_.Include })
    if ($selected.Count -eq 0) { throw "No selected rows in $SourceName." }
    $existing = Get-NinjaOneCustomFieldDefinitions
    $lookup = New-ExistingCustomFieldLookup -DefinitionObjects $existing
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $toCreate = [System.Collections.ArrayList]::new()
    $skipped = 0
    $invalid = 0
    foreach ($row in $selected) {
        try {
            Validate-Row -Row $row
            $payload = ConvertTo-CustomFieldApiBody -Row $row
            $row.Validation = ''
        } catch {
            $row.Validation = $_.Exception.Message
            $invalid++
            continue
        }
        $fn = [string]$payload.fieldName
        $lk = Get-NormalizedCustomFieldLabelKey -Label ([string]$payload.label)
        if ($seen.Contains($fn) -or $lookup.FieldNames.Contains($fn) -or $lookup.LabelKeys.Contains($lk)) {
            $row.Validation = 'Skipped (duplicate/existing)'
            $skipped++
            continue
        }
        [void]$seen.Add($fn)
        $row.Validation = 'Ready'
        [void]$toCreate.Add($payload)
    }
    $dgBulk.Items.Refresh()
    $dgCsv.Items.Refresh()
    if ($invalid -gt 0) { throw "$invalid selected row(s) are invalid. Fix Validation column and retry." }
    if ($toCreate.Count -eq 0) {
        Set-Status "No new fields to create from $SourceName. All rows duplicate existing fields."
        return
    }
    [void](Invoke-NinjaApi -Method 'POST' -Endpoint 'custom-fields/bulk' -Body @{ customFields = @($toCreate) })
    Set-Status "Created $($toCreate.Count) field(s) from $SourceName. Skipped $skipped duplicate/existing row(s)."
}

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
Title="NinjaOne Custom Field Manager" Height="800" Width="1320" WindowStartupLocation="CenterScreen">
  <Grid Margin="10">
    <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <TabControl x:Name="tabMain" Grid.Row="0">
      <TabItem Header="Authentication">
        <Grid Margin="10">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <Grid.ColumnDefinitions><ColumnDefinition Width="170"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
          <TextBlock Grid.Row="0" Grid.Column="0" Margin="0,6,8,6" Text="NinjaOne Instance"/>
          <TextBox x:Name="txtInstance" Grid.Row="0" Grid.Column="1" Margin="0,4,0,4"/>
          <TextBlock Grid.Row="1" Grid.Column="0" Margin="0,6,8,6" Text="Client ID"/>
          <TextBox x:Name="txtClientId" Grid.Row="1" Grid.Column="1" Margin="0,4,0,4"/>
          <TextBlock Grid.Row="2" Grid.Column="0" Margin="0,6,8,6" Text="Scopes"/>
          <TextBox x:Name="txtScope" Grid.Row="2" Grid.Column="1" Margin="0,4,0,4" Text="monitoring management offline_access"/>
          <StackPanel Grid.Row="3" Grid.ColumnSpan="2" Orientation="Horizontal" Margin="0,12,0,0">
            <Button x:Name="btnConnect" Width="140" Height="30" Content="Authenticate PKCE" Margin="0,0,8,0"/>
            <Button x:Name="btnUnlockSession" Width="120" Height="30" Content="Unlock Saved"/>
            <Button x:Name="btnClearSaved" Width="120" Height="30" Content="Clear Saved" Margin="8,0,0,0"/>
          </StackPanel>
        </Grid>
      </TabItem>
      <TabItem Header="CSV Import">
        <Grid Margin="10">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,8">
            <TextBox x:Name="txtCsvPath" Width="740"/>
            <Button x:Name="btnBrowseCsv" Width="120" Height="28" Content="Browse CSV" Margin="8,0,0,0"/>
            <Button x:Name="btnLoadCsv" Width="120" Height="28" Content="Load CSV" Margin="8,0,0,0"/>
          </StackPanel>
          <TextBlock Grid.Row="1" Text="Rows loaded from CSV. Toggle Include for submit." Margin="0,0,0,8"/>
          <DataGrid x:Name="dgCsv" Grid.Row="2" AutoGenerateColumns="True" CanUserAddRows="False" IsReadOnly="False"/>
          <Button x:Name="btnCreateCsv" Grid.Row="3" Width="220" Height="30" Content="Create Selected CSV Rows" HorizontalAlignment="Left" Margin="0,8,0,0"/>
        </Grid>
      </TabItem>
      <TabItem Header="Bulk Create">
        <Grid Margin="10">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <WrapPanel Grid.Row="0">
            <Button x:Name="btnAddRow" Width="120" Height="28" Content="Add Row"/>
            <Button x:Name="btnRemoveRows" Width="170" Height="28" Content="Remove Selected Rows" Margin="8,0,0,0"/>
          </WrapPanel>
          <WrapPanel Grid.Row="1" Margin="0,8,0,8">
            <TextBlock VerticalAlignment="Center" Margin="0,0,8,0" Text="Global Type"/>
            <ComboBox x:Name="cbGlobalType" Width="200" Margin="0,0,12,0"/>
            <TextBlock VerticalAlignment="Center" Margin="0,0,8,0" Text="Global Scope"/>
            <ComboBox x:Name="cbGlobalScope" Width="200" Margin="0,0,12,0"/>
            <TextBlock VerticalAlignment="Center" Margin="0,0,8,0" Text="Tech Perm"/>
            <ComboBox x:Name="cbGlobalTechPerm" Width="140" Margin="0,0,8,0"/>
            <TextBlock VerticalAlignment="Center" Margin="0,0,8,0" Text="Script Perm"/>
            <ComboBox x:Name="cbGlobalScriptPerm" Width="140" Margin="0,0,8,0"/>
            <TextBlock VerticalAlignment="Center" Margin="0,0,8,0" Text="API Perm"/>
            <ComboBox x:Name="cbGlobalApiPerm" Width="140" Margin="0,0,8,0"/>
            <Button x:Name="btnApplyGlobal" Width="180" Height="28" Content="Apply To Selected Rows"/>
          </WrapPanel>
          <DataGrid x:Name="dgBulk" Grid.Row="2" AutoGenerateColumns="True" IsReadOnly="False"/>
          <Button x:Name="btnCreateBulk" Grid.Row="3" Width="220" Height="30" Content="Create Selected Bulk Rows" HorizontalAlignment="Left" Margin="0,8,0,0"/>
        </Grid>
      </TabItem>
    </TabControl>
    <TextBlock x:Name="txtStatus" Grid.Row="1" Margin="4,8,4,0" Text="Ready." TextWrapping="Wrap"/>
  </Grid>
</Window>
"@

$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$xaml)))

$txtInstance = $window.FindName('txtInstance')
$txtClientId = $window.FindName('txtClientId')
$txtScope = $window.FindName('txtScope')
$btnConnect = $window.FindName('btnConnect')
$btnUnlockSession = $window.FindName('btnUnlockSession')
$btnClearSaved = $window.FindName('btnClearSaved')
$txtCsvPath = $window.FindName('txtCsvPath')
$btnBrowseCsv = $window.FindName('btnBrowseCsv')
$btnLoadCsv = $window.FindName('btnLoadCsv')
$btnCreateCsv = $window.FindName('btnCreateCsv')
$dgCsv = $window.FindName('dgCsv')
$btnAddRow = $window.FindName('btnAddRow')
$btnRemoveRows = $window.FindName('btnRemoveRows')
$cbGlobalType = $window.FindName('cbGlobalType')
$cbGlobalScope = $window.FindName('cbGlobalScope')
$cbGlobalTechPerm = $window.FindName('cbGlobalTechPerm')
$cbGlobalScriptPerm = $window.FindName('cbGlobalScriptPerm')
$cbGlobalApiPerm = $window.FindName('cbGlobalApiPerm')
$btnApplyGlobal = $window.FindName('btnApplyGlobal')
$dgBulk = $window.FindName('dgBulk')
$btnCreateBulk = $window.FindName('btnCreateBulk')
$txtStatus = $window.FindName('txtStatus')

$cbGlobalType.ItemsSource = $script:ValidTypes
$cbGlobalScope.ItemsSource = $script:DefinitionScopes
$cbGlobalTechPerm.ItemsSource = $script:PermissionLevels
$cbGlobalScriptPerm.ItemsSource = $script:PermissionLevels
$cbGlobalApiPerm.ItemsSource = $script:PermissionLevels
$cbGlobalType.SelectedItem = $script:DefaultType
$cbGlobalScope.SelectedItem = $script:DefaultScope
$cbGlobalTechPerm.SelectedItem = 'NONE'
$cbGlobalScriptPerm.SelectedItem = 'NONE'
$cbGlobalApiPerm.SelectedItem = 'NONE'

$dgBulk.ItemsSource = $script:BulkRows
$dgCsv.ItemsSource = $script:CsvRows
[void]$script:BulkRows.Add((New-RowObject))

$saved = Get-AppConfig
$txtInstance.Text = [string]$saved.NinjaInstance
$txtClientId.Text = [string]$saved.ClientId
$script:MasterPasswordVerifier = [string]$saved.MasterPasswordVerifier

$btnConnect.Add_Click({
    try {
        $script:NinjaBaseUrl = Resolve-BaseUrl -Instance $txtInstance.Text
        $script:NinjaClientId = $txtClientId.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($script:NinjaClientId)) { throw 'Client ID is required.' }
        $verifier = New-PkceVerifier
        $challenge = Get-PkceChallenge -Verifier $verifier
        $state = New-OAuthState
        $authUrl = "$($script:NinjaBaseUrl)/ws/oauth/authorize?response_type=code&client_id=$([System.Uri]::EscapeDataString($script:NinjaClientId))&redirect_uri=$([System.Uri]::EscapeDataString($script:AuthRedirectUri))&scope=$([System.Uri]::EscapeDataString($txtScope.Text))&state=$([System.Uri]::EscapeDataString($state))&code_challenge=$([System.Uri]::EscapeDataString($challenge))&code_challenge_method=S256"
        $listener = [System.Net.HttpListener]::new()
        $listener.Prefixes.Add($script:AuthRedirectUri)
        $listener.Start()
        Start-Process $authUrl
        Set-Status 'Waiting for browser sign-in...'
        $ctx = $listener.GetContext()
        $qs = [System.Web.HttpUtility]::ParseQueryString($ctx.Request.Url.Query)
        $code = $qs['code']
        $stateBack = $qs['state']
        $okHtml = '<html><body><h3>Authentication complete. You can close this window.</h3></body></html>'
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($okHtml)
        $ctx.Response.ContentType = 'text/html'
        $ctx.Response.OutputStream.Write($bytes,0,$bytes.Length)
        $ctx.Response.OutputStream.Close()
        $listener.Stop()
        if ($stateBack -ne $state) { throw 'OAuth state mismatch.' }
        if ([string]::IsNullOrWhiteSpace($code)) { throw 'No authorization code returned.' }
        $tokenResp = Invoke-RestMethod -Uri "$($script:NinjaBaseUrl)/ws/oauth/token" -Method Post -UseBasicParsing -ContentType 'application/x-www-form-urlencoded' -Body @{
            grant_type = 'authorization_code'
            code = $code
            redirect_uri = $script:AuthRedirectUri
            client_id = $script:NinjaClientId
            code_verifier = $verifier
        }
        Update-TokensFromResponse -Response $tokenResp
        $pwd = Show-MasterPasswordPrompt -Title 'Save Session' -Message 'Set a master password to encrypt and save your refresh token.' -IsNewPassword
        if ($pwd) {
            $script:MasterPassword = $pwd
            $script:MasterPasswordVerifier = New-MasterPasswordVerifier -MasterPwd $pwd
            Save-CurrentSession
        } else {
            Save-AppConfig -Instance $script:NinjaBaseUrl -ClientIdValue $script:NinjaClientId -EncryptedRefreshToken '' -Verifier ''
        }
        Set-Status 'Authentication successful.'
    } catch {
        Set-Status "Authentication failed: $($_.Exception.Message)"
    } finally {
        if ($listener -and $listener.IsListening) { $listener.Stop() }
    }
})

$btnUnlockSession.Add_Click({
    try {
        $cfg = Get-AppConfig
        if ([string]::IsNullOrWhiteSpace($cfg.EncryptedRefreshToken)) { throw 'No saved session found.' }
        $pwd = Show-MasterPasswordPrompt -Title 'Unlock Saved Session' -Message 'Enter your master password.'
        if (-not $pwd) { return }
        if (-not (Test-MasterPasswordValid -MasterPwd $pwd -Verifier $cfg.MasterPasswordVerifier)) { throw 'Invalid master password.' }
        $plain = Unprotect-String -CipherText $cfg.EncryptedRefreshToken -MasterPwd $pwd
        $script:MasterPassword = $pwd
        $script:MasterPasswordVerifier = $cfg.MasterPasswordVerifier
        $script:RefreshToken = ConvertTo-SecureToken $plain
        $script:NinjaBaseUrl = if ($cfg.NinjaInstance) { Resolve-BaseUrl -Instance ([string]$cfg.NinjaInstance) } else { Resolve-BaseUrl -Instance $txtInstance.Text }
        $script:NinjaClientId = if ($cfg.ClientId) { [string]$cfg.ClientId } else { $txtClientId.Text.Trim() }
        if ($cfg.NinjaInstance) { $txtInstance.Text = [string]$cfg.NinjaInstance }
        if ($cfg.ClientId) { $txtClientId.Text = [string]$cfg.ClientId }
        Invoke-TokenRefresh
        Set-Status 'Saved session unlocked and token refreshed.'
    } catch {
        Set-Status "Unlock failed: $($_.Exception.Message)"
    }
})

$btnClearSaved.Add_Click({
    Clear-SavedSession
    Set-Status 'Saved session cleared.'
})

$btnBrowseCsv.Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    if ($dlg.ShowDialog() -eq $true) { $txtCsvPath.Text = $dlg.FileName }
})

$btnLoadCsv.Add_Click({
    try {
        $path = $txtCsvPath.Text
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'Select a valid CSV file first.' }
        $rows = Import-Csv -LiteralPath $path
        $script:CsvRows.Clear()
        foreach ($r in $rows) {
            if ([string]::IsNullOrWhiteSpace([string]$r.Label)) { continue }
            $obj = New-RowObject
            $obj.Label = [string]$r.Label
            $obj.FieldName = ConvertTo-CamelCaseFromLabel -Label $obj.Label
            $obj.Type = if ($r.Type) { ([string]$r.Type).ToUpperInvariant() } else { 'TEXT' }
            if ($obj.Type -notin $script:ValidTypes) { $obj.Type = 'TEXT' }
            $obj.DefinitionScope = if ($r.DefinitionScope) { [string]$r.DefinitionScope } else { 'NODE' }
            $obj.Description = [string]$r.Description
            $obj.DefaultValue = [string]$r.DefaultValue
            $obj.TechnicianPermission = if ($r.TechnicianPermission) { [string]$r.TechnicianPermission } else { 'NONE' }
            $obj.ScriptPermission = if ($r.ScriptPermission) { [string]$r.ScriptPermission } else { 'NONE' }
            $obj.ApiPermission = if ($r.ApiPermission) { [string]$r.ApiPermission } else { 'NONE' }
            $obj.DropdownValues = [string]$r.DropdownValues
            $obj.Validation = ''
            [void]$script:CsvRows.Add($obj)
        }
        $dgCsv.Items.Refresh()
        Set-Status "Loaded $($script:CsvRows.Count) CSV row(s)."
    } catch {
        Set-Status "CSV load failed: $($_.Exception.Message)"
    }
})

$btnCreateCsv.Add_Click({
    try { Submit-RowCollection -Rows $script:CsvRows -SourceName 'CSV Import' } catch { Set-Status "CSV create failed: $($_.Exception.Message)" }
})

$btnAddRow.Add_Click({ [void]$script:BulkRows.Add((New-RowObject)) })

$btnRemoveRows.Add_Click({
    $selected = @($dgBulk.SelectedItems)
    foreach ($item in $selected) { [void]$script:BulkRows.Remove($item) }
    if ($script:BulkRows.Count -eq 0) { [void]$script:BulkRows.Add((New-RowObject)) }
})

$btnApplyGlobal.Add_Click({
    foreach ($row in @($script:BulkRows | Where-Object { $_.Include })) {
        $row.Type = [string]$cbGlobalType.SelectedItem
        $row.DefinitionScope = [string]$cbGlobalScope.SelectedItem
        $row.TechnicianPermission = [string]$cbGlobalTechPerm.SelectedItem
        $row.ScriptPermission = [string]$cbGlobalScriptPerm.SelectedItem
        $row.ApiPermission = [string]$cbGlobalApiPerm.SelectedItem
    }
    $dgBulk.Items.Refresh()
    Set-Status 'Applied global settings to selected bulk rows.'
})

$btnCreateBulk.Add_Click({
    try { Submit-RowCollection -Rows $script:BulkRows -SourceName 'Bulk Create' } catch { Set-Status "Bulk create failed: $($_.Exception.Message)" }
})

[void]$window.ShowDialog()

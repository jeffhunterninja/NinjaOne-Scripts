#Requires -Version 5.1
<#
.SYNOPSIS
  Collects end-user Quiet Hours (Do-Not-Disturb) preferences via a simple WPF popup.

.DESCRIPTION
  - Asks the user to choose quiet hours for Weekdays and Weekends (or specific days).
  - Saves preferences to C:\RMM\NinjaOne-QuietHours\quiet_hours.json.
  - If Ninja-Property-Set exists, also writes JSON to a Ninja custom field (default: quietHours).
  - Provides Test-QuietHours helper to let other scripts respect quiet windows.
  - Quiet-window logic uses the machine's local time.

.PARAMETER PreferencePath
  Full path to the JSON file where quiet hours preferences are saved. Default: C:\RMM\NinjaOne-QuietHours\quiet_hours.json

.PARAMETER NinjaCustomField
  Name of the NinjaOne custom field to write the JSON to when Ninja-Property-Set is available Can be overridden by $env:quietHoursCustomField. Default: quietHours

.EXAMPLE
  .\Set-QuietHours.ps1
  Opens the Quiet Hours survey dialog with default paths and saves to the default JSON file and Ninja custom field (if available).

.EXIT CODES
  0 = Script completed (user closed dialog). Save vs Cancel is not distinguished. This is a UI-only script.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [string]$PreferencePath = "C:\RMM\NinjaOne-QuietHours\quiet_hours.json",
  [Parameter(Mandatory = $false)]
  [string]$NinjaCustomField = $(if ($env:quietHoursCustomField) { $env:quietHoursCustomField } else { "quietHours" })
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Helpers
function Test-CommandExists {
  param([Parameter(Mandatory)] [string]$Name)
  if (Get-Command -Name $Name -ErrorAction SilentlyContinue) { return $true }
  return $false
}

# Creates the directory for the preference file if missing. If access is denied (e.g. C:\RMM), an admin should create C:\RMM\NinjaOne-QuietHours and grant Users Modify.
function Ensure-Dir { param([Parameter(Mandatory)] [string]$Path)
  $dir = [System.IO.Path]::GetDirectoryName($Path)
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
}

function Get-QuietPrefs {
  param([string]$Path)
  if (Test-Path $Path) {
    try {
      return Get-Content -Raw -Path $Path | ConvertFrom-Json
    } catch {
      Write-Verbose "Could not load quiet hours prefs: $($_.Exception.Message)"
      return $null
    }
  }
  return $null
}

function Save-QuietPrefs {
  param(
    [Parameter(Mandatory)] $Prefs,
    [Parameter(Mandatory)] [string]$Path
  )
  Ensure-Dir $Path
  $Prefs.updated = (Get-Date).ToUniversalTime().ToString("o")
  $json = $Prefs | ConvertTo-Json -Depth 5
  Set-Content -Path $Path -Value $json -Encoding UTF8
  Write-Verbose "Saving preferences to $Path"

  if (Test-CommandExists -Name 'Ninja-Property-Set') {
    try {
      Ninja-Property-Set $NinjaCustomField $json
      Write-Verbose "Writing to Ninja custom field: $NinjaCustomField"
    } catch {
      Write-Warning "Ninja-Property-Set failed: $($_.Exception.Message)"
    }
  }
}

function Test-TimeInRange {
  param(
   [Parameter(Mandatory)] [datetime]$Now,
    [Parameter(Mandatory)] [string]$Start,
    [Parameter(Mandatory)] [string]$End
  )
  # Accepts "HH:mm" 24-hour format. Handles wrap over midnight.
  $today = $Now.Date
  $startTime = [datetime]::ParseExact($Start, 'HH:mm', $null)
  $endTime   = [datetime]::ParseExact($End  , 'HH:mm', $null)

  $startDt = $today.AddHours($startTime.Hour).AddMinutes($startTime.Minute)
  $endDt   = $today.AddHours($endTime.Hour  ).AddMinutes($endTime.Minute)

  if ($endDt -le $startDt) {
    # Quiet window crosses midnight, e.g., 21:00 -> 07:00
    return (($Now -ge $startDt) -or ($Now -lt $endDt.AddDays(1)))
  } else {
    return ($Now -ge $startDt -and $Now -lt $endDt)
  }
}

function Test-QuietHours {
  <#
    .SYNOPSIS
      Returns $true if the current local time falls within the user's Quiet Hours.
    .DESCRIPTION
      Accepts a Quiet Hours JSON object (as returned by Get-QuietPrefs) or reads from default path.
      Supports single range per “weekday” and per “weekend”, plus optional per-day overrides.
  #>
  param(
    $Prefs = $(Get-QuietPrefs -Path $PreferencePath),
    [datetime]$Now = (Get-Date)
  )
  if (-not $Prefs) { return $false }

  $dow = [int]$Now.DayOfWeek  # Sunday=0
  $isWeekend = ($dow -in 0,6)

  # Prefer per-day override if present
  $perDay = $Prefs.perDay | Where-Object { $_.day -eq $dow }
  if ($perDay) {
    return (Test-TimeInRange -Now $Now -Start $perDay.start -End $perDay.end)
  }

  if ($isWeekend -and $Prefs.weekend) {
    return (Test-TimeInRange -Now $Now -Start $Prefs.weekend.start -End $Prefs.weekend.end)
  }
  if (-not $isWeekend -and $Prefs.weekdays) {
    return (Test-TimeInRange -Now $Now -Start $Prefs.weekdays.start -End $Prefs.weekdays.end)
  }
  return $false
}
#endregion Helpers

#region WPF UI
Add-Type -AssemblyName PresentationFramework, PresentationCore

# Build a small XAML UI; declare xmlns:x so x:Name bindings work.
$xaml = @"
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="Set Quiet Hours" Height="440" Width="520"
  WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock FontSize="18" FontWeight="Bold" Text="When should we avoid bothering you?" />
    <TextBlock Grid.Row="1" Margin="0,8,0,12" TextWrapping="Wrap"
               Text="Choose quiet hours. We'll avoid prompts or disruptive tasks during these windows (where possible). Times use 24-hour HH:mm." />

    <StackPanel Grid.Row="2" Orientation="Vertical">

      <!-- Weekdays block -->
      <GroupBox Header="Weekdays (Mon–Fri)" Margin="0,0,0,8">
        <StackPanel Margin="8">
          <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
            <TextBlock Width="80" VerticalAlignment="Center">Start</TextBlock>
            <TextBox x:Name="tbWeekdayStart" Width="80" Text="21:00" />
            <TextBlock Margin="16,0,0,0" Width="80" VerticalAlignment="Center">End</TextBlock>
            <TextBox x:Name="tbWeekdayEnd" Width="80" Text="07:00" />
          </StackPanel>
          <TextBlock FontSize="11" Foreground="Gray" Text="Example: 21:00 to 07:00 (crosses midnight)" />
        </StackPanel>
      </GroupBox>

      <!-- Weekend block -->
      <GroupBox Header="Weekend (Sat–Sun)" Margin="0,0,0,8">
        <StackPanel Margin="8">
          <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
            <TextBlock Width="80" VerticalAlignment="Center">Start</TextBlock>
            <TextBox x:Name="tbWeekendStart" Width="80" Text="22:00" />
            <TextBlock Margin="16,0,0,0" Width="80" VerticalAlignment="Center">End</TextBlock>
            <TextBox x:Name="tbWeekendEnd" Width="80" Text="08:00" />
          </StackPanel>
          <TextBlock FontSize="11" Foreground="Gray" Text="Tip: Weekends often start a bit later." />
        </StackPanel>
      </GroupBox>

      <!-- Optional per-day overrides -->
      <GroupBox Header="Per-day override (optional)" Margin="0,0,0,8">
        <StackPanel Margin="8">
          <TextBlock FontSize="11" Foreground="Gray" Margin="0,0,0,6"
                     Text="Check any day you want a custom time. Unchecked days use the weekday/weekend defaults." />
          <UniformGrid Columns="2" Rows="7" Margin="0,0,0,8">
            <CheckBox x:Name="cb0" Content="Sunday" />
            <TextBox  x:Name="tx0" Width="120" IsEnabled="False" Text="22:00-08:00" />
            <CheckBox x:Name="cb1" Content="Monday" />
            <TextBox  x:Name="tx1" Width="120" IsEnabled="False" Text="21:00-07:00" />
            <CheckBox x:Name="cb2" Content="Tuesday" />
            <TextBox  x:Name="tx2" Width="120" IsEnabled="False" Text="21:00-07:00" />
            <CheckBox x:Name="cb3" Content="Wednesday" />
            <TextBox  x:Name="tx3" Width="120" IsEnabled="False" Text="21:00-07:00" />
            <CheckBox x:Name="cb4" Content="Thursday" />
            <TextBox  x:Name="tx4" Width="120" IsEnabled="False" Text="21:00-07:00" />
            <CheckBox x:Name="cb5" Content="Friday" />
            <TextBox  x:Name="tx5" Width="120" IsEnabled="False" Text="22:00-07:00" />
            <CheckBox x:Name="cb6" Content="Saturday" />
            <TextBox  x:Name="tx6" Width="120" IsEnabled="False" Text="22:00-08:00" />
          </UniformGrid>
          <TextBlock FontSize="11" Foreground="Gray" Text="Format for overrides: HH:mm-HH:mm (24-hour). Example: 20:30-06:30" />
        </StackPanel>
      </GroupBox>

    </StackPanel>

    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="btnSave" Width="120" Height="30" Margin="0,0,8,0">Save</Button>
      <Button x:Name="btnCancel" Width="120" Height="30">Cancel</Button>
    </StackPanel>
  </Grid>
</Window>
"@

# Create UI objects
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$tbWeekdayStart = $window.FindName('tbWeekdayStart')
$tbWeekdayEnd   = $window.FindName('tbWeekdayEnd')
$tbWeekendStart = $window.FindName('tbWeekendStart')
$tbWeekendEnd   = $window.FindName('tbWeekendEnd')
$btnSave        = $window.FindName('btnSave')
$btnCancel      = $window.FindName('btnCancel')

# Day checkboxes + textboxes
$dayCbs = 0..6 | ForEach-Object { $window.FindName("cb$_") }
$dayTbs = 0..6 | ForEach-Object { $window.FindName("tx$_") }

# Enable/disable per-day text when checkbox toggled
for ($i=0; $i -lt $dayCbs.Count; $i++) {
  $idx = $i
  $dayCbs[$idx].Add_Checked({ $dayTbs[$idx].IsEnabled = $true })
  $dayCbs[$idx].Add_Unchecked({ $dayTbs[$idx].IsEnabled = $false })
}

# Preload existing prefs, if any
Write-Verbose "Loading preferences from $PreferencePath"
$existing = Get-QuietPrefs -Path $PreferencePath
if ($existing) {
  if ($existing.weekdays) {
    $tbWeekdayStart.Text = $existing.weekdays.start
    $tbWeekdayEnd.Text   = $existing.weekdays.end
  }
  if ($existing.weekend) {
    $tbWeekendStart.Text = $existing.weekend.start
    $tbWeekendEnd.Text   = $existing.weekend.end
  }
  if ($existing.perDay) {
    foreach ($pd in $existing.perDay) {
      $d = [int]$pd.day
      if ($d -ge 0 -and $d -le 6) {
        $dayCbs[$d].IsChecked = $true
        $dayTbs[$d].IsEnabled = $true
        $dayTbs[$d].Text = ("{0}-{1}" -f $pd.start, $pd.end)
      }
    }
  }
}
#endregion WPF UI

#region Validation
function Test-HHMM { param([string]$t) return [bool]($t -match '^(?:[01]\d|2[0-3]):[0-5]\d$') }
function Parse-Override { 
  param([string]$s)
  if ($s -match '^(?<start>(?:[01]\d|2[0-3]):[0-5]\d)-(?<end>(?:[01]\d|2[0-3]):[0-5]\d)$') {
    return @{start=$Matches.start; end=$Matches.end}
  }
  return $null
}
#endregion Validation

#region Event Handlers
$btnSave.Add_Click({
  # Validate HH:mm
  $errors = @()
  if (-not (Test-HHMM $tbWeekdayStart.Text)) { $errors += "Weekday Start must be HH:mm" }
  if (-not (Test-HHMM $tbWeekdayEnd.Text))   { $errors += "Weekday End must be HH:mm" }
  if (-not (Test-HHMM $tbWeekendStart.Text)) { $errors += "Weekend Start must be HH:mm" }
  if (-not (Test-HHMM $tbWeekendEnd.Text))   { $errors += "Weekend End must be HH:mm" }

  $perDay = @()
  for ($i=0; $i -lt 7; $i++) {
    if ($dayCbs[$i].IsChecked) {
      $ov = Parse-Override $dayTbs[$i].Text
      if (-not $ov) { $errors += "Invalid override for $($dayCbs[$i].Content): use format HH:mm-HH:mm (e.g. 21:00-07:00)" }
      else { $perDay += @{ day=$i; start=$ov.start; end=$ov.end } }
    }
  }

  if ($errors.Count) {
    [System.Windows.MessageBox]::Show(($errors -join "`r`n"), "Fix these", 'OK', 'Error') | Out-Null
    return
  }

  $prefs = [ordered]@{
    weekdays = @{ start = $tbWeekdayStart.Text; end = $tbWeekdayEnd.Text }
    weekend  = @{ start = $tbWeekendStart.Text; end = $tbWeekendEnd.Text }
  }
  if ($perDay.Count) { $prefs.perDay = $perDay }

  Save-QuietPrefs -Prefs $prefs -Path $PreferencePath

  [System.Windows.MessageBox]::Show("Saved. We'll avoid prompts/tasks during your quiet hours when possible.", "Quiet Hours", 'OK', 'Information') | Out-Null
  $window.DialogResult = $true
  $window.Close()
})

$btnCancel.Add_Click({
  $window.DialogResult = $false
  $window.Close()
})
#endregion Event Handlers

# Show window (STA required)
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
  $t = New-Object Threading.Thread({
    $window.ShowDialog() | Out-Null
  })
  $t.SetApartmentState([Threading.ApartmentState]::STA)
  $t.Start()
  $t.Join()
} else {
  $window.ShowDialog() | Out-Null
}

# Example usage from other scripts:
# $prefs = Get-QuietPrefs -Path "C:\RMM\NinjaOne-QuietHours\quiet_hours.json"
# if (Test-QuietHours -Prefs $prefs) {
#   Write-Host "[QuietHours] Skipping disruptive actions at $(Get-Date)."
#   return
# }
# ...else continue with normal work...


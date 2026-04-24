[CmdletBinding()]
param(
    [string]$StateRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptRoot)) {
    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $scriptRoot = Split-Path -Parent $PSCommandPath
    } else {
        $scriptRoot = (Get-Location).Path
    }
}

if ([string]::IsNullOrWhiteSpace($StateRoot)) {
    $StateRoot = Join-Path $scriptRoot 'state'
}

$enginePath = Join-Path $scriptRoot 'WinDefState.ps1'
if (-not (Test-Path -LiteralPath $enginePath)) {
    throw "WinDefState.ps1 was not found next to this GUI script: $enginePath"
}

function Test-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

$isWindowsPlatform = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT
if (-not $isWindowsPlatform) {
    throw 'WinDefState GUI requires Windows PowerShell with WPF.'
}

$requiresRelaunch = -not (Test-Administrator) -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA'
if ($requiresRelaunch) {
    $arguments = @(
        '-NoProfile'
        '-ExecutionPolicy'
        'Bypass'
        '-Sta'
        '-File'
        ('"{0}"' -f $PSCommandPath)
        '-StateRoot'
        ('"{0}"' -f $StateRoot)
    )

    if (Test-Administrator) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList ($arguments -join ' ')
    } else {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList ($arguments -join ' ')
    }
    return
}

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase

$script:EnginePath = $enginePath
$script:StateRoot = [IO.Path]::GetFullPath($StateRoot)
$script:SnapshotPath = $null
$script:Rows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'

function Ensure-Directory {
    param([Parameter(Mandatory)] [string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function ConvertTo-ShortText {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return '<null>'
    }

    if ($Value -is [string] -or $Value -is [bool] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
        $text = [string]$Value
    } else {
        try {
            $text = ConvertTo-Json -InputObject $Value -Depth 5 -Compress
        } catch {
            $text = [string]$Value
        }
    }

    if ($text.Length -gt 180) {
        return ($text.Substring(0, 177) + '...')
    }

    $text
}

function Get-ItemCount {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return 0
    }

    $count = 0
    foreach ($item in @($Value)) {
        $count++
    }

    $count
}

function Get-EntryCategory {
    param([Parameter(Mandatory)] [object]$Entry)

    $id = [string]$Entry.Id
    if ($id -match '^([^.]+)\.') {
        return $matches[1]
    }

    switch ([string]$Entry.Type) {
        'WdacPolicies' { 'wdac' }
        'BitLockerVolumes' { 'bitlocker' }
        default { 'other' }
    }
}

function Get-EntryCurrentSummary {
    param([Parameter(Mandatory)] [object]$Entry)

    switch ([string]$Entry.Type) {
        'RegistryValue' {
            if ($Entry.Exists) { return (ConvertTo-ShortText -Value $Entry.CurrentValue) }
            return '<absent>'
        }
        'MpPreferenceValue' {
            return (ConvertTo-ShortText -Value $Entry.CurrentValue)
        }
        'MpPreferenceList' {
            return ('{0} item(s)' -f (Get-ItemCount -Value $Entry.CurrentValue))
        }
        'ServiceConfig' {
            return ('{0} / {1}' -f $Entry.CurrentValue.StartMode, $Entry.CurrentValue.State)
        }
        'FirewallProfiles' {
            return ('{0} profile(s)' -f (Get-ItemCount -Value $Entry.CurrentValue.Profiles))
        }
        'FirewallRules' {
            return ('{0} rule(s)' -f (Get-ItemCount -Value $Entry.CurrentValue.Rules))
        }
        'LoadedUserRegistryValues' {
            return ('{0} value(s), {1} issue(s)' -f (Get-ItemCount -Value $Entry.CurrentValue.Entries), (Get-ItemCount -Value $Entry.CurrentValue.CaptureIssues))
        }
        'WdacPolicies' {
            return ('{0} policy item(s), {1} file(s)' -f (Get-ItemCount -Value $Entry.CurrentValue.Policies), (Get-ItemCount -Value $Entry.CurrentValue.Files))
        }
        'BitLockerVolumes' {
            return ('{0} volume(s)' -f (Get-ItemCount -Value $Entry.CurrentValue.Volumes))
        }
        default {
            return (ConvertTo-ShortText -Value $Entry.CurrentValue)
        }
    }
}

function Get-EntryBadges {
    param([Parameter(Mandatory)] [object]$Entry)

    $badges = New-Object System.Collections.Generic.List[string]
    if ($Entry.RequiresReboot) {
        $badges.Add('Requires reboot')
    }

    if ($Entry.PSObject.Properties['InvalidEntries'] -and (Get-ItemCount -Value $Entry.InvalidEntries) -gt 0) {
        $badges.Add('Partial')
    }

    if ($Entry.PSObject.Properties['Captured'] -and -not [bool]$Entry.Captured) {
        $badges.Add('Partial')
    }

    if ($Entry.PSObject.Properties['CommandAvailable'] -and -not [bool]$Entry.CommandAvailable) {
        $badges.Add('Provider missing')
    }

    if ($Entry.PSObject.Properties['CurrentValue'] -and $null -ne $Entry.CurrentValue) {
        $state = $Entry.CurrentValue
        if ($state.PSObject.Properties['CaptureIssues'] -and (Get-ItemCount -Value $state.CaptureIssues) -gt 0) {
            $badges.Add('Partial')
        }
        if ($state.PSObject.Properties['LocalMatchesEffective'] -and -not [bool]$state.LocalMatchesEffective) {
            $badges.Add('Partial')
        }
        if ($state.PSObject.Properties['TimedOutMountPoints'] -and (Get-ItemCount -Value $state.TimedOutMountPoints) -gt 0) {
            $badges.Add('Partial')
        }
        if ($state.PSObject.Properties['Policies']) {
            foreach ($policy in @($state.Policies)) {
                $platformProperty = $policy.PSObject.Properties['Platform Policy']
                if ($null -eq $platformProperty) {
                    $platformProperty = $policy.PSObject.Properties['PlatformPolicy']
                }
                if ($null -ne $platformProperty -and [string]$platformProperty.Value -match '^(True|1)$') {
                    $badges.Add('Platform managed')
                    break
                }
            }
        }
    }

    @($badges | Sort-Object -Unique) -join ', '
}

function New-SnapshotRow {
    param([Parameter(Mandatory)] [object]$Entry)

    [PSCustomObject]@{
        Selected      = $false
        Category      = Get-EntryCategory -Entry $Entry
        Id            = [string]$Entry.Id
        Type          = [string]$Entry.Type
        Reboot        = if ($Entry.RequiresReboot) { 'Yes' } else { 'No' }
        Badges        = Get-EntryBadges -Entry $Entry
        Current       = Get-EntryCurrentSummary -Entry $Entry
        Action        = 'Permissive target'
        ActionOptions = @('Permissive target', 'Restore captured')
    }
}

function Get-LatestSnapshotPath {
    $snapshotDir = Join-Path $script:StateRoot 'snapshots'
    if (-not (Test-Path -LiteralPath $snapshotDir)) {
        return $null
    }

    $latest = Get-ChildItem -LiteralPath $snapshotDir -Filter '*.json' -File -ErrorAction SilentlyContinue |
        Sort-Object -Property LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -eq $latest) {
        return $null
    }

    $latest.FullName
}

function Set-GuiStatus {
    param([Parameter(Mandatory)] [string]$Message)

    $StatusText.Text = $Message
    $StatusText.ToolTip = $Message
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Add-GuiLog {
    param([Parameter(Mandatory)] [string]$Message)

    $LogBox.AppendText($Message.TrimEnd() + [Environment]::NewLine)
    $LogBox.ScrollToEnd()
}

function Load-Snapshot {
    param([Parameter(Mandatory)] [string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $fullPath)) {
        throw "Snapshot was not found: $fullPath"
    }

    $snapshot = Get-Content -LiteralPath $fullPath -Raw | ConvertFrom-Json
    $script:Rows.Clear()
    foreach ($entry in @($snapshot.Settings)) {
        $script:Rows.Add((New-SnapshotRow -Entry $entry)) | Out-Null
    }

    $script:SnapshotPath = $fullPath
    $SnapshotPathText.Text = $fullPath
    Set-GuiStatus ("Loaded {0} setting(s) from {1}" -f (Get-ItemCount -Value $script:Rows), [IO.Path]::GetFileName($fullPath))
}

function Get-SelectedRows {
    @($script:Rows | Where-Object { $_.Selected })
}

function Invoke-WinDefState {
    param(
        [Parameter(Mandatory)] [ValidateSet('Snapshot', 'Permissive', 'Restore')] [string]$Command,
        [string]$SnapshotPath,
        [string[]]$IncludeId
    )

    $arguments = @(
        '-NoProfile'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        $script:EnginePath
        '-Command'
        $Command
        '-StateRoot'
        $script:StateRoot
    )

    if (-not [string]::IsNullOrWhiteSpace($SnapshotPath)) {
        $arguments += @('-SnapshotPath', $SnapshotPath)
    }

    if ((Get-ItemCount -Value $IncludeId) -gt 0) {
        $arguments += '-IncludeId'
        $arguments += @($IncludeId)
    }

    Add-GuiLog ("> powershell.exe {0}" -f ($arguments -join ' '))
    $output = & powershell.exe @arguments 2>&1 | Out-String
    if (-not [string]::IsNullOrWhiteSpace($output)) {
        Add-GuiLog $output
    }

    if ($LASTEXITCODE -ne 0) {
        throw "WinDefState exited with code $LASTEXITCODE."
    }
}

function Invoke-GuiOperation {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [scriptblock]$Operation
    )

    try {
        $Window.Cursor = [System.Windows.Input.Cursors]::Wait
        Set-GuiStatus $Name
        & $Operation
    } catch {
        Set-GuiStatus $_.Exception.Message
        Add-GuiLog ("ERROR: {0}" -f $_.Exception.Message)
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'WinDefState', 'OK', 'Error') | Out-Null
    } finally {
        $Window.Cursor = $null
    }
}

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="WinDefState" Width="1180" Height="760" MinWidth="940" MinHeight="560"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="12">
  <Grid Background="#F4F6F8">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="150"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <DockPanel Grid.Row="0" LastChildFill="True" Margin="12">
      <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">
        <Button x:Name="SnapshotButton" Content="Snapshot" Width="92" Margin="0,0,8,0"/>
        <Button x:Name="LoadLatestButton" Content="Latest" Width="76" Margin="0,0,8,0"/>
        <Button x:Name="BrowseButton" Content="Load..." Width="76" Margin="0,0,8,0"/>
        <Button x:Name="SelectAllButton" Content="Select all" Width="82" Margin="12,0,8,0"/>
        <Button x:Name="ClearButton" Content="Clear" Width="64" Margin="0,0,8,0"/>
        <Button x:Name="RunSelectedButton" Content="Run selected" Width="108" Margin="12,0,8,0"/>
        <Button x:Name="OpenReportButton" Content="Report" Width="70" Margin="12,0,8,0"/>
        <Button x:Name="OpenStateButton" Content="State folder" Width="92"/>
      </StackPanel>
      <TextBlock x:Name="SnapshotPathText" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="16,0,0,0"/>
    </DockPanel>

    <DataGrid x:Name="SettingsGrid" Grid.Row="1" Margin="12,0,12,8"
              AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
              IsReadOnly="False" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
              SelectionMode="Extended" Background="White" AlternatingRowBackground="#F8FAFC"
              RowHeaderWidth="0">
      <DataGrid.Columns>
        <DataGridCheckBoxColumn Header="Use" Width="42" Binding="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
        <DataGridTextColumn Header="Category" Width="110" IsReadOnly="True" Binding="{Binding Category}"/>
        <DataGridTextColumn Header="Id" Width="260" IsReadOnly="True" Binding="{Binding Id}"/>
        <DataGridTextColumn Header="Type" Width="175" IsReadOnly="True" Binding="{Binding Type}"/>
        <DataGridTextColumn Header="Reboot" Width="72" IsReadOnly="True" Binding="{Binding Reboot}"/>
        <DataGridTextColumn Header="Badges" Width="180" IsReadOnly="True" Binding="{Binding Badges}"/>
        <DataGridTextColumn Header="Current" Width="*" MinWidth="220" IsReadOnly="True" Binding="{Binding Current}"/>
        <DataGridTemplateColumn Header="Action" Width="150">
          <DataGridTemplateColumn.CellTemplate>
            <DataTemplate>
              <ComboBox ItemsSource="{Binding ActionOptions}" SelectedItem="{Binding Action, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" MinWidth="128"/>
            </DataTemplate>
          </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>
      </DataGrid.Columns>
    </DataGrid>

    <TextBox x:Name="LogBox" Grid.Row="2" Margin="12,0,12,8" IsReadOnly="True"
             VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
             TextWrapping="NoWrap" FontFamily="Consolas" FontSize="11" Background="#101820" Foreground="#E6EDF3"/>

    <Border Grid.Row="3" Background="#FFFFFF" BorderBrush="#D7DEE7" BorderThickness="1,1,0,0">
      <TextBlock x:Name="StatusText" Margin="12,7" Text="Ready" TextTrimming="CharacterEllipsis"/>
    </Border>
  </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

$SnapshotButton = $Window.FindName('SnapshotButton')
$LoadLatestButton = $Window.FindName('LoadLatestButton')
$BrowseButton = $Window.FindName('BrowseButton')
$SelectAllButton = $Window.FindName('SelectAllButton')
$ClearButton = $Window.FindName('ClearButton')
$RunSelectedButton = $Window.FindName('RunSelectedButton')
$OpenReportButton = $Window.FindName('OpenReportButton')
$OpenStateButton = $Window.FindName('OpenStateButton')
$SettingsGrid = $Window.FindName('SettingsGrid')
$SnapshotPathText = $Window.FindName('SnapshotPathText')
$LogBox = $Window.FindName('LogBox')
$StatusText = $Window.FindName('StatusText')

$SettingsGrid.ItemsSource = $script:Rows
Ensure-Directory -Path $script:StateRoot

$SnapshotButton.Add_Click({
    Invoke-GuiOperation -Name 'Taking snapshot...' -Operation {
        Invoke-WinDefState -Command Snapshot
        $latest = Get-LatestSnapshotPath
        if ($null -ne $latest) {
            Load-Snapshot -Path $latest
        }
    }
})

$LoadLatestButton.Add_Click({
    Invoke-GuiOperation -Name 'Loading latest snapshot...' -Operation {
        $latest = Get-LatestSnapshotPath
        if ($null -eq $latest) {
            throw 'No snapshots were found.'
        }
        Load-Snapshot -Path $latest
    }
})

$BrowseButton.Add_Click({
    Invoke-GuiOperation -Name 'Loading snapshot...' -Operation {
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = 'WinDefState snapshots (*.json)|*.json|All files (*.*)|*.*'
        $snapshotDir = Join-Path $script:StateRoot 'snapshots'
        if (Test-Path -LiteralPath $snapshotDir) {
            $dialog.InitialDirectory = $snapshotDir
        }
        if ($dialog.ShowDialog()) {
            Load-Snapshot -Path $dialog.FileName
        }
    }
})

$SelectAllButton.Add_Click({
    foreach ($row in $script:Rows) {
        $row.Selected = $true
    }
    $SettingsGrid.Items.Refresh()
    Set-GuiStatus ("Selected {0} setting(s)." -f (Get-ItemCount -Value $script:Rows))
})

$ClearButton.Add_Click({
    foreach ($row in $script:Rows) {
        $row.Selected = $false
    }
    $SettingsGrid.Items.Refresh()
    Set-GuiStatus 'Selection cleared.'
})

$RunSelectedButton.Add_Click({
    Invoke-GuiOperation -Name 'Running selected action...' -Operation {
        $selectedRows = @(Get-SelectedRows)
        if ((Get-ItemCount -Value $selectedRows) -eq 0) {
            throw 'Select at least one setting first.'
        }

        $actions = @($selectedRows | ForEach-Object { [string]$_.Action } | Sort-Object -Unique)
        if ((Get-ItemCount -Value $actions) -ne 1) {
            throw 'Selected rows must use one action at a time.'
        }

        $ids = @($selectedRows | ForEach-Object { [string]$_.Id })
        switch ($actions[0]) {
            'Permissive target' {
                Invoke-WinDefState -Command Permissive -IncludeId $ids
                $latest = Get-LatestSnapshotPath
                if ($null -ne $latest) {
                    Load-Snapshot -Path $latest
                }
            }
            'Restore captured' {
                if ([string]::IsNullOrWhiteSpace($script:SnapshotPath)) {
                    throw 'Load a snapshot before restoring selected settings.'
                }
                Invoke-WinDefState -Command Restore -SnapshotPath $script:SnapshotPath -IncludeId $ids
                Load-Snapshot -Path $script:SnapshotPath
            }
            default {
                throw "Unknown action: $($actions[0])"
            }
        }
    }
})

$OpenReportButton.Add_Click({
    Invoke-GuiOperation -Name 'Opening report...' -Operation {
        if ([string]::IsNullOrWhiteSpace($script:SnapshotPath)) {
            throw 'Load a snapshot first.'
        }
        $reportPath = [IO.Path]::ChangeExtension($script:SnapshotPath, 'txt')
        if (-not (Test-Path -LiteralPath $reportPath)) {
            throw "Report was not found: $reportPath"
        }
        Invoke-Item -LiteralPath $reportPath
    }
})

$OpenStateButton.Add_Click({
    Invoke-GuiOperation -Name 'Opening state folder...' -Operation {
        Ensure-Directory -Path $script:StateRoot
        Invoke-Item -LiteralPath $script:StateRoot
    }
})

$latestSnapshot = Get-LatestSnapshotPath
if ($null -ne $latestSnapshot) {
    try {
        Load-Snapshot -Path $latestSnapshot
    } catch {
        Set-GuiStatus $_.Exception.Message
    }
} else {
    $SnapshotPathText.Text = $script:StateRoot
    Set-GuiStatus 'Ready. Take a snapshot to populate the grid.'
}

$Window.ShowDialog() | Out-Null

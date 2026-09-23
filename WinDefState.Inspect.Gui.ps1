#Requires -Version 5.1
[CmdletBinding()]
param([switch]$ValidateOnly, [string]$BaselinePath, [string]$PreviewPath)

$ErrorActionPreference = 'Stop'
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'The inspection dashboard requires Windows with WPF.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Launch with powershell.exe -NoProfile -Sta -File .\WinDefState.Inspect.Gui.ps1' }
. (Join-Path $PSScriptRoot 'WinDefState.Environment.ps1')
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$script:InspectionRoot = $PSScriptRoot
$script:InspectionReport = $null
$script:InspectionDiff = $null
$script:InspectionCapture = $null

$inspectionXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="WinDefState - Environment" Width="1260" Height="880" MinWidth="960" MinHeight="680"
 WindowStartupLocation="CenterScreen" Background="#F5F6F7" FontFamily="Segoe UI" FontSize="13"
 Foreground="#202830" UseLayoutRounding="True">
 <Window.Resources>
  <Style TargetType="Button">
   <Setter Property="Padding" Value="12,6"/><Setter Property="Margin" Value="0,0,6,0"/>
   <Setter Property="Background" Value="White"/><Setter Property="Foreground" Value="#202830"/>
   <Setter Property="BorderBrush" Value="#BDC4CB"/><Setter Property="Cursor" Value="Hand"/>
  </Style>
  <Style TargetType="DataGrid">
   <Setter Property="AutoGenerateColumns" Value="False"/><Setter Property="IsReadOnly" Value="True"/>
   <Setter Property="CanUserAddRows" Value="False"/><Setter Property="CanUserDeleteRows" Value="False"/>
   <Setter Property="HeadersVisibility" Value="Column"/><Setter Property="GridLinesVisibility" Value="Horizontal"/>
   <Setter Property="HorizontalGridLinesBrush" Value="#EDF0F2"/><Setter Property="BorderThickness" Value="0"/>
   <Setter Property="RowBackground" Value="White"/><Setter Property="AlternatingRowBackground" Value="#FAFBFC"/>
   <Setter Property="Background" Value="White"/>
   <Setter Property="SelectionMode" Value="Single"/><Setter Property="EnableRowVirtualization" Value="True"/>
   <Setter Property="EnableColumnVirtualization" Value="True"/><Setter Property="MinRowHeight" Value="28"/>
  </Style>
  <Style TargetType="DataGridColumnHeader">
   <Setter Property="Background" Value="#EEF1F4"/><Setter Property="Foreground" Value="#384653"/>
   <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Padding" Value="8,7"/>
   <Setter Property="BorderThickness" Value="0"/>
  </Style>
  <Style TargetType="DataGridCell"><Setter Property="Padding" Value="8,4"/><Setter Property="BorderThickness" Value="0"/></Style>
  <Style TargetType="TabItem"><Setter Property="Padding" Value="16,7"/></Style>
  <Style x:Key="EmptyMessage" TargetType="TextBlock">
   <Setter Property="Foreground" Value="#52606D"/><Setter Property="HorizontalAlignment" Value="Center"/>
   <Setter Property="VerticalAlignment" Value="Center"/><Setter Property="TextWrapping" Value="Wrap"/>
   <Setter Property="Margin" Value="24"/><Setter Property="IsHitTestVisible" Value="False"/>
  </Style>
 </Window.Resources>
 <Grid Background="#F5F6F7">
  <Grid.RowDefinitions><RowDefinition Height="46"/><RowDefinition Height="*"/><RowDefinition Height="32"/></Grid.RowDefinitions>
  <Border Background="White" BorderBrush="#D5DAE0" BorderThickness="0,0,0,1" Padding="16,8">
   <DockPanel>
    <TextBlock DockPanel.Dock="Right" Text="Read-only" Foreground="#52606D" VerticalAlignment="Center"/>
    <StackPanel Orientation="Horizontal"><TextBlock Text="WinDefState" FontSize="18" FontWeight="SemiBold"/>
     <TextBlock Text="Environment" Foreground="#52606D" Margin="16,0,0,0" VerticalAlignment="Center"/>
    </StackPanel>
   </DockPanel>
  </Border>
  <Grid Grid.Row="1" Margin="16,12,16,10">
   <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="32"/>
    <RowDefinition Height="*"/><RowDefinition Height="6"/><RowDefinition Height="126"/>
   </Grid.RowDefinitions>
   <DockPanel Margin="0,0,0,12">
    <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">
     <Button x:Name="CaptureButton" Content="_Capture" Background="#176CAF" Foreground="White" BorderBrush="#176CAF" ToolTip="Capture this computer's environment"/>
     <Button x:Name="CancelButton" Content="Cancel" IsEnabled="False"/>
     <Button x:Name="OpenButton" Content="_Open..." ToolTip="Open a saved JSON baseline"/>
     <Button x:Name="CompareButton" Content="C_ompare..." ToolTip="Choose the earlier baseline to compare with the loaded capture" IsEnabled="False"/>
     <Button x:Name="SaveButton" Content="_Save baseline..." IsEnabled="False"/>
    </StackPanel>
    <TextBlock DockPanel.Dock="Left" Text="Search" Margin="12,0,8,0" VerticalAlignment="Center"/>
    <TextBox x:Name="SearchBox" Padding="8,5" VerticalContentAlignment="Center"
     ToolTip="Search the current view (Ctrl+F)" AutomationProperties.Name="Search checks or inventory"/>
   </DockPanel>
   <StackPanel Grid.Row="1" Margin="0,0,0,6">
    <TextBlock x:Name="HostText" Text="No baseline loaded" FontSize="16" FontWeight="SemiBold"/>
    <TextBlock x:Name="CaptureText" Text="Capture this computer or open a saved baseline." Foreground="#52606D" Margin="0,3,0,0" TextTrimming="CharacterEllipsis"/>
   </StackPanel>
   <WrapPanel Grid.Row="2" VerticalAlignment="Center">
    <TextBlock Margin="0,0,24,0"><Run Text="TCP / UDP: "/><Run x:Name="ListenersCount" Text="--" FontWeight="SemiBold"/></TextBlock>
    <TextBlock Margin="0,0,24,0"><Run Text="Firewall rules: "/><Run x:Name="RulesCount" Text="--" FontWeight="SemiBold"/></TextBlock>
    <TextBlock Margin="0,0,24,0" Foreground="#8C4700"><Run Text="Health findings: "/><Run x:Name="AttentionCount" Text="--" FontWeight="SemiBold"/></TextBlock>
    <TextBlock Foreground="#5D4BB0"><Run Text="Unavailable sections: "/><Run x:Name="UnknownCount" Text="--" FontWeight="SemiBold"/></TextBlock>
   </WrapPanel>
   <TabControl x:Name="InspectionTabs" Grid.Row="3" SelectedIndex="0" Background="White" BorderBrush="#D5DAE0">
    <TabItem Header="Inventory">
     <Grid Margin="8"><Grid.ColumnDefinitions><ColumnDefinition Width="225"/><ColumnDefinition Width="8"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
      <ListBox x:Name="SectionList" BorderBrush="#D5DAE0" ScrollViewer.HorizontalScrollBarVisibility="Disabled" AutomationProperties.Name="Inventory sections">
       <ListBox.ItemTemplate><DataTemplate><TextBlock Text="{Binding Label}" TextWrapping="Wrap" Padding="6,5"/></DataTemplate></ListBox.ItemTemplate>
      </ListBox>
      <DataGrid x:Name="InventoryGrid" Grid.Column="2" AutomationProperties.Name="Inventory records"/>
      <TextBlock x:Name="InventoryEmpty" Grid.Column="2" Style="{StaticResource EmptyMessage}" Text="Capture this computer or open a baseline."/>
     </Grid>
    </TabItem>
    <TabItem Header="Security">
     <Grid Margin="8">
      <DataGrid x:Name="HealthGrid" AutomationProperties.Name="Protection checks">
       <DataGrid.RowStyle><Style TargetType="DataGridRow"><Style.Triggers>
        <DataTrigger Binding="{Binding Status}" Value="Attention"><Setter Property="Foreground" Value="#8C4700"/></DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Unknown"><Setter Property="Foreground" Value="#5D4BB0"/></DataTrigger>
       </Style.Triggers></Style></DataGrid.RowStyle>
       <DataGrid.Columns>
        <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="105"/>
        <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="120"/>
        <DataGridTextColumn Header="Check" Binding="{Binding Name}" Width="2*"/>
        <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="2*"/>
       </DataGrid.Columns>
      </DataGrid>
      <TextBlock x:Name="HealthEmpty" Style="{StaticResource EmptyMessage}" Text="Capture this computer or open a baseline."/>
     </Grid>
    </TabItem>
    <TabItem Header="Changes">
     <Grid Margin="8">
      <DataGrid x:Name="DiffGrid" AutomationProperties.Name="Environment changes">
       <DataGrid.Columns>
        <DataGridTextColumn Header="Change" Binding="{Binding Change}" Width="100"/>
        <DataGridTextColumn Header="Section" Binding="{Binding Section}" Width="210"/>
        <DataGridTextColumn Header="Identity" Binding="{Binding Key}" Width="*"/>
       </DataGrid.Columns>
      </DataGrid>
      <TextBlock x:Name="DiffEmpty" Style="{StaticResource EmptyMessage}" Text="Choose Compare to load an earlier baseline."/>
     </Grid>
    </TabItem>
   </TabControl>
   <GridSplitter Grid.Row="4" HorizontalAlignment="Stretch" Background="Transparent"/>
   <Border Grid.Row="5" Background="White" BorderBrush="#D5DAE0" BorderThickness="1" Padding="10,8">
    <DockPanel><TextBlock DockPanel.Dock="Top" Text="Details" FontWeight="SemiBold" Foreground="#52606D" Margin="0,0,0,4"/>
     <TextBox x:Name="DetailBox" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
      Background="Transparent" BorderThickness="0" FontFamily="Consolas" FontSize="12" Text="Select a row to view its details."/>
    </DockPanel>
   </Border>
  </Grid>
  <Border Grid.Row="2" Background="#EEF1F4" BorderBrush="#D5DAE0" BorderThickness="0,1,0,0" Padding="16,6">
   <DockPanel>
    <TextBlock x:Name="RowsText" DockPanel.Dock="Right" Text="0 records" MinWidth="140" TextAlignment="Right" Foreground="#52606D" Margin="12,0,0,0"/>
    <ProgressBar x:Name="CaptureProgress" DockPanel.Dock="Right" Width="100" Height="5" IsIndeterminate="False" Margin="12,0,0,0"/>
    <TextBlock x:Name="StatusText" Text="Ready" Foreground="#52606D" TextTrimming="CharacterEllipsis"/>
   </DockPanel>
  </Border>
 </Grid>
</Window>
'@
$reader = New-Object Xml.XmlNodeReader ([xml]$inspectionXaml)
try { $Window = [Windows.Markup.XamlReader]::Load($reader) } finally { $reader.Close() }
$controls = @{}
foreach ($name in @('CaptureButton', 'CancelButton', 'HostText', 'CaptureText', 'ListenersCount', 'RulesCount', 'AttentionCount', 'UnknownCount', 'OpenButton', 'CompareButton', 'SaveButton', 'SearchBox', 'InspectionTabs', 'SectionList', 'InventoryGrid', 'HealthGrid', 'DiffGrid', 'DetailBox', 'CaptureProgress', 'StatusText', 'RowsText', 'InventoryEmpty', 'HealthEmpty', 'DiffEmpty')) {
    $controls[$name] = $Window.FindName($name)
    if ($null -eq $controls[$name]) { throw "Missing inspection control: $name" }
}
foreach ($name in @('InventoryGrid', 'HealthGrid', 'DiffGrid', 'SectionList')) { $controls[$name].FontWeight = [Windows.FontWeights]::Normal }

function Update-InspectionFilter {
    $controls.DetailBox.Text = 'Select a row to view its details.'
    $query = [string]$controls.SearchBox.Text
    $tokens = @($query.Trim() -split '\s+' | Where-Object { $_ })
    $matchesQuery = { param($value) foreach ($token in $tokens) { if ($value.IndexOf($token, [StringComparison]::OrdinalIgnoreCase) -lt 0) { return $false } }; return $true }
    $section = $controls.SectionList.SelectedItem
    $controls.InventoryGrid.ItemsSource = @(if ($null -ne $section) { foreach ($item in $section.Section.Items) {
        $summary = ConvertTo-Json -InputObject $item.Data -Depth 12 -Compress
        if (& $matchesQuery ($item.Key + ' ' + $summary)) {
            $display = [ordered]@{}
            foreach ($property in $item.Data.PSObject.Properties) {
                $display[$property.Name] = if ($property.Value -is [array]) { ConvertTo-Json -InputObject $property.Value -Depth 8 -Compress } else { [string]$property.Value }
            }
            [pscustomobject]@{ Key = $item.Key; Summary = $summary; Data = $item.Data; Display = [pscustomobject]$display }
        }
    } })
    $controls.HealthGrid.ItemsSource = @(if ($null -ne $script:InspectionReport) { $script:InspectionReport.Health.Checks | Where-Object { & $matchesQuery (ConvertTo-Json -InputObject $_ -Compress) } })
    $controls.DiffGrid.ItemsSource = @(if ($null -ne $script:InspectionDiff) { $script:InspectionDiff.Changes | Where-Object { & $matchesQuery (ConvertTo-Json -InputObject $_ -Depth 14 -Compress) } })
    $controls.InventoryEmpty.Visibility = if ($controls.InventoryGrid.Items.Count -eq 0) { 'Visible' } else { 'Collapsed' }
    $controls.InventoryEmpty.Text = if ($null -eq $script:InspectionReport) { 'Capture this computer or open a baseline.' } elseif ($null -ne $section -and $section.Section.Status -eq 'Unknown') { 'Section unavailable. See Details.' } elseif ($tokens.Count -gt 0) { 'No matching records.' } else { 'No records in this section.' }
    $controls.HealthEmpty.Visibility = if ($controls.HealthGrid.Items.Count -eq 0) { 'Visible' } else { 'Collapsed' }
    $controls.HealthEmpty.Text = if ($null -eq $script:InspectionReport) { 'Capture this computer or open a baseline.' } else { 'No matching checks.' }
    $controls.DiffEmpty.Visibility = if ($controls.DiffGrid.Items.Count -eq 0) { 'Visible' } else { 'Collapsed' }
    $controls.DiffEmpty.Text = if ($null -eq $script:InspectionDiff) { 'Choose Compare to load an earlier baseline.' } elseif (@($script:InspectionDiff.Changes).Count -eq 0) { 'No inventory differences found.' } else { 'No matching changes.' }
    $total = 0; $visible = 0
    switch ($controls.InspectionTabs.SelectedIndex) {
        0 { $visible = $controls.InventoryGrid.Items.Count; if ($null -ne $section) { $total = @($section.Section.Items).Count } }
        1 { $visible = $controls.HealthGrid.Items.Count; if ($null -ne $script:InspectionReport) { $total = @($script:InspectionReport.Health.Checks).Count } }
        2 { $visible = $controls.DiffGrid.Items.Count; if ($null -ne $script:InspectionDiff) { $total = @($script:InspectionDiff.Changes).Count } }
    }
    $controls.RowsText.Text = '{0} of {1} records' -f $visible, $total
    $controls.SaveButton.Content = if ($controls.InspectionTabs.SelectedIndex -eq 2 -and $null -ne $script:InspectionDiff) { '_Save comparison...' } else { '_Save baseline...' }
}

function Set-InspectionColumns {
    $controls.InventoryGrid.Columns.Clear()
    $selected = $controls.SectionList.SelectedItem
    if ($null -eq $selected -or @($selected.Section.Items).Count -eq 0) { return }
    $fields = @($selected.Section.Items[0].Data.PSObject.Properties.Name)
    $priority = switch ($selected.Section.Id) {
        'network.tcp' { @('Protocol', 'LocalAddress', 'LocalPort', 'ProcessName', 'ExecutablePath', 'ProcessId', 'OwnerStatus') }
        'network.udp' { @('Protocol', 'LocalAddress', 'LocalPort', 'ProcessName', 'ExecutablePath', 'ProcessId', 'OwnerStatus') }
        'firewall.rules' { @('DisplayName', 'Enabled', 'Direction', 'Action', 'Profile', 'Name') }
        'system.services' { @('Name', 'State', 'StartMode', 'StartName', 'PathName') }
        default { @() }
    }
    $orderedFields = @(@($priority | Where-Object { $_ -in $fields }) + @($fields | Where-Object { $_ -notin $priority }))
    foreach ($field in $orderedFields) {
        $column = New-Object Windows.Controls.DataGridTextColumn
        $column.Header = $field -creplace '([a-z])([A-Z])', '$1 $2'
        $column.Binding = New-Object Windows.Data.Binding ('Display.' + $field)
        $column.Width = [Windows.Controls.DataGridLength]::new(150)
        if ($field -in @('Protocol', 'LocalPort', 'ProcessId', 'Enabled', 'Action')) { $column.Width = [Windows.Controls.DataGridLength]::new(90) }
        if ($field -in @('ExecutablePath', 'PathName', 'DisplayName')) { $column.Width = [Windows.Controls.DataGridLength]::new(240) }
        $controls.InventoryGrid.Columns.Add($column)
    }
}

function Set-InspectionReport {
    param($Report)
    Assert-EnvironmentBaseline $Report
    if ($null -eq $Report.Health -or $Report.Health.ReportType -ne 'WinDefState.Health' -or $null -eq $Report.Health.Windows -or $null -eq $Report.Health.Summary) { throw 'This baseline is missing the health metadata required by the dashboard.' }
    $script:InspectionReport = $Report
    $script:InspectionDiff = $null
    $controls.HostText.Text = $Report.ComputerName + ' / ' + $Report.Health.Windows.Family + ' ' + $Report.Health.Windows.Release
    $captureTime = [datetimeoffset]::MinValue
    $captured = if ([datetimeoffset]::TryParse([string]$Report.CapturedAtUtc, [ref]$captureTime)) { $captureTime.UtcDateTime.ToString('yyyy-MM-dd HH:mm:ss') + ' UTC' } else { [string]$Report.CapturedAtUtc }
    $access = if ($Report.Elevated) { 'Administrator' } else { 'Standard user' }
    $controls.CaptureText.Text = 'Captured ' + $captured + ' | ' + $access
    $counts = foreach ($id in @('network.tcp', 'network.udp')) {
        $section = @($Report.Sections | Where-Object Id -eq $id)
        if ($section.Count -eq 1 -and $section[0].Status -eq 'Captured') { [string]@($section[0].Items).Count } else { '?' }
    }
    $controls.ListenersCount.Text = $counts -join ' / '
    $rules = @($Report.Sections | Where-Object Id -eq 'firewall.rules')
    $controls.RulesCount.Text = if ($rules.Count -eq 1 -and $rules[0].Status -eq 'Captured') { [string]@($rules[0].Items).Count } else { '?' }
    $controls.AttentionCount.Text = [string]([int]$Report.Health.Summary.Attention + [int]$Report.Health.Summary.Unknown)
    $controls.AttentionCount.ToolTip = '{0} need attention; {1} unknown checks.' -f $Report.Health.Summary.Attention, $Report.Health.Summary.Unknown
    $controls.UnknownCount.Text = [string]@($Report.Sections | Where-Object Status -eq Unknown).Count
    $controls.SectionList.ItemsSource = @($Report.Sections | ForEach-Object { [pscustomobject]@{ Label = '{0} ({1})' -f $_.Name, $(if ($_.Status -eq 'Captured') { @($_.Items).Count } else { '?' }); Section = $_ } })
    $controls.SectionList.SelectedIndex = 0
    $controls.CompareButton.IsEnabled = $true
    $controls.SaveButton.IsEnabled = $true
    Update-InspectionFilter
}

function Clear-InspectionCapture {
    if ($null -eq $script:InspectionCapture) { return }
    $script:InspectionCapture.Process.Dispose()
    foreach ($path in @($script:InspectionCapture.Path, $script:InspectionCapture.ErrorPath)) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
    $script:InspectionCapture = $null
}

function Set-InspectionBusy {
    param([bool]$Busy)
    $controls.CaptureButton.IsEnabled = -not $Busy
    $controls.CancelButton.IsEnabled = $Busy
    $controls.OpenButton.IsEnabled = -not $Busy
    $controls.CompareButton.IsEnabled = -not $Busy -and $null -ne $script:InspectionReport
    $controls.SaveButton.IsEnabled = -not $Busy -and $null -ne $script:InspectionReport
    $controls.CaptureProgress.IsIndeterminate = $Busy
}

$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(250)
$timer.Add_Tick({
    if ($null -eq $script:InspectionCapture) { return }
    try {
        $capture = $script:InspectionCapture
        if (-not $capture.Process.HasExited) {
            if (((Get-Date) - $capture.Started).TotalSeconds -gt 180) {
                $capture.Process.Kill()
                $capture.Cancelled = $true
                $controls.StatusText.Text = 'Capture timed out after 180 seconds. The previous baseline remains loaded.'
            }
            return
        }
        $timer.Stop()
        if (-not $capture.Cancelled) {
            if ($capture.Process.ExitCode -ne 0) {
                $message = if (Test-Path -LiteralPath $capture.ErrorPath) { Get-Content -LiteralPath $capture.ErrorPath -Raw } else { 'Capture did not return a baseline.' }
                throw $message
            }
            Set-InspectionReport (Read-EnvironmentBaseline $capture.Path)
            $controls.StatusText.Text = 'Capture complete. Save the baseline before testing.'
        }
    } catch { $controls.StatusText.Text = $_.Exception.Message; $controls.DetailBox.Text = $_.Exception.Message }
    finally {
        if ($null -ne $script:InspectionCapture -and $script:InspectionCapture.Process.HasExited) { Clear-InspectionCapture; Set-InspectionBusy $false }
    }
})
$controls.CaptureButton.Add_Click({
    try {
        $path = Join-Path ([IO.Path]::GetTempPath()) ('WinDefState-environment-' + [guid]::NewGuid().ToString('N') + '.json')
        $errorPath = $path + '.error'
        $engine = (Join-Path $script:InspectionRoot 'WinDefState.Environment.ps1').Replace("'", "''")
        $escapedPath = $path.Replace("'", "''")
        $escapedError = $errorPath.Replace("'", "''")
        $command = "try { & '$engine' -OutputPath '$escapedPath' | Out-Null; exit 0 } catch { [IO.File]::WriteAllText('$escapedError', `$_.Exception.Message); exit 1 }"
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $start.Arguments = '-NoProfile -NonInteractive -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
        $start.UseShellExecute = $false; $start.CreateNoWindow = $true; $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        $process = [Diagnostics.Process]::Start($start)
        $script:InspectionCapture = [pscustomobject]@{ Process = $process; Path = $path; ErrorPath = $errorPath; Started = Get-Date; Cancelled = $false }
        Set-InspectionBusy $true
        $controls.StatusText.Text = 'Capturing environment...'
        $timer.Start()
    } catch { $controls.StatusText.Text = $_.Exception.Message; Set-InspectionBusy $false }
})
$controls.CancelButton.Add_Click({
    if ($null -ne $script:InspectionCapture -and -not $script:InspectionCapture.Process.HasExited) {
        $script:InspectionCapture.Cancelled = $true
        $script:InspectionCapture.Process.Kill()
        $controls.StatusText.Text = 'Capture cancelled. Previous baseline retained.'
    }
})
$controls.OpenButton.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = 'Environment baseline (*.json)|*.json'
    if ($dialog.ShowDialog($Window)) {
        try { Set-InspectionReport (Read-EnvironmentBaseline $dialog.FileName); $controls.StatusText.Text = 'Loaded ' + $dialog.FileName }
        catch { $controls.StatusText.Text = $_.Exception.Message }
    }
})
$controls.CompareButton.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Title = 'Select the earlier baseline (before testing)'
    $dialog.Filter = 'Environment baseline (*.json)|*.json'
    if ($dialog.ShowDialog($Window)) {
        try {
            $script:InspectionDiff = Compare-EnvironmentBaseline (Read-EnvironmentBaseline $dialog.FileName) $script:InspectionReport
            Update-InspectionFilter
            $controls.InspectionTabs.SelectedIndex = 2
            $controls.StatusText.Text = 'Compared with ' + [IO.Path]::GetFileName($dialog.FileName) + '. The loaded capture is the after-state.'
        } catch { $controls.StatusText.Text = $_.Exception.Message }
    }
})
$controls.SaveButton.Add_Click({
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = 'JSON (*.json)|*.json|HTML report (*.html)|*.html'
    $dialog.FileName = 'WinDefState-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    $dialog.OverwritePrompt = $false
    if ($dialog.ShowDialog($Window)) {
        try {
            $report = if ($controls.InspectionTabs.SelectedIndex -eq 2 -and $null -ne $script:InspectionDiff) { $script:InspectionDiff } else { $script:InspectionReport }
            $format = if ($dialog.FilterIndex -eq 2) { 'Html' } else { 'Json' }
            $saved = Export-EnvironmentReport $report $dialog.FileName $format
            $controls.StatusText.Text = 'Saved ' + $saved
        } catch { $controls.StatusText.Text = $_.Exception.Message }
    }
})
$controls.SearchBox.Add_TextChanged({ Update-InspectionFilter })
$controls.InspectionTabs.Add_SelectionChanged({
    param($eventControl, $selectionEvent)
    if ($selectionEvent.Source -eq $controls.InspectionTabs) { Update-InspectionFilter }
})
$Window.Add_PreviewKeyDown({
    param($eventControl, $keyEvent)
    if ($keyEvent.Key -eq [Windows.Input.Key]::F -and [Windows.Input.Keyboard]::Modifiers -eq [Windows.Input.ModifierKeys]::Control) {
        $null = $controls.SearchBox.Focus()
        $controls.SearchBox.SelectAll()
        $keyEvent.Handled = $true
    }
})
$controls.SectionList.Add_SelectionChanged({
    Set-InspectionColumns
    Update-InspectionFilter
    if ($null -ne $controls.SectionList.SelectedItem) {
        $section = $controls.SectionList.SelectedItem.Section
        $controls.DetailBox.Text = $section.Name + ' / ' + $section.Status + ' / ' + $section.DurationMs + ' ms' + [Environment]::NewLine + $section.Error
    }
})
$controls.InventoryGrid.Add_SelectionChanged({ if ($null -ne $controls.InventoryGrid.SelectedItem) { $controls.DetailBox.Text = ConvertTo-Json -InputObject $controls.InventoryGrid.SelectedItem.Data -Depth 14 } })
$controls.HealthGrid.Add_SelectionChanged({ if ($null -ne $controls.HealthGrid.SelectedItem) { $controls.DetailBox.Text = ConvertTo-Json -InputObject $controls.HealthGrid.SelectedItem -Depth 8 } })
$controls.DiffGrid.Add_SelectionChanged({ if ($null -ne $controls.DiffGrid.SelectedItem) { $controls.DetailBox.Text = ConvertTo-Json -InputObject $controls.DiffGrid.SelectedItem -Depth 14 } })
$Window.Add_Closing({
    $timer.Stop()
    if ($null -ne $script:InspectionCapture) {
        if (-not $script:InspectionCapture.Process.HasExited) { $script:InspectionCapture.Process.Kill(); $null = $script:InspectionCapture.Process.WaitForExit(3000) }
        if ($script:InspectionCapture.Process.HasExited) { Clear-InspectionCapture }
    }
})
if ($BaselinePath) { Set-InspectionReport (Read-EnvironmentBaseline $BaselinePath) }
if ($ValidateOnly) {
    if ($PreviewPath) {
        $null = [Windows.Interop.WindowInteropHelper]::new($Window).EnsureHandle()
        $Window.Content.Measure([Windows.Size]::new($Window.Width, $Window.Height))
        $Window.Content.Arrange([Windows.Rect]::new(0, 0, $Window.Width, $Window.Height))
        $Window.Content.UpdateLayout()
        $null = $Window.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::Loaded)
        $Window.Content.UpdateLayout()
        $bitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new([int]$Window.Width, [int]$Window.Height, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
        $bitmap.Render($Window.Content)
        $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
        $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $stream = [IO.File]::Open($PreviewPath, [IO.FileMode]::CreateNew)
        try { $encoder.Save($stream) } finally { $stream.Dispose() }
    }
    $Window.Close()
    Write-Output 'Inspection dashboard WPF validation passed.'
    return
}
$null = $Window.ShowDialog()

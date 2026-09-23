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
 Title="WinDefState | Environment inspection" Width="1260" Height="880" MinWidth="960" MinHeight="680" WindowStartupLocation="CenterScreen"
 Background="#F3F6FA" FontFamily="Segoe UI" FontSize="14" Foreground="#172C3E" UseLayoutRounding="True">
 <Window.Resources>
  <Style TargetType="Button"><Setter Property="Padding" Value="16,10"/><Setter Property="Margin" Value="0,0,10,0"/><Setter Property="Background" Value="White"/><Setter Property="Foreground" Value="#172C3E"/><Setter Property="BorderBrush" Value="#B3C3D1"/><Setter Property="Cursor" Value="Hand"/></Style>
  <Style TargetType="DataGrid"><Setter Property="AutoGenerateColumns" Value="False"/><Setter Property="IsReadOnly" Value="True"/><Setter Property="CanUserAddRows" Value="False"/><Setter Property="CanUserDeleteRows" Value="False"/><Setter Property="HeadersVisibility" Value="Column"/><Setter Property="GridLinesVisibility" Value="Horizontal"/><Setter Property="HorizontalGridLinesBrush" Value="#E3EAF1"/><Setter Property="BorderThickness" Value="0"/><Setter Property="RowBackground" Value="White"/><Setter Property="AlternatingRowBackground" Value="#F8FAFC"/><Setter Property="SelectionMode" Value="Single"/><Setter Property="EnableRowVirtualization" Value="True"/><Setter Property="EnableColumnVirtualization" Value="True"/><Setter Property="MinRowHeight" Value="38"/></Style>
  <Style TargetType="DataGridColumnHeader"><Setter Property="Background" Value="#E8EEF5"/><Setter Property="Foreground" Value="#425870"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Padding" Value="10,12"/><Setter Property="BorderThickness" Value="0"/></Style>
  <Style TargetType="DataGridCell"><Setter Property="Padding" Value="10,8"/><Setter Property="BorderThickness" Value="0"/></Style>
  <Style TargetType="TabItem"><Setter Property="Padding" Value="18,10"/><Setter Property="FontWeight" Value="SemiBold"/></Style>
 </Window.Resources>
 <Grid Background="#F3F6FA">
  <Grid.RowDefinitions><RowDefinition Height="126"/><RowDefinition Height="*"/><RowDefinition Height="44"/></Grid.RowDefinitions>
  <Border Background="#132C40" Padding="30,20">
   <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
    <StackPanel><TextBlock Text="WINDEFSTATE / INSPECT" Foreground="#79D5C5" FontSize="12" FontWeight="Bold"/><TextBlock Text="Know your starting point." Foreground="White" FontSize="30" FontWeight="SemiBold" Margin="0,5,0,0"/><TextBlock Text="Capture the environment. Review protection. Compare what changed." Foreground="#BCD0DF"/></StackPanel>
    <Border Grid.Column="1" Background="#264556" CornerRadius="6" Padding="14,8" VerticalAlignment="Top"><TextBlock Text="READ ONLY" Foreground="#A5EAD8" FontSize="12" FontWeight="Bold"/></Border>
   </Grid>
  </Border>
  <Grid Grid.Row="1" Margin="28,20,28,16">
   <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="6"/><RowDefinition Height="142"/></Grid.RowDefinitions>
   <DockPanel Margin="0,0,0,14"><StackPanel DockPanel.Dock="Right" Orientation="Horizontal"><Button x:Name="CaptureButton" Content="_Capture environment" Background="#147B74" Foreground="White" BorderBrush="#147B74"/><Button x:Name="CancelButton" Content="Cancel" IsEnabled="False" Margin="0"/></StackPanel><StackPanel><TextBlock x:Name="HostText" Text="Environment baseline" FontSize="22" FontWeight="SemiBold"/><TextBlock x:Name="CaptureText" Text="Capture this PC or open a saved JSON baseline." Foreground="#425870" Margin="0,4,12,0" TextWrapping="Wrap"/></StackPanel></DockPanel>
   <UniformGrid Grid.Row="1" Columns="4" Margin="0,0,0,16">
    <Border Background="White" BorderBrush="#D4DFE8" BorderThickness="1" CornerRadius="8" Padding="16,12" Margin="0,0,12,0"><StackPanel><TextBlock x:Name="ListenersCount" Text="--" FontSize="28" FontWeight="SemiBold"/><TextBlock Text="TCP listeners / UDP endpoints" Foreground="#425870" FontSize="12"/></StackPanel></Border>
    <Border Background="White" BorderBrush="#D4DFE8" BorderThickness="1" CornerRadius="8" Padding="16,12" Margin="0,0,12,0"><StackPanel><TextBlock x:Name="RulesCount" Text="--" FontSize="28" FontWeight="SemiBold"/><TextBlock Text="Effective firewall rules" Foreground="#425870" FontSize="12"/></StackPanel></Border>
    <Border Background="White" BorderBrush="#D4DFE8" BorderThickness="1" CornerRadius="8" Padding="16,12" Margin="0,0,12,0"><StackPanel><TextBlock x:Name="AttentionCount" Text="--" FontSize="28" FontWeight="SemiBold" Foreground="#8C4700"/><TextBlock Text="Health checks to review" Foreground="#425870" FontSize="12"/></StackPanel></Border>
    <Border Background="White" BorderBrush="#D4DFE8" BorderThickness="1" CornerRadius="8" Padding="16,12"><StackPanel><TextBlock x:Name="UnknownCount" Text="--" FontSize="28" FontWeight="SemiBold" Foreground="#5D4BB0"/><TextBlock Text="Unreadable inventory sections" Foreground="#425870" FontSize="12"/></StackPanel></Border>
   </UniformGrid>
   <DockPanel Grid.Row="2" Margin="0,0,0,12"><StackPanel DockPanel.Dock="Left" Orientation="Horizontal"><Button x:Name="OpenButton" Content="_Open baseline"/><Button x:Name="CompareButton" Content="_Compare with before..." IsEnabled="False"/><Button x:Name="SaveButton" Content="_Export..." IsEnabled="False"/></StackPanel><TextBlock DockPanel.Dock="Left" Text="Search" Margin="8,0,10,0" VerticalAlignment="Center"/><TextBox x:Name="SearchBox" Padding="10,8" VerticalContentAlignment="Center" ToolTip="Search all fields in the current view" AutomationProperties.Name="Search checks or inventory"/></DockPanel>
   <TabControl x:Name="InspectionTabs" Grid.Row="3" Background="White" BorderBrush="#D4DFE8">
    <TabItem Header="Environment inventory"><Grid Margin="12"><Grid.ColumnDefinitions><ColumnDefinition Width="235"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
     <ListBox x:Name="SectionList" BorderBrush="#D4DFE8" ScrollViewer.HorizontalScrollBarVisibility="Disabled" AutomationProperties.Name="Inventory sections"><ListBox.ItemTemplate><DataTemplate><TextBlock Text="{Binding Label}" TextWrapping="Wrap" Padding="6,8"/></DataTemplate></ListBox.ItemTemplate></ListBox>
     <DataGrid x:Name="InventoryGrid" Grid.Column="2" AutomationProperties.Name="Inventory records"><DataGrid.Columns><DataGridTextColumn Header="Identity" Binding="{Binding Key}" Width="2*"/><DataGridTextColumn Header="Observed configuration" Binding="{Binding Summary}" Width="3*"/></DataGrid.Columns></DataGrid>
    </Grid></TabItem>
    <TabItem Header="Protection health"><DataGrid x:Name="HealthGrid" Margin="12" AutomationProperties.Name="Protection checks"><DataGrid.RowStyle><Style TargetType="DataGridRow"><Style.Triggers><DataTrigger Binding="{Binding Status}" Value="Attention"><Setter Property="Foreground" Value="#8C4700"/></DataTrigger><DataTrigger Binding="{Binding Status}" Value="Unknown"><Setter Property="Foreground" Value="#5D4BB0"/></DataTrigger></Style.Triggers></Style></DataGrid.RowStyle><DataGrid.Columns><DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="105"/><DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="110"/><DataGridTextColumn Header="Check" Binding="{Binding Name}" Width="2*"/><DataGridTextColumn Header="Observed value" Binding="{Binding Value}" Width="2*"/></DataGrid.Columns></DataGrid></TabItem>
    <TabItem Header="Changes since baseline"><DataGrid x:Name="DiffGrid" Margin="12" AutomationProperties.Name="Environment changes"><DataGrid.Columns><DataGridTextColumn Header="Change" Binding="{Binding Change}" Width="100"/><DataGridTextColumn Header="Section" Binding="{Binding Section}" Width="210"/><DataGridTextColumn Header="Identity" Binding="{Binding Key}" Width="*"/></DataGrid.Columns></DataGrid></TabItem>
   </TabControl>
   <GridSplitter Grid.Row="4" HorizontalAlignment="Stretch" Background="Transparent"/>
   <Border Grid.Row="5" Background="#E8EEF5" CornerRadius="6" Padding="12"><DockPanel><TextBlock DockPanel.Dock="Top" Text="EVIDENCE &amp; CONTEXT" FontSize="11" FontWeight="Bold" Foreground="#425870" Margin="0,0,0,6"/><TextBox x:Name="DetailBox" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Background="Transparent" BorderThickness="0" FontFamily="Consolas" FontSize="12" Text="Select a record to inspect its evidence. Local listeners do not prove remote reachability. Keep a VM checkpoint or disk backup for full recovery."/></DockPanel></Border>
  </Grid>
  <Border Grid.Row="2" Background="#E6EDF4" Padding="28,10"><DockPanel><ProgressBar x:Name="CaptureProgress" DockPanel.Dock="Right" Width="120" Height="6" IsIndeterminate="False" Margin="12,0,0,0"/><TextBlock x:Name="StatusText" Text="Ready / No system settings are changed by this dashboard." Foreground="#425870" TextTrimming="CharacterEllipsis"/></DockPanel></Border>
 </Grid>
</Window>
'@
$reader = New-Object Xml.XmlNodeReader ([xml]$inspectionXaml)
try { $Window = [Windows.Markup.XamlReader]::Load($reader) } finally { $reader.Close() }
$controls = @{}
foreach ($name in @('CaptureButton', 'CancelButton', 'HostText', 'CaptureText', 'ListenersCount', 'RulesCount', 'AttentionCount', 'UnknownCount', 'OpenButton', 'CompareButton', 'SaveButton', 'SearchBox', 'InspectionTabs', 'SectionList', 'InventoryGrid', 'HealthGrid', 'DiffGrid', 'DetailBox', 'CaptureProgress', 'StatusText')) {
    $controls[$name] = $Window.FindName($name)
    if ($null -eq $controls[$name]) { throw "Missing inspection control: $name" }
}
foreach ($name in @('InventoryGrid', 'HealthGrid', 'DiffGrid', 'SectionList')) { $controls[$name].FontWeight = [Windows.FontWeights]::Normal }

function Update-InspectionFilter {
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
    $controls.CaptureText.Text = 'Captured ' + $Report.CapturedAtUtc + ' / Elevated: ' + $Report.Elevated
    $counts = foreach ($id in @('network.tcp', 'network.udp')) {
        $section = @($Report.Sections | Where-Object Id -eq $id)
        if ($section.Count -eq 1 -and $section[0].Status -eq 'Captured') { [string]@($section[0].Items).Count } else { '?' }
    }
    $controls.ListenersCount.Text = $counts -join ' / '
    $rules = @($Report.Sections | Where-Object Id -eq 'firewall.rules')
    $controls.RulesCount.Text = if ($rules.Count -eq 1 -and $rules[0].Status -eq 'Captured') { [string]@($rules[0].Items).Count } else { '?' }
    $controls.AttentionCount.Text = [string]$Report.Health.Summary.Attention
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
            $controls.StatusText.Text = 'Capture complete. Export JSON to retain this baseline before testing.'
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
        $controls.StatusText.Text = 'Capturing local environment. Some providers require an elevated session; unreadable sections remain visible.'
        $timer.Start()
    } catch { $controls.StatusText.Text = $_.Exception.Message; Set-InspectionBusy $false }
})
$controls.CancelButton.Add_Click({
    if ($null -ne $script:InspectionCapture -and -not $script:InspectionCapture.Process.HasExited) {
        $script:InspectionCapture.Cancelled = $true
        $script:InspectionCapture.Process.Kill()
        $controls.StatusText.Text = 'Read-only capture cancelled. The previous baseline remains loaded.'
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
    $dialog.Title = 'Choose the BEFORE baseline; the currently loaded capture is AFTER'
    $dialog.Filter = 'Environment baseline (*.json)|*.json'
    if ($dialog.ShowDialog($Window)) {
        try {
            $script:InspectionDiff = Compare-EnvironmentBaseline (Read-EnvironmentBaseline $dialog.FileName) $script:InspectionReport
            Update-InspectionFilter
            $controls.InspectionTabs.SelectedIndex = 2
            $controls.StatusText.Text = 'Comparison: ' + @($script:InspectionDiff.Changes).Count + ' differences or unreadable sections. Export saves this comparison while the Changes tab is active.'
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

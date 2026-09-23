#Requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$demoPath = Join-Path $PSScriptRoot 'fixtures/inspection-demo.json'
. (Join-Path $repositoryRoot 'WinDefState.Inspect.Gui.ps1') -ValidateOnly -BaselinePath $demoPath

if ($controls.InventoryGrid.Items.Count -ne 3 -or $controls.InventoryGrid.Columns.Count -ne 7) {
    throw 'The initial TCP inventory did not populate its rows and columns.'
}
$controls.SearchBox.Text = 'labservice 8080'
if ($controls.InventoryGrid.Items.Count -ne 1) { throw 'Multi-word inventory search failed.' }
$controls.InventoryGrid.SelectedIndex = 0
if ($controls.DetailBox.Text -notmatch 'LabService') { throw 'Selected record evidence was not displayed.' }
$controls.SearchBox.Text = 'no-such-inventory-value'
if ($controls.InventoryGrid.Items.Count -ne 0) { throw 'Search retained unrelated rows.' }
$controls.SearchBox.Text = ''
$controls.SectionList.SelectedIndex = 2
if ($controls.InventoryGrid.Items.Count -ne 1 -or $controls.InventoryGrid.Columns[0].Header -ne 'Display Name') {
    throw 'Changing sections did not update the inventory columns.'
}
$controls.InspectionTabs.SelectedIndex = 1
if ($controls.HealthGrid.Items.Count -ne $script:InspectionReport.Health.Checks.Count) { throw 'Health rows were not populated.' }
$controls.SearchBox.Text = 'tamper'
if ($controls.HealthGrid.Items.Count -ne 1) { throw 'Health search failed.' }
$controls.SearchBox.Text = ''
$before = Read-EnvironmentBaseline $demoPath
$script:InspectionReport.Sections[2].Items[0].Data.Action = 'Block'
$script:InspectionDiff = Compare-EnvironmentBaseline $before $script:InspectionReport
Update-InspectionFilter
if (@($controls.DiffGrid.Items | Where-Object Change -eq Changed).Count -ne 1) { throw 'The comparison grid did not display the changed rule.' }
$script:InspectionReport.Health.Summary.Unknown = 2
Set-InspectionReport $script:InspectionReport
if ($controls.AttentionCount.Text -ne '2') { throw 'Unknown health checks were omitted from the review count.' }
Write-Output 'Inspection dashboard search, section, evidence and comparison checks passed.'

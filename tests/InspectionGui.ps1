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
if ($controls.RowsText.Text -ne '3 of 3 records' -or $controls.InventoryEmpty.Visibility -ne 'Collapsed') { throw 'The initial inventory count or empty state is incorrect.' }
$controls.SearchBox.Text = 'labservice 8080'
if ($controls.InventoryGrid.Items.Count -ne 1) { throw 'Multi-word inventory search failed.' }
if ($controls.RowsText.Text -ne '1 of 3 records') { throw 'The filtered row count is incorrect.' }
$controls.InventoryGrid.SelectedIndex = 0
if ($controls.DetailBox.Text -notmatch 'LabService') { throw 'Selected record evidence was not displayed.' }
$controls.SearchBox.Text = 'no-such-inventory-value'
if ($controls.InventoryGrid.Items.Count -ne 0) { throw 'Search retained unrelated rows.' }
if ($controls.InventoryEmpty.Visibility -ne 'Visible' -or $controls.InventoryEmpty.Text -ne 'No matching records.') { throw 'The filtered empty state is missing.' }
$controls.SearchBox.Text = ''
$controls.SectionList.SelectedIndex = 2
if ($controls.InventoryGrid.Items.Count -ne 1 -or $controls.InventoryGrid.Columns[0].Header -ne 'Display Name') {
    throw 'Changing sections did not update the inventory columns.'
}
$controls.InspectionTabs.SelectedIndex = 1
if ($controls.HealthGrid.Items.Count -ne $script:InspectionReport.Health.Checks.Count) { throw 'Health rows were not populated.' }
if ($controls.DetailBox.Text -ne 'Select a row to view its details.') { throw 'Switching tabs retained details from another view.' }
$controls.SearchBox.Text = 'tamper'
if ($controls.HealthGrid.Items.Count -ne 1) { throw 'Health search failed.' }
$controls.SearchBox.Text = ''
$controls.InspectionTabs.SelectedIndex = 2
if ($controls.DiffEmpty.Visibility -ne 'Visible' -or $controls.SaveButton.Content -ne '_Save baseline...') { throw 'A comparison was offered before one was loaded.' }
$before = Read-EnvironmentBaseline $demoPath
$script:InspectionReport.Sections[2].Items[0].Data.Action = 'Block'
$script:InspectionDiff = Compare-EnvironmentBaseline $before $script:InspectionReport
Update-InspectionFilter
if (@($controls.DiffGrid.Items | Where-Object Change -eq Changed).Count -ne 1) { throw 'The comparison grid did not display the changed rule.' }
if ($controls.SaveButton.Content -ne '_Save comparison...' -or $controls.DiffEmpty.Visibility -ne 'Collapsed') { throw 'The active comparison was not identified as the save target.' }
$controls.InspectionTabs.SelectedIndex = 0
if ($controls.SaveButton.Content -ne '_Save baseline...') { throw 'Returning to inventory did not restore the baseline save target.' }
$script:InspectionReport.Health.Summary.Unknown = 2
$script:InspectionReport.Sections[0].Status = 'Unknown'
$script:InspectionReport.Sections[0].Items = @()
$script:InspectionReport.Sections[0].Error = 'Access denied'
Set-InspectionReport $script:InspectionReport
if ($controls.AttentionCount.Text -ne '2') { throw 'Unknown health checks were omitted from the review count.' }
if ($controls.InventoryEmpty.Text -ne 'Section unavailable. See Details.') { throw 'An unreadable section was presented as an empty successful capture.' }
Write-Output 'Inspection dashboard search, section, evidence and comparison checks passed.'

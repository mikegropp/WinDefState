[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$settingsPath = Join-Path $repositoryRoot 'PSScriptAnalyzerSettings.psd1'
$scriptPaths = @(
    Join-Path $repositoryRoot 'WinDefState.ps1'
    Join-Path $repositoryRoot 'WinDefState.Gui.ps1'
    Join-Path $repositoryRoot 'WinDefState.Health.ps1'
    Join-Path $repositoryRoot 'WinDefState.Environment.ps1'
    Join-Path $repositoryRoot 'WinDefState.Inspect.Gui.ps1'
    Join-Path $repositoryRoot 'build/Build-Release.ps1'
    Join-Path $repositoryRoot 'tests/WindowsSnapshot.ps1'
    Join-Path $repositoryRoot 'tests/InspectionGui.ps1'
)

if ($null -eq (Get-Command Invoke-ScriptAnalyzer -ErrorAction SilentlyContinue)) {
    throw 'PSScriptAnalyzer is required. Install version 1.24.0, then rerun tests\Analyze.ps1.'
}

$findings = @(
    foreach ($scriptPath in $scriptPaths) {
        Invoke-ScriptAnalyzer -Path $scriptPath -Settings $settingsPath
    }
)

if ($findings.Count -gt 0) {
    $findings |
        Sort-Object -Property ScriptName, Line, RuleName |
        Format-Table -AutoSize -Property RuleName, Severity, ScriptName, Line, Message |
        Out-Host
    throw "PSScriptAnalyzer found $($findings.Count) blocking issue(s)."
}

Write-Host 'PSScriptAnalyzer checks passed.'

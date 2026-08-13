[CmdletBinding()]
param(
    [string]$OutputRoot = (Join-Path $env:TEMP ("WinDefState-Integration-{0}" -f [guid]::NewGuid().ToString('N')))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($env:OS -ne 'Windows_NT') {
    throw 'WindowsSnapshot.ps1 must run on Windows.'
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'WindowsSnapshot.ps1 requires an elevated Windows PowerShell session.'
}

$enginePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'WinDefState.ps1'
$invocationStartedAtUtc = (Get-Date).ToUniversalTime()
$invocationStopwatch = [Diagnostics.Stopwatch]::StartNew()
try {
    & $enginePath -Command Snapshot -StateRoot $OutputRoot -ConsoleReport None
} finally {
    $invocationStopwatch.Stop()
}
$invocationCompletedAtUtc = (Get-Date).ToUniversalTime()
$invocationDurationMs = [Math]::Round($invocationStopwatch.Elapsed.TotalMilliseconds, 1)

$snapshot = Get-ChildItem -LiteralPath (Join-Path $OutputRoot 'snapshots') -Filter '*.json' -File |
    Sort-Object -Property LastWriteTimeUtc -Descending |
    Select-Object -First 1
if ($null -eq $snapshot) {
    throw 'The integration run did not create a snapshot JSON file.'
}

$state = Get-Content -LiteralPath $snapshot.FullName -Raw | ConvertFrom-Json
if ([int]$state.SchemaVersion -ne 2) {
    throw "Expected snapshot schema 2, received '$($state.SchemaVersion)'."
}
if (@($state.Settings).Count -ne 94) {
    throw "Expected 94 setting entries, received '$(@($state.Settings).Count)'."
}
if (-not $state.PSObject.Properties['CaptureMetrics'] -or $null -eq $state.CaptureMetrics) {
    throw 'Snapshot capture metrics are missing.'
}
if (-not $state.PSObject.Properties['Producer'] -or $null -eq $state.Producer -or [string]::IsNullOrWhiteSpace([string]$state.Producer.ScriptSha256)) {
    throw 'Snapshot producer provenance is missing.'
}
$engineHash = (Get-FileHash -LiteralPath $enginePath -Algorithm SHA256).Hash
if (-not [string]::Equals([string]$state.Producer.ScriptSha256, $engineHash, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw 'Snapshot producer hash does not match the engine that created it.'
}

$providerQueries = @($state.CaptureMetrics.ProviderQueries)
$settingTimings = @($state.CaptureMetrics.Settings)
$snapshotIds = @($state.Settings | ForEach-Object { [string]$_.Id } | Sort-Object -Unique)
$timingIds = @($settingTimings | ForEach-Object { [string]$_.Id } | Sort-Object -Unique)
if ($settingTimings.Count -ne 94 -or $timingIds.Count -ne 94) {
    throw "Expected one unique timing for each of 94 settings, received '$($settingTimings.Count)' timing row(s) across '$($timingIds.Count)' unique ID(s)."
}
$missingTimingIds = @(
    Compare-Object -ReferenceObject $snapshotIds -DifferenceObject $timingIds |
        Where-Object { [string]$_.SideIndicator -eq '<=' } |
        ForEach-Object { [string]$_.InputObject }
)
if ($missingTimingIds.Count -gt 0) {
    throw "Snapshot settings are missing timing rows: $($missingTimingIds -join ', ')"
}
if ([int]$state.CaptureMetrics.ProviderQueryCount -ne $providerQueries.Count) {
    throw "ProviderQueryCount '$($state.CaptureMetrics.ProviderQueryCount)' does not match the '$($providerQueries.Count)' persisted provider timing row(s)."
}
if ([double]$state.CaptureMetrics.DurationMs -le 0) {
    throw 'Snapshot capture duration was not recorded as a positive value.'
}
$captureDurationMs = [double]$state.CaptureMetrics.DurationMs
if ($invocationDurationMs -lt $captureDurationMs) {
    throw "End-to-end invocation duration '$invocationDurationMs' ms is shorter than recorded capture duration '$captureDurationMs' ms."
}
$nonCaptureOverheadMs = [Math]::Round(($invocationDurationMs - $captureDurationMs), 1)
if ([int]$state.CaptureMetrics.CacheHitCount -le 0) {
    throw 'The full snapshot did not report any shared provider reads. The provider cache may not have been exercised.'
}
$duplicateProviderKeys = @(
    $providerQueries |
        Where-Object { -not ([string]$_.Key).StartsWith('cleanup:', [System.StringComparison]::OrdinalIgnoreCase) } |
        Group-Object -Property Key |
        Where-Object { $_.Count -gt 1 }
)
if ($duplicateProviderKeys.Count -gt 0) {
    throw "Shared provider keys were queried more than once: $(@($duplicateProviderKeys.Name) -join ', ')"
}
$defenderQueries = @($providerQueries | Where-Object { [string]$_.Key -eq 'defender.preferences' })
$serviceQueries = @($providerQueries | Where-Object { [string]$_.Key -eq 'services' })
if ($defenderQueries.Count -ne 1) {
    throw "Expected exactly one shared Defender preference query, received '$($defenderQueries.Count)'."
}
if ($serviceQueries.Count -ne 1) {
    throw "Expected exactly one shared tracked-service query, received '$($serviceQueries.Count)'."
}

$slowestSettings = @($settingTimings | Sort-Object -Property DurationMs -Descending | Select-Object -First 10)
$slowestProviderQueries = @($providerQueries | Sort-Object -Property DurationMs -Descending | Select-Object -First 10)
$performanceSummaryPath = Join-Path $OutputRoot 'performance-summary.json'
$performanceSummary = [PSCustomObject]@{
    CapturedAtUtc        = [string]$state.CapturedAtUtc
    InvocationStartedAtUtc = $invocationStartedAtUtc.ToString('o')
    InvocationCompletedAtUtc = $invocationCompletedAtUtc.ToString('o')
    EngineSha256        = [string]$state.Producer.ScriptSha256
    PowerShellVersion   = [string]$state.Producer.PowerShellVersion
    InvocationDurationMs = $invocationDurationMs
    CaptureDurationMs   = $captureDurationMs
    NonCaptureOverheadMs = $nonCaptureOverheadMs
    SettingCount        = @($state.Settings).Count
    ProviderQueryCount  = [int]$state.CaptureMetrics.ProviderQueryCount
    CacheHitCount       = [int]$state.CaptureMetrics.CacheHitCount
    SlowestSettings     = @($slowestSettings)
    SlowestProviders    = @($slowestProviderQueries)
}
$performanceSummary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $performanceSummaryPath -Encoding UTF8

Write-Host ('Snapshot: {0}' -f $snapshot.FullName)
Write-Host ('Performance summary: {0}' -f $performanceSummaryPath)
Write-Host ('Engine SHA-256: {0}' -f [string]$state.Producer.ScriptSha256)
Write-Host ('End-to-end duration: {0:N2} seconds' -f ($invocationDurationMs / 1000))
Write-Host ('Capture duration: {0:N2} seconds' -f ($captureDurationMs / 1000))
Write-Host ('Non-capture overhead: {0:N2} seconds' -f ($nonCaptureOverheadMs / 1000))
Write-Host ('Provider queries: {0}' -f [int]$state.CaptureMetrics.ProviderQueryCount)
Write-Host ('Shared reads reused: {0}' -f [int]$state.CaptureMetrics.CacheHitCount)
Write-Host 'Slowest settings:'
foreach ($timing in $slowestSettings) {
    Write-Host ('  {0}: {1:N1} ms' -f [string]$timing.Id, [double]$timing.DurationMs)
}
Write-Host 'Slowest provider queries:'
foreach ($timing in $slowestProviderQueries) {
    Write-Host ('  {0}: {1:N1} ms' -f [string]$timing.Key, [double]$timing.DurationMs)
}

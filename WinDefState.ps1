<#
.SYNOPSIS
Snapshots, loosens, and restores supported Windows defense settings.

.DESCRIPTION
WinDefState is designed around three elevated PowerShell calls. Use the call you need at that point in the workflow:

1. Snapshot only: capture the current host state and write a report.
2. Permissive: snapshot first, then apply supported permissive test settings.
3. Restore: restore and verify from the saved operation journal, or from an explicit snapshot path.

The GUI is optional. The safest and simplest workflow is this single script plus these three commands.

Use powershell.exe with -ExecutionPolicy Bypass when running a freshly downloaded copy on hosts that block direct .ps1 execution.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Snapshot

Capture the current state only.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Permissive

Capture the current state, write current-operation.json, then apply permissive settings.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Restore

Restore from current-operation.json after a previous Permissive run.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Restore -SnapshotPath .\state\snapshots\HOST-20260420-120000.json

Restore from a specific snapshot file instead of the current operation journal.

.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Snapshot -IncludeCategory defender,firewall

Capture only the Defender and firewall categories. Setting ID wildcards such as rdp.* are also supported.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [ValidateSet('Snapshot', 'Permissive', 'Restore')]
    [string]$Command,

    [string]$SnapshotPath,

    [string]$StateRoot,

    [string[]]$IncludeId,

    [string[]]$ExcludeId,

    [string[]]$IncludeCategory,

    [string[]]$ExcludeCategory,

    [switch]$EmitProgress,

    [switch]$AllowDifferentComputer,

    [string]$CancellationPath,

    [string]$MutationApprovalPath,

    [ValidateSet('Summary', 'Full', 'None')]
    [string]$ConsoleReport = 'Summary'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:IsDotSourced = $MyInvocation.InvocationName -eq '.'

$scriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptRoot)) {
    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $scriptRoot = Split-Path -Parent $PSCommandPath
    } elseif ($null -ne $MyInvocation.MyCommand -and -not [string]::IsNullOrWhiteSpace([string]$MyInvocation.MyCommand.Path)) {
        $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        $scriptRoot = (Get-Location).Path
    }
}

if ([string]::IsNullOrWhiteSpace($StateRoot)) {
    $StateRoot = if (-not [string]::IsNullOrWhiteSpace([string]$env:ProgramData)) {
        Join-Path $env:ProgramData 'WinDefState'
    } else {
        Join-Path $scriptRoot 'state'
    }
}

$script:WinDefStateScriptPath = if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
    [IO.Path]::GetFullPath($PSCommandPath)
} else {
    [IO.Path]::GetFullPath((Join-Path $scriptRoot 'WinDefState.ps1'))
}
$script:WinDefStateRuntimeInfo = $null
$script:WinDefStateCommandCache = @{}
$script:WinDefStateDefinitionCache = $null
$script:WinDefStateDefinitionMapCache = $null
$script:WinDefStateSnapshotAssetCache = @{}

#region Core runtime, persistence, and operation state

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Run this script from an elevated PowerShell session.'
    }
}

function Enter-WinDefStateOperationLock {
    param([int]$TimeoutSeconds = 2)

    $mutex = New-Object System.Threading.Mutex($false, 'Global\WinDefState.Operation')
    $acquired = $false
    try {
        try {
            $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds))
        } catch [System.Threading.AbandonedMutexException] {
            $acquired = $true
        }

        if (-not $acquired) {
            throw 'Another WinDefState operation is already running on this computer.'
        }

        [PSCustomObject]@{
            Mutex    = $mutex
            Acquired = $true
        }
    } catch {
        $mutex.Dispose()
        throw
    }
}

function Exit-WinDefStateOperationLock {
    param([AllowNull()] [object]$Lock)

    if ($null -eq $Lock -or -not $Lock.PSObject.Properties['Mutex'] -or $null -eq $Lock.Mutex) {
        return
    }

    try {
        if ($Lock.PSObject.Properties['Acquired'] -and [bool]$Lock.Acquired) {
            $Lock.Mutex.ReleaseMutex()
        }
    } finally {
        $Lock.Mutex.Dispose()
    }
}

function Invoke-ProtectedWinDefStateOperation {
    param([Parameter(Mandatory)] [scriptblock]$ScriptBlock)

    Assert-Administrator
    Protect-StateRoot -Path $StateRoot
    Clear-SnapshotAssetCache
    $operationLock = Enter-WinDefStateOperationLock
    try {
        & $ScriptBlock
    } finally {
        Exit-WinDefStateOperationLock -Lock $operationLock
    }
}

function ConvertTo-OperationProgressField {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    (([string]$Value) -replace '[\r\n|]+', ' ').Trim()
}

function Write-OperationProgress {
    param(
        [Parameter(Mandatory)] [string]$Phase,
        [Parameter(Mandatory)] [int]$Current,
        [Parameter(Mandatory)] [int]$Total,
        [string]$Id
    )

    $safePhase = ConvertTo-OperationProgressField -Value $Phase
    $safeId = ConvertTo-OperationProgressField -Value $Id
    $verboseText = if ([string]::IsNullOrWhiteSpace($safeId)) {
        "[{0} {1}/{2}]" -f $safePhase, $Current, $Total
    } else {
        "[{0} {1}/{2}] {3}" -f $safePhase, $Current, $Total, $safeId
    }
    Write-Verbose $verboseText

    $activity = "WinDefState $safePhase"
    $status = if ([string]::IsNullOrWhiteSpace($safeId)) { $verboseText } else { $safeId }
    $percentComplete = if ($Total -gt 0) {
        [Math]::Min(100, [Math]::Max(0, [int](($Current / [double]$Total) * 100)))
    } else {
        0
    }
    Write-Progress -Id 0 -Activity $activity -Status $status -PercentComplete $percentComplete
    if ($Total -gt 0 -and $Current -ge $Total) {
        Write-Progress -Id 0 -Activity $activity -Completed
    }

    if ($EmitProgress) {
        Write-Host ("WDS_PROGRESS|{0}|{1}|{2}|{3}" -f $safePhase, $Current, $Total, $safeId)
    }
}

function Write-OperationResult {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [string]$Value
    )

    if (-not $EmitProgress) {
        return
    }

    $safeName = ConvertTo-OperationProgressField -Value $Name
    $safeValue = ConvertTo-OperationProgressField -Value $Value
    Write-Host ("WDS_RESULT|{0}|{1}" -f $safeName, $safeValue)
}

function Test-OperationCancellationRequested {
    param([AllowNull()] [string]$Path)

    -not [string]::IsNullOrWhiteSpace($Path) -and (Test-Path -LiteralPath $Path -PathType Leaf)
}

function Assert-OperationNotCancelled {
    param(
        [AllowNull()] [string]$Path,
        [Parameter(Mandatory)] [string]$Stage
    )

    if (-not (Test-OperationCancellationRequested -Path $Path)) {
        return
    }

    $safeStage = ConvertTo-OperationProgressField -Value $Stage
    if ($EmitProgress) {
        Write-Host ("WDS_CANCELLED|{0}" -f $safeStage)
    }
    throw [System.OperationCanceledException]::new("WinDefState operation cancelled at a safe boundary: $safeStage. No defense setting was changed by the cancelled operation.")
}

function Wait-MutationApproval {
    param(
        [AllowNull()] [string]$Path,
        [Parameter(Mandatory)] [ValidateSet('Permissive', 'Restore')] [string]$Action,
        [Parameter(Mandatory)] [string]$SnapshotPath,
        [Parameter(Mandatory)] [string]$TrustedRoot,
        [AllowNull()] [string]$CancellationPath,
        [ValidateRange(1, 3600)] [int]$TimeoutSeconds = 900
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    $fullPath = [IO.Path]::GetFullPath($Path)
    $trustedRoot = [IO.Path]::GetFullPath($TrustedRoot).TrimEnd([char[]]@('\', '/'))
    $approvalParent = [IO.Path]::GetDirectoryName($fullPath).TrimEnd([char[]]@('\', '/'))
    $approvalLeaf = [IO.Path]::GetFileName($fullPath)
    if (
        -not [string]::Equals($approvalParent, $trustedRoot, [System.StringComparison]::OrdinalIgnoreCase) -or
        $approvalLeaf -notmatch '^\.review-[0-9a-fA-F]{32}\.approve$'
    ) {
        throw "Mutation approval path must be a one-time .review-<guid>.approve marker directly under the protected state root: $trustedRoot"
    }
    if (Test-Path -LiteralPath $fullPath) {
        throw "Mutation approval marker already exists before review began: $fullPath"
    }

    $safeAction = ConvertTo-OperationProgressField -Value $Action
    $safeSnapshotPath = ConvertTo-OperationProgressField -Value ([IO.Path]::GetFullPath($SnapshotPath))
    Write-Host ("WDS_REVIEW|{0}|{1}" -f $safeAction, $safeSnapshotPath)
    try {
        [Console]::Out.Flush()
    } catch {
        Write-Verbose "The host output stream could not be flushed explicitly; asynchronous output collection remains active."
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        while ($stopwatch.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
            Assert-OperationNotCancelled -Path $CancellationPath -Stage 'pre-change review'
            if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
                $decision = ([IO.File]::ReadAllText($fullPath)).Trim()
                if ([string]::IsNullOrWhiteSpace($decision)) {
                    Start-Sleep -Milliseconds 100
                    continue
                }
                if ([string]::Equals($decision, 'CANCEL', [System.StringComparison]::Ordinal)) {
                    Remove-Item -LiteralPath $fullPath -Force -ErrorAction SilentlyContinue
                    Write-Host 'WDS_CANCELLED|pre-change review rejected'
                    throw [System.OperationCanceledException]::new('WinDefState pre-change review was rejected. The baseline remains saved, but no defense setting was changed.')
                }
                if (-not [string]::Equals($decision, 'APPROVE', [System.StringComparison]::Ordinal)) {
                    throw "Mutation approval marker contained an invalid decision. Expected APPROVE or CANCEL but received '$decision'."
                }

                Remove-Item -LiteralPath $fullPath -Force -ErrorAction SilentlyContinue
                Write-Host ("WDS_APPROVED|{0}" -f $safeAction)
                return
            }

            Start-Sleep -Milliseconds 100
        }
    } finally {
        $stopwatch.Stop()
    }

    if ($EmitProgress) {
        Write-Host 'WDS_CANCELLED|pre-change review timeout'
    }
    throw [System.OperationCanceledException]::new("WinDefState pre-change review timed out after $TimeoutSeconds seconds. The baseline was saved, but no defense setting was changed.")
}

function Ensure-Directory {
    param([Parameter(Mandatory)] [string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Clear-SnapshotAssetCache {
    $script:WinDefStateSnapshotAssetCache = @{}
}

function Get-Sha256HashFromBytes {
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [byte[]]$Content)

    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        ([BitConverter]::ToString($algorithm.ComputeHash($Content))).Replace('-', '')
    } finally {
        $algorithm.Dispose()
    }
}

function Get-SnapshotAssetCacheRecord {
    param([Parameter(Mandatory)] [string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    $cacheKey = "asset|$fullPath"
    if ($script:WinDefStateSnapshotAssetCache.ContainsKey($cacheKey)) {
        return $script:WinDefStateSnapshotAssetCache[$cacheKey]
    }
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw "Snapshot asset is missing: $fullPath"
    }

    $bytes = [IO.File]::ReadAllBytes($fullPath)
    $record = [PSCustomObject]@{
        Path   = $fullPath
        Bytes  = [byte[]]$bytes
        Sha256 = Get-Sha256HashFromBytes -Content $bytes
        Text   = $null
    }
    $script:WinDefStateSnapshotAssetCache[$cacheKey] = $record
    $record
}

function Assert-SnapshotAssetHash {
    param(
        [Parameter(Mandatory)] [object]$Record,
        [AllowNull()] [string]$ExpectedSha256,
        [string]$Description = 'Snapshot asset'
    )

    if ([string]::IsNullOrWhiteSpace($ExpectedSha256)) {
        return
    }
    if ($ExpectedSha256 -notmatch '^[a-fA-F0-9]{64}$') {
        throw "$Description contains an invalid recorded SHA-256: $ExpectedSha256"
    }
    if (-not [string]::Equals($ExpectedSha256, [string]$Record.Sha256, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "$Description no longer matches its recorded SHA-256: $($Record.Path)"
    }
}

function Read-SnapshotAssetText {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [AllowNull()] [string]$ExpectedSha256,
        [string]$Description = 'Snapshot asset'
    )

    $record = Get-SnapshotAssetCacheRecord -Path $Path
    Assert-SnapshotAssetHash -Record $record -ExpectedSha256 $ExpectedSha256 -Description $Description
    if ($null -eq $record.Text) {
        $stream = [System.IO.MemoryStream]::new([byte[]]$record.Bytes, $false)
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
        try {
            $record.Text = $reader.ReadToEnd()
        } finally {
            $reader.Dispose()
        }
    }

    [string]$record.Text
}

function Read-SnapshotAssetBytes {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [AllowNull()] [string]$ExpectedSha256,
        [string]$Description = 'Snapshot asset'
    )

    $record = Get-SnapshotAssetCacheRecord -Path $Path
    Assert-SnapshotAssetHash -Record $record -ExpectedSha256 $ExpectedSha256 -Description $Description
    return ,([byte[]]$record.Bytes)
}

function Get-WinDefStateRuntimeInfo {
    if ($null -ne $script:WinDefStateRuntimeInfo) {
        return $script:WinDefStateRuntimeInfo
    }

    $scriptHash = $null
    if (Test-Path -LiteralPath $script:WinDefStateScriptPath -PathType Leaf) {
        $scriptHash = (Get-FileHash -LiteralPath $script:WinDefStateScriptPath -Algorithm SHA256).Hash
    }

    $powerShellEdition = if ($PSVersionTable.ContainsKey('PSEdition')) { [string]$PSVersionTable.PSEdition } else { 'Desktop' }
    $script:WinDefStateRuntimeInfo = [PSCustomObject]@{
        ScriptFileName      = [IO.Path]::GetFileName($script:WinDefStateScriptPath)
        ScriptSha256        = $scriptHash
        PowerShellVersion   = [string]$PSVersionTable.PSVersion
        PowerShellEdition   = $powerShellEdition
        ProcessArchitecture = [string]$env:PROCESSOR_ARCHITECTURE
    }

    $script:WinDefStateRuntimeInfo
}

# Capture provenance before any operation can outlive or replace the script file on disk.
$null = Get-WinDefStateRuntimeInfo

function Protect-StateRoot {
    param([Parameter(Mandatory)] [string]$Path)

    Ensure-Directory -Path $Path
    if ($env:OS -ne 'Windows_NT') {
        return
    }

    $fullPath = Resolve-FileSystemPath -Path $Path
    $administrators = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')
    $system = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
    $inheritance = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
    $propagation = [System.Security.AccessControl.PropagationFlags]::None
    $allow = [System.Security.AccessControl.AccessControlType]::Allow

    $acl = New-Object System.Security.AccessControl.DirectorySecurity
    $acl.SetAccessRuleProtection($true, $false)
    $acl.SetOwner($administrators)
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($administrators, 'FullControl', $inheritance, $propagation, $allow)))
    $acl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($system, 'FullControl', $inheritance, $propagation, $allow)))

    try {
        [System.IO.Directory]::SetAccessControl($fullPath, $acl)
    } catch {
        throw "WinDefState could not secure state root '$fullPath' for Administrators and SYSTEM only. No operation was started. $($_.Exception.Message)"
    }
}

function Resolve-FileSystemPath {
    param([Parameter(Mandatory)] [string]$Path)

    $expandedPath = [Environment]::ExpandEnvironmentVariables($Path)

    try {
        [IO.Path]::GetFullPath($expandedPath)
    } catch {
        $expandedPath
    }
}

function Resolve-ContainedFileSystemPath {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [string]$RelativePath,
        [string]$Description = 'relative path'
    )

    if ([string]::IsNullOrWhiteSpace($RelativePath) -or [IO.Path]::IsPathRooted($RelativePath)) {
        throw "Invalid ${Description}: $RelativePath"
    }

    $fullRoot = [IO.Path]::GetFullPath((Resolve-FileSystemPath -Path $Root)).TrimEnd([char[]]@('\', '/'))
    $rootPrefix = $fullRoot + [IO.Path]::DirectorySeparatorChar
    $candidate = [IO.Path]::GetFullPath((Join-Path $fullRoot $RelativePath))
    if (-not $candidate.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "${Description} escapes its trusted root: $RelativePath"
    }

    $candidate
}

function Get-AtomicTempPath {
    param([Parameter(Mandatory)] [string]$DestinationPath)

    $parent = Split-Path -Parent $DestinationPath
    $leaf = Split-Path -Leaf $DestinationPath
    Join-Path $parent ('.{0}.{1}.tmp' -f $leaf, [guid]::NewGuid().ToString('N'))
}

function Publish-FileAtomic {
    param(
        [Parameter(Mandatory)] [string]$TempPath,
        [Parameter(Mandatory)] [string]$DestinationPath
    )

    $backupPath = $null
    try {
        if (Test-Path -LiteralPath $DestinationPath -PathType Leaf) {
            try {
                $backupPath = Get-AtomicTempPath -DestinationPath ($DestinationPath + '.backup')
                [System.IO.File]::Replace($TempPath, $DestinationPath, $backupPath, $true)
                return
            } catch {
                $replaceError = if ($null -ne $_.Exception.InnerException) { $_.Exception.InnerException } else { $_.Exception }
                if ($replaceError -isnot [System.PlatformNotSupportedException] -and $replaceError -isnot [System.IO.IOException]) {
                    throw
                }
                # Some removable and non-NTFS filesystems do not support File.Replace.
            }
        }

        Move-Item -LiteralPath $TempPath -Destination $DestinationPath -Force
    } finally {
        if (Test-Path -LiteralPath $TempPath) {
            Remove-Item -LiteralPath $TempPath -Force -ErrorAction SilentlyContinue
        }
        if (-not [string]::IsNullOrWhiteSpace($backupPath) -and (Test-Path -LiteralPath $backupPath)) {
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Write-JsonAtomic {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [object]$InputObject
    )

    $resolvedPath = Resolve-FileSystemPath -Path $Path
    $parent = Split-Path -Parent $resolvedPath
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Ensure-Directory -Path $parent
    }

    $tempPath = Get-AtomicTempPath -DestinationPath $resolvedPath
    try {
        $json = $InputObject | ConvertTo-Json -Depth 12
        [System.IO.File]::WriteAllText($tempPath, $json, [System.Text.UTF8Encoding]::new($false))
        Publish-FileAtomic -TempPath $tempPath -DestinationPath $resolvedPath
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Write-SnapshotJsonAtomic {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [object]$Snapshot
    )

    $resolvedPath = Resolve-FileSystemPath -Path $Path
    $parent = Split-Path -Parent $resolvedPath
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Ensure-Directory -Path $parent
    }

    $tempPath = Get-AtomicTempPath -DestinationPath $resolvedPath
    $stream = $null
    $writer = $null
    try {
        $encoding = [System.Text.UTF8Encoding]::new($false)
        $stream = [System.IO.File]::Open($tempPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $writer = New-Object System.IO.StreamWriter($stream, $encoding)

        try {
            $writer.WriteLine('{')
            $writer.WriteLine(('  "SchemaVersion": {0},' -f (ConvertTo-Json -InputObject $Snapshot.SchemaVersion -Compress)))
            $writer.WriteLine(('  "Tool": {0},' -f (ConvertTo-Json -InputObject $Snapshot.Tool -Compress)))
            if ($Snapshot.PSObject.Properties['Producer']) {
                $writer.WriteLine(('  "Producer": {0},' -f (ConvertTo-Json -InputObject $Snapshot.Producer -Depth 4 -Compress)))
            }
            $writer.WriteLine(('  "ComputerName": {0},' -f (ConvertTo-Json -InputObject $Snapshot.ComputerName -Compress)))
            $writer.WriteLine(('  "CapturedAtUtc": {0},' -f (ConvertTo-Json -InputObject $Snapshot.CapturedAtUtc -Compress)))
            if ($Snapshot.PSObject.Properties['CaptureMetrics']) {
                $writer.WriteLine(('  "CaptureMetrics": {0},' -f (ConvertTo-Json -InputObject $Snapshot.CaptureMetrics -Depth 8 -Compress)))
            }
            if ($Snapshot.PSObject.Properties['CaptureScope']) {
                $writer.WriteLine(('  "CaptureScope": {0},' -f (ConvertTo-Json -InputObject $Snapshot.CaptureScope -Depth 4 -Compress)))
            }
            $writer.WriteLine('  "Settings": [')

            $settings = @($Snapshot.Settings)
            for ($i = 0; $i -lt $settings.Count; $i++) {
                $entry = $settings[$i]
                $entryId = if ($null -ne $entry -and $entry.PSObject.Properties['Id']) { [string]$entry.Id } else { "<entry-$i>" }
                Write-Verbose ("[json {0}/{1}] Serializing {2}" -f ($i + 1), $settings.Count, $entryId)
                $entryJson = ConvertTo-Json -InputObject $entry -Depth 12 -Compress
                $suffix = if ($i -lt ($settings.Count - 1)) { ',' } else { '' }
                $writer.WriteLine(('    {0}{1}' -f $entryJson, $suffix))
            }

            $writer.WriteLine('  ]')
            $writer.WriteLine('}')
        } finally {
            if ($null -ne $writer) {
                $writer.Dispose()
                $writer = $null
                $stream = $null
            } elseif ($null -ne $stream) {
                $stream.Dispose()
                $stream = $null
            }
        }

        Publish-FileAtomic -TempPath $tempPath -DestinationPath $resolvedPath
    } finally {
        if ($null -ne $writer) {
            $writer.Dispose()
        } elseif ($null -ne $stream) {
            $stream.Dispose()
        }
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Write-TextAtomic {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Content
    )

    $resolvedPath = Resolve-FileSystemPath -Path $Path
    $parent = Split-Path -Parent $resolvedPath
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Ensure-Directory -Path $parent
    }

    $tempPath = Get-AtomicTempPath -DestinationPath $resolvedPath
    try {
        [System.IO.File]::WriteAllText($tempPath, $Content, [System.Text.UTF8Encoding]::new($false))
        Publish-FileAtomic -TempPath $tempPath -DestinationPath $resolvedPath
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Write-BytesAtomic {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [byte[]]$Content
    )

    $resolvedPath = Resolve-FileSystemPath -Path $Path
    $parent = Split-Path -Parent $resolvedPath
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Ensure-Directory -Path $parent
    }

    $tempPath = Get-AtomicTempPath -DestinationPath $resolvedPath
    try {
        [System.IO.File]::WriteAllBytes($tempPath, $Content)
        Publish-FileAtomic -TempPath $tempPath -DestinationPath $resolvedPath
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-TemporaryFilePath {
    param([string]$Extension = '.tmp')

    if (-not $Extension.StartsWith('.')) {
        $Extension = ".$Extension"
    }

    Join-Path ([System.IO.Path]::GetTempPath()) ("WinDefState-{0}{1}" -f ([guid]::NewGuid().ToString('N')), $Extension)
}

function Clear-WinDefStateCommandCache {
    $script:WinDefStateCommandCache.Clear()
}

function Get-WinDefStateCommand {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [switch]$Refresh
    )

    $cacheKey = $Name.Trim().ToLowerInvariant()
    if (-not $Refresh -and $script:WinDefStateCommandCache.ContainsKey($cacheKey)) {
        return $script:WinDefStateCommandCache[$cacheKey]
    }

    # Module auto-loading can emit hundreds of export records under top-level -Verbose.
    $command = & {
        $VerbosePreference = 'SilentlyContinue'
        Get-Command -Name $Name -ErrorAction SilentlyContinue -Verbose:$false | Select-Object -First 1
    }
    $script:WinDefStateCommandCache[$cacheKey] = $command
    $command
}

function Test-CommandAvailable {
    param([Parameter(Mandatory)] [string]$Name)

    $null -ne (Get-WinDefStateCommand -Name $Name)
}

function New-WinDefStatePreflightCheck {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [ValidateSet('Pass', 'Warning', 'Info')] [string]$Status,
        [Parameter(Mandatory)] [string]$Value,
        [string]$Detail
    )

    [PSCustomObject]@{
        Name   = $Name
        Status = $Status
        Value  = $Value
        Detail = $Detail
    }
}

function Get-WinDefStatePendingRebootState {
    if ($env:OS -ne 'Windows_NT') {
        return [PSCustomObject]@{
            Supported = $false
            Pending   = $false
            Reasons   = @()
        }
    }

    $reasons = [System.Collections.Generic.List[string]]::new()
    foreach ($path in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    )) {
        if (Test-Path -LiteralPath $path) {
            $reasons.Add($path) | Out-Null
        }
    }

    try {
        $sessionManager = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction Stop
        if ($sessionManager.PSObject.Properties['PendingFileRenameOperations'] -and $null -ne $sessionManager.PendingFileRenameOperations) {
            $reasons.Add('PendingFileRenameOperations') | Out-Null
        }
    } catch [System.Management.Automation.ItemNotFoundException] {
        Write-Verbose 'PendingFileRenameOperations is not present.'
    } catch [System.Management.Automation.PSArgumentException] {
        Write-Verbose 'PendingFileRenameOperations is not present.'
    }

    [PSCustomObject]@{
        Supported = $true
        Pending   = $reasons.Count -gt 0
        Reasons   = @($reasons)
    }
}

function Get-WinDefStatePreflightProviderRequirements {
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Items,
        [Parameter(Mandatory)] [ValidateSet('Snapshot', 'Permissive', 'Restore')] [string]$Action
    )

    $commandsByProviderType = @{
        DefenderRuntimeStatus    = @('Get-MpComputerStatus')
        MpPreferenceValue        = @('Get-MpPreference')
        MpPreferenceList         = @('Get-MpPreference')
        AsrRules                 = @('Get-MpPreference')
        ServiceConfig            = @('Get-CimInstance')
        LocalUser                = @('Get-CimInstance')
        LoadedUserRegistryValues = @('Get-CimInstance', 'reg.exe')
        NetBiosAdapters          = @('Get-CimInstance')
        WsManValue               = @('Get-WSManInstance')
        WinRmListeners           = @('Get-WSManInstance')
        SmbClientConfig          = @('Get-SmbClientConfiguration', 'powershell.exe')
        SmbServerConfig          = @('Get-SmbServerConfiguration', 'powershell.exe')
        BitLockerVolumes         = @('Get-BitLockerVolume', 'powershell.exe')
        AppLockerPolicy          = @('Get-AppLockerPolicy')
        ExploitProtectionPolicy  = @('Get-ProcessMitigation')
        FirewallProfiles         = @('Get-NetFirewallProfile')
        FirewallRules            = @('Get-NetFirewallRule')
    }
    $mutationCommandsByProviderType = @{
        MpPreferenceValue       = @('Set-MpPreference')
        MpPreferenceList        = @('Add-MpPreference', 'Remove-MpPreference')
        AsrRules                = @('Add-MpPreference', 'Remove-MpPreference')
        ServiceConfig           = @('sc.exe')
        WsManValue              = @('Set-WSManInstance')
        WinRmListeners          = @('New-WSManInstance', 'Remove-WSManInstance')
        SmbClientConfig         = @('Set-SmbClientConfiguration')
        SmbServerConfig         = @('Set-SmbServerConfiguration')
        BitLockerVolumes        = @('Suspend-BitLocker', 'Resume-BitLocker', 'manage-bde.exe')
        AppLockerPolicy         = @('Set-AppLockerPolicy')
        ExploitProtectionPolicy = @('Set-ProcessMitigation')
        WdacPolicies             = @('CiTool.exe')
        FirewallProfiles        = @('Set-NetFirewallProfile')
        FirewallRules           = @('Set-NetFirewallRule')
    }

    $providersByCommand = @{}
    foreach ($item in @($Items)) {
        if ($null -eq $item -or -not $item.PSObject.Properties['Type']) {
            continue
        }
        $providerType = [string]$item.Type
        $commands = [System.Collections.Generic.List[string]]::new()
        if ($commandsByProviderType.ContainsKey($providerType)) {
            foreach ($commandName in @($commandsByProviderType[$providerType])) {
                $commands.Add([string]$commandName) | Out-Null
            }
        }
        $needsMutationCommands = if ($Action -eq 'Permissive') {
            Test-DefinitionHasPermissiveAction -Definition $item
        } elseif ($Action -eq 'Restore') {
            Test-DefinitionHasRestoreAction -Definition $item
        } else {
            $false
        }
        if ($needsMutationCommands -and $mutationCommandsByProviderType.ContainsKey($providerType)) {
            foreach ($commandName in @($mutationCommandsByProviderType[$providerType])) {
                $commands.Add([string]$commandName) | Out-Null
            }
        }

        foreach ($commandName in $commands) {
            $commandKey = $commandName.ToLowerInvariant()
            if (-not $providersByCommand.ContainsKey($commandKey)) {
                $providersByCommand[$commandKey] = [PSCustomObject]@{
                    Command   = $commandName
                    Providers = [System.Collections.Generic.List[string]]::new()
                }
            }
            if (-not $providersByCommand[$commandKey].Providers.Contains($providerType)) {
                $providersByCommand[$commandKey].Providers.Add($providerType) | Out-Null
            }
        }
    }

    @(
        foreach ($record in @($providersByCommand.Values | Sort-Object -Property Command)) {
            [PSCustomObject]@{
                Command   = [string]$record.Command
                Providers = @($record.Providers | Sort-Object)
            }
        }
    )
}

function Get-WinDefStatePreflight {
    param(
        [Parameter(Mandatory)] [ValidateSet('Snapshot', 'Permissive', 'Restore')] [string]$Action,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Items,
        [Parameter(Mandatory)] [string]$Root,
        [string]$SnapshotPath
    )

    $checks = [System.Collections.Generic.List[object]]::new()
    $languageMode = [string]$ExecutionContext.SessionState.LanguageMode
    $languageStatus = if ([string]::Equals($languageMode, 'FullLanguage', [System.StringComparison]::OrdinalIgnoreCase)) { 'Pass' } else { 'Warning' }
    $checks.Add((New-WinDefStatePreflightCheck -Name 'PowerShell language mode' -Status $languageStatus -Value $languageMode -Detail 'Restricted language modes can prevent provider and .NET operations.')) | Out-Null

    $runtime = Get-WinDefStateRuntimeInfo
    $runtimeValue = "PowerShell $($runtime.PowerShellVersion) $($runtime.PowerShellEdition), $($runtime.ProcessArchitecture)"
    $checks.Add((New-WinDefStatePreflightCheck -Name 'Runtime' -Status Pass -Value $runtimeValue -Detail 'WinDefState targets Windows PowerShell 5.1 compatibility.')) | Out-Null

    try {
        $effectivePolicy = [string](Get-ExecutionPolicy -ErrorAction Stop)
        $checks.Add((New-WinDefStatePreflightCheck -Name 'Effective execution policy' -Status Info -Value $effectivePolicy -Detail 'The current process has already loaded WinDefState; this value is diagnostic only.')) | Out-Null
    } catch {
        $checks.Add((New-WinDefStatePreflightCheck -Name 'Effective execution policy' -Status Warning -Value '<unavailable>' -Detail $_.Exception.Message)) | Out-Null
    }

    $pendingReboot = Get-WinDefStatePendingRebootState
    if (-not $pendingReboot.Supported) {
        $checks.Add((New-WinDefStatePreflightCheck -Name 'Pending reboot' -Status Info -Value 'Not evaluated' -Detail 'Pending-reboot detection is available only on Windows.')) | Out-Null
    } elseif ($pendingReboot.Pending) {
        $checks.Add((New-WinDefStatePreflightCheck -Name 'Pending reboot' -Status Warning -Value 'Yes' -Detail (@($pendingReboot.Reasons) -join '; '))) | Out-Null
    } else {
        $checks.Add((New-WinDefStatePreflightCheck -Name 'Pending reboot' -Status Pass -Value 'No' -Detail 'No common reboot-pending markers were found.')) | Out-Null
    }

    $rootPath = [IO.Path]::GetFullPath($Root)
    $rootReady = Test-Path -LiteralPath $rootPath -PathType Container
    $checks.Add((New-WinDefStatePreflightCheck -Name 'Protected state root' -Status $(if ($rootReady) { 'Pass' } else { 'Warning' }) -Value $rootPath -Detail $(if ($rootReady) { 'The operation state directory is available.' } else { 'The operation state directory is not available.' }))) | Out-Null

    try {
        $operation = Get-OperationState -Root $Root
        if ($null -eq $operation) {
            $journalStatus = if ($Action -eq 'Restore' -and [string]::IsNullOrWhiteSpace($SnapshotPath)) { 'Warning' } else { 'Pass' }
            $journalDetail = if ($journalStatus -eq 'Warning') { 'Restore needs an active journal or an explicit snapshot path.' } else { 'No active operation journal exists.' }
            $checks.Add((New-WinDefStatePreflightCheck -Name 'Operation journal' -Status $journalStatus -Value 'None' -Detail $journalDetail)) | Out-Null
        } else {
            $status = if ($operation.PSObject.Properties['Status']) { [string]$operation.Status } else { 'Legacy' }
            $journalStatus = if ($Action -eq 'Permissive') { 'Warning' } else { 'Info' }
            $checks.Add((New-WinDefStatePreflightCheck -Name 'Operation journal' -Status $journalStatus -Value $status -Detail ([string]$operation.SnapshotPath))) | Out-Null
        }
    } catch {
        $checks.Add((New-WinDefStatePreflightCheck -Name 'Operation journal' -Status Warning -Value '<unreadable>' -Detail $_.Exception.Message)) | Out-Null
    }

    $checks.Add((New-WinDefStatePreflightCheck -Name 'Selected setting scope' -Status Info -Value ("{0} setting(s)" -f @($Items).Count) -Detail $Action)) | Out-Null
    $requirements = @(Get-WinDefStatePreflightProviderRequirements -Items $Items -Action $Action)
    foreach ($requirement in $requirements) {
        $available = Test-CommandAvailable -Name ([string]$requirement.Command)
        $checks.Add((New-WinDefStatePreflightCheck `
            -Name ("Provider command: {0}" -f $requirement.Command) `
            -Status $(if ($available) { 'Pass' } else { 'Warning' }) `
            -Value $(if ($available) { 'Available' } else { 'Missing' }) `
            -Detail (@($requirement.Providers) -join ', '))) | Out-Null
    }

    $warningChecks = @($checks | Where-Object { [string]$_.Status -eq 'Warning' })
    [PSCustomObject]@{
        Tool          = 'WinDefState'
        Action        = $Action
        ComputerName  = $env:COMPUTERNAME
        CheckedAtUtc  = (Get-Date).ToUniversalTime().ToString('o')
        OverallStatus = if ($warningChecks.Count -gt 0) { 'Warning' } else { 'Ready' }
        WarningCount  = $warningChecks.Count
        Checks        = @($checks)
    }
}

function Get-WinDefStatePreflightReportLines {
    param([Parameter(Mandatory)] [object]$Preflight)

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('WinDefState preflight report') | Out-Null
    $lines.Add(('Action: {0}' -f $Preflight.Action)) | Out-Null
    $lines.Add(('Computer: {0}' -f $Preflight.ComputerName)) | Out-Null
    $lines.Add(('Checked at UTC: {0}' -f $Preflight.CheckedAtUtc)) | Out-Null
    $lines.Add(('Overall status: {0}' -f $Preflight.OverallStatus)) | Out-Null
    $lines.Add(('Warnings: {0}' -f $Preflight.WarningCount)) | Out-Null
    $lines.Add('') | Out-Null
    foreach ($check in @($Preflight.Checks)) {
        $lines.Add(('[{0}] {1}: {2}' -f ([string]$check.Status).ToUpperInvariant(), $check.Name, $check.Value)) | Out-Null
        if (-not [string]::IsNullOrWhiteSpace([string]$check.Detail)) {
            $lines.Add(('  {0}' -f $check.Detail)) | Out-Null
        }
    }
    @($lines)
}

function Invoke-WinDefStatePreflight {
    param(
        [Parameter(Mandatory)] [ValidateSet('Snapshot', 'Permissive', 'Restore')] [string]$Action,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Items,
        [Parameter(Mandatory)] [string]$Root,
        [string]$SnapshotPath
    )

    $preflight = Get-WinDefStatePreflight -Action $Action -Items $Items -Root $Root -SnapshotPath $SnapshotPath
    $reportPath = Get-PreflightReportPath -Root $Root -Action $Action
    $reportLines = Get-WinDefStatePreflightReportLines -Preflight $preflight
    Write-TextAtomic -Path $reportPath -Content ($reportLines -join [Environment]::NewLine)
    $preflight | Add-Member -NotePropertyName ReportPath -NotePropertyValue $reportPath -Force
    Write-OperationResult -Name 'PreflightReport' -Value $reportPath

    if ($preflight.WarningCount -gt 0) {
        $warningNames = @($preflight.Checks | Where-Object { [string]$_.Status -eq 'Warning' } | ForEach-Object { [string]$_.Name })
        Write-Warning ("Preflight found {0} warning(s): {1}. Report: {2}" -f $preflight.WarningCount, ($warningNames -join ', '), $reportPath)
    } else {
        Write-Host "Preflight ready. Report saved to: $reportPath"
    }
    $preflight
}

function New-CaptureSession {
    param([Parameter(Mandatory)] [ValidateSet('Snapshot', 'Permissive', 'Restore', 'Verification', 'PermissiveVerification')] [string]$Phase)

    [PSCustomObject]@{
        Phase           = $Phase
        StartedAtUtc    = (Get-Date).ToUniversalTime().ToString('o')
        Stopwatch       = [System.Diagnostics.Stopwatch]::StartNew()
        Cache           = @{}
        CacheHitCount   = 0
        ProviderTimings = [System.Collections.Generic.List[object]]::new()
        SettingTimings  = [System.Collections.Generic.List[object]]::new()
        Resources       = [System.Collections.Generic.List[object]]::new()
        UserRegistryDefinitionsRemaining = 0
        WinRmMutationDefinitionsRemaining = 0
        ServiceNames    = @()
    }
}

function Get-CaptureSessionValue {
    param(
        [AllowNull()] [object]$Session,
        [Parameter(Mandatory)] [string]$Key,
        [Parameter(Mandatory)] [scriptblock]$Factory
    )

    if ($null -eq $Session) {
        return (& $Factory)
    }

    if ($Session.Cache.ContainsKey($Key)) {
        $Session.CacheHitCount = [int]$Session.CacheHitCount + 1
        return $Session.Cache[$Key]
    }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $errorMessage = $null
    try {
        $value = & $Factory
        $Session.Cache[$Key] = $value
        return $value
    } catch {
        $errorMessage = $_.Exception.Message
        throw
    } finally {
        $stopwatch.Stop()
        $Session.ProviderTimings.Add([PSCustomObject]@{
            Key        = $Key
            DurationMs = [Math]::Round($stopwatch.Elapsed.TotalMilliseconds, 1)
            Succeeded  = [string]::IsNullOrWhiteSpace($errorMessage)
            Error      = $errorMessage
        }) | Out-Null
    }
}

function Register-CaptureSessionResource {
    param(
        [AllowNull()] [object]$Session,
        [Parameter(Mandatory)] [string]$Kind,
        [Parameter(Mandatory)] [object]$Value
    )

    if ($null -eq $Session) {
        return
    }

    $Session.Resources.Add([PSCustomObject]@{
        Kind  = $Kind
        Value = $Value
    }) | Out-Null
}

function Add-CaptureSessionSettingTiming {
    param(
        [AllowNull()] [object]$Session,
        [Parameter(Mandatory)] [string]$Id,
        [Parameter(Mandatory)] [string]$Type,
        [Parameter(Mandatory)] [double]$DurationMs,
        [bool]$Succeeded = $true,
        [string]$ErrorMessage
    )

    if ($null -eq $Session) {
        return
    }

    $Session.SettingTimings.Add([PSCustomObject]@{
        Id         = $Id
        Type       = $Type
        DurationMs = [Math]::Round($DurationMs, 1)
        Succeeded  = $Succeeded
        Error      = $ErrorMessage
    }) | Out-Null
}

function Release-CaptureSessionResources {
    param(
        [Parameter(Mandatory)] [object]$Session,
        [string]$Kind,
        [switch]$ThrowOnError
    )

    for ($i = $Session.Resources.Count - 1; $i -ge 0; $i--) {
        $resource = $Session.Resources[$i]
        if (-not [string]::IsNullOrWhiteSpace($Kind) -and [string]$resource.Kind -ne $Kind) {
            continue
        }

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $errorMessage = $null
        $releaseError = $null
        try {
            switch ([string]$resource.Kind) {
                'UserRegistryTarget' { Close-UserRegistryTarget -Target $resource.Value }
                'WinRmWriteScope' { Exit-WinRmServiceWriteScope -Scope $resource.Value }
            }
        } catch {
            $releaseError = $_
            $errorMessage = $_.Exception.Message
            Write-Warning ("Failed to release capture resource {0}: {1}" -f ([string]$resource.Kind), $errorMessage)
        } finally {
            $stopwatch.Stop()
            $Session.ProviderTimings.Add([PSCustomObject]@{
                Key        = "cleanup:$([string]$resource.Kind)"
                DurationMs = [Math]::Round($stopwatch.Elapsed.TotalMilliseconds, 1)
                Succeeded  = [string]::IsNullOrWhiteSpace($errorMessage)
                Error      = $errorMessage
            }) | Out-Null
        }
        $Session.Resources.RemoveAt($i)
        if ($ThrowOnError -and $null -ne $releaseError) {
            throw $releaseError
        }
    }
}

function Complete-UserRegistrySessionDefinition {
    param([AllowNull()] [object]$Session)

    if ($null -eq $Session -or [int]$Session.UserRegistryDefinitionsRemaining -le 0) {
        return
    }

    $Session.UserRegistryDefinitionsRemaining = [int]$Session.UserRegistryDefinitionsRemaining - 1
    if ([int]$Session.UserRegistryDefinitionsRemaining -eq 0) {
        Release-CaptureSessionResources -Session $Session -Kind 'UserRegistryTarget'
    }
}

function Complete-WinRmMutationSessionDefinition {
    param([AllowNull()] [object]$Session)

    if ($null -eq $Session -or [int]$Session.WinRmMutationDefinitionsRemaining -le 0) {
        return
    }

    $Session.WinRmMutationDefinitionsRemaining = [int]$Session.WinRmMutationDefinitionsRemaining - 1
    if ([int]$Session.WinRmMutationDefinitionsRemaining -eq 0) {
        Release-CaptureSessionResources -Session $Session -Kind 'WinRmWriteScope' -ThrowOnError
        [void]$Session.Cache.Remove('winrm.write-scope')
    }
}

function Complete-MutationSessionDefinition {
    param(
        [AllowNull()] [object]$Session,
        [Parameter(Mandatory)] [string]$Type
    )

    if ($Type -eq 'LoadedUserRegistryValues') {
        Complete-UserRegistrySessionDefinition -Session $Session
    }
    if ($Type -in @('WsManValue', 'WinRmListeners')) {
        Complete-WinRmMutationSessionDefinition -Session $Session
    }
}

function Complete-CaptureSession {
    param([Parameter(Mandatory)] [object]$Session)

    Release-CaptureSessionResources -Session $Session

    if ($Session.Stopwatch.IsRunning) {
        $Session.Stopwatch.Stop()
    }

    [PSCustomObject]@{
        Phase              = [string]$Session.Phase
        StartedAtUtc       = [string]$Session.StartedAtUtc
        CompletedAtUtc     = (Get-Date).ToUniversalTime().ToString('o')
        DurationMs         = [Math]::Round($Session.Stopwatch.Elapsed.TotalMilliseconds, 1)
        ProviderQueryCount = $Session.ProviderTimings.Count
        CacheHitCount      = [int]$Session.CacheHitCount
        ProviderQueries    = @($Session.ProviderTimings)
        Settings           = @($Session.SettingTimings)
    }
}

function Get-NormalizedIdFilter {
    param([AllowNull()] [string[]]$Ids)

    @(
        foreach ($rawId in @($Ids)) {
            foreach ($id in @(([string]$rawId) -split ',')) {
                if ([string]::IsNullOrWhiteSpace($id)) {
                    continue
                }

                ([string]$id).Trim()
            }
        }
    ) | Sort-Object -Unique
}

function Test-SettingIdMatchesFilter {
    param(
        [Parameter(Mandatory)] [string]$Id,
        [Parameter(Mandatory)] [string[]]$Pattern
    )

    foreach ($candidate in @($Pattern)) {
        if ($Id -like $candidate) {
            return $true
        }
    }

    $false
}

function Test-SettingIdFilterMatchesAny {
    param(
        [Parameter(Mandatory)] [string]$Pattern,
        [Parameter(Mandatory)] [string[]]$AvailableId
    )

    foreach ($id in @($AvailableId)) {
        if (Test-SettingIdMatchesFilter -Id $id -Pattern @($Pattern)) {
            return $true
        }
    }

    $false
}

function Get-SettingCategory {
    param([Parameter(Mandatory)] [string]$Id)

    $separatorIndex = $Id.IndexOf('.')
    if ($separatorIndex -lt 1) {
        return $Id.ToLowerInvariant()
    }

    $Id.Substring(0, $separatorIndex).ToLowerInvariant()
}

function Get-NormalizedCategoryFilter {
    param([AllowNull()] [string[]]$Category)

    @(
        foreach ($rawCategory in @($Category)) {
            foreach ($item in @(([string]$rawCategory) -split ',')) {
                if ([string]::IsNullOrWhiteSpace($item)) {
                    continue
                }

                ([string]$item).Trim().ToLowerInvariant()
            }
        }
    ) | Sort-Object -Unique
}

function ConvertTo-CategoryIdFilter {
    param([AllowNull()] [string[]]$Category)

    $categories = @(Get-NormalizedCategoryFilter -Category $Category)
    if ($categories.Count -eq 0) {
        return @()
    }

    $availableCategories = @(
        Get-DefenseDefinitions |
            ForEach-Object { Get-SettingCategory -Id ([string]$_.Id) } |
            Sort-Object -Unique
    )
    $unknownCategories = @($categories | Where-Object { $_ -notin $availableCategories })
    if ($unknownCategories.Count -gt 0) {
        throw "Unknown setting category/categories: $($unknownCategories -join ', '). Available categories: $($availableCategories -join ', ')"
    }

    @($categories | ForEach-Object { "$_.*" })
}

function Merge-SettingIdFilter {
    param(
        [AllowNull()] [string[]]$Id,
        [AllowNull()] [string[]]$Category
    )

    @(
        @(
            @(Get-NormalizedIdFilter -Ids $Id)
            @(ConvertTo-CategoryIdFilter -Category $Category)
        ) | Sort-Object -Unique
    )
}

function Test-SettingIdIncluded {
    param(
        [Parameter(Mandatory)] [string]$Id,
        [AllowNull()] [string[]]$IncludeId,
        [AllowNull()] [string[]]$ExcludeId
    )

    $includedIds = @(Get-NormalizedIdFilter -Ids $IncludeId)
    $excludedIds = @(Get-NormalizedIdFilter -Ids $ExcludeId)

    if ($includedIds.Count -gt 0 -and -not (Test-SettingIdMatchesFilter -Id $Id -Pattern $includedIds)) {
        return $false
    }

    if ($excludedIds.Count -gt 0 -and (Test-SettingIdMatchesFilter -Id $Id -Pattern $excludedIds)) {
        return $false
    }

    $true
}

function Test-IdFilterActive {
    param(
        [AllowNull()] [string[]]$IncludeId,
        [AllowNull()] [string[]]$ExcludeId
    )

    (@(Get-NormalizedIdFilter -Ids $IncludeId).Count -gt 0 -or @(Get-NormalizedIdFilter -Ids $ExcludeId).Count -gt 0)
}

function Assert-ValidSettingIdFilter {
    param(
        [Parameter(Mandatory)] [string[]]$AvailableId,
        [AllowNull()] [string[]]$IncludeId,
        [AllowNull()] [string[]]$ExcludeId
    )

    $availableIds = @($AvailableId | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
    $includedIds = @(Get-NormalizedIdFilter -Ids $IncludeId)
    $excludedIds = @(Get-NormalizedIdFilter -Ids $ExcludeId)
    $unmatchedPatterns = @(
        @($includedIds + $excludedIds) |
            Where-Object { -not (Test-SettingIdFilterMatchesAny -Pattern $_ -AvailableId $availableIds) } |
            Sort-Object -Unique
    )
    if ($unmatchedPatterns.Count -gt 0) {
        throw "Unknown setting ID or unmatched pattern(s): $($unmatchedPatterns -join ', ')"
    }

    $overlap = @($includedIds | Where-Object { $_ -in $excludedIds })
    if ($overlap.Count -gt 0) {
        throw "Setting ID(s) cannot be both included and excluded: $($overlap -join ', ')"
    }

    if (Test-IdFilterActive -IncludeId $IncludeId -ExcludeId $ExcludeId) {
        $selectedIds = @($availableIds | Where-Object { Test-SettingIdIncluded -Id $_ -IncludeId $IncludeId -ExcludeId $ExcludeId })
        if ($selectedIds.Count -eq 0) {
            throw 'The setting filter selects no settings.'
        }
    }
}

function Get-SelectedDefenseDefinitions {
    param(
        [AllowNull()] [string[]]$IncludeId,
        [AllowNull()] [string[]]$ExcludeId
    )

    $definitions = @(Get-DefenseDefinitions)
    Assert-ValidSettingIdFilter -AvailableId @($definitions.Id) -IncludeId $IncludeId -ExcludeId $ExcludeId
    @(
        $definitions | Where-Object {
            Test-SettingIdIncluded -Id ([string]$_.Id) -IncludeId $IncludeId -ExcludeId $ExcludeId
        }
    )
}

function Assert-ValidDefenseSnapshot {
    param(
        [Parameter(Mandatory)] [object]$Snapshot,
        [switch]$AllowDifferentComputer
    )

    foreach ($propertyName in @('SchemaVersion', 'Tool', 'ComputerName', 'Settings')) {
        if (-not $Snapshot.PSObject.Properties[$propertyName]) {
            throw "Snapshot is missing required property '$propertyName'."
        }
    }

    $schemaVersion = try { [int]$Snapshot.SchemaVersion } catch { -1 }
    if ($schemaVersion -notin @(1, 2)) {
        throw "Unsupported snapshot schema version '$($Snapshot.SchemaVersion)'."
    }
    if ([string]$Snapshot.Tool -ne 'WinDefState') {
        throw "The file is not a WinDefState snapshot (Tool='$($Snapshot.Tool)')."
    }

    $currentComputerName = if (-not [string]::IsNullOrWhiteSpace([string]$env:COMPUTERNAME)) {
        [string]$env:COMPUTERNAME
    } else {
        [Environment]::MachineName
    }
    if (
        -not $AllowDifferentComputer -and
        -not [string]::IsNullOrWhiteSpace([string]$Snapshot.ComputerName) -and
        -not [string]::Equals([string]$Snapshot.ComputerName, $currentComputerName, [System.StringComparison]::OrdinalIgnoreCase)
    ) {
        throw "Snapshot computer '$($Snapshot.ComputerName)' does not match this computer '$currentComputerName'. Use -AllowDifferentComputer only for an intentional cross-host restore."
    }

    $entries = @($Snapshot.Settings)
    if ($entries.Count -eq 0) {
        throw 'Snapshot contains no setting entries.'
    }

    $definitions = Get-DefenseDefinitionMap
    $seenIds = @{}
    $immutableFieldsByType = @{
        RegistryValue          = @('Path', 'Name')
        RegistryKeyFlat        = @('Path')
        MpPreferenceValue      = @('Property')
        MpPreferenceList       = @('Property')
        MachineEnvironmentValue = @('Name')
        ServiceConfig          = @('Name')
        AuditPolicy            = @('Subcategory')
        WsManValue             = @('Path')
    }

    foreach ($entry in $entries) {
        if ($null -eq $entry -or -not $entry.PSObject.Properties['Id'] -or [string]::IsNullOrWhiteSpace([string]$entry.Id)) {
            throw 'Snapshot contains a setting entry without an ID.'
        }
        $entryId = [string]$entry.Id
        if ($seenIds.ContainsKey($entryId)) {
            throw "Snapshot contains duplicate setting ID '$entryId'."
        }
        $seenIds[$entryId] = $true

        if (-not $definitions.ContainsKey($entryId)) {
            throw "Snapshot contains unknown setting ID '$entryId'."
        }
        if (-not $entry.PSObject.Properties['Type'] -or [string]::IsNullOrWhiteSpace([string]$entry.Type)) {
            throw "Snapshot setting '$entryId' is missing its type."
        }

        $definition = $definitions[$entryId]
        if ([string]$entry.Type -ne [string]$definition.Type) {
            throw "Snapshot setting '$entryId' has type '$($entry.Type)', expected '$($definition.Type)'."
        }

        $immutableFields = @(if ($immutableFieldsByType.ContainsKey([string]$entry.Type)) { $immutableFieldsByType[[string]$entry.Type] })
        foreach ($field in $immutableFields) {
            if (-not $entry.PSObject.Properties[$field] -or -not $definition.PSObject.Properties[$field]) {
                throw "Snapshot setting '$entryId' is missing immutable target '$field'."
            }
            if (-not [string]::Equals([string]$entry.$field, [string]$definition.$field, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Snapshot setting '$entryId' targets unexpected $field '$($entry.$field)'."
            }
        }
    }
}

function Get-CiToolCommand {
    foreach ($name in @('CiTool.exe', 'CiTool')) {
        $command = Get-WinDefStateCommand -Name $name
        if ($null -ne $command) {
            return $command
        }
    }

    $null
}

function Invoke-ChildPowerShell {
    param(
        [Parameter(Mandatory)] [string]$ScriptText,
        [int]$TimeoutSeconds = 15
    )

    $powershellCommand = Get-WinDefStateCommand -Name 'powershell.exe'
    if ($null -eq $powershellCommand) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            TimedOut         = $false
            ExitCode         = $null
            StdOut           = ''
            StdErr           = 'powershell.exe was not found.'
        }
    }

    $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($ScriptText))
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $powershellCommand.Source
    $startInfo.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encodedCommand"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        [void]$process.Start()
        $stdoutTask = $process.StandardOutput.ReadToEndAsync()
        $stderrTask = $process.StandardError.ReadToEndAsync()

        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            try {
                $process.Kill()
            } catch {
                Write-Verbose ("Unable to terminate timed-out child PowerShell process: {0}" -f $_.Exception.Message)
            }

            $terminated = $process.WaitForExit(5000)
            $stdout = ''
            $stderr = ''
            if ($terminated) {
                $stdoutAwaiter = $stdoutTask.GetAwaiter()
                $stderrAwaiter = $stderrTask.GetAwaiter()
                $stdout = $stdoutAwaiter.GetResult()
                $stderr = $stderrAwaiter.GetResult()
            }

            return [PSCustomObject]@{
                CommandAvailable = $true
                TimedOut         = $true
                ExitCode         = $null
                StdOut           = $stdout
                StdErr           = $stderr
            }
        }

        $process.WaitForExit()
        $stdoutAwaiter = $stdoutTask.GetAwaiter()
        $stderrAwaiter = $stderrTask.GetAwaiter()
        return [PSCustomObject]@{
            CommandAvailable = $true
            TimedOut         = $false
            ExitCode         = $process.ExitCode
            StdOut           = $stdoutAwaiter.GetResult()
            StdErr           = $stderrAwaiter.GetResult()
        }
    } finally {
        $process.Dispose()
    }
}

function Invoke-ChildPowerShellBatch {
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Requests)

    $requestItems = @($Requests)
    if ($requestItems.Count -eq 0) {
        return
    }

    $powershellCommand = Get-WinDefStateCommand -Name 'powershell.exe'
    if ($null -eq $powershellCommand) {
        foreach ($request in $requestItems) {
            [PSCustomObject]@{
                Key              = [string]$request.Key
                CommandAvailable = $false
                TimedOut         = $false
                ExitCode         = $null
                StdOut           = ''
                StdErr           = 'powershell.exe was not found.'
                DurationMs       = 0
            }
        }
        return
    }

    $invocations = [System.Collections.Generic.List[object]]::new()
    try {
        foreach ($request in $requestItems) {
            $timeoutSeconds = if ($request.PSObject.Properties['TimeoutSeconds']) { [int]$request.TimeoutSeconds } else { 15 }
            if ($timeoutSeconds -lt 1) {
                throw "Child PowerShell request '$([string]$request.Key)' has an invalid timeout."
            }

            $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes([string]$request.ScriptText))
            $startInfo = New-Object System.Diagnostics.ProcessStartInfo
            $startInfo.FileName = $powershellCommand.Source
            $startInfo.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encodedCommand"
            $startInfo.UseShellExecute = $false
            $startInfo.CreateNoWindow = $true
            $startInfo.RedirectStandardOutput = $true
            $startInfo.RedirectStandardError = $true

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $startInfo
            $processStarted = $false
            try {
                [void]$process.Start()
                $processStarted = $true
                $invocations.Add([PSCustomObject]@{
                    Key            = [string]$request.Key
                    TimeoutSeconds = $timeoutSeconds
                    Process        = $process
                    StdOutTask     = $process.StandardOutput.ReadToEndAsync()
                    StdErrTask     = $process.StandardError.ReadToEndAsync()
                    Stopwatch      = [System.Diagnostics.Stopwatch]::StartNew()
                }) | Out-Null
                $process = $null
            } finally {
                if ($null -ne $process) {
                    if ($processStarted -and -not $process.HasExited) {
                        try {
                            $process.Kill()
                            [void]$process.WaitForExit(5000)
                        } catch {
                            Write-Verbose ("Unable to terminate child PowerShell process after startup failure: {0}" -f $_.Exception.Message)
                        }
                    }
                    $process.Dispose()
                }
            }
        }

        foreach ($invocation in $invocations) {
            $process = $invocation.Process
            $timeoutMs = [int]([int]$invocation.TimeoutSeconds * 1000)
            $remainingMs = [Math]::Max(0, $timeoutMs - [int]$invocation.Stopwatch.ElapsedMilliseconds)
            $terminated = $process.HasExited -or $process.WaitForExit($remainingMs)
            $timedOut = -not $terminated
            if ($timedOut) {
                try {
                    $process.Kill()
                } catch {
                    Write-Verbose ("Unable to terminate timed-out child PowerShell batch process: {0}" -f $_.Exception.Message)
                }
                $terminated = $process.WaitForExit(5000)
            }

            $stdout = ''
            $stderr = ''
            $exitCode = $null
            if ($terminated) {
                $process.WaitForExit()
                $stdout = $invocation.StdOutTask.GetAwaiter().GetResult()
                $stderr = $invocation.StdErrTask.GetAwaiter().GetResult()
                if (-not $timedOut) {
                    $exitCode = $process.ExitCode
                }
            }
            $invocation.Stopwatch.Stop()

            [PSCustomObject]@{
                Key              = [string]$invocation.Key
                CommandAvailable = $true
                TimedOut         = $timedOut
                ExitCode         = $exitCode
                StdOut           = $stdout
                StdErr           = $stderr
                DurationMs       = [Math]::Round($invocation.Stopwatch.Elapsed.TotalMilliseconds, 1)
            }
        }
    } finally {
        foreach ($invocation in $invocations) {
            if ($null -ne $invocation.Process) {
                if (-not $invocation.Process.HasExited) {
                    try {
                        $invocation.Process.Kill()
                        [void]$invocation.Process.WaitForExit(5000)
                    } catch {
                        Write-Verbose ("Unable to terminate child PowerShell batch process during cleanup: {0}" -f $_.Exception.Message)
                    }
                }
                $invocation.Process.Dispose()
            }
        }
    }
}

function Read-JsonFile {
    param([Parameter(Mandatory)] [string]$Path)

    $resolvedPath = Resolve-FileSystemPath -Path $Path
    try {
        $value = [System.IO.File]::ReadAllText($resolvedPath) | ConvertFrom-Json -ErrorAction Stop
        if ($null -eq $value) {
            throw 'The file did not contain a JSON object.'
        }
        return $value
    } catch {
        throw "Could not read JSON file '$resolvedPath': $($_.Exception.Message)"
    }
}

function Get-AvailableArtifactPath {
    param(
        [Parameter(Mandatory)] [string]$Directory,
        [Parameter(Mandatory)] [string]$BaseName,
        [Parameter(Mandatory)] [string]$Extension
    )

    Ensure-Directory -Path $Directory
    if (-not $Extension.StartsWith('.')) {
        $Extension = ".$Extension"
    }

    for ($sequence = 0; $sequence -lt 10000; $sequence++) {
        $suffix = if ($sequence -eq 0) { '' } else { "-$sequence" }
        $candidate = Join-Path $Directory ("{0}{1}{2}" -f $BaseName, $suffix, $Extension)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    throw "Could not allocate an unused artifact path for '$BaseName' in '$Directory'."
}

function Get-DefaultSnapshotPath {
    param([Parameter(Mandatory)] [string]$Root)

    $snapshotDir = Join-Path $Root 'snapshots'
    $baseName = "$($env:COMPUTERNAME)-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Get-AvailableArtifactPath -Directory $snapshotDir -BaseName $baseName -Extension '.json'
}

function Get-SnapshotReportPath {
    param([Parameter(Mandatory)] [string]$SnapshotPath)

    [IO.Path]::ChangeExtension([IO.Path]::GetFullPath($SnapshotPath), 'txt')
}

function Get-SnapshotAssetRoot {
    param([Parameter(Mandatory)] [string]$SnapshotPath)

    $fullPath = [IO.Path]::GetFullPath($SnapshotPath)
    $snapshotDir = Split-Path -Parent $fullPath
    $baseName = [IO.Path]::GetFileNameWithoutExtension($fullPath)
    Join-Path $snapshotDir ($baseName + '.assets')
}

function Get-VerificationReportPath {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $verificationDir = Join-Path $Root 'verification'
    $baseName = [IO.Path]::GetFileNameWithoutExtension($SnapshotPath)
    $reportBaseName = "$baseName-restore-check-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Get-AvailableArtifactPath -Directory $verificationDir -BaseName $reportBaseName -Extension '.txt'
}

function Get-PermissiveVerificationReportPath {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $verificationDir = Join-Path $Root 'verification'
    $baseName = [IO.Path]::GetFileNameWithoutExtension($SnapshotPath)
    $reportBaseName = "$baseName-permissive-check-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Get-AvailableArtifactPath -Directory $verificationDir -BaseName $reportBaseName -Extension '.txt'
}

function Get-WdacVerificationReportPath {
    param([Parameter(Mandatory)] [string]$VerificationPath)

    $fullPath = [IO.Path]::GetFullPath($VerificationPath)
    $parent = Split-Path -Parent $fullPath
    $baseName = [IO.Path]::GetFileNameWithoutExtension($fullPath)
    Join-Path $parent ($baseName + '-wdac.txt')
}

function Get-PreflightReportPath {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [ValidateSet('Snapshot', 'Permissive', 'Restore')] [string]$Action
    )

    $preflightDir = Join-Path $Root 'preflight'
    $computerName = if (-not [string]::IsNullOrWhiteSpace([string]$env:COMPUTERNAME)) { [string]$env:COMPUTERNAME } else { [Environment]::MachineName }
    $baseName = "{0}-{1}-{2}" -f $computerName, (Get-Date -Format 'yyyyMMdd-HHmmss'), $Action.ToLowerInvariant()
    Get-AvailableArtifactPath -Directory $preflightDir -BaseName $baseName -Extension '.txt'
}

function Get-OperationPath {
    param([Parameter(Mandatory)] [string]$Root)

    Ensure-Directory -Path $Root
    Join-Path $Root 'current-operation.json'
}

function Get-SnapshotIntegrityManifest {
    param([Parameter(Mandatory)] [string]$SnapshotPath)

    $fullSnapshotPath = [IO.Path]::GetFullPath($SnapshotPath)
    if (-not (Test-Path -LiteralPath $fullSnapshotPath -PathType Leaf)) {
        throw "Snapshot file was not found: $fullSnapshotPath"
    }

    $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $fullSnapshotPath
    $assets = @()
    if (Test-Path -LiteralPath $assetRoot -PathType Container) {
        $assets = @(
            Get-ChildItem -LiteralPath $assetRoot -File -Recurse -ErrorAction Stop |
                Sort-Object -Property FullName |
                ForEach-Object {
                    $record = Get-SnapshotAssetCacheRecord -Path $_.FullName
                    [PSCustomObject]@{
                        RelativePath = $_.FullName.Substring($assetRoot.Length).TrimStart([char[]]@('\', '/'))
                        Sha256       = [string]$record.Sha256
                    }
                }
        )
    }

    [PSCustomObject]@{
        SnapshotSha256 = (Get-FileHash -LiteralPath $fullSnapshotPath -Algorithm SHA256).Hash
        Assets         = @($assets)
    }
}

function Assert-OperationSnapshotIntegrity {
    param([Parameter(Mandatory)] [object]$Operation)

    if (-not $Operation.PSObject.Properties['SnapshotIntegrity'] -or $null -eq $Operation.SnapshotIntegrity) {
        return
    }

    $snapshotPath = [string]$Operation.SnapshotPath
    $actual = Get-SnapshotIntegrityManifest -SnapshotPath $snapshotPath
    $expected = $Operation.SnapshotIntegrity
    if (
        -not $expected.PSObject.Properties['SnapshotSha256'] -or
        -not [string]::Equals([string]$expected.SnapshotSha256, [string]$actual.SnapshotSha256, [System.StringComparison]::OrdinalIgnoreCase)
    ) {
        throw "The snapshot no longer matches the operation journal: $snapshotPath"
    }

    $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $snapshotPath
    $actualAssetsByPath = @{}
    foreach ($asset in @($actual.Assets)) {
        $actualPath = Resolve-ContainedFileSystemPath `
            -Root $assetRoot `
            -RelativePath ([string]$asset.RelativePath) `
            -Description 'actual snapshot asset path'
        $actualAssetsByPath[$actualPath] = $asset
    }
    $expectedAssets = @(if ($expected.PSObject.Properties['Assets']) { $expected.Assets })
    foreach ($asset in $expectedAssets) {
        $relativePath = [string]$asset.RelativePath
        $assetPath = Resolve-ContainedFileSystemPath `
            -Root $assetRoot `
            -RelativePath $relativePath `
            -Description 'operation journal snapshot asset path'
        if (-not $actualAssetsByPath.ContainsKey($assetPath)) {
            throw "A snapshot sidecar asset recorded by the operation journal is missing: $relativePath"
        }

        $actualHash = [string]$actualAssetsByPath[$assetPath].Sha256
        if (-not [string]::Equals([string]$asset.Sha256, [string]$actualHash, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "A snapshot sidecar asset no longer matches the operation journal: $relativePath"
        }
    }
}

function Test-OperationTargetsSnapshot {
    param(
        [AllowNull()] [object]$Operation,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    if ($null -eq $Operation -or -not $Operation.PSObject.Properties['SnapshotPath']) {
        return $false
    }

    $operationPath = [IO.Path]::GetFullPath([string]$Operation.SnapshotPath).TrimEnd([char[]]@('\', '/'))
    $candidatePath = [IO.Path]::GetFullPath($SnapshotPath).TrimEnd([char[]]@('\', '/'))
    [string]::Equals($operationPath, $candidatePath, [System.StringComparison]::OrdinalIgnoreCase)
}

function Write-OperationState {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [string]$SnapshotPath,
        [Parameter(Mandatory)] [string]$Mode,
        [AllowNull()] [string[]]$IncludeId,
        [AllowNull()] [string[]]$ExcludeId
    )

    $state = [PSCustomObject]@{
        SchemaVersion     = 1
        OperationId       = [guid]::NewGuid().ToString('D')
        Mode              = $Mode
        Status            = 'Applying'
        SnapshotPath      = [IO.Path]::GetFullPath($SnapshotPath)
        SnapshotIntegrity = Get-SnapshotIntegrityManifest -SnapshotPath $SnapshotPath
        IncludeId         = @(Get-NormalizedIdFilter -Ids $IncludeId)
        ExcludeId         = @(Get-NormalizedIdFilter -Ids $ExcludeId)
        StartedAtUtc      = (Get-Date).ToUniversalTime().ToString('o')
        UpdatedAtUtc      = (Get-Date).ToUniversalTime().ToString('o')
        ComputerName      = $env:COMPUTERNAME
        Producer          = Get-WinDefStateRuntimeInfo
    }

    Write-JsonAtomic -Path (Get-OperationPath -Root $Root) -InputObject $state
    $state
}

function Update-OperationStateStatus {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [object]$Operation,
        [Parameter(Mandatory)] [ValidateSet('Applying', 'Applied', 'AppliedVerified', 'AppliedPendingReboot', 'ApplyFailed', 'ApplyVerificationFailed', 'Restoring', 'RestoreFailed', 'RestoreVerificationFailed', 'PartiallyRestored')] [string]$Status,
        [AllowNull()] [object]$PermissiveVerification,
        [string]$PermissiveVerificationReportPath
    )

    if (-not $Operation.PSObject.Properties['Status']) {
        $Operation | Add-Member -NotePropertyName Status -NotePropertyValue $Status
    } else {
        $Operation.Status = $Status
    }
    if (-not $Operation.PSObject.Properties['UpdatedAtUtc']) {
        $Operation | Add-Member -NotePropertyName UpdatedAtUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o'))
    } else {
        $Operation.UpdatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    }

    if ($PSBoundParameters.ContainsKey('PermissiveVerification') -and $null -ne $PermissiveVerification) {
        $verificationSummary = [PSCustomObject]@{
            VerifiedAtUtc       = [string]$PermissiveVerification.VerifiedAtUtc
            VerifiedCount       = [int]$PermissiveVerification.VerifiedCount
            PendingRebootCount  = [int]$PermissiveVerification.PendingRebootCount
            MismatchCount       = [int]$PermissiveVerification.MismatchCount
            MutationDurationMs  = if ($PermissiveVerification.PSObject.Properties['MutationMetrics']) { [double]$PermissiveVerification.MutationMetrics.DurationMs } else { $null }
            ReportPath          = if (-not [string]::IsNullOrWhiteSpace($PermissiveVerificationReportPath)) { [IO.Path]::GetFullPath($PermissiveVerificationReportPath) } else { $null }
            MismatchedIds       = @($PermissiveVerification.Results | Where-Object { [string]$_.Status -eq 'Mismatch' } | ForEach-Object { [string]$_.Id })
            PendingRebootIds    = @($PermissiveVerification.Results | Where-Object { [string]$_.Status -eq 'ConfiguredPendingReboot' } | ForEach-Object { [string]$_.Id })
        }
        if (-not $Operation.PSObject.Properties['PermissiveVerification']) {
            $Operation | Add-Member -NotePropertyName PermissiveVerification -NotePropertyValue $verificationSummary
        } else {
            $Operation.PermissiveVerification = $verificationSummary
        }
    }

    Write-JsonAtomic -Path (Get-OperationPath -Root $Root) -InputObject $Operation
    $Operation
}

function Get-NormalizedRestoreCheckpointIds {
    param([AllowNull()] [object[]]$Id)

    $result = [System.Collections.Generic.List[string]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($candidate in @($Id)) {
        $value = ([string]$candidate).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value) -and $seen.Add($value)) {
            $result.Add($value) | Out-Null
        }
    }
    @($result)
}

function Initialize-RestoreCheckpoint {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [object]$Operation,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]]$RequestedIds
    )

    $existing = if ($Operation.PSObject.Properties['RestoreCheckpoint']) { $Operation.RestoreCheckpoint } else { $null }
    $completedIds = if ($null -ne $existing -and $existing.PSObject.Properties['CompletedIds']) {
        @(Get-NormalizedRestoreCheckpointIds -Id @($existing.CompletedIds))
    } else {
        @()
    }
    $attemptNumber = if ($null -ne $existing -and $existing.PSObject.Properties['AttemptNumber']) {
        try { [int]$existing.AttemptNumber + 1 } catch { 1 }
    } else {
        1
    }
    $now = (Get-Date).ToUniversalTime().ToString('o')
    $checkpoint = [PSCustomObject]@{
        AttemptNumber        = $attemptNumber
        AttemptId            = [guid]::NewGuid().ToString('D')
        AttemptStartedAtUtc  = $now
        UpdatedAtUtc         = $now
        RequestedIds         = @(Get-NormalizedRestoreCheckpointIds -Id $RequestedIds)
        CompletedIds         = @($completedIds)
        VerifiedCompletedIds = @()
        CurrentWorkItemId    = $null
        CurrentIds           = @()
        LastFailure          = $null
    }

    if ($Operation.PSObject.Properties['RestoreCheckpoint']) {
        $Operation.RestoreCheckpoint = $checkpoint
    } else {
        $Operation | Add-Member -NotePropertyName RestoreCheckpoint -NotePropertyValue $checkpoint
    }
    Write-JsonAtomic -Path (Get-OperationPath -Root $Root) -InputObject $Operation
    $Operation
}

function Set-RestoreCheckpointVerifiedIds {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [object]$Operation,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]]$VerifiedIds
    )

    $Operation.RestoreCheckpoint.VerifiedCompletedIds = @(Get-NormalizedRestoreCheckpointIds -Id $VerifiedIds)
    $Operation.RestoreCheckpoint.UpdatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    Write-JsonAtomic -Path (Get-OperationPath -Root $Root) -InputObject $Operation
    $Operation
}

function Start-RestoreCheckpointWorkItem {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [object]$Operation,
        [Parameter(Mandatory)] [string]$WorkItemId,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]]$SettingIds
    )

    $Operation.RestoreCheckpoint.CurrentWorkItemId = $WorkItemId
    $Operation.RestoreCheckpoint.CurrentIds = @(Get-NormalizedRestoreCheckpointIds -Id $SettingIds)
    $Operation.RestoreCheckpoint.LastFailure = $null
    $Operation.RestoreCheckpoint.UpdatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    Write-JsonAtomic -Path (Get-OperationPath -Root $Root) -InputObject $Operation
    $Operation
}

function Complete-RestoreCheckpointWorkItem {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [object]$Operation,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [string[]]$SettingIds
    )

    $Operation.RestoreCheckpoint.CompletedIds = @(
        Get-NormalizedRestoreCheckpointIds -Id (@($Operation.RestoreCheckpoint.CompletedIds) + @($SettingIds))
    )
    $Operation.RestoreCheckpoint.CurrentWorkItemId = $null
    $Operation.RestoreCheckpoint.CurrentIds = @()
    $Operation.RestoreCheckpoint.UpdatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    Write-JsonAtomic -Path (Get-OperationPath -Root $Root) -InputObject $Operation
    $Operation
}

function Set-RestoreCheckpointFailure {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [object]$Operation,
        [Parameter(Mandatory)] [string]$Message
    )

    $safeMessage = if ($Message.Length -gt 2048) { $Message.Substring(0, 2048) } else { $Message }
    $Operation.RestoreCheckpoint.LastFailure = [PSCustomObject]@{
        FailedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        WorkItemId  = $Operation.RestoreCheckpoint.CurrentWorkItemId
        SettingIds  = @($Operation.RestoreCheckpoint.CurrentIds)
        Message     = $safeMessage
    }
    $Operation.RestoreCheckpoint.UpdatedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    Write-JsonAtomic -Path (Get-OperationPath -Root $Root) -InputObject $Operation
    $Operation
}

function Get-OperationState {
    param([Parameter(Mandatory)] [string]$Root)

    $path = Get-OperationPath -Root $Root
    if (-not (Test-Path -LiteralPath $path)) {
        return $null
    }

    $operation = Read-JsonFile -Path $path
    if (
        -not $operation.PSObject.Properties['SnapshotPath'] -or
        [string]::IsNullOrWhiteSpace([string]$operation.SnapshotPath)
    ) {
        throw "The operation journal is missing its snapshot path: $path"
    }

    $operation
}

function Assert-NoActiveOperation {
    param([Parameter(Mandatory)] [string]$Root)

    $operation = Get-OperationState -Root $Root
    if ($null -eq $operation) {
        return
    }

    $snapshotPath = if ($operation.PSObject.Properties['SnapshotPath']) { [string]$operation.SnapshotPath } else { '<unknown>' }
    $status = if ($operation.PSObject.Properties['Status']) { [string]$operation.Status } else { 'Legacy journal' }
    throw "An active WinDefState operation already exists (Status=$status, Snapshot=$snapshotPath). Restore it before starting another permissive operation."
}

function Clear-OperationState {
    param([Parameter(Mandatory)] [string]$Root)

    $path = Get-OperationPath -Root $Root
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

#endregion

#region Registry, service, and user-profile providers

function Ensure-RegistryPath {
    param([Parameter(Mandatory)] [string]$Path)

    if (-not (Test-Path -Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
}

function Remove-RegistryKeyIfExists {
    param([Parameter(Mandatory)] [string]$Path)

    if (Test-Path -Path $Path) {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-RegistryValueCaptureState {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [string]$DefaultValueKind,
        [AllowNull()] [object]$CaptureSession
    )

    $cacheKey = "registry.key:{0}" -f $Path.ToLowerInvariant()
    $keyState = Get-CaptureSessionValue -Session $CaptureSession -Key $cacheKey -Factory {
        try {
            $key = Get-Item -Path $Path -ErrorAction Stop
            [PSCustomObject]@{
                Captured   = $true
                Error      = $null
                Key        = $key
                Properties = Get-ItemProperty -Path $Path -ErrorAction Stop
            }
        } catch {
            if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                return [PSCustomObject]@{
                    Captured   = $true
                    Error      = $null
                    Key        = $null
                    Properties = $null
                }
            }

            [PSCustomObject]@{
                Captured   = $false
                Error      = $_.Exception.Message
                Key        = $null
                Properties = $null
            }
        }
    }

    $property = if ($keyState.Captured -and $null -ne $keyState.Properties) { $keyState.Properties.PSObject.Properties[$Name] } else { $null }
    $exists = $null -ne $property
    $valueKind = if ($exists -and $null -ne $keyState.Key) {
        try { $keyState.Key.GetValueKind($Name).ToString() } catch { $DefaultValueKind }
    } else {
        $DefaultValueKind
    }

    [PSCustomObject]@{
        Captured     = [bool]$keyState.Captured
        Error        = [string]$keyState.Error
        Exists       = $exists
        CurrentValue = if ($exists) { $property.Value } else { $null }
        ValueKind    = $valueKind
    }
}

function Get-ServiceCaptureState {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [AllowNull()] [object]$CaptureSession
    )

    if ($null -eq $CaptureSession) {
        $escapedName = $Name.Replace("'", "''")
        return Get-CimInstance -ClassName Win32_Service -Filter "Name='$escapedName'" -ErrorAction SilentlyContinue
    }

    $serviceState = Get-CaptureSessionValue -Session $CaptureSession -Key 'services' -Factory {
        $serviceNames = @(
            @(@($CaptureSession.ServiceNames) + @($Name)) |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )
        $filterParts = @(
            foreach ($serviceName in $serviceNames) {
                $escapedName = ([string]$serviceName).Replace("'", "''")
                "Name='$escapedName'"
            }
        )
        $services = @(
            if ($filterParts.Count -gt 0) {
                Get-CimInstance -ClassName Win32_Service -Filter ($filterParts -join ' OR ') -ErrorAction SilentlyContinue
            }
        )
        $byName = @{}
        $queriedNames = @{}
        foreach ($serviceName in $serviceNames) {
            $queriedNames[[string]$serviceName] = $true
        }
        foreach ($service in $services) {
            if ($null -ne $service -and -not [string]::IsNullOrWhiteSpace([string]$service.Name)) {
                $byName[[string]$service.Name] = $service
            }
        }

        [PSCustomObject]@{
            ByName       = $byName
            QueriedNames = $queriedNames
        }
    }

    if ($serviceState.ByName.ContainsKey($Name)) {
        return $serviceState.ByName[$Name]
    }

    if (-not $serviceState.QueriedNames.ContainsKey($Name)) {
        $escapedName = $Name.Replace("'", "''")
        $additionalService = Get-CaptureSessionValue -Session $CaptureSession -Key ("service:{0}" -f $Name.ToLowerInvariant()) -Factory {
            Get-CimInstance -ClassName Win32_Service -Filter "Name='$escapedName'" -ErrorAction SilentlyContinue
        }
        $serviceState.QueriedNames[$Name] = $true
        if ($null -ne $additionalService) {
            $serviceState.ByName[$Name] = $additionalService
            return $additionalService
        }
    }

    $null
}

function Get-RegistryValueEntries {
    param([Parameter(Mandatory)] [string]$Path)

    if (-not (Test-Path -Path $Path -ErrorAction Stop)) {
        return @()
    }

    $item = Get-ItemProperty -Path $Path -ErrorAction Stop
    $key = Get-Item -Path $Path -ErrorAction Stop

    $values = foreach ($property in $item.PSObject.Properties) {
        if ($property.Name -like 'PS*') {
            continue
        }

        [PSCustomObject]@{
            Name      = $property.Name
            Value     = $property.Value
            ValueKind = $key.GetValueKind($property.Name).ToString()
        }
    }

    @($values)
}

function Get-RegistryKeyFlatCaptureState {
    param([Parameter(Mandatory)] [string]$Path)

    try {
        $exists = Test-Path -Path $Path -ErrorAction Stop
        [PSCustomObject]@{
            Captured = $true
            Error    = $null
            Exists   = $exists
            Values   = if ($exists) { @(Get-RegistryValueEntries -Path $Path) } else { @() }
        }
    } catch {
        [PSCustomObject]@{
            Captured = $false
            Error    = $_.Exception.Message
            Exists   = $false
            Values   = @()
        }
    }
}

function Set-RegistryKeyValuesExact {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [object[]]$Values
    )

    Ensure-RegistryPath -Path $Path

    foreach ($existingValue in @(Get-RegistryValueEntries -Path $Path)) {
        Remove-ItemProperty -Path $Path -Name $existingValue.Name -ErrorAction SilentlyContinue
    }

    foreach ($value in @($Values)) {
        New-ItemProperty -Path $Path -Name $value.Name -PropertyType $value.ValueKind -Value $value.Value -Force | Out-Null
    }
}

function Capture-RegistryKeyFlatState {
    param(
        [Parameter(Mandatory)] [string]$Id,
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [bool]$RequiresReboot
    )

    $state = Get-RegistryKeyFlatCaptureState -Path $Path
    [PSCustomObject]@{
        Id             = $Id
        Type           = 'RegistryKeyFlat'
        Path           = $Path
        Captured       = $state.Captured
        CaptureError   = $state.Error
        Exists         = $state.Exists
        CurrentValue   = @($state.Values)
        RequiresReboot = $RequiresReboot
    }
}

function Restore-RegistryKeyFlatState {
    param([Parameter(Mandatory)] [object]$Entry)

    if ($Entry.Exists) {
        Set-RegistryKeyValuesExact -Path $Entry.Path -Values @($Entry.CurrentValue)
    } else {
        Remove-RegistryKeyIfExists -Path $Entry.Path
    }
}

function Capture-PowerShellModuleLoggingState {
    $basePath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ModuleLogging'
    $moduleNamesPath = Join-Path $basePath 'ModuleNames'
    $baseState = Get-RegistryKeyFlatCaptureState -Path $basePath
    $moduleNamesState = Get-RegistryKeyFlatCaptureState -Path $moduleNamesPath
    $captureErrors = @(
        if (-not $baseState.Captured) { [string]$baseState.Error }
        if (-not $moduleNamesState.Captured) { [string]$moduleNamesState.Error }
    )

    [PSCustomObject]@{
        Id             = 'powershell.module_logging'
        Type           = 'PowerShellModuleLogging'
        BasePath       = $basePath
        Captured       = $baseState.Captured -and $moduleNamesState.Captured
        CaptureError   = $captureErrors -join '; '
        Exists         = $baseState.Exists
        CurrentValue   = [PSCustomObject]@{
            BaseValues        = @($baseState.Values)
            ModuleNamesExists = $moduleNamesState.Exists
            ModuleNamesValues = @($moduleNamesState.Values)
        }
        RequiresReboot = $false
    }
}

function Set-Permissive-PowerShellModuleLogging {
    Remove-RegistryKeyIfExists -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ModuleLogging'
}

function Restore-PowerShellModuleLogging {
    param([Parameter(Mandatory)] [object]$Entry)

    $basePath = [string]$Entry.BasePath
    $moduleNamesPath = Join-Path $basePath 'ModuleNames'

    if (-not $Entry.Exists) {
        Remove-RegistryKeyIfExists -Path $basePath
        return
    }

    Set-RegistryKeyValuesExact -Path $basePath -Values @($Entry.CurrentValue.BaseValues)

    if ($Entry.CurrentValue.ModuleNamesExists) {
        Set-RegistryKeyValuesExact -Path $moduleNamesPath -Values @($Entry.CurrentValue.ModuleNamesValues)
    } else {
        Remove-RegistryKeyIfExists -Path $moduleNamesPath
    }
}

function Get-LoadedUserSids {
    $sids = foreach ($key in Get-ChildItem -Path Registry::HKEY_USERS -ErrorAction SilentlyContinue) {
        if ($key.PSChildName -match '^S-\d-\d+-.+' -and $key.PSChildName -notmatch '_Classes$') {
            $key.PSChildName
        }
    }

    @($sids)
}

function Get-UserProfileRegistryTargets {
    $targetsBySid = @{}
    $loadedSidSet = @{}

    foreach ($sid in @(Get-LoadedUserSids)) {
        $loadedSidSet[[string]$sid] = $true
    }

    if (Test-CommandAvailable -Name 'Get-CimInstance') {
        try {
            foreach ($userProfile in @(Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop)) {
                $sid = if ($null -ne $userProfile.PSObject.Properties['SID']) { [string]$userProfile.SID } else { '' }
                if ([string]::IsNullOrWhiteSpace($sid) -or $sid -notmatch '^S-\d-\d+-.+') {
                    continue
                }

                $isSpecial = $false
                if ($null -ne $userProfile.PSObject.Properties['Special'] -and $null -ne $userProfile.Special) {
                    $isSpecial = [bool]$userProfile.Special
                }

                if ($isSpecial) {
                    continue
                }

                $profilePath = $null
                if ($null -ne $userProfile.PSObject.Properties['LocalPath'] -and -not [string]::IsNullOrWhiteSpace([string]$userProfile.LocalPath)) {
                    $profilePath = Resolve-FileSystemPath -Path ([string]$userProfile.LocalPath)
                }

                $targetsBySid[$sid] = [PSCustomObject]@{
                    Sid         = $sid
                    ProfilePath = $profilePath
                    HivePath    = if (-not [string]::IsNullOrWhiteSpace($profilePath)) { Join-Path $profilePath 'NTUSER.DAT' } else { $null }
                    Loaded      = ($loadedSidSet.ContainsKey($sid) -or ($null -ne $userProfile.PSObject.Properties['Loaded'] -and [bool]$userProfile.Loaded))
                }
            }
        } catch {
            Write-Verbose ("Failed to enumerate Win32_UserProfile instances: {0}" -f $_.Exception.Message)
        }
    }

    foreach ($profileKey in @(Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' -ErrorAction SilentlyContinue)) {
        $sid = [string]$profileKey.PSChildName
        if ([string]::IsNullOrWhiteSpace($sid) -or $sid -notmatch '^S-\d-\d+-.+') {
            continue
        }

        if ($targetsBySid.ContainsKey($sid)) {
            continue
        }

        $profilePath = $null
        $properties = Get-ItemProperty -Path $profileKey.PSPath -Name 'ProfileImagePath' -ErrorAction SilentlyContinue
        if ($null -ne $properties -and -not [string]::IsNullOrWhiteSpace([string]$properties.ProfileImagePath)) {
            $profilePath = Resolve-FileSystemPath -Path ([string]$properties.ProfileImagePath)
        }

        $targetsBySid[$sid] = [PSCustomObject]@{
            Sid         = $sid
            ProfilePath = $profilePath
            HivePath    = if (-not [string]::IsNullOrWhiteSpace($profilePath)) { Join-Path $profilePath 'NTUSER.DAT' } else { $null }
            Loaded      = $loadedSidSet.ContainsKey($sid)
        }
    }

    foreach ($sid in @($loadedSidSet.Keys)) {
        if ($targetsBySid.ContainsKey($sid)) {
            continue
        }

        $targetsBySid[$sid] = [PSCustomObject]@{
            Sid         = $sid
            ProfilePath = $null
            HivePath    = $null
            Loaded      = $true
        }
    }

    @(
        foreach ($entry in @($targetsBySid.GetEnumerator() | Sort-Object -Property Name)) {
            $entry.Value
        }
    )
}

function Open-UserRegistryTarget {
    param([Parameter(Mandatory)] [object]$Target)

    $sid = [string]$Target.Sid
    if ([string]::IsNullOrWhiteSpace($sid)) {
        throw 'User registry target SID is missing.'
    }

    $loadedRootPath = "Registry::HKEY_USERS\$sid"
    if (Test-Path -LiteralPath $loadedRootPath) {
        return [PSCustomObject]@{
            Sid           = $sid
            ProfilePath   = if ($null -ne $Target.PSObject.Properties['ProfilePath']) { [string]$Target.ProfilePath } else { $null }
            HivePath      = if ($null -ne $Target.PSObject.Properties['HivePath']) { [string]$Target.HivePath } else { $null }
            RootPath      = $loadedRootPath
            MountName     = $sid
            MountedByTool = $false
        }
    }

    $hivePath = if ($null -ne $Target.PSObject.Properties['HivePath']) { [string]$Target.HivePath } else { $null }
    if ([string]::IsNullOrWhiteSpace($hivePath)) {
        throw ("User profile hive path is unavailable for SID {0}." -f $sid)
    }

    if (-not (Test-Path -LiteralPath $hivePath)) {
        throw ("User profile hive file was not found for SID {0} at {1}." -f $sid, $hivePath)
    }

    $mountName = "WinDefState_{0}" -f ([guid]::NewGuid().ToString('N'))
    $loadOutput = & reg.exe load "HKU\$mountName" $hivePath 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = (@($loadOutput | ForEach-Object { $_.ToString().Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = "reg.exe load failed with exit code $LASTEXITCODE."
        }

        throw ("Failed to load user hive for SID {0} from {1}. {2}" -f $sid, $hivePath, $message)
    }

    [PSCustomObject]@{
        Sid           = $sid
        ProfilePath   = if ($null -ne $Target.PSObject.Properties['ProfilePath']) { [string]$Target.ProfilePath } else { $null }
        HivePath      = $hivePath
        RootPath      = "Registry::HKEY_USERS\$mountName"
        MountName     = $mountName
        MountedByTool = $true
    }
}

function Close-UserRegistryTarget {
    param([AllowNull()] [object]$Target)

    if ($null -eq $Target) {
        return
    }

    $mountedByTool = $false
    if ($null -ne $Target.PSObject.Properties['MountedByTool'] -and $Target.MountedByTool) {
        $mountedByTool = $true
    }

    if (-not $mountedByTool) {
        return
    }

    $mountName = if ($null -ne $Target.PSObject.Properties['MountName']) { [string]$Target.MountName } else { $null }
    if ([string]::IsNullOrWhiteSpace($mountName)) {
        return
    }

    $unloadOutput = & reg.exe unload "HKU\$mountName" 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = (@($unloadOutput | ForEach-Object { $_.ToString().Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = "reg.exe unload failed with exit code $LASTEXITCODE."
        }

        Write-Warning ("Failed to unload temporary user hive mount HKU\{0} for SID {1}. {2}" -f $mountName, ([string]$Target.Sid), $message)
    }
}

function New-UserRegistryCaptureIssue {
    param(
        [Parameter(Mandatory)] [string]$Sid,
        [string]$ProfilePath,
        [string]$HivePath,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        Sid         = $Sid
        ProfilePath = $ProfilePath
        HivePath    = $HivePath
        Message     = $Message
    }
}

function Normalize-UserRegistryValueState {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return [PSCustomObject]@{
            Entries       = @()
            CaptureIssues = @()
        }
    }

    $entries = @()
    $captureIssues = @()

    if ($null -ne $State.PSObject.Properties['Entries']) {
        $entries = @($State.Entries)
    } else {
        $entries = @($State)
    }

    if ($null -ne $State.PSObject.Properties['CaptureIssues']) {
        $captureIssues = @($State.CaptureIssues)
    }

    [PSCustomObject]@{
        Entries       = @($entries)
        CaptureIssues = @($captureIssues)
    }
}

function Get-UserRegistryTargetAccessState {
    param(
        [Parameter(Mandatory)] [object]$Target,
        [AllowNull()] [object]$CaptureSession
    )

    if ($null -eq $CaptureSession) {
        try {
            return [PSCustomObject]@{
                Captured = $true
                Access   = Open-UserRegistryTarget -Target $Target
                Error    = $null
            }
        } catch {
            return [PSCustomObject]@{
                Captured = $false
                Access   = $null
                Error    = $_.Exception.Message
            }
        }
    }

    Get-CaptureSessionValue `
        -Session $CaptureSession `
        -Key ("user.registry.access:{0}" -f ([string]$Target.Sid)) `
        -Factory {
            try {
                $openedAccess = Open-UserRegistryTarget -Target $Target
                if ($openedAccess.PSObject.Properties['MountedByTool'] -and [bool]$openedAccess.MountedByTool) {
                    Register-CaptureSessionResource -Session $CaptureSession -Kind 'UserRegistryTarget' -Value $openedAccess
                }
                return [PSCustomObject]@{
                    Captured = $true
                    Access   = $openedAccess
                    Error    = $null
                }
            } catch {
                return [PSCustomObject]@{
                    Captured = $false
                    Access   = $null
                    Error    = $_.Exception.Message
                }
            }
        }
}

function Get-LoadedUserRegistryValueStates {
    param(
        [Parameter(Mandatory)] [object[]]$Items,
        [AllowNull()] [object]$CaptureSession
    )

    $entries = @()
    $captureIssues = @()
    $targets = @(Get-CaptureSessionValue -Session $CaptureSession -Key 'user.registry.targets' -Factory {
        @(Get-UserProfileRegistryTargets)
    })

    foreach ($target in $targets) {
        $access = $null
        $closeAfterCapture = $null -eq $CaptureSession

        try {
            $accessState = Get-UserRegistryTargetAccessState -Target $target -CaptureSession $CaptureSession
            if (-not $accessState.Captured) {
                throw [System.InvalidOperationException]::new([string]$accessState.Error)
            }
            $access = $accessState.Access

            foreach ($item in @($Items)) {
                $path = "$($access.RootPath)\$($item.RelativePath)"
                $property = Get-ItemProperty -Path $path -Name $item.Name -ErrorAction SilentlyContinue
                $exists = $null -ne $property -and $null -ne $property.$($item.Name)
                $kind = if ($exists) {
                    try { (Get-Item -Path $path).GetValueKind($item.Name).ToString() } catch { $item.ValueKind }
                } else {
                    $item.ValueKind
                }

                $entries += [PSCustomObject]@{
                    Sid          = [string]$target.Sid
                    ProfilePath  = if ($null -ne $target.PSObject.Properties['ProfilePath']) { [string]$target.ProfilePath } else { $null }
                    HivePath     = if ($null -ne $access.PSObject.Properties['HivePath']) { [string]$access.HivePath } else { $null }
                    RelativePath = $item.RelativePath
                    Name         = $item.Name
                    Exists       = $exists
                    CurrentValue = if ($exists) { $property.$($item.Name) } else { $null }
                    ValueKind    = $kind
                }
            }
        } catch {
            $message = $_.Exception.Message
            $profilePath = if ($null -ne $target.PSObject.Properties['ProfilePath']) { [string]$target.ProfilePath } else { $null }
            $hivePath = if ($null -ne $target.PSObject.Properties['HivePath']) { [string]$target.HivePath } else { $null }
            $captureIssues += New-UserRegistryCaptureIssue -Sid ([string]$target.Sid) -ProfilePath $profilePath -HivePath $hivePath -Message $message
            Write-Warning ("Failed to capture user-scoped registry values for SID {0}: {1}" -f ([string]$target.Sid), $message)
        } finally {
            if ($closeAfterCapture) {
                Close-UserRegistryTarget -Target $access
            }
        }
    }

    [PSCustomObject]@{
        Entries       = @($entries)
        CaptureIssues = @($captureIssues)
    }
}

function Set-Permissive-LoadedUserRegistryValues {
    param(
        [Parameter(Mandatory)] [object[]]$Items,
        [AllowNull()] [object]$CaptureSession
    )

    $targets = @(Get-CaptureSessionValue -Session $CaptureSession -Key 'user.registry.targets' -Factory {
        @(Get-UserProfileRegistryTargets)
    })
    foreach ($target in $targets) {
        $access = $null
        $closeAfterMutation = $null -eq $CaptureSession

        try {
            $accessState = Get-UserRegistryTargetAccessState -Target $target -CaptureSession $CaptureSession
            if (-not $accessState.Captured) {
                throw [System.InvalidOperationException]::new([string]$accessState.Error)
            }
            $access = $accessState.Access

            foreach ($item in @($Items)) {
                $path = "$($access.RootPath)\$($item.RelativePath)"

                if ($item.PermissiveExists) {
                    Ensure-RegistryPath -Path $path
                    New-ItemProperty -Path $path -Name $item.Name -PropertyType $item.ValueKind -Value $item.PermissiveValue -Force | Out-Null
                } else {
                    Remove-ItemProperty -Path $path -Name $item.Name -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-Warning ("Failed to set permissive user-scoped registry values for SID {0}: {1}" -f ([string]$target.Sid), $_.Exception.Message)
        } finally {
            if ($closeAfterMutation) {
                Close-UserRegistryTarget -Target $access
            }
        }
    }
}

function Restore-LoadedUserRegistryValues {
    param(
        [Parameter(Mandatory)] [object[]]$Entries,
        [AllowNull()] [object]$CaptureSession
    )

    $targetsBySid = @{}
    $targets = @(Get-CaptureSessionValue -Session $CaptureSession -Key 'user.registry.targets' -Factory {
        @(Get-UserProfileRegistryTargets)
    })
    foreach ($target in $targets) {
        $targetsBySid[[string]$target.Sid] = $target
    }

    foreach ($group in @($Entries | Group-Object -Property Sid)) {
        $groupEntries = @($group.Group)
        if ($groupEntries.Count -eq 0) {
            continue
        }

        $sid = [string]$group.Name
        $sampleEntry = $groupEntries[0]
        $target = if ($targetsBySid.ContainsKey($sid)) {
            $targetsBySid[$sid]
        } else {
            [PSCustomObject]@{
                Sid         = $sid
                ProfilePath = if ($null -ne $sampleEntry.PSObject.Properties['ProfilePath']) { [string]$sampleEntry.ProfilePath } else { $null }
                HivePath    = if ($null -ne $sampleEntry.PSObject.Properties['HivePath']) { [string]$sampleEntry.HivePath } else { $null }
                Loaded      = $false
            }
        }

        $access = $null
        $closeAfterMutation = $null -eq $CaptureSession
        try {
            $accessState = Get-UserRegistryTargetAccessState -Target $target -CaptureSession $CaptureSession
            if (-not $accessState.Captured) {
                throw [System.InvalidOperationException]::new([string]$accessState.Error)
            }
            $access = $accessState.Access

            foreach ($entry in $groupEntries) {
                $path = "$($access.RootPath)\$($entry.RelativePath)"

                if ($entry.Exists) {
                    Ensure-RegistryPath -Path $path
                    New-ItemProperty -Path $path -Name $entry.Name -PropertyType $entry.ValueKind -Value $entry.CurrentValue -Force | Out-Null
                } else {
                    Remove-ItemProperty -Path $path -Name $entry.Name -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-Warning ("Failed to restore user-scoped registry values for SID {0}: {1}" -f $sid, $_.Exception.Message)
        } finally {
            if ($closeAfterMutation) {
                Close-UserRegistryTarget -Target $access
            }
        }
    }
}

function Resolve-LocalUserTarget {
    param([Parameter(Mandatory)] [object]$Reference)

    $sidProperty = $Reference.PSObject.Properties['Sid']
    if ($null -ne $sidProperty -and -not [string]::IsNullOrWhiteSpace([string]$sidProperty.Value)) {
        try {
            return Get-LocalUser -SID ([Security.Principal.SecurityIdentifier]$sidProperty.Value) -ErrorAction SilentlyContinue
        } catch {
            Write-Verbose ("Unable to resolve local user directly by SID; falling back to RID/name matching: {0}" -f $_.Exception.Message)
        }
    }

    $ridProperty = $Reference.PSObject.Properties['Rid']
    if ($null -ne $ridProperty -and $null -ne $ridProperty.Value) {
        $ridSuffix = "-$([int]$ridProperty.Value)$"
        $user = Get-LocalUser -ErrorAction SilentlyContinue |
            Where-Object { $null -ne $_.SID -and $_.SID.Value -match $ridSuffix } |
            Select-Object -First 1

        if ($null -ne $user) {
            return $user
        }
    }

    $nameProperty = $Reference.PSObject.Properties['Name']
    if ($null -ne $nameProperty -and -not [string]::IsNullOrWhiteSpace([string]$nameProperty.Value)) {
        return Get-LocalUser -Name ([string]$nameProperty.Value) -ErrorAction SilentlyContinue
    }

    $null
}

function Set-LocalUserEnabledState {
    param(
        [Parameter(Mandatory)] [object]$Reference,
        [Parameter(Mandatory)] [bool]$Enabled
    )

    $user = Resolve-LocalUserTarget -Reference $Reference
    if ($null -eq $user -or $null -eq $user.SID) {
        throw 'The target local user could not be resolved.'
    }

    if ($Enabled) {
        Enable-LocalUser -SID $user.SID -ErrorAction Stop
    } else {
        Disable-LocalUser -SID $user.SID -ErrorAction Stop
    }
}

#endregion

#region Remote management and network providers

function ConvertTo-WsManTextValue {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [bool]) {
        return $Value.ToString().ToLowerInvariant()
    }

    switch -Exact (([string]$Value).ToLowerInvariant()) {
        'true' { 'true' }
        'false' { 'false' }
        default { [string]$Value }
    }
}

function ConvertFrom-WsManTextValue {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    switch -Exact (([string]$Value).Trim().ToLowerInvariant()) {
        'true' { return $true }
        'false' { return $false }
        '1' { return $true }
        '0' { return $false }
        default { return ([string]$Value).Trim() }
    }
}

function ConvertTo-NullableBoolean {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [bool]) {
        return [bool]$Value
    }

    switch -Exact (([string]$Value).Trim().ToLowerInvariant()) {
        'true' { return $true }
        'false' { return $false }
        '1' { return $true }
        '0' { return $false }
        default { return $null }
    }
}

function Resolve-WsManConfigTarget {
    param([Parameter(Mandatory)] [string]$Path)

    $prefix = 'WSMan:\localhost\'
    if (-not $Path.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsupported WSMan path: $Path"
    }

    $relativePath = $Path.Substring($prefix.Length)
    $segments = @($relativePath -split '\\')
    if ($segments.Count -lt 2) {
        throw "Unsupported WSMan path: $Path"
    }

    $role = [string]$segments[0]
    $resourceRole = switch -Exact ($role.ToLowerInvariant()) {
        'service' { 'service' }
        'client' { 'client' }
        default { throw "Unsupported WSMan path role: $Path" }
    }

    if ($segments.Count -eq 2) {
        $name = [string]$segments[1]
        switch -Exact ($name) {
            'AllowUnencrypted' {
                return [PSCustomObject]@{
                    ResourceUri = "winrm/config/$resourceRole"
                    Property    = 'AllowUnencrypted'
                    ValueKind   = 'Boolean'
                }
            }
            'TrustedHosts' {
                if ($resourceRole -ne 'client') {
                    throw "Unsupported WSMan path for TrustedHosts: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/client'
                    Property    = 'TrustedHosts'
                    ValueKind   = 'String'
                }
            }
            'IPv4Filter' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for IPv4Filter: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'IPv4Filter'
                    ValueKind   = 'String'
                }
            }
            'IPv6Filter' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for IPv6Filter: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'IPv6Filter'
                    ValueKind   = 'String'
                }
            }
            'EnableCompatibilityHttpListener' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for EnableCompatibilityHttpListener: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'EnableCompatibilityHttpListener'
                    ValueKind   = 'Boolean'
                }
            }
            'EnableCompatibilityHttpsListener' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for EnableCompatibilityHttpsListener: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'EnableCompatibilityHttpsListener'
                    ValueKind   = 'Boolean'
                }
            }
            'CbtHardeningLevel' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for CbtHardeningLevel: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'CbtHardeningLevel'
                    ValueKind   = 'String'
                }
            }
            default {
                throw "Unsupported WSMan path: $Path"
            }
        }
    }

    if ($segments.Count -eq 3 -and [string]$segments[1] -eq 'Auth') {
        $property = [string]$segments[2]
        $supportedAuthProperties = @(if ($resourceRole -eq 'client') {
            'Basic', 'Digest', 'Kerberos', 'Negotiate', 'Certificate', 'CredSSP'
        } else {
            'Basic', 'Kerberos', 'Negotiate', 'Certificate', 'CredSSP'
        })

        if ($supportedAuthProperties -notcontains $property) {
            throw "Unsupported WSMan auth path: $Path"
        }

        return [PSCustomObject]@{
            ResourceUri = "winrm/config/$resourceRole/auth"
            Property    = $property
            ValueKind   = 'Boolean'
        }
    }

    throw "Unsupported WSMan path: $Path"
}

function ConvertFrom-WsManPropertyValue {
    param(
        [AllowNull()] [object]$Value,
        [Parameter(Mandatory)] [string]$ValueKind
    )

    switch ($ValueKind) {
        'Boolean' { return (ConvertTo-NullableBoolean -Value $Value) }
        'String' {
            if ($null -eq $Value) {
                return $null
            }

            return [string]$Value
        }
        default {
            throw "Unsupported WSMan value kind: $ValueKind"
        }
    }
}

function ConvertTo-WsManSetValue {
    param(
        [AllowNull()] [object]$Value,
        [Parameter(Mandatory)] [string]$ValueKind
    )

    switch ($ValueKind) {
        'Boolean' {
            $boolValue = ConvertTo-NullableBoolean -Value $Value
            if ($null -eq $boolValue) {
                throw "Unsupported WSMan boolean value: $Value"
            }

            return (ConvertTo-WsManTextValue -Value $boolValue)
        }
        'String' {
            if ($null -eq $Value) {
                return ''
            }

            return [string]$Value
        }
        default {
            throw "Unsupported WSMan value kind: $ValueKind"
        }
    }
}

function Get-WsManConfigValueState {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [AllowNull()] [object]$CaptureSession
    )

    $target = Resolve-WsManConfigTarget -Path $Path
    $resourceState = Get-CaptureSessionValue -Session $CaptureSession -Key ("winrm.instance:{0}" -f $target.ResourceUri) -Factory {
        if (-not (Test-CommandAvailable -Name 'Get-WSManInstance')) {
            return [PSCustomObject]@{
                CommandAvailable = $false
                Captured         = $false
                Value            = $null
                Error            = 'Get-WSManInstance was not found.'
            }
        }

        try {
            return [PSCustomObject]@{
                CommandAvailable = $true
                Captured         = $true
                Value            = Get-WSManInstance -ResourceURI $target.ResourceUri -ErrorAction Stop
                Error            = $null
            }
        } catch {
            return [PSCustomObject]@{
                CommandAvailable = $true
                Captured         = $false
                Value            = $null
                Error            = $_.Exception.Message
            }
        }
    }

    if (-not $resourceState.Captured -or $null -eq $resourceState.Value) {
        return [PSCustomObject]@{
            CommandAvailable = [bool]$resourceState.CommandAvailable
            Captured         = $false
            Value            = $null
            Error            = $resourceState.Error
        }
    }

    $instance = $resourceState.Value
    $property = $instance.PSObject.Properties[$target.Property]
    if ($null -eq $property) {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $false
            Value            = $null
            Error            = "WSMan property $($target.Property) was not returned for $Path."
        }
    }

    [PSCustomObject]@{
        CommandAvailable = $true
        Captured         = $true
        Value            = ConvertFrom-WsManPropertyValue -Value $property.Value -ValueKind $target.ValueKind
        Error            = $null
    }
}

function Enter-WinRmServiceWriteScope {
    $service = Get-CimInstance Win32_Service -Filter "Name='WinRM'" -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        return [PSCustomObject]@{
            ServiceFound     = $false
            OriginalStartMode = $null
            ChangedStartMode = $false
            StartedService   = $false
        }
    }

    $originalStartMode = [string]$service.StartMode
    $wasRunning = ([string]$service.State -eq 'Running')
    $changedStartMode = $false
    $startedService = $false

    try {
        if (-not $wasRunning) {
            if ($originalStartMode -eq 'Disabled') {
                Set-ServiceStartModeValue -Name 'WinRM' -StartModeValue 'demand'
                $changedStartMode = $true
            }

            Start-Service -Name 'WinRM' -ErrorAction Stop
            $startedService = $true
        }

        [PSCustomObject]@{
            ServiceFound      = $true
            OriginalStartMode = $originalStartMode
            ChangedStartMode  = $changedStartMode
            StartedService    = $startedService
        }
    } catch {
        $scopeError = $_
        $partialScope = [PSCustomObject]@{
            ServiceFound      = $true
            OriginalStartMode = $originalStartMode
            ChangedStartMode  = $changedStartMode
            StartedService    = $startedService
        }
        try {
            Exit-WinRmServiceWriteScope -Scope $partialScope
        } catch {
            Write-Warning ("WinRM write scope startup failed and cleanup also failed: {0}" -f $_.Exception.Message)
        }
        throw $scopeError
    }
}

function Exit-WinRmServiceWriteScope {
    param([AllowNull()] [object]$Scope)

    if ($null -eq $Scope -or -not $Scope.PSObject.Properties['ServiceFound'] -or -not [bool]$Scope.ServiceFound) {
        return
    }

    $cleanupErrors = [System.Collections.Generic.List[string]]::new()
    if ($Scope.PSObject.Properties['StartedService'] -and [bool]$Scope.StartedService) {
        try {
            Set-ServiceRunningState -Name 'WinRM' -Running $false
        } catch {
            $cleanupErrors.Add(("Could not stop the temporary WinRM service: {0}" -f $_.Exception.Message)) | Out-Null
        }
    }

    if ($Scope.PSObject.Properties['ChangedStartMode'] -and [bool]$Scope.ChangedStartMode) {
        $originalStartMode = if ($Scope.PSObject.Properties['OriginalStartMode']) { [string]$Scope.OriginalStartMode } else { $null }
        if (-not [string]::IsNullOrWhiteSpace($originalStartMode)) {
            try {
                $startModeValue = Convert-ServiceStartModeToScValue -StartMode $originalStartMode
                Set-ServiceStartModeValue -Name 'WinRM' -StartModeValue $startModeValue
            } catch {
                $cleanupErrors.Add(("Could not restore the WinRM startup mode: {0}" -f $_.Exception.Message)) | Out-Null
            }
        }
    }

    if ($cleanupErrors.Count -gt 0) {
        throw [System.InvalidOperationException]::new(($cleanupErrors -join ' '))
    }
}

function Invoke-WithTemporaryWinRmServiceForWrite {
    param(
        [Parameter(Mandatory)] [scriptblock]$ScriptBlock,
        [AllowNull()] [object]$CaptureSession
    )

    if ($null -eq $CaptureSession) {
        $scope = Enter-WinRmServiceWriteScope
        try {
            & $ScriptBlock
        } finally {
            Exit-WinRmServiceWriteScope -Scope $scope
        }
        return
    }

    $null = Get-CaptureSessionValue -Session $CaptureSession -Key 'winrm.write-scope' -Factory {
        $scope = Enter-WinRmServiceWriteScope
        try {
            Register-CaptureSessionResource -Session $CaptureSession -Kind 'WinRmWriteScope' -Value $scope
        } catch {
            Exit-WinRmServiceWriteScope -Scope $scope
            throw
        }
        $scope
    }
    & $ScriptBlock
}

function Get-WsManConfigValue {
    param([Parameter(Mandatory)] [string]$Path)

    (Get-WsManConfigValueState -Path $Path).Value
}

function Set-WsManConfigValues {
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Items,
        [AllowNull()] [object]$CaptureSession
    )

    if ($Items.Count -eq 0) {
        return
    }

    if (-not (Test-CommandAvailable -Name 'Set-WSManInstance')) {
        throw 'Set-WSManInstance was not found.'
    }

    $resourceUri = $null
    $valueSet = @{}
    foreach ($item in @($Items)) {
        if ($null -eq $item -or $null -eq $item.PSObject.Properties['Path']) {
            throw 'A WSMan mutation item is missing its path.'
        }

        $path = [string]$item.Path
        $target = Resolve-WsManConfigTarget -Path $path
        if ($null -eq $resourceUri) {
            $resourceUri = [string]$target.ResourceUri
        } elseif (-not [string]::Equals($resourceUri, [string]$target.ResourceUri, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "WSMan mutation batch mixes resource URIs '$resourceUri' and '$($target.ResourceUri)'."
        }

        if ($valueSet.ContainsKey([string]$target.Property)) {
            throw "WSMan mutation batch contains duplicate property '$($target.Property)'."
        }
        $value = if ($item.PSObject.Properties['Value']) { $item.Value } else { $null }
        $valueSet[[string]$target.Property] = ConvertTo-WsManSetValue -Value $value -ValueKind $target.ValueKind
    }

    Invoke-WithTemporaryWinRmServiceForWrite -CaptureSession $CaptureSession -ScriptBlock {
        Set-WSManInstance -ResourceURI $resourceUri -ValueSet $valueSet -ErrorAction Stop | Out-Null
    }
}

function Set-WsManConfigValue {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [AllowNull()] [object]$Value,
        [AllowNull()] [object]$CaptureSession
    )

    Set-WsManConfigValues -Items @(
        [PSCustomObject]@{ Path = $Path; Value = $Value }
    ) -CaptureSession $CaptureSession
}

function Get-WinRmListenerStates {
    param([AllowNull()] [object]$CaptureSession)

    $listenerState = Get-CaptureSessionValue -Session $CaptureSession -Key 'winrm.listeners' -Factory {
        if (-not (Test-CommandAvailable -Name 'Get-WSManInstance')) {
            return [PSCustomObject]@{
                CommandAvailable = $false
                Captured         = $false
                Value            = @()
                Error            = 'Get-WSManInstance was not found.'
            }
        }

        try {
            return [PSCustomObject]@{
                CommandAvailable = $true
                Captured         = $true
                Value            = @(Get-WSManInstance -ResourceURI 'winrm/config/listener' -Enumerate -ErrorAction Stop)
                Error            = $null
            }
        } catch {
            return [PSCustomObject]@{
                CommandAvailable = $true
                Captured         = $false
                Value            = @()
                Error            = $_.Exception.Message
            }
        }
    }

    if (-not $listenerState.Captured) {
        return [PSCustomObject]@{
            CommandAvailable = [bool]$listenerState.CommandAvailable
            Captured         = $false
            Listeners        = @()
            Error            = $listenerState.Error
        }
    }

    $listenerItems = @($listenerState.Value)
    $listeners = @(
        foreach ($listener in $listenerItems) {
            [PSCustomObject]@{
                Address               = if ($listener.PSObject.Properties['Address']) { [string]$listener.Address } else { $null }
                Transport             = if ($listener.PSObject.Properties['Transport']) { [string]$listener.Transport } else { $null }
                Port                  = if ($listener.PSObject.Properties['Port'] -and -not [string]::IsNullOrWhiteSpace([string]$listener.Port)) { [int]$listener.Port } else { $null }
                Hostname              = if ($listener.PSObject.Properties['Hostname']) { [string]$listener.Hostname } else { $null }
                Enabled               = if ($listener.PSObject.Properties['Enabled']) { ConvertTo-NullableBoolean -Value $listener.Enabled } else { $null }
                URLPrefix             = if ($listener.PSObject.Properties['URLPrefix']) { [string]$listener.URLPrefix } else { $null }
                CertificateThumbprint = if ($listener.PSObject.Properties['CertificateThumbprint']) { [string]$listener.CertificateThumbprint } else { $null }
            }
        }
    ) | Sort-Object -Property Address, Transport, Port, Hostname, URLPrefix, CertificateThumbprint

    [PSCustomObject]@{
        CommandAvailable = $true
        Captured         = $true
        Listeners        = @($listeners)
        Error            = $null
    }
}

function Remove-WinRmListener {
    param([Parameter(Mandatory)] [object]$Listener)

    if (-not (Test-CommandAvailable -Name 'Remove-WSManInstance')) {
        throw 'Remove-WSManInstance was not found.'
    }

    $selectorSet = @{
        Address   = [string]$Listener.Address
        Transport = [string]$Listener.Transport
    }

    Remove-WSManInstance -ResourceURI 'winrm/config/listener' -SelectorSet $selectorSet -ErrorAction Stop | Out-Null
}

function New-WinRmListener {
    param([Parameter(Mandatory)] [object]$Listener)

    if (-not (Test-CommandAvailable -Name 'New-WSManInstance')) {
        throw 'New-WSManInstance was not found.'
    }

    $selectorSet = @{
        Address   = [string]$Listener.Address
        Transport = [string]$Listener.Transport
    }

    $valueSet = @{}
    if ($null -ne $Listener.PSObject.Properties['Port'] -and $null -ne $Listener.Port) {
        $valueSet['Port'] = [string][int]$Listener.Port
    }

    if ($null -ne $Listener.PSObject.Properties['Hostname'] -and -not [string]::IsNullOrWhiteSpace([string]$Listener.Hostname)) {
        $valueSet['Hostname'] = [string]$Listener.Hostname
    }

    if ($null -ne $Listener.PSObject.Properties['Enabled'] -and $null -ne $Listener.Enabled) {
        $valueSet['Enabled'] = ConvertTo-WsManTextValue -Value ([bool]$Listener.Enabled)
    }

    if ($null -ne $Listener.PSObject.Properties['URLPrefix'] -and -not [string]::IsNullOrWhiteSpace([string]$Listener.URLPrefix)) {
        $valueSet['URLPrefix'] = [string]$Listener.URLPrefix
    }

    if (
        [string]$Listener.Transport -eq 'HTTPS' -and
        $null -ne $Listener.PSObject.Properties['CertificateThumbprint'] -and
        -not [string]::IsNullOrWhiteSpace([string]$Listener.CertificateThumbprint)
    ) {
        $valueSet['CertificateThumbprint'] = [string]$Listener.CertificateThumbprint
    }

    New-WSManInstance -ResourceURI 'winrm/config/listener' -SelectorSet $selectorSet -ValueSet $valueSet -ErrorAction Stop | Out-Null
}

function Set-WinRmListenersExact {
    param(
        [Parameter(Mandatory)] [object[]]$Listeners,
        [AllowNull()] [object]$CaptureSession
    )

    Invoke-WithTemporaryWinRmServiceForWrite -CaptureSession $CaptureSession -ScriptBlock {
        $liveState = Get-WinRmListenerStates
        if (-not $liveState.CommandAvailable) {
            throw 'Get-WSManInstance was not found.'
        }

        if (-not $liveState.Captured) {
            throw $liveState.Error
        }

        foreach ($listener in @($liveState.Listeners)) {
            Remove-WinRmListener -Listener $listener
        }

        foreach ($listener in @($Listeners)) {
            New-WinRmListener -Listener $listener
        }
    }
}

function Set-Permissive-WinRmListeners {
    param(
        [Parameter(Mandatory)] [object[]]$Listeners,
        [AllowNull()] [object]$CaptureSession
    )

    Set-WinRmListenersExact -Listeners @($Listeners) -CaptureSession $CaptureSession
}

function Restore-WinRmListeners {
    param(
        [Parameter(Mandatory)] [object[]]$Listeners,
        [AllowNull()] [object]$CaptureSession
    )

    Set-WinRmListenersExact -Listeners @($Listeners) -CaptureSession $CaptureSession
}

function Normalize-SmbConfigState {
    param([AllowNull()] [object]$Value)

    if ($Value -is [bool] -or $Value -is [string] -or $Value -is [int] -or $Value -is [long]) {
        $legacyValue = ConvertTo-NullableBoolean -Value $Value
        if ($null -ne $legacyValue) {
            return [PSCustomObject]@{
                CommandAvailable         = $true
                TimedOut                 = $false
                RequireSecuritySignature = $legacyValue
            }
        }
    }

    if ($null -eq $Value) {
        return [PSCustomObject]@{
            CommandAvailable         = $false
            TimedOut                 = $false
            RequireSecuritySignature = $null
        }
    }

    $commandAvailable = if ($Value.PSObject.Properties['CommandAvailable']) { [bool]$Value.CommandAvailable } else { $true }
    $timedOut = if ($Value.PSObject.Properties['TimedOut']) { [bool]$Value.TimedOut } else { $false }
    $requireSecuritySignature = if ($Value.PSObject.Properties['RequireSecuritySignature']) {
        ConvertTo-NullableBoolean -Value $Value.RequireSecuritySignature
    } else {
        ConvertTo-NullableBoolean -Value $Value
    }

    [PSCustomObject]@{
        CommandAvailable         = $commandAvailable
        TimedOut                 = $timedOut
        RequireSecuritySignature = $requireSecuritySignature
    }
}

function Get-SmbConfigurationStates {
    $result = Invoke-ChildPowerShell -TimeoutSeconds 12 -ScriptText @'
$ErrorActionPreference = 'Stop'
$states = [ordered]@{}
foreach ($role in @('Client', 'Server')) {
    $commandName = if ($role -eq 'Client') { 'Get-SmbClientConfiguration' } else { 'Get-SmbServerConfiguration' }
    $command = Get-Command -Name $commandName -ErrorAction SilentlyContinue -Verbose:$false | Select-Object -First 1
    if ($null -eq $command) {
        $states[$role] = [PSCustomObject]@{
            CommandAvailable         = $false
            Captured                 = $false
            Error                    = "$commandName was not found."
            RequireSecuritySignature = $null
        }
        continue
    }

    try {
        $config = & $commandName -ErrorAction Stop
        $states[$role] = [PSCustomObject]@{
            CommandAvailable         = $true
            Captured                 = $true
            Error                    = $null
            RequireSecuritySignature = [bool]$config.RequireSecuritySignature
        }
    } catch {
        $states[$role] = [PSCustomObject]@{
            CommandAvailable         = $true
            Captured                 = $false
            Error                    = $_.Exception.Message
            RequireSecuritySignature = $null
        }
    }
}

$json = [PSCustomObject]$states | ConvertTo-Json -Compress -Depth 5
[Console]::Out.WriteLine('WDS_SMB_JSON:' + $json)
'@

    $unavailableState = Normalize-SmbConfigState -Value $null
    if (-not $result.CommandAvailable) {
        return [PSCustomObject]@{ Client = $unavailableState; Server = $unavailableState }
    }
    if ($result.TimedOut) {
        Write-Warning 'SMB client/server snapshot timed out. Skipping both settings.'
        $timedOutState = Normalize-SmbConfigState -Value ([PSCustomObject]@{ CommandAvailable = $true; TimedOut = $true; RequireSecuritySignature = $null })
        return [PSCustomObject]@{ Client = $timedOutState; Server = $timedOutState }
    }
    if ($result.ExitCode -ne 0) {
        $message = if (-not [string]::IsNullOrWhiteSpace([string]$result.StdErr)) { $result.StdErr.Trim() } else { 'The child process failed without error output.' }
        Write-Warning ("SMB client/server snapshot failed: {0}" -f $message)
        $failedState = Normalize-SmbConfigState -Value ([PSCustomObject]@{ CommandAvailable = $true; TimedOut = $false; RequireSecuritySignature = $null })
        return [PSCustomObject]@{ Client = $failedState; Server = $failedState }
    }

    $payloadLine = @(([string]$result.StdOut -split "`r?`n") | Where-Object { $_.StartsWith('WDS_SMB_JSON:') } | Select-Object -Last 1)
    if ($payloadLine.Count -eq 0) {
        Write-Warning 'SMB client/server snapshot returned no framed data. Skipping both settings.'
        return [PSCustomObject]@{ Client = $unavailableState; Server = $unavailableState }
    }

    try {
        $parsed = $payloadLine[0].Substring('WDS_SMB_JSON:'.Length) | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-Warning 'SMB client/server snapshot returned unparsable data. Skipping both settings.'
        return [PSCustomObject]@{ Client = $unavailableState; Server = $unavailableState }
    }

    $normalizedStates = @{}
    foreach ($role in @('Client', 'Server')) {
        $roleState = if ($parsed.PSObject.Properties[$role]) { $parsed.$role } else { $null }
        if ($null -ne $roleState -and $roleState.PSObject.Properties['Error'] -and -not [string]::IsNullOrWhiteSpace([string]$roleState.Error)) {
            Write-Warning ("SMB {0} snapshot failed: {1}" -f $role.ToLowerInvariant(), [string]$roleState.Error)
        }
        $captured = $null -ne $roleState -and $roleState.PSObject.Properties['Captured'] -and [bool]$roleState.Captured
        $normalizedStates[$role] = Normalize-SmbConfigState -Value ([PSCustomObject]@{
            CommandAvailable         = $null -ne $roleState -and $roleState.PSObject.Properties['CommandAvailable'] -and [bool]$roleState.CommandAvailable
            TimedOut                 = $false
            RequireSecuritySignature = if ($captured -and $roleState.PSObject.Properties['RequireSecuritySignature']) { $roleState.RequireSecuritySignature } else { $null }
        })
    }

    [PSCustomObject]@{
        Client = $normalizedStates.Client
        Server = $normalizedStates.Server
    }
}

function Get-SmbConfigurationState {
    param([Parameter(Mandatory)] [ValidateSet('Client', 'Server')] [string]$Role)

    (Get-SmbConfigurationStates).$Role
}

function Get-SmbClientConfigurationState {
    Get-SmbConfigurationState -Role 'Client'
}

function Get-SmbServerConfigurationState {
    Get-SmbConfigurationState -Role 'Server'
}

function Get-NetBiosAdapterCaptureState {
    if (-not (Test-CommandAvailable -Name 'Get-CimInstance')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Captured         = $false
            Error            = 'Get-CimInstance was not found.'
            Adapters         = @()
        }
    }

    try {
        $adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = TRUE' -ErrorAction Stop
        $states = foreach ($adapter in @($adapters)) {
            [PSCustomObject]@{
                Index               = [int]$adapter.Index
                Description         = $adapter.Description
                TcpipNetbiosOptions = [int]$adapter.TcpipNetbiosOptions
            }
        }

        [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $true
            Error            = $null
            Adapters         = @($states)
        }
    } catch {
        [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $false
            Error            = $_.Exception.Message
            Adapters         = @()
        }
    }
}

function Get-NetBiosAdapterStates {
    @((Get-NetBiosAdapterCaptureState).Adapters)
}

function Set-NetBiosAdapterOption {
    param(
        [Parameter(Mandatory)] [int]$Index,
        [Parameter(Mandatory)] [uint32]$Option
    )

    Set-NetBiosAdapterOptions -Targets @(
        [PSCustomObject]@{ Index = $Index; Option = $Option }
    )
}

function Set-NetBiosAdapterOptions {
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Targets)

    if ($Targets.Count -eq 0) {
        return
    }

    $targetsByIndex = @{}
    foreach ($target in @($Targets)) {
        if (
            $null -eq $target -or
            $null -eq $target.PSObject.Properties['Index'] -or
            $null -eq $target.PSObject.Properties['Option']
        ) {
            throw 'A NetBIOS mutation target is missing its adapter index or option.'
        }

        $index = [int]$target.Index
        if ($targetsByIndex.ContainsKey($index)) {
            throw "NetBIOS mutation target contains duplicate adapter index $index."
        }
        $targetsByIndex[$index] = [uint32]$target.Option
    }

    $filterParts = @($targetsByIndex.Keys | Sort-Object | ForEach-Object { "Index = $_" })
    $adapters = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter ($filterParts -join ' OR ') -ErrorAction Stop)
    $adaptersByIndex = @{}
    foreach ($adapter in $adapters) {
        $adaptersByIndex[[int]$adapter.Index] = $adapter
    }

    foreach ($index in @($targetsByIndex.Keys | Sort-Object)) {
        if (-not $adaptersByIndex.ContainsKey($index)) {
            throw "Network adapter index $index was not found."
        }
    }

    foreach ($index in @($targetsByIndex.Keys | Sort-Object)) {
        Invoke-CimMethod -InputObject $adaptersByIndex[$index] -MethodName SetTcpipNetbios -Arguments @{ TcpipNetbiosOptions = $targetsByIndex[$index] } -ErrorAction Stop | Out-Null
    }
}

function Set-Permissive-NetBiosAdapters {
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Adapters)

    $targets = @(
        foreach ($adapter in @($Adapters)) {
            [PSCustomObject]@{ Index = [int]$adapter.Index; Option = [uint32]1 }
        }
    )
    Set-NetBiosAdapterOptions -Targets $targets
}

function Restore-NetBiosAdapters {
    param([Parameter(Mandatory)] [object[]]$Adapters)

    $targets = @(
        foreach ($adapter in @($Adapters)) {
            [PSCustomObject]@{ Index = [int]$adapter.Index; Option = [uint32]$adapter.TcpipNetbiosOptions }
        }
    )
    Set-NetBiosAdapterOptions -Targets $targets
}

#endregion

#region Defender provider

function Get-AsrRuleCatalog {
    @{
        '01443614-cd74-433a-b99e-2ecdc07bfc25' = 'Block executable files by prevalence, age, or trusted list'
        '26190899-1602-49e8-8b27-eb1d0a1ce869' = 'Block Office communication apps from creating child processes'
        '3b576869-a4ec-4529-8536-b80a7769e899' = 'Block Office applications from creating executable content'
        '56a863a9-875e-4185-98a7-b882c64b5ce5' = 'Block abused vulnerable signed drivers'
        '5beb7efe-fd9a-4556-801d-275e5ffc04cc' = 'Block execution of potentially obfuscated scripts'
        '75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84' = 'Block Office apps from injecting code into other processes'
        '7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c' = 'Block Adobe Reader from creating child processes'
        '92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b' = 'Block Win32 API calls from Office macros'
        '9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2' = 'Block credential stealing from LSASS'
        'a8f5898e-1dc8-49a9-9878-85004b8a61e6' = 'Block webshell creation for servers'
        'b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4' = 'Block untrusted and unsigned processes from USB'
        'be9ba2d9-53ea-4cdc-84e5-9b1eeee46550' = 'Block executable content from email and webmail'
        'c0033c00-d16d-4114-a5a0-dc9b3a7d2ceb' = 'Block copied or impersonated system tools'
        'c1db55ab-c21a-4637-bb3f-a12568109d35' = 'Use advanced protection against ransomware'
        'd1e49aac-8f56-4280-b9ba-993a6d77406c' = 'Block process creation from PSExec and WMI'
        'd3e037e1-3eb8-44c8-a917-57927947596d' = 'Block JS or VBS from launching downloaded executables'
        'd4f940ab-401b-4efc-aadc-ad5f3c50688a' = 'Block all Office apps from creating child processes'
        'e6db77e5-3df2-4cf1-b95a-636979351e5b' = 'Block persistence through WMI event subscription'
        '33ddedf1-c6e0-47cb-833e-de6133960387' = 'Block rebooting in Safe Mode'
    }
}

function Get-AsrActionLabel {
    param([Parameter(Mandatory)] [AllowNull()] [object]$Action)

    switch ([string]$Action) {
        '0' { 'Disabled' }
        '1' { 'Block' }
        '2' { 'Audit' }
        '5' { 'Not configured' }
        '6' { 'Warn' }
        'Disabled' { 'Disabled' }
        'Enabled' { 'Block' }
        'Block' { 'Block' }
        'AuditMode' { 'Audit' }
        'Audit' { 'Audit' }
        'NotConfigured' { 'Not configured' }
        'Warn' { 'Warn' }
        default { "Unknown ($Action)" }
    }
}

function Get-AsrRestoreAction {
    param([Parameter(Mandatory)] [AllowNull()] [object]$Action)

    switch ([string]$Action) {
        '0' { 'Disabled' }
        '1' { 'Enabled' }
        '2' { 'AuditMode' }
        '5' { 'NotConfigured' }
        '6' { 'Warn' }
        'Disabled' { 'Disabled' }
        'Enabled' { 'Enabled' }
        'Block' { 'Enabled' }
        'AuditMode' { 'AuditMode' }
        'Audit' { 'AuditMode' }
        'NotConfigured' { 'NotConfigured' }
        'Warn' { 'Warn' }
        default { $null }
    }
}

function Test-AsrRuleId {
    param([AllowNull()] [object]$Id)

    $text = [string]$Id
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $false
    }

    $guid = [guid]::Empty
    [guid]::TryParse($text, [ref]$guid)
}

function Get-AsrRuleCaptureState {
    param([AllowNull()] [object]$CaptureSession)

    $catalog = Get-AsrRuleCatalog
    $preferenceState = Get-MpPreferenceCaptureState -CaptureSession $CaptureSession
    $mp = if ($preferenceState.Captured) { $preferenceState.Value } else { $null }
    if ($null -eq $mp) {
        return [PSCustomObject]@{
            CommandAvailable = $preferenceState.CommandAvailable
            Captured         = $false
            Error            = $preferenceState.Error
            Rules          = @()
            InvalidEntries = @()
        }
    }

    $ids = @($mp.AttackSurfaceReductionRules_Ids)
    $actions = @($mp.AttackSurfaceReductionRules_Actions)
    $validRules = @()
    $invalidEntries = @()
    $entryCount = [Math]::Max($ids.Count, $actions.Count)

    for ($i = 0; $i -lt $entryCount; $i++) {
        $id = if ($ids.Count -gt $i) { [string]$ids[$i] } else { $null }
        $action = if ($actions.Count -gt $i) { [string]$actions[$i] } else { $null }
        $actionLabel = Get-AsrActionLabel -Action $action

        $restoreAction = Get-AsrRestoreAction -Action $action
        if (-not (Test-AsrRuleId -Id $id) -or $null -eq $restoreAction) {
            $invalidEntries += [PSCustomObject]@{
                Id          = $id
                Action      = $action
                ActionLabel = $actionLabel
            }
            continue
        }

        $validRules += [PSCustomObject]@{
            Id          = $id
            Name        = if ($catalog.ContainsKey($id)) { $catalog[$id] } else { 'Unknown / custom rule' }
            Action      = $action
            ActionLabel = $actionLabel
        }
    }

    [PSCustomObject]@{
        CommandAvailable = $true
        Captured         = $true
        Error            = $null
        Rules          = @($validRules)
        InvalidEntries = @($invalidEntries)
    }
}

function Get-AsrInvalidEntriesFromEntry {
    param([AllowNull()] [object]$Entry)

    if ($null -eq $Entry) {
        return @()
    }

    $invalidEntries = @()

    if ($Entry.PSObject.Properties['InvalidEntries']) {
        foreach ($rule in @($Entry.InvalidEntries)) {
            if ($null -eq $rule) {
                continue
            }

            $action = if ($rule.PSObject.Properties['Action']) { [string]$rule.Action } else { $null }
            $actionLabel = if ($rule.PSObject.Properties['ActionLabel']) { [string]$rule.ActionLabel } else { Get-AsrActionLabel -Action $action }

            $invalidEntries += [PSCustomObject]@{
                Id          = if ($rule.PSObject.Properties['Id']) { [string]$rule.Id } else { $null }
                Action      = $action
                ActionLabel = $actionLabel
            }
        }
    }

    foreach ($rule in @($Entry.CurrentValue)) {
        if ($null -eq $rule) {
            continue
        }

        $id = if ($rule.PSObject.Properties['Id']) { [string]$rule.Id } else { $null }
        $action = if ($rule.PSObject.Properties['Action']) { [string]$rule.Action } else { $null }
        if ((Test-AsrRuleId -Id $id) -and $null -ne (Get-AsrRestoreAction -Action $action)) {
            continue
        }

        $actionLabel = if ($rule.PSObject.Properties['ActionLabel']) { [string]$rule.ActionLabel } else { Get-AsrActionLabel -Action $action }
        $invalidEntries += [PSCustomObject]@{
            Id          = $id
            Action      = $action
            ActionLabel = $actionLabel
        }
    }

    @($invalidEntries)
}

function Get-ConfiguredAsrRules {
    @((Get-AsrRuleCaptureState).Rules)
}

function Disable-ConfiguredAsrRules {
    param([AllowNull()] [object]$CaptureSession)

    $preferenceState = Get-MpPreferenceCaptureState -CaptureSession $CaptureSession
    if (-not $preferenceState.CommandAvailable) {
        throw 'Get-MpPreference was not found.'
    }
    if (-not $preferenceState.Captured -or $null -eq $preferenceState.Value) {
        $detail = if ([string]::IsNullOrWhiteSpace([string]$preferenceState.Error)) { 'Get-MpPreference returned no Defender preference state.' } else { [string]$preferenceState.Error }
        throw $detail
    }

    $mp = $preferenceState.Value
    $ids = @($mp.AttackSurfaceReductionRules_Ids)
    $currentActions = @($mp.AttackSurfaceReductionRules_Actions)
    if ($ids.Count -ne $currentActions.Count) {
        throw "Cannot safely clear ASR rules because Defender returned $($ids.Count) ID(s) and $($currentActions.Count) action(s)."
    }

    $validIds = [System.Collections.Generic.List[string]]::new()
    $validActions = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -lt $ids.Count; $i++) {
        $id = [string]$ids[$i]
        $action = Get-AsrRestoreAction -Action $currentActions[$i]
        if (-not (Test-AsrRuleId -Id $id) -or $null -eq $action) {
            throw "Cannot safely clear malformed ASR rule entry at index $i."
        }

        $validIds.Add($id) | Out-Null
        $validActions.Add($action) | Out-Null
    }

    if ($validIds.Count -gt 0) {
        Remove-MpPreference `
            -AttackSurfaceReductionRules_Ids @($validIds) `
            -AttackSurfaceReductionRules_Actions @($validActions) `
            -ErrorAction Stop
    }
}

function Restore-AsrRules {
    param(
        [Parameter(Mandatory)] [object[]]$Rules,
        [AllowNull()] [object]$CaptureSession
    )

    Disable-ConfiguredAsrRules -CaptureSession $CaptureSession

    $ids = @()
    $actions = @()

    foreach ($rule in @($Rules)) {
        if (-not (Test-AsrRuleId -Id $rule.Id)) {
            continue
        }

        $restoreAction = Get-AsrRestoreAction -Action $rule.Action
        if ($null -eq $restoreAction) {
            continue
        }

        $ids += [string]$rule.Id
        $actions += $restoreAction
    }

    if ($ids.Count -gt 0) {
        Add-MpPreference -AttackSurfaceReductionRules_Ids $ids -AttackSurfaceReductionRules_Actions $actions
    }
}

function Set-MpPreferencePropertyValues {
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Items)

    if ($Items.Count -eq 0) {
        return
    }
    $command = Get-WinDefStateCommand -Name 'Set-MpPreference'
    if ($null -eq $command) {
        return
    }

    $params = @{}
    foreach ($item in @($Items)) {
        if ($null -eq $item -or $null -eq $item.PSObject.Properties['Property']) {
            throw 'A Defender preference mutation item is missing its property name.'
        }

        $property = [string]$item.Property
        if ($params.ContainsKey($property)) {
            throw "Defender preference mutation batch contains duplicate property '$property'."
        }

        $value = if ($item.PSObject.Properties['Value']) { $item.Value } else { $null }
        if ($null -eq $value -or -not $command.Parameters.ContainsKey($property)) {
            continue
        }
        $params[$property] = $value
    }

    if ($params.Count -eq 0) {
        return
    }

    Set-MpPreference @params
}

function Set-MpPreferencePropertyValue {
    param(
        [Parameter(Mandatory)] [string]$Property,
        [Parameter(Mandatory)] [AllowNull()] [object]$Value
    )

    Set-MpPreferencePropertyValues -Items @(
        [PSCustomObject]@{ Property = $Property; Value = $Value }
    )
}

function Get-MpPreferenceCaptureState {
    param([AllowNull()] [object]$CaptureSession)

    Get-CaptureSessionValue -Session $CaptureSession -Key 'defender.preferences' -Factory {
        if (-not (Test-CommandAvailable -Name 'Get-MpPreference')) {
            return [PSCustomObject]@{
                CommandAvailable = $false
                Captured         = $false
                Value            = $null
                Error            = 'Get-MpPreference was not found.'
            }
        }

        try {
            $preference = Get-MpPreference -ErrorAction Stop
            return [PSCustomObject]@{
                CommandAvailable = $true
                Captured         = $true
                Value            = $preference
                Error            = $null
            }
        } catch {
            return [PSCustomObject]@{
                CommandAvailable = $true
                Captured         = $false
                Value            = $null
                Error            = $_.Exception.Message
            }
        }
    }
}

function Get-MpPreferencePropertyState {
    param(
        [Parameter(Mandatory)] [string]$Property,
        [AllowNull()] [object]$CaptureSession
    )

    $state = Get-MpPreferenceCaptureState -CaptureSession $CaptureSession
    if (-not $state.Captured -or $null -eq $state.Value) {
        return [PSCustomObject]@{
            CommandAvailable = $state.CommandAvailable
            Captured         = $false
            Value            = $null
            Error            = $state.Error
        }
    }

    $mp = $state.Value
    $propertyInfo = $mp.PSObject.Properties[$Property]
    if ($null -eq $propertyInfo) {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $false
            Value            = $null
            Error            = "Get-MpPreference did not return property '$Property'."
        }
    }

    [PSCustomObject]@{
        CommandAvailable = $true
        Captured         = $true
        Value            = $propertyInfo.Value
        Error            = $null
    }
}

function Get-MpPreferencePropertyRawValue {
    param(
        [Parameter(Mandatory)] [string]$Property,
        [AllowNull()] [object]$CaptureSession
    )

    (Get-MpPreferencePropertyState -Property $Property -CaptureSession $CaptureSession).Value
}

function Resolve-MpPreferenceValue {
    param(
        [Parameter(Mandatory)] [object]$Definition,
        [Parameter(Mandatory)] [AllowNull()] [object]$Value
    )

    $valueMapProperty = $Definition.PSObject.Properties['ValueMap']
    if ($null -eq $valueMapProperty -or $null -eq $valueMapProperty.Value) {
        return $Value
    }

    $valueMap = $valueMapProperty.Value
    $key = [string]$Value
    if ($valueMap.ContainsKey($key)) {
        return $valueMap[$key]
    }

    return $Value
}

function Normalize-MpPreferenceListItems {
    param([AllowNull()] [object]$Value)

    @(
        foreach ($item in @($Value)) {
            $text = [string]$item
            if ([string]::IsNullOrWhiteSpace($text)) {
                continue
            }

            $text.Trim()
        }
    ) | Sort-Object -Unique
}

function Get-MpPreferenceListState {
    param(
        [Parameter(Mandatory)] [string]$Property,
        [AllowNull()] [object]$CaptureSession
    )

    $state = Get-MpPreferenceCaptureState -CaptureSession $CaptureSession
    if (-not $state.Captured -or $null -eq $state.Value) {
        return [PSCustomObject]@{
            CommandAvailable = [bool]$state.CommandAvailable
            Captured         = $false
            Items            = @()
            Error            = $state.Error
        }
    }

    $mp = $state.Value
    $propertyInfo = $mp.PSObject.Properties[$Property]
    if ($null -eq $propertyInfo) {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $false
            Items            = @()
            Error            = "Get-MpPreference does not expose property '$Property'."
        }
    }

    [PSCustomObject]@{
        CommandAvailable = $true
        Captured         = $true
        Items            = @(Normalize-MpPreferenceListItems -Value $propertyInfo.Value)
        Error            = $null
    }
}

function Test-MpPreferenceListCapturedExactly {
    param([Parameter(Mandatory)] [object]$Entry)

    $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
    $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
    ($commandAvailable -and $captured)
}

function Set-MpPreferenceListValue {
    param(
        [Parameter(Mandatory)] [string]$Property,
        [AllowNull()] [object[]]$DesiredItems,
        [AllowNull()] [object]$CaptureSession
    )

    $addCommand = Get-WinDefStateCommand -Name 'Add-MpPreference'
    $removeCommand = Get-WinDefStateCommand -Name 'Remove-MpPreference'
    if (
        $null -eq $addCommand -or
        $null -eq $removeCommand -or
        -not $addCommand.Parameters.ContainsKey($Property) -or
        -not $removeCommand.Parameters.ContainsKey($Property)
    ) {
        return
    }

    $currentState = Get-MpPreferenceListState -Property $Property -CaptureSession $CaptureSession
    if (-not ($currentState.CommandAvailable -and $currentState.Captured)) {
        return
    }

    $currentItems = @(Normalize-MpPreferenceListItems -Value $currentState.Items)
    $targetItems = @(Normalize-MpPreferenceListItems -Value $DesiredItems)
    $itemsToRemove = @($currentItems | Where-Object { $_ -notin $targetItems })
    $itemsToAdd = @($targetItems | Where-Object { $_ -notin $currentItems })

    if ($itemsToRemove.Count -gt 0) {
        $removeParams = @{}
        $removeParams[$Property] = @($itemsToRemove)
        Remove-MpPreference @removeParams
    }

    if ($itemsToAdd.Count -gt 0) {
        $addParams = @{}
        $addParams[$Property] = @($itemsToAdd)
        Add-MpPreference @addParams
    }
}

function Get-DefenderRuntimeStatus {
    if (-not (Test-CommandAvailable -Name 'Get-MpComputerStatus')) {
        return [PSCustomObject]@{
            CommandAvailable          = $false
            Captured                  = $false
            AMRunningMode             = $null
            RealTimeProtectionEnabled = $null
            AntivirusEnabled          = $null
            IsTamperProtected         = $null
            Error                     = $null
        }
    }

    try {
        $status = Get-MpComputerStatus -ErrorAction Stop
    } catch {
        return [PSCustomObject]@{
            CommandAvailable          = $true
            Captured                  = $false
            AMRunningMode             = $null
            RealTimeProtectionEnabled = $null
            AntivirusEnabled          = $null
            IsTamperProtected         = $null
            Error                     = $_.Exception.Message
        }
    }

    [PSCustomObject]@{
        CommandAvailable          = $true
        Captured                  = $true
        AMRunningMode             = if ($status.PSObject.Properties['AMRunningMode']) { [string]$status.AMRunningMode } else { $null }
        RealTimeProtectionEnabled = if ($status.PSObject.Properties['RealTimeProtectionEnabled']) { ConvertTo-NullableBoolean -Value $status.RealTimeProtectionEnabled } else { $null }
        AntivirusEnabled          = if ($status.PSObject.Properties['AntivirusEnabled']) { ConvertTo-NullableBoolean -Value $status.AntivirusEnabled } else { $null }
        IsTamperProtected         = if ($status.PSObject.Properties['IsTamperProtected']) { ConvertTo-NullableBoolean -Value $status.IsTamperProtected } else { $null }
        Error                     = $null
    }
}

function Test-DefenderRuntimeStatusCapturedExactly {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $false
    }

    $commandAvailable = if ($State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
    $captured = if ($State.PSObject.Properties['Captured']) { [bool]$State.Captured } else { $true }
    ($commandAvailable -and $captured)
}

#endregion

#region BitLocker provider

function ConvertTo-BitLockerProtectionStatusLabel {
    param([AllowNull()] [object]$Value)

    switch ([string]$Value) {
        '0' { 'Off' }
        '1' { 'On' }
        '2' { 'Unknown' }
        'Off' { 'Off' }
        'On' { 'On' }
        'Unknown' { 'Unknown' }
        default { [string]$Value }
    }
}

function Test-BitLockerProtectionEnabled {
    param([AllowNull()] [object]$Value)

    (ConvertTo-BitLockerProtectionStatusLabel -Value $Value) -eq 'On'
}

function Get-BitLockerTimedOutMountPoints {
    param([AllowNull()] [object]$State)

    if ($null -eq $State -or -not $State.PSObject.Properties['TimedOutMountPoints']) {
        return @()
    }

    @(
        foreach ($mountPoint in @($State.TimedOutMountPoints)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$mountPoint)) {
                [string]$mountPoint
            }
        }
    )
}

function Get-BitLockerCommandAvailableFlag {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $false
    }

    if ($State.PSObject.Properties['CommandAvailable']) {
        return [bool]$State.CommandAvailable
    }

    $true
}

function New-BitLockerCaptureIssue {
    param(
        [string]$MountPoint,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        MountPoint = $MountPoint
        Message    = $Message
    }
}

function Get-BitLockerCaptureIssues {
    param([AllowNull()] [object]$State)

    if ($null -eq $State -or -not $State.PSObject.Properties['CaptureIssues']) {
        return @()
    }

    @(
        foreach ($issue in @($State.CaptureIssues)) {
            if ($null -eq $issue) {
                continue
            }

            [PSCustomObject]@{
                MountPoint = if ($issue.PSObject.Properties['MountPoint']) { [string]$issue.MountPoint } else { $null }
                Message    = if ($issue.PSObject.Properties['Message']) { [string]$issue.Message } else { $null }
            }
        }
    )
}

function Get-BitLockerProtectionModeLabel {
    param(
        [AllowNull()] [object]$ProtectionStatus,
        [AllowNull()] [object]$VolumeStatus,
        [AllowNull()] [object]$EncryptionPercentage,
        [AllowNull()] [object]$KeyProtectorCount
    )

    $protectionLabel = ConvertTo-BitLockerProtectionStatusLabel -Value $ProtectionStatus
    if ($protectionLabel -eq 'On') {
        return 'Protected'
    }

    if ([string]$VolumeStatus -eq 'FullyDecrypted' -or ($null -ne $EncryptionPercentage -and [int]$EncryptionPercentage -eq 0)) {
        return 'Decrypted'
    }

    if ($protectionLabel -eq 'Off') {
        if ($null -ne $KeyProtectorCount -and [int]$KeyProtectorCount -gt 0) {
            return 'Suspended'
        }

        if ([string]$VolumeStatus -match 'Encrypt') {
            return 'Suspended'
        }

        return 'ProtectionOff'
    }

    if ([string]::IsNullOrWhiteSpace([string]$protectionLabel)) {
        return 'Unknown'
    }

    return [string]$protectionLabel
}

function Test-BitLockerAutoUnlockSupportedVolume {
    param([AllowNull()] [object]$Volume)

    if ($null -eq $Volume) {
        return $false
    }

    $mountPoint = if ($Volume.PSObject.Properties['MountPoint']) { [string]$Volume.MountPoint } else { $null }
    if ([string]::IsNullOrWhiteSpace($mountPoint)) {
        return $false
    }

    $volumeType = if ($Volume.PSObject.Properties['VolumeType']) { [string]$Volume.VolumeType } else { $null }
    if ([string]::Equals($volumeType, 'OperatingSystem', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    $true
}

function Normalize-BitLockerVolumeState {
    param([AllowNull()] [object]$Volume)

    if ($null -eq $Volume) {
        return $null
    }

    $keyProtectors = @()
    if ($Volume.PSObject.Properties['KeyProtectors']) {
        $keyProtectors = @(
            foreach ($protector in @($Volume.KeyProtectors)) {
                if ($null -eq $protector) {
                    continue
                }

                [PSCustomObject]@{
                    KeyProtectorId   = if ($protector.PSObject.Properties['KeyProtectorId']) { [string]$protector.KeyProtectorId } else { $null }
                    KeyProtectorType = if ($protector.PSObject.Properties['KeyProtectorType']) { [string]$protector.KeyProtectorType } else { $null }
                }
            }
        )
    }

    $keyProtectorCount = if ($Volume.PSObject.Properties['KeyProtectorCount'] -and $null -ne $Volume.KeyProtectorCount) {
        [int]$Volume.KeyProtectorCount
    } elseif ($Volume.PSObject.Properties['KeyProtectors']) {
        @($keyProtectors).Count
    } else {
        $null
    }

    $encryptionPercentage = if ($Volume.PSObject.Properties['EncryptionPercentage'] -and $null -ne $Volume.EncryptionPercentage) {
        [int]$Volume.EncryptionPercentage
    } else {
        $null
    }

    $protectionStatus = ConvertTo-BitLockerProtectionStatusLabel -Value $(if ($Volume.PSObject.Properties['ProtectionStatus']) { $Volume.ProtectionStatus } else { $null })
    $volumeStatus = if ($Volume.PSObject.Properties['VolumeStatus']) { [string]$Volume.VolumeStatus } else { $null }

    [PSCustomObject]@{
        MountPoint           = if ($Volume.PSObject.Properties['MountPoint']) { [string]$Volume.MountPoint } else { $null }
        VolumeType           = if ($Volume.PSObject.Properties['VolumeType']) { [string]$Volume.VolumeType } else { $null }
        ProtectionStatus     = $protectionStatus
        ProtectionMode       = if ($Volume.PSObject.Properties['ProtectionMode'] -and -not [string]::IsNullOrWhiteSpace([string]$Volume.ProtectionMode)) {
            [string]$Volume.ProtectionMode
        } else {
            Get-BitLockerProtectionModeLabel -ProtectionStatus $protectionStatus -VolumeStatus $volumeStatus -EncryptionPercentage $encryptionPercentage -KeyProtectorCount $keyProtectorCount
        }
        VolumeStatus         = $volumeStatus
        LockStatus           = if ($Volume.PSObject.Properties['LockStatus']) { [string]$Volume.LockStatus } else { $null }
        EncryptionMethod     = if ($Volume.PSObject.Properties['EncryptionMethod']) { [string]$Volume.EncryptionMethod } else { $null }
        EncryptionPercentage = $encryptionPercentage
        KeyProtectorCount    = $keyProtectorCount
        KeyProtectors        = @($keyProtectors)
        AutoUnlockEnabled    = if ($Volume.PSObject.Properties['AutoUnlockEnabled']) { ConvertTo-NullableBoolean -Value $Volume.AutoUnlockEnabled } else { $null }
    }
}

function Normalize-BitLockerState {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return [PSCustomObject]@{
            CommandAvailable    = $false
            TimedOutMountPoints = @()
            CaptureIssues       = @()
            Volumes             = @()
        }
    }

    [PSCustomObject]@{
        CommandAvailable    = Get-BitLockerCommandAvailableFlag -State $State
        TimedOutMountPoints = @(Get-BitLockerTimedOutMountPoints -State $State)
        CaptureIssues       = @(Get-BitLockerCaptureIssues -State $State)
        Volumes             = @(
            foreach ($volume in @($State.Volumes)) {
                Normalize-BitLockerVolumeState -Volume $volume
            }
        )
    }
}

function Test-BitLockerStateHasExtendedFields {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $false
    }

    if ($State.PSObject.Properties['CaptureIssues']) {
        return $true
    }

    foreach ($volume in @($State.Volumes)) {
        if ($null -eq $volume) {
            continue
        }

        foreach ($propertyName in @('LockStatus', 'ProtectionMode', 'KeyProtectors')) {
            if ($volume.PSObject.Properties[$propertyName]) {
                return $true
            }
        }
    }

    $false
}

function Get-BitLockerProtectorTypeSummary {
    param([AllowNull()] [object]$Volume)

    if ($null -eq $Volume) {
        return '<unknown>'
    }

    $types = @(
        foreach ($protector in @($Volume.KeyProtectors)) {
            if ($null -eq $protector) {
                continue
            }

            $type = if ($protector.PSObject.Properties['KeyProtectorType']) { [string]$protector.KeyProtectorType } else { $null }
            if ([string]::IsNullOrWhiteSpace($type)) {
                continue
            }

            $type
        }
    ) | Sort-Object -Unique

    if (@($types).Count -gt 0) {
        return ($types -join ', ')
    }

    if ($Volume.PSObject.Properties['KeyProtectorCount'] -and $null -ne $Volume.KeyProtectorCount) {
        $count = [int]$Volume.KeyProtectorCount
        if ($count -eq 0) {
            return '<none>'
        }

        return ('<count={0}>' -f $count)
    }

    '<unknown>'
}

function ConvertTo-ComparableBitLockerState {
    param(
        [AllowNull()] [object]$State,
        [AllowNull()] [object]$ReferenceState
    )

    $normalizedState = Normalize-BitLockerState -State $State
    $comparisonReference = if ($null -ne $ReferenceState) { $ReferenceState } else { $State }
    $useExtendedFields = Test-BitLockerStateHasExtendedFields -State $comparisonReference
    $comparisonCommandAvailable = $normalizedState.CommandAvailable
    if (-not $useExtendedFields -and @($normalizedState.CaptureIssues).Count -gt 0) {
        $comparisonCommandAvailable = $false
    }

    $comparableState = [ordered]@{
        CommandAvailable    = $comparisonCommandAvailable
        TimedOutMountPoints = @(
            foreach ($mountPoint in @($normalizedState.TimedOutMountPoints | Sort-Object -Unique)) {
                [string]$mountPoint
            }
        )
    }

    if ($useExtendedFields) {
        $comparableState['CaptureIssues'] = @(
            foreach ($issue in @($normalizedState.CaptureIssues | Sort-Object -Property MountPoint, Message)) {
                [PSCustomObject]@{
                    MountPoint = if (-not [string]::IsNullOrWhiteSpace([string]$issue.MountPoint)) { [string]$issue.MountPoint } else { $null }
                    Message    = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Message)) { [string]$issue.Message } else { $null }
                }
            }
        )
    }

    $comparableState['Volumes'] = @(
        foreach ($volume in @($normalizedState.Volumes | Sort-Object -Property MountPoint)) {
            if ($useExtendedFields) {
                [PSCustomObject]@{
                    MountPoint           = [string]$volume.MountPoint
                    VolumeType           = if (-not [string]::IsNullOrWhiteSpace([string]$volume.VolumeType)) { [string]$volume.VolumeType } else { $null }
                    ProtectionStatus     = [string]$volume.ProtectionStatus
                    ProtectionMode       = if (-not [string]::IsNullOrWhiteSpace([string]$volume.ProtectionMode)) { [string]$volume.ProtectionMode } else { $null }
                    VolumeStatus         = if (-not [string]::IsNullOrWhiteSpace([string]$volume.VolumeStatus)) { [string]$volume.VolumeStatus } else { $null }
                    LockStatus           = if (-not [string]::IsNullOrWhiteSpace([string]$volume.LockStatus)) { [string]$volume.LockStatus } else { $null }
                    EncryptionMethod     = if (-not [string]::IsNullOrWhiteSpace([string]$volume.EncryptionMethod)) { [string]$volume.EncryptionMethod } else { $null }
                    EncryptionPercentage = $volume.EncryptionPercentage
                    KeyProtectorCount    = $volume.KeyProtectorCount
                    KeyProtectors        = @(
                        foreach ($protector in @($volume.KeyProtectors | Sort-Object -Property KeyProtectorId, KeyProtectorType)) {
                            [PSCustomObject]@{
                                KeyProtectorId   = if (-not [string]::IsNullOrWhiteSpace([string]$protector.KeyProtectorId)) { [string]$protector.KeyProtectorId } else { $null }
                                KeyProtectorType = if (-not [string]::IsNullOrWhiteSpace([string]$protector.KeyProtectorType)) { [string]$protector.KeyProtectorType } else { $null }
                            }
                        }
                    )
                    AutoUnlockEnabled    = $volume.AutoUnlockEnabled
                }
                continue
            }

            [PSCustomObject]@{
                MountPoint           = [string]$volume.MountPoint
                ProtectionStatus     = [string]$volume.ProtectionStatus
                VolumeStatus         = [string]$volume.VolumeStatus
                EncryptionPercentage = $volume.EncryptionPercentage
                KeyProtectorCount    = $volume.KeyProtectorCount
            }
        }
    )

    $comparableState
}

function Get-ManageBdeCommand {
    foreach ($name in @('manage-bde.exe', 'manage-bde')) {
        $command = Get-WinDefStateCommand -Name $name
        if ($null -ne $command) {
            return $command
        }
    }

    $null
}

function Set-BitLockerAutoUnlockState {
    param(
        [Parameter(Mandatory)] [string]$MountPoint,
        [Parameter(Mandatory)] [bool]$Enabled
    )

    $manageBde = Get-ManageBdeCommand
    if ($null -eq $manageBde) {
        throw 'manage-bde was not found.'
    }

    $action = if ($Enabled) { '-enable' } else { '-disable' }
    $output = (& $manageBde.Source -autounlock $action $MountPoint 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) {
        $message = "manage-bde -autounlock $action $MountPoint failed with exit code $LASTEXITCODE"
        if (-not [string]::IsNullOrWhiteSpace($output)) {
            $message = "$message. Output: $output"
        }

        throw $message
    }
}

function Test-BitLockerStateCapturedExactly {
    param([AllowNull()] [object]$State)

    $normalized = Normalize-BitLockerState -State $State
    if ($null -eq $normalized) {
        return $false
    }

    if (-not $normalized.CommandAvailable -or @($normalized.TimedOutMountPoints).Count -ne 0 -or @($normalized.CaptureIssues).Count -ne 0) {
        return $false
    }

    $seenMountPoints = @{}
    foreach ($volume in @($normalized.Volumes)) {
        if ($null -eq $volume -or [string]::IsNullOrWhiteSpace([string]$volume.MountPoint)) {
            return $false
        }

        $mountPoint = [string]$volume.MountPoint
        if ($seenMountPoints.ContainsKey($mountPoint)) {
            return $false
        }
        $seenMountPoints[$mountPoint] = $true
    }

    return $true
}

function Get-BitLockerVolumeStates {
    if (-not (Test-CommandAvailable -Name 'Get-BitLockerVolume')) {
        return [PSCustomObject]@{
            CommandAvailable   = $false
            TimedOutMountPoints = @()
            CaptureIssues      = @()
            Volumes            = @()
        }
    }

    $mountPoints = @(
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 2 OR DriveType = 3' -ErrorAction SilentlyContinue |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.DeviceID) } |
            ForEach-Object { [string]$_.DeviceID } |
            Sort-Object -Unique
    )

    $probeRequests = @(
        foreach ($mountPoint in $mountPoints) {
            $escapedMountPoint = $mountPoint.Replace("'", "''")
            [PSCustomObject]@{
                Key            = $mountPoint
                TimeoutSeconds = 12
                ScriptText     = @"
`$ErrorActionPreference = 'Stop'
`$volume = Get-BitLockerVolume -MountPoint '$escapedMountPoint' -ErrorAction SilentlyContinue
if (`$null -eq `$volume) {
    return
}

[bool]`$keyProtectorsCaptured = `$false
`$keyProtectors = @()
if (`$volume.PSObject.Properties['KeyProtector']) {
    `$keyProtectorsCaptured = `$true
    `$keyProtectors = @(
        foreach (`$protector in @(`$volume.KeyProtector)) {
            if (`$null -eq `$protector) {
                continue
            }

            [PSCustomObject]@{
                KeyProtectorId   = if (`$protector.PSObject.Properties['KeyProtectorId']) { [string]`$protector.KeyProtectorId } else { `$null }
                KeyProtectorType = if (`$protector.PSObject.Properties['KeyProtectorType']) { [string]`$protector.KeyProtectorType } else { `$null }
            }
        }
    )
}

[bool]`$autoUnlockCaptured = `$false
`$autoUnlockEnabled = `$null
if (`$volume.PSObject.Properties['AutoUnlockEnabled']) {
    `$autoUnlockCaptured = `$true
    if (`$null -ne `$volume.AutoUnlockEnabled) {
        `$autoUnlockEnabled = [bool]`$volume.AutoUnlockEnabled
    }
}

[PSCustomObject]@{
    MountPoint           = [string]`$volume.MountPoint
    VolumeType           = [string]`$volume.VolumeType
    ProtectionStatus     = [string]`$volume.ProtectionStatus
    VolumeStatus         = [string]`$volume.VolumeStatus
    LockStatus           = if (`$volume.PSObject.Properties['LockStatus']) { [string]`$volume.LockStatus } else { `$null }
    EncryptionMethod     = [string]`$volume.EncryptionMethod
    EncryptionPercentage = if (`$null -ne `$volume.EncryptionPercentage) { [int]`$volume.EncryptionPercentage } else { `$null }
    KeyProtectorCount    = @(`$keyProtectors).Count
    KeyProtectors        = @(`$keyProtectors)
    KeyProtectorsCaptured = `$keyProtectorsCaptured
    AutoUnlockEnabled    = `$autoUnlockEnabled
    AutoUnlockCaptured   = `$autoUnlockCaptured
} | ConvertTo-Json -Compress -Depth 6
"@
            }
        }
    )
    $probeResultsByMountPoint = @{}
    foreach ($result in @(Invoke-ChildPowerShellBatch -Requests $probeRequests)) {
        $probeResultsByMountPoint[[string]$result.Key] = $result
    }

    $timedOutMountPoints = [System.Collections.Generic.List[string]]::new()
    $captureIssues = [System.Collections.Generic.List[object]]::new()
    $states = foreach ($mountPoint in $mountPoints) {
        Write-Verbose ("BitLocker snapshot mount point {0}" -f $mountPoint)
        if (-not $probeResultsByMountPoint.ContainsKey($mountPoint)) {
            $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker snapshot did not return a child-process result for this mount point.')) | Out-Null
            continue
        }
        $result = $probeResultsByMountPoint[$mountPoint]

        if ($result.TimedOut) {
            $timedOutMountPoints.Add($mountPoint) | Out-Null
            Write-Warning "BitLocker snapshot timed out on mount point $mountPoint. Skipping it."
            continue
        }

        if ($result.ExitCode -ne 0) {
            if (-not [string]::IsNullOrWhiteSpace([string]$result.StdErr)) {
                Write-Warning ("BitLocker snapshot failed on mount point {0}: {1}" -f $mountPoint, $result.StdErr.Trim())
                $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message $result.StdErr.Trim())) | Out-Null
            } else {
                $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker snapshot command failed for this mount point.')) | Out-Null
            }
            continue
        }

        $json = [string]$result.StdOut
        if ([string]::IsNullOrWhiteSpace($json)) {
            Write-Warning "BitLocker snapshot returned no data on mount point $mountPoint. Skipping it."
            $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker snapshot returned no data for this mount point.')) | Out-Null
            continue
        }

        try {
            $volume = $json | ConvertFrom-Json -ErrorAction Stop
        } catch {
            Write-Warning "BitLocker snapshot returned unparsable output on mount point $mountPoint. Skipping it."
            $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker snapshot returned unparsable output for this mount point.')) | Out-Null
            continue
        }

        $normalizedVolume = Normalize-BitLockerVolumeState -Volume $volume
        if (-not ($volume.PSObject.Properties['KeyProtectorsCaptured'] -and [bool]$volume.KeyProtectorsCaptured)) {
            $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker key protector inventory could not be captured for this mount point.')) | Out-Null
        }

        if ((Test-BitLockerAutoUnlockSupportedVolume -Volume $normalizedVolume) -and -not ($volume.PSObject.Properties['AutoUnlockCaptured'] -and [bool]$volume.AutoUnlockCaptured)) {
            $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker auto-unlock state could not be captured for this mount point.')) | Out-Null
        }

        $normalizedVolume
    }

    [PSCustomObject]@{
        CommandAvailable    = $true
        TimedOutMountPoints = @($timedOutMountPoints)
        CaptureIssues       = @($captureIssues)
        Volumes             = @($states)
    }
}

function Set-Permissive-BitLockerVolumes {
    param([AllowNull()] [object]$State)

    if (-not (Test-CommandAvailable -Name 'Suspend-BitLocker')) {
        return
    }

    if ($null -eq $State) {
        if (-not (Test-CommandAvailable -Name 'Get-BitLockerVolume')) {
            return
        }

        $State = Get-BitLockerVolumeStates
    }

    if (-not (Test-BitLockerStateCapturedExactly -State $State)) {
        return
    }

    $normalizedState = Normalize-BitLockerState -State $State

    foreach ($volume in @($normalizedState.Volumes)) {
        if (-not (Test-BitLockerProtectionEnabled -Value $volume.ProtectionStatus)) {
            continue
        }

        $mountPoint = [string]$volume.MountPoint

        if ([string]::IsNullOrWhiteSpace($mountPoint)) {
            continue
        }

        Suspend-BitLocker -MountPoint $mountPoint -RebootCount 0 -ErrorAction Stop | Out-Null
    }

    foreach ($volume in @($normalizedState.Volumes)) {
        if (-not (Test-BitLockerAutoUnlockSupportedVolume -Volume $volume)) {
            continue
        }

        if ($null -eq $volume.AutoUnlockEnabled -or $volume.AutoUnlockEnabled) {
            continue
        }

        if ([string]$volume.ProtectionMode -eq 'Decrypted') {
            continue
        }

        Set-BitLockerAutoUnlockState -MountPoint ([string]$volume.MountPoint) -Enabled $true
    }
}

function Restore-BitLockerVolumes {
    param([Parameter(Mandatory)] [object]$State)

    if (-not (Test-BitLockerStateCapturedExactly -State $State)) {
        return
    }

    if (-not (Test-CommandAvailable -Name 'Get-BitLockerVolume')) {
        return
    }

    $normalizedState = Normalize-BitLockerState -State $State
    $targetVolumes = @($normalizedState.Volumes)
    if ($targetVolumes.Count -eq 0) {
        return
    }

    $mountPoints = @($targetVolumes | ForEach-Object { [string]$_.MountPoint })
    $liveVolumes = @(Get-BitLockerVolume -MountPoint $mountPoints -ErrorAction Stop)
    $liveVolumesByMountPoint = @{}
    foreach ($liveVolume in $liveVolumes) {
        $liveMountPoint = [string]$liveVolume.MountPoint
        if (-not [string]::IsNullOrWhiteSpace($liveMountPoint)) {
            $liveVolumesByMountPoint[$liveMountPoint] = $liveVolume
        }
    }

    foreach ($mountPoint in $mountPoints) {
        if (-not $liveVolumesByMountPoint.ContainsKey($mountPoint)) {
            throw "BitLocker volume '$mountPoint' was not returned during restore."
        }
    }

    foreach ($volume in $targetVolumes) {
        $mountPoint = [string]$volume.MountPoint
        $liveVolume = $liveVolumesByMountPoint[$mountPoint]

        $targetProtection = ConvertTo-BitLockerProtectionStatusLabel -Value $volume.ProtectionStatus
        $liveProtection = ConvertTo-BitLockerProtectionStatusLabel -Value $liveVolume.ProtectionStatus
        if ($targetProtection -eq 'On' -and $liveProtection -ne 'On') {
            if (Test-CommandAvailable -Name 'Resume-BitLocker') {
                Resume-BitLocker -MountPoint $mountPoint -ErrorAction Stop | Out-Null
            } else {
                throw 'Resume-BitLocker was not found.'
            }
            continue
        }

        if ($targetProtection -eq 'Off' -and $liveProtection -ne 'Off' -and [string]$volume.VolumeStatus -ne 'FullyDecrypted') {
            if (-not (Test-CommandAvailable -Name 'Suspend-BitLocker')) {
                throw 'Suspend-BitLocker was not found.'
            }
            Suspend-BitLocker -MountPoint $mountPoint -RebootCount 0 -ErrorAction Stop | Out-Null
        }
    }

    foreach ($volume in $targetVolumes) {
        if (-not (Test-BitLockerAutoUnlockSupportedVolume -Volume $volume)) {
            continue
        }

        if ($null -eq $volume.AutoUnlockEnabled) {
            continue
        }

        Set-BitLockerAutoUnlockState -MountPoint ([string]$volume.MountPoint) -Enabled ([bool]$volume.AutoUnlockEnabled)
    }
}

#endregion

#region Application control, exploit protection, and audit providers

function New-AppLockerCaptureIssue {
    param(
        [string]$Scope,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        Scope   = $Scope
        Message = $Message
    }
}

function ConvertTo-AppLockerXmlText {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return ([string]$Value).Trim()
    }

    ((@($Value) | ForEach-Object { [string]$_ }) -join [Environment]::NewLine).Trim()
}

function Normalize-AppLockerXml {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml)) {
        return ''
    }

    try {
        $document = New-Object System.Xml.XmlDocument
        $document.PreserveWhitespace = $false
        $document.LoadXml($Xml)
        $document.OuterXml
    } catch {
        $Xml.Trim()
    }
}

function Get-AppLockerCollectionSummaries {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml)) {
        return @()
    }

    try {
        $document = New-Object System.Xml.XmlDocument
        $document.PreserveWhitespace = $false
        $document.LoadXml($Xml)

        @(
            foreach ($ruleCollection in @($document.SelectNodes('/AppLockerPolicy/RuleCollection'))) {
                if ($null -eq $ruleCollection) {
                    continue
                }

                $ruleCount = @(
                    foreach ($childNode in @($ruleCollection.ChildNodes)) {
                        if (
                            $childNode -is [System.Xml.XmlElement] -and
                            $childNode.LocalName -like '*Rule'
                        ) {
                            $childNode
                        }
                    }
                ).Count

                $servicesNode = $ruleCollection.SelectSingleNode('RuleCollectionExtensions/ThresholdExtensions/Services')
                $systemAppsNode = $ruleCollection.SelectSingleNode('RuleCollectionExtensions/RedstoneExtensions/SystemApps')

                [PSCustomObject]@{
                    Type                = if ($ruleCollection.Attributes['Type']) { [string]$ruleCollection.Attributes['Type'].Value } else { $null }
                    EnforcementMode     = if ($ruleCollection.Attributes['EnforcementMode']) { [string]$ruleCollection.Attributes['EnforcementMode'].Value } else { 'NotConfigured' }
                    RuleCount           = $ruleCount
                    ServicesEnforcement = if ($null -ne $servicesNode -and $servicesNode.Attributes['EnforcementMode']) { [string]$servicesNode.Attributes['EnforcementMode'].Value } else { $null }
                    SystemAppsAllow     = if ($null -ne $systemAppsNode -and $systemAppsNode.Attributes['Allow']) { [string]$systemAppsNode.Attributes['Allow'].Value } else { $null }
                }
            }
        )
    } catch {
        @()
    }
}

function Get-AppLockerPolicyXml {
    param(
        [AllowNull()] [object]$State,
        [ValidateSet('Local', 'Effective')] [string]$PolicyScope = 'Effective',
        [string]$SnapshotPath
    )

    if ($null -eq $State) {
        return ''
    }

    $inlineProperty = if ($PolicyScope -eq 'Local') { 'LocalXml' } else { 'EffectiveXml' }
    $assetProperty = if ($PolicyScope -eq 'Local') { 'LocalSnapshotAssetRelativePath' } else { 'EffectiveSnapshotAssetRelativePath' }
    $assetHashProperty = if ($PolicyScope -eq 'Local') { 'LocalSnapshotAssetSha256' } else { 'EffectiveSnapshotAssetSha256' }

    if (
        $State.PSObject.Properties[$inlineProperty] -and
        -not [string]::IsNullOrWhiteSpace([string]$State.PSObject.Properties[$inlineProperty].Value)
    ) {
        return [string]$State.PSObject.Properties[$inlineProperty].Value
    }

    if (
        $State.PSObject.Properties[$assetProperty] -and
        -not [string]::IsNullOrWhiteSpace([string]$State.PSObject.Properties[$assetProperty].Value) -and
        -not [string]::IsNullOrWhiteSpace($SnapshotPath)
    ) {
        $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $SnapshotPath
        $assetPath = Resolve-ContainedFileSystemPath `
            -Root $assetRoot `
            -RelativePath ([string]$State.PSObject.Properties[$assetProperty].Value) `
            -Description 'AppLocker snapshot asset path'
        if (-not (Test-Path -LiteralPath $assetPath)) {
            throw "AppLocker snapshot asset is missing: $assetPath"
        }

        $expectedSha256 = if ($State.PSObject.Properties[$assetHashProperty]) { [string]$State.PSObject.Properties[$assetHashProperty].Value } else { $null }
        return Read-SnapshotAssetText `
            -Path $assetPath `
            -ExpectedSha256 $expectedSha256 `
            -Description 'AppLocker snapshot asset'
    }

    ''
}

function Get-AppLockerPolicyState {
    if (-not (Test-CommandAvailable -Name 'Get-AppLockerPolicy')) {
        return [PSCustomObject]@{
            CommandAvailable      = $false
            LocalCaptured         = $false
            EffectiveCaptured     = $false
            LocalMatchesEffective = $false
            CaptureIssues         = @()
            CollectionSummaries   = @()
            LocalXml              = $null
            EffectiveXml          = $null
        }
    }

    $captureIssues = [System.Collections.Generic.List[object]]::new()
    $localXml = ''
    $effectiveXml = ''
    $localCaptured = $false
    $effectiveCaptured = $false

    try {
        $localXml = ConvertTo-AppLockerXmlText -Value (Get-AppLockerPolicy -Local -Xml -ErrorAction Stop)
        $localCaptured = $true
    } catch {
        $captureIssues.Add((New-AppLockerCaptureIssue -Scope 'Local' -Message $_.Exception.Message)) | Out-Null
    }

    try {
        $effectiveXml = ConvertTo-AppLockerXmlText -Value (Get-AppLockerPolicy -Effective -Xml -ErrorAction Stop)
        $effectiveCaptured = $true
    } catch {
        $captureIssues.Add((New-AppLockerCaptureIssue -Scope 'Effective' -Message $_.Exception.Message)) | Out-Null
    }

    $normalizedLocalXml = Normalize-AppLockerXml -Xml $localXml
    $normalizedEffectiveXml = Normalize-AppLockerXml -Xml $effectiveXml

    [PSCustomObject]@{
        CommandAvailable      = $true
        LocalCaptured         = $localCaptured
        EffectiveCaptured     = $effectiveCaptured
        LocalMatchesEffective = (
            $localCaptured -and
            $effectiveCaptured -and
            [string]::Equals($normalizedLocalXml, $normalizedEffectiveXml, [System.StringComparison]::Ordinal)
        )
        CaptureIssues         = @($captureIssues)
        CollectionSummaries   = @(Get-AppLockerCollectionSummaries -Xml $effectiveXml)
        LocalXml              = if ($localCaptured) { $localXml } else { $null }
        EffectiveXml          = if ($effectiveCaptured) { $effectiveXml } else { $null }
    }
}

function Test-AppLockerPolicyCapturedExactly {
    param(
        [AllowNull()] [object]$State,
        [string]$SnapshotPath
    )

    if ($null -eq $State) {
        return $false
    }

    $commandAvailable = if ($State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
    $localCaptured = if ($State.PSObject.Properties['LocalCaptured']) { [bool]$State.LocalCaptured } else { $false }
    $effectiveCaptured = if ($State.PSObject.Properties['EffectiveCaptured']) { [bool]$State.EffectiveCaptured } else { $false }
    $localMatchesEffective = if ($State.PSObject.Properties['LocalMatchesEffective']) { [bool]$State.LocalMatchesEffective } else { $false }
    $captureIssues = @(
        if ($State.PSObject.Properties['CaptureIssues']) {
            $State.CaptureIssues
        }
    )
    if (
        -not $commandAvailable -or
        -not $localCaptured -or
        -not $effectiveCaptured -or
        -not $localMatchesEffective -or
        @($captureIssues).Count -ne 0
    ) {
        return $false
    }

    $localXml = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $State -PolicyScope Local -SnapshotPath $SnapshotPath)
    $effectiveXml = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $State -PolicyScope Effective -SnapshotPath $SnapshotPath)

    (
        -not [string]::IsNullOrWhiteSpace($localXml) -and
        -not [string]::IsNullOrWhiteSpace($effectiveXml)
    )
}

function Get-EmptyAppLockerPolicyXml {
    '<AppLockerPolicy Version="1" />'
}

function Apply-AppLockerPolicyXml {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml) -or -not (Test-CommandAvailable -Name 'Set-AppLockerPolicy')) {
        return
    }

    $tempPath = New-TemporaryFilePath -Extension '.xml'
    try {
        [System.IO.File]::WriteAllText($tempPath, $Xml, [System.Text.UTF8Encoding]::new($false))
        Set-AppLockerPolicy -XmlPolicy $tempPath | Out-Null
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Set-Permissive-AppLockerPolicy {
    param(
        [AllowNull()] [object]$State,
        [string]$SnapshotPath
    )

    if ($null -eq $State) {
        $State = Get-AppLockerPolicyState
    }

    if (-not (Test-AppLockerPolicyCapturedExactly -State $State -SnapshotPath $SnapshotPath)) {
        return
    }

    Apply-AppLockerPolicyXml -Xml (Get-EmptyAppLockerPolicyXml)
}

function Restore-AppLockerPolicy {
    param(
        [Parameter(Mandatory)] [object]$State,
        [string]$SnapshotPath
    )

    if (-not (Test-AppLockerPolicyCapturedExactly -State $State -SnapshotPath $SnapshotPath)) {
        return
    }

    Apply-AppLockerPolicyXml -Xml (Get-AppLockerPolicyXml -State $State -PolicyScope Local -SnapshotPath $SnapshotPath)
}

function Get-ExploitProtectionPolicyState {
    if (-not (Test-CommandAvailable -Name 'Get-ProcessMitigation')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Xml              = $null
        }
    }

    $tempPath = New-TemporaryFilePath -Extension '.xml'
    try {
        Get-ProcessMitigation -RegistryConfigFilePath $tempPath | Out-Null
        $xml = if (Test-Path -LiteralPath $tempPath) { Get-Content -LiteralPath $tempPath -Raw } else { $null }

        [PSCustomObject]@{
            CommandAvailable = $true
            Xml              = $xml
        }
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Normalize-ExploitProtectionXml {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml)) {
        return ''
    }

    try {
        $document = New-Object System.Xml.XmlDocument
        $document.PreserveWhitespace = $false
        $document.LoadXml($Xml)

        $mitigationPolicyNode = $document.SelectSingleNode('/MitigationPolicy')
        $systemConfigNode = if ($null -ne $mitigationPolicyNode) { $mitigationPolicyNode.SelectSingleNode('SystemConfig') } else { $null }
        if ($null -ne $systemConfigNode) {
            $childElements = @($systemConfigNode.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] })
            if ($childElements.Count -eq 1 -and $childElements[0].Name -eq 'ASLR') {
                $aslrNode = [System.Xml.XmlElement]$childElements[0]
                $expectedAttributes = [ordered]@{
                    ForceRelocateImages = 'false'
                    RequireInfo         = 'false'
                    BottomUp            = 'false'
                    HighEntropy         = 'false'
                }

                $isDefaultNoOpSystemConfig = ($aslrNode.Attributes.Count -eq $expectedAttributes.Count)
                if ($isDefaultNoOpSystemConfig) {
                    foreach ($attributeName in $expectedAttributes.Keys) {
                        if (
                            -not $aslrNode.HasAttribute($attributeName) -or
                            -not [string]::Equals($aslrNode.GetAttribute($attributeName), $expectedAttributes[$attributeName], [System.StringComparison]::OrdinalIgnoreCase)
                        ) {
                            $isDefaultNoOpSystemConfig = $false
                            break
                        }
                    }
                }

                if ($isDefaultNoOpSystemConfig) {
                    $null = $mitigationPolicyNode.RemoveChild($systemConfigNode)
                }
            }
        }

        return $document.OuterXml
    } catch {
        return $Xml.Trim()
    }
}

function Get-ExploitProtectionPolicyXml {
    param(
        [AllowNull()] [object]$State,
        [string]$SnapshotPath
    )

    if ($null -eq $State) {
        return ''
    }

    if ($State.PSObject.Properties['Xml'] -and -not [string]::IsNullOrWhiteSpace([string]$State.Xml)) {
        return [string]$State.Xml
    }

    if (
        $State.PSObject.Properties['SnapshotAssetRelativePath'] -and
        -not [string]::IsNullOrWhiteSpace([string]$State.SnapshotAssetRelativePath) -and
        -not [string]::IsNullOrWhiteSpace($SnapshotPath)
    ) {
        $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $SnapshotPath
        $assetPath = Resolve-ContainedFileSystemPath `
            -Root $assetRoot `
            -RelativePath ([string]$State.SnapshotAssetRelativePath) `
            -Description 'Exploit protection snapshot asset path'
        if (-not (Test-Path -LiteralPath $assetPath)) {
            throw "Exploit protection snapshot asset is missing: $assetPath"
        }

        $expectedSha256 = if ($State.PSObject.Properties['SnapshotAssetSha256']) { [string]$State.SnapshotAssetSha256 } else { $null }
        return Read-SnapshotAssetText `
            -Path $assetPath `
            -ExpectedSha256 $expectedSha256 `
            -Description 'Exploit protection snapshot asset'
    }

    ''
}

function Test-ExploitProtectionPolicyCapturedExactly {
    param(
        [AllowNull()] [object]$State,
        [string]$SnapshotPath
    )

    $commandAvailable = if ($null -ne $State -and $State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
    return (
        $null -ne $State -and
        $commandAvailable -and
        -not [string]::IsNullOrWhiteSpace((Get-ExploitProtectionPolicyXml -State $State -SnapshotPath $SnapshotPath))
    )
}

function Apply-ExploitProtectionPolicyXml {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml) -or -not (Test-CommandAvailable -Name 'Set-ProcessMitigation')) {
        return
    }

    $tempPath = New-TemporaryFilePath -Extension '.xml'
    try {
        [System.IO.File]::WriteAllText($tempPath, $Xml, [System.Text.UTF8Encoding]::new($false))
        Set-ProcessMitigation -PolicyFilePath $tempPath | Out-Null
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Reset-ExploitProtectionSystemConfig {
    if (-not (Test-CommandAvailable -Name 'Set-ProcessMitigation')) {
        return
    }

    try {
        Set-ProcessMitigation -System -Reset | Out-Null
    } catch {
        Write-Verbose ("Set-ProcessMitigation -System -Reset was unavailable or rejected: {0}" -f $_.Exception.Message)
    }
}

function Set-Permissive-ExploitProtection {
    Reset-ExploitProtectionSystemConfig

    foreach ($mitigation in @('DEP', 'EmulateAtlThunks', 'CFG', 'StrictCFG', 'SuppressExports', 'ForceRelocateImages', 'BottomUp', 'HighEntropy', 'SEHOP', 'SEHOPTelemetry')) {
        try {
            Set-ProcessMitigation -System -Disable $mitigation | Out-Null
        } catch {
            Write-Verbose ("Unable to disable exploit mitigation '{0}': {1}" -f $mitigation, $_.Exception.Message)
        }
    }
}

function Get-WdacCodeIntegrityRoot {
    Resolve-FileSystemPath -Path (Join-Path $env:windir 'System32\CodeIntegrity')
}

function Get-WdacPolicyFileBackups {
    $root = Get-WdacCodeIntegrityRoot
    $files = @()

    $activeDir = Join-Path $root 'CiPolicies\Active'
    if (Test-Path -LiteralPath $activeDir) {
        $files += @(
            foreach ($file in @(Get-ChildItem -LiteralPath $activeDir -Filter '*.cip' -File -ErrorAction SilentlyContinue)) {
                [PSCustomObject]@{
                    RelativePath = Join-Path 'CiPolicies\Active' $file.Name
                    FileName     = $file.Name
                    Sha256       = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
                    Base64       = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file.FullName))
                }
            }
        )
    }

    $singlePolicyPath = Join-Path $root 'SiPolicy.p7b'
    if (Test-Path -LiteralPath $singlePolicyPath) {
        $files += [PSCustomObject]@{
            RelativePath = 'SiPolicy.p7b'
            FileName     = 'SiPolicy.p7b'
            Sha256       = (Get-FileHash -LiteralPath $singlePolicyPath -Algorithm SHA256).Hash
            Base64       = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($singlePolicyPath))
        }
    }

    @($files)
}

function Get-WdacPolicyDestinationPath {
    param([Parameter(Mandatory)] [object]$File)

    if (-not $File.PSObject.Properties['RelativePath']) {
        throw 'A WDAC snapshot file is missing its relative path.'
    }

    $relativePath = ([string]$File.RelativePath).Replace('/', '\')
    $fileName = if ($File.PSObject.Properties['FileName']) { [string]$File.FileName } else { $null }
    $canonicalRelativePath = $null
    $expectedFileName = $null

    if ([string]::Equals($relativePath, 'SiPolicy.p7b', [System.StringComparison]::OrdinalIgnoreCase)) {
        $canonicalRelativePath = 'SiPolicy.p7b'
        $expectedFileName = 'SiPolicy.p7b'
    } elseif ($relativePath -match '^CiPolicies\\Active\\([^\\/:*?"<>|]+\.cip)$') {
        $expectedFileName = [string]$matches[1]
        $canonicalRelativePath = Join-Path (Join-Path 'CiPolicies' 'Active') $expectedFileName
    } else {
        throw "A WDAC snapshot file targets an unsupported policy path: $relativePath"
    }

    if (
        -not [string]::IsNullOrWhiteSpace($fileName) -and
        -not [string]::Equals($fileName, $expectedFileName, [System.StringComparison]::OrdinalIgnoreCase)
    ) {
        throw "A WDAC snapshot file name does not match its relative path: $fileName"
    }

    Resolve-ContainedFileSystemPath -Root (Get-WdacCodeIntegrityRoot) -RelativePath $canonicalRelativePath -Description 'WDAC policy path'
}

function Get-WdacPolicyIdentifiers {
    param([Parameter(Mandatory)] [object[]]$Policies)

    @(
        foreach ($policy in @($Policies)) {
            $policyId = Get-WdacPolicyIdentifierFromObject -Policy $policy
            if (-not [string]::IsNullOrWhiteSpace($policyId)) {
                $policyId
            }
        }
    ) | Sort-Object -Unique
}

function Get-WdacPolicyIdentifierFromObject {
    param([Parameter(Mandatory)] [object]$Policy)

    foreach ($name in @('PolicyID', 'PolicyId', 'PolicyGuid', 'PolicyGUID', 'Id', 'ID')) {
        $property = $Policy.PSObject.Properties[$name]
        if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return [string]$property.Value
        }
    }

    return $null
}

function Get-WdacNormalizedPolicyId {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = $text.TrimStart('{').TrimEnd('}')
    $guid = [guid]::Empty
    if ([guid]::TryParse($text, [ref]$guid)) {
        return $guid.ToString().ToLowerInvariant()
    }

    return $text.ToLowerInvariant()
}

function Get-WdacPolicyIdFromFileName {
    param([AllowNull()] [string]$FileName)

    if ([string]::IsNullOrWhiteSpace($FileName)) {
        return $null
    }

    $leaf = [System.IO.Path]::GetFileName($FileName)
    if ($leaf -match '^\{?([0-9A-Fa-f-]{36})\}?\.cip$') {
        return Get-WdacNormalizedPolicyId -Value $matches[1]
    }

    return $null
}

function Get-WdacPolicyPlatformManagementInfo {
    param([Parameter(Mandatory)] [object]$Policy)

    $policyId = Get-WdacPolicyIdentifierFromObject -Policy $Policy
    $normalizedPolicyId = Get-WdacNormalizedPolicyId -Value $policyId
    $friendlyName = if (
        $Policy.PSObject.Properties['FriendlyName'] -and
        -not [string]::IsNullOrWhiteSpace([string]$Policy.FriendlyName)
    ) {
        [string]$Policy.FriendlyName
    } else {
        '<unknown>'
    }

    $platformPolicy = $null
    foreach ($propertyName in @('Platform Policy', 'PlatformPolicy')) {
        if ($Policy.PSObject.Properties[$propertyName]) {
            $platformPolicy = ConvertTo-NullableBoolean -Value $Policy.PSObject.Properties[$propertyName].Value
            if ($null -ne $platformPolicy) {
                break
            }
        }
    }

    $knownInboxPolicies = @{
        '0283ac0f-fff1-49ae-ada1-8a933130cad6' = 'Inbox Smart App Control base policy'
        '1283ac0f-fff1-49ae-ada1-8a933130cad6' = 'Inbox Smart App Control evaluation base policy'
        '1678656c-05ef-481f-bc5b-ebd8c991502d' = 'Inbox Smart App Control flight supplemental policy'
        '2678656c-05ef-481f-bc5b-ebd8c991502d' = 'Inbox Smart App Control evaluation flight supplemental policy'
        '0939ed82-bfd5-4d32-b58e-d31d3c49715a' = 'Inbox Smart App Control test supplemental policy'
        '1939ed82-bfd5-4d32-b58e-d31d3c49715a' = 'Inbox Smart App Control evaluation test supplemental policy'
        'd2bda982-ccf6-4344-ac5b-0b44427b6816' = 'Inbox Microsoft Windows Driver Policy'
        'a072029f-588b-4b5e-b7f9-05aad67df687' = 'Inbox Microsoft Windows Virtualization Based Security policy'
        '82443e1e-8a39-4b4a-96a8-f40ddc00b9f3' = 'Inbox Windows 11 SE lockdown base policy'
        '5dac656c-21ad-4a02-ab49-649917162e70' = 'Inbox Windows 11 SE flight supplemental policy'
        'cdd5cb55-db68-4d71-aa38-3df2b6473a52' = 'Inbox Windows 11 SE test supplemental policy'
        '5951a96a-e0b5-4d3d-8fb8-3e5b61030784' = 'Inbox Windows 10 S lockdown base policy'
        '784c4414-79f4-4c32-a6a5-f0fb42a51d0d' = 'Inbox Microsoft Code Integrity cross-certificates exception policy'
    }

    $classification = $null
    if (-not [string]::IsNullOrWhiteSpace($normalizedPolicyId) -and $knownInboxPolicies.ContainsKey($normalizedPolicyId)) {
        $classification = $knownInboxPolicies[$normalizedPolicyId]
    } elseif ($friendlyName -match '^VerifiedAndReputableDesktop') {
        $classification = 'Inbox Smart App Control policy'
    } elseif ($friendlyName -eq 'Microsoft Windows Driver Policy') {
        $classification = 'Inbox Microsoft Windows Driver Policy'
    } elseif ($friendlyName -eq 'Microsoft Windows Virtualization Based Security Policy') {
        $classification = 'Inbox Microsoft Windows Virtualization Based Security policy'
    } elseif ($friendlyName -match '^Windows(E|10S)_Lockdown') {
        $classification = 'Inbox Windows lockdown policy'
    } elseif ($platformPolicy -eq $true) {
        $classification = 'Platform-managed WDAC policy'
    }

    [PSCustomObject]@{
        PolicyId           = $policyId
        FriendlyName       = $friendlyName
        NormalizedPolicyId = $normalizedPolicyId
        PlatformPolicy     = $platformPolicy
        IsPlatformManaged  = -not [string]::IsNullOrWhiteSpace([string]$classification)
        Classification     = $classification
    }
}

function Get-WdacPlatformManagedPolicies {
    param([AllowNull()] [object]$State)

    if ($null -eq $State -or -not $State.PSObject.Properties['Policies']) {
        return @()
    }

    @(
        foreach ($policy in @($State.Policies)) {
            $info = Get-WdacPolicyPlatformManagementInfo -Policy $policy
            if ($info.IsPlatformManaged) {
                $info
            }
        }
    )
}

function Get-WdacPolicyReportRows {
    param([AllowNull()] [object]$State)

    if ($null -eq $State -or -not $State.PSObject.Properties['Policies']) {
        return @()
    }

    @(
        foreach ($policy in @($State.Policies)) {
            $policyInfo = Get-WdacPolicyPlatformManagementInfo -Policy $policy

            $hasFileOnDisk = $null
            if ($policy.PSObject.Properties['HasFileOnDisk']) {
                $hasFileOnDisk = ConvertTo-NullableBoolean -Value $policy.HasFileOnDisk
            }

            $isEnforced = $null
            foreach ($name in @('IsCurrentlyEnforced', 'IsEnforced')) {
                if ($policy.PSObject.Properties[$name]) {
                    $isEnforced = ConvertTo-NullableBoolean -Value $policy.$name
                    if ($null -ne $isEnforced) {
                        break
                    }
                }
            }

            $presence = if ($hasFileOnDisk -eq $true -and $isEnforced -eq $true) {
                'Active + on-disk'
            } elseif ($hasFileOnDisk -eq $true -and $isEnforced -eq $false) {
                'On-disk only / pending reboot'
            } elseif ($hasFileOnDisk -eq $false -and $isEnforced -eq $true) {
                'Active only'
            } elseif ($hasFileOnDisk -eq $false -and $isEnforced -eq $false) {
                'Present but inactive'
            } elseif ($hasFileOnDisk -eq $true) {
                'On-disk state only'
            } elseif ($isEnforced -eq $true) {
                'Active state only'
            } else {
                'Unknown state'
            }

            [PSCustomObject]@{
                PolicyId          = $policyInfo.PolicyId
                FriendlyName      = $policyInfo.FriendlyName
                Classification    = $policyInfo.Classification
                IsPlatformManaged = $policyInfo.IsPlatformManaged
                HasFileOnDisk     = $hasFileOnDisk
                IsEnforced        = $isEnforced
                Presence          = $presence
            }
        }
    )
}

function ConvertTo-ComparableWdacState {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $null
    }

    $platformPolicyIds = @(
        foreach ($policy in @($State.Policies)) {
            $managementInfo = Get-WdacPolicyPlatformManagementInfo -Policy $policy
            if ($managementInfo.IsPlatformManaged -and -not [string]::IsNullOrWhiteSpace([string]$managementInfo.NormalizedPolicyId)) {
                [string]$managementInfo.NormalizedPolicyId
            }
        }
    )
    $customPolicies = @(
        @($State.Policies) |
            Where-Object { -not (Get-WdacPolicyPlatformManagementInfo -Policy $_).IsPlatformManaged } |
            Sort-Object -Property @{ Expression = { Get-WdacNormalizedPolicyId -Value (Get-WdacPolicyIdentifierFromObject -Policy $_) } }, @{ Expression = { [string]$_.FriendlyName } }
    )
    $customFiles = @(
        foreach ($file in @($State.Files)) {
            $filePolicyId = Get-WdacPolicyIdFromFileName -FileName ([string]$file.FileName)
            if (-not [string]::IsNullOrWhiteSpace($filePolicyId) -and $filePolicyId -in $platformPolicyIds) {
                continue
            }
            $file
        }
    )

    [ordered]@{
        Captured      = if ($State.PSObject.Properties['Captured']) { [bool]$State.Captured } else { $true }
        CaptureIssues = @(
            if ($State.PSObject.Properties['CaptureIssues']) {
                @([string[]]$State.CaptureIssues) | Sort-Object
            }
        )
        Policies      = @(
            foreach ($policy in $customPolicies) {
                $ordered = [ordered]@{}
                foreach ($property in @($policy.PSObject.Properties | Sort-Object -Property Name)) {
                    $ordered[$property.Name] = $property.Value
                }
                [PSCustomObject]$ordered
            }
        )
        Files         = @(
            foreach ($file in @($customFiles | Sort-Object -Property RelativePath, FileName)) {
                [PSCustomObject]@{
                    RelativePath = [string]$file.RelativePath
                    FileName     = [string]$file.FileName
                    Sha256       = [string]$file.Sha256
                }
            }
        )
    }
}

function Test-WdacPolicyRemovalCandidate {
    param([Parameter(Mandatory)] [object]$Policy)

    $hasFileOnDisk = $null
    if ($Policy.PSObject.Properties['HasFileOnDisk']) {
        $hasFileOnDisk = ConvertTo-NullableBoolean -Value $Policy.HasFileOnDisk
    }

    $isEnforced = $null
    foreach ($name in @('IsCurrentlyEnforced', 'IsEnforced')) {
        if ($Policy.PSObject.Properties[$name]) {
            $isEnforced = ConvertTo-NullableBoolean -Value $Policy.$name
            if ($null -ne $isEnforced) {
                break
            }
        }
    }

    return (($hasFileOnDisk -eq $true) -or ($isEnforced -eq $true))
}

function Test-WdacPolicyPresent {
    param(
        [Parameter(Mandatory)] [object]$State,
        [AllowNull()] [string]$PolicyId
    )

    $normalizedTarget = Get-WdacNormalizedPolicyId -Value $PolicyId
    if ([string]::IsNullOrWhiteSpace($normalizedTarget)) {
        return $false
    }

    foreach ($policy in @($State.Policies)) {
        $normalizedPolicyId = Get-WdacNormalizedPolicyId -Value (Get-WdacPolicyIdentifierFromObject -Policy $policy)
        if (-not [string]::IsNullOrWhiteSpace($normalizedPolicyId) -and $normalizedPolicyId -eq $normalizedTarget) {
            return $true
        }
    }

    return $false
}

function Test-WdacSnapshotFileMatchesLiveFile {
    param(
        [Parameter(Mandatory)] [object]$File,
        [string]$SnapshotPath
    )

    $destination = Get-WdacPolicyDestinationPath -File $File
    if (-not (Test-Path -LiteralPath $destination)) {
        return $false
    }

    $expectedHash = if ($File.PSObject.Properties['Sha256']) { [string]$File.Sha256 } else { $null }
    if ([string]::IsNullOrWhiteSpace($expectedHash)) {
        $expectedHash = (Get-FileHash -InputStream ([System.IO.MemoryStream]::new((Get-WdacSnapshotFileBytes -File $File -SnapshotPath $SnapshotPath))) -Algorithm SHA256).Hash
    }

    $currentHash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash
    return [string]::Equals($expectedHash, $currentHash, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-WdacPolicyState {
    $ciTool = Get-CiToolCommand
    $ciToolAvailable = $null -ne $ciTool
    $policies = @()
    $captureIssues = [System.Collections.Generic.List[string]]::new()

    if ($ciToolAvailable) {
        try {
            $json = (& $ciTool.Source -lp -json 2>&1 | Out-String).Trim()
            if ($LASTEXITCODE -ne 0) {
                throw "CiTool policy inventory exited with code ${LASTEXITCODE}: $json"
            }
            if ([string]::IsNullOrWhiteSpace($json)) {
                throw 'CiTool policy inventory returned no JSON.'
            }

            $parsed = $json | ConvertFrom-Json
            $policyItems = @(if ($parsed.PSObject.Properties['Policies']) { $parsed.Policies } else { $parsed })
            $policies = @(
                foreach ($policy in @($policyItems)) {
                    $ordered = [ordered]@{}
                    foreach ($property in @($policy.PSObject.Properties | Sort-Object -Property Name)) {
                        $ordered[$property.Name] = $property.Value
                    }
                    [PSCustomObject]$ordered
                }
            )
        } catch {
            $policies = @()
            $captureIssues.Add($_.Exception.Message)
        }
    }

    $files = @()
    try {
        $files = @(Get-WdacPolicyFileBackups)
    } catch {
        $captureIssues.Add("WDAC policy files could not be captured: $($_.Exception.Message)")
    }

    [PSCustomObject]@{
        CiToolAvailable = $ciToolAvailable
        Captured        = $captureIssues.Count -eq 0
        CaptureIssues   = @($captureIssues)
        Policies        = @($policies)
        Files           = @($files)
    }
}

function Test-WdacPolicyStateCapturedExactly {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $false
    }

    $captured = if ($State.PSObject.Properties['Captured']) { [bool]$State.Captured } else { $true }
    $captureIssues = @(
        if ($State.PSObject.Properties['CaptureIssues']) {
            $State.CaptureIssues
        }
    )
    $captured -and $captureIssues.Count -eq 0
}

function Remove-WdacPolicies {
    param([Parameter(Mandatory)] [object]$State)

    $platformPolicyIds = @(
        foreach ($policy in @($State.Policies)) {
            $managementInfo = Get-WdacPolicyPlatformManagementInfo -Policy $policy
            if ($managementInfo.IsPlatformManaged -and -not [string]::IsNullOrWhiteSpace([string]$managementInfo.NormalizedPolicyId)) {
                [string]$managementInfo.NormalizedPolicyId
            }
        }
    )
    $policyIds = @(@(
        foreach ($policy in @($State.Policies)) {
            if (-not (Test-WdacPolicyRemovalCandidate -Policy $policy)) {
                continue
            }

            $managementInfo = Get-WdacPolicyPlatformManagementInfo -Policy $policy
            if ($managementInfo.IsPlatformManaged) {
                Write-Verbose ("Skipping platform-managed WDAC policy {0} ({1})." -f $managementInfo.PolicyId, $managementInfo.Classification)
                continue
            }

            $policyId = Get-WdacPolicyIdentifierFromObject -Policy $policy
            if ([string]::IsNullOrWhiteSpace($policyId)) {
                throw 'A removable WDAC policy has no usable policy identifier. No raw policy files were deleted.'
            }

            $policyId
        }
    ) | Sort-Object -Unique)

    if ($policyIds.Count -eq 0) {
        $unclassifiedFiles = @(
            foreach ($file in @($State.Files)) {
                $filePolicyId = Get-WdacPolicyIdFromFileName -FileName ([string]$file.FileName)
                if (-not [string]::IsNullOrWhiteSpace($filePolicyId) -and $filePolicyId -in $platformPolicyIds) {
                    continue
                }
                $file
            }
        )
        if ($unclassifiedFiles.Count -gt 0) {
            throw 'WDAC policy files are present without a safely removable policy identifier. No raw policy files were deleted.'
        }
        return
    }

    $ciTool = if ($State.PSObject.Properties['CiToolAvailable'] -and [bool]$State.CiToolAvailable) { Get-CiToolCommand } else { $null }
    if ($null -eq $ciTool) {
        throw 'CiTool is unavailable, so WinDefState cannot safely remove non-platform WDAC policies. No raw policy files were deleted.'
    }

    foreach ($policyId in @($policyIds)) {
        $output = (& $ciTool.Source -rp $policyId -json 2>&1 | Out-String).Trim()
        if ($LASTEXITCODE -ne 0) {
            $message = "CiTool could not remove WDAC policy '$policyId' (exit ${LASTEXITCODE})."
            if (-not [string]::IsNullOrWhiteSpace($output)) {
                $message = "$message Output: $output"
            }
            throw $message
        }
    }

    $refreshOutput = (& $ciTool.Source -r 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) {
        $message = "CiTool could not refresh WDAC policy state (exit ${LASTEXITCODE})."
        if (-not [string]::IsNullOrWhiteSpace($refreshOutput)) {
            $message = "$message Output: $refreshOutput"
        }
        throw $message
    }
}

function Get-WdacRestoreRemovalState {
    param(
        [Parameter(Mandatory)] [object]$BaselineState,
        [Parameter(Mandatory)] [object]$LiveState
    )

    $baselinePolicyIds = @(
        @(
            @(Get-WdacPolicyIdentifiers -Policies @($BaselineState.Policies))
            foreach ($file in @($BaselineState.Files)) {
                Get-WdacPolicyIdFromFileName -FileName ([string]$file.FileName)
            }
        ) |
            ForEach-Object { Get-WdacNormalizedPolicyId -Value $_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            Sort-Object -Unique
    )

    $policiesToRemove = @(
        foreach ($policy in @($LiveState.Policies)) {
            if (-not (Test-WdacPolicyRemovalCandidate -Policy $policy)) {
                continue
            }

            $managementInfo = Get-WdacPolicyPlatformManagementInfo -Policy $policy
            if ($managementInfo.IsPlatformManaged) {
                continue
            }

            $policyId = Get-WdacNormalizedPolicyId -Value (Get-WdacPolicyIdentifierFromObject -Policy $policy)
            if ([string]::IsNullOrWhiteSpace($policyId)) {
                throw 'A live non-platform WDAC policy has no usable identifier. Restore left it untouched.'
            }
            if ($policyId -notin $baselinePolicyIds) {
                $policy
            }
        }
    )
    $removalPolicyIds = @(Get-WdacPolicyIdentifiers -Policies $policiesToRemove | ForEach-Object { Get-WdacNormalizedPolicyId -Value $_ })
    $filesToRemove = @(
        foreach ($file in @($LiveState.Files)) {
            $filePolicyId = Get-WdacPolicyIdFromFileName -FileName ([string]$file.FileName)
            if (-not [string]::IsNullOrWhiteSpace($filePolicyId) -and $filePolicyId -in $removalPolicyIds) {
                $file
            }
        }
    )

    [PSCustomObject]@{
        CiToolAvailable = if ($LiveState.PSObject.Properties['CiToolAvailable']) { [bool]$LiveState.CiToolAvailable } else { $false }
        Policies        = @($policiesToRemove)
        Files           = @($filesToRemove)
    }
}

function Get-WdacSnapshotFileBytes {
    param(
        [Parameter(Mandatory)] [object]$File,
        [string]$SnapshotPath
    )

    $expectedSha256 = if ($File.PSObject.Properties['Sha256']) { [string]$File.Sha256 } else { $null }
    if ($File.PSObject.Properties['Base64'] -and -not [string]::IsNullOrWhiteSpace([string]$File.Base64)) {
        $bytes = [Convert]::FromBase64String([string]$File.Base64)
        $record = [PSCustomObject]@{
            Path = '<inline WDAC snapshot content>'
            Bytes = [byte[]]$bytes
            Sha256 = Get-Sha256HashFromBytes -Content $bytes
            Text = $null
        }
        Assert-SnapshotAssetHash -Record $record -ExpectedSha256 $expectedSha256 -Description 'WDAC snapshot content'
        return ,([byte[]]$bytes)
    }

    if (
        $File.PSObject.Properties['SnapshotAssetRelativePath'] -and
        -not [string]::IsNullOrWhiteSpace([string]$File.SnapshotAssetRelativePath) -and
        -not [string]::IsNullOrWhiteSpace($SnapshotPath)
    ) {
        $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $SnapshotPath
        $assetPath = Resolve-ContainedFileSystemPath -Root $assetRoot -RelativePath ([string]$File.SnapshotAssetRelativePath) -Description 'WDAC snapshot asset path'
        if (-not (Test-Path -LiteralPath $assetPath)) {
            throw "WDAC snapshot asset is missing: $assetPath"
        }

        return Read-SnapshotAssetBytes `
            -Path $assetPath `
            -ExpectedSha256 $expectedSha256 `
            -Description 'WDAC snapshot asset'
    }

    throw "WDAC snapshot content is missing for $([string]$File.RelativePath)"
}

function Write-WdacPolicyFiles {
    param(
        [Parameter(Mandatory)] [object[]]$Files,
        [string]$SnapshotPath
    )

    foreach ($file in @($Files)) {
        $destination = Get-WdacPolicyDestinationPath -File $file
        Write-BytesAtomic -Path $destination -Content (Get-WdacSnapshotFileBytes -File $file -SnapshotPath $SnapshotPath)
    }
}

function Get-WdacCiToolStagingFileName {
    param([Parameter(Mandatory)] [object]$File)

    $candidates = @()
    if ($File.PSObject.Properties['FileName']) {
        $candidates += [string]$File.FileName
    }

    if ($File.PSObject.Properties['RelativePath']) {
        $candidates += Split-Path -Leaf ([string]$File.RelativePath)
    }

    foreach ($candidate in $candidates) {
        $safeName = [System.IO.Path]::GetFileName([string]$candidate)
        if (-not [string]::IsNullOrWhiteSpace($safeName)) {
            return $safeName
        }
    }

    return ('{{{0}}}.cip' -f ([guid]::NewGuid().ToString().ToUpperInvariant()))
}

function Restore-WdacPolicies {
    param(
        [Parameter(Mandatory)] [object]$State,
        [string]$SnapshotPath
    )

    $liveState = Get-WdacPolicyState
    $removalState = Get-WdacRestoreRemovalState -BaselineState $State -LiveState $liveState
    if (@($removalState.Policies).Count -gt 0) {
        Remove-WdacPolicies -State $removalState
    }

    $files = @($State.Files)
    if ($files.Count -eq 0) {
        return
    }

    $ciPolicyFiles = @($files | Where-Object { ([string]$_.FileName).ToLowerInvariant().EndsWith('.cip') })
    $singlePolicyFiles = @($files | Where-Object { ([string]$_.FileName) -eq 'SiPolicy.p7b' })

    $ciTool = Get-CiToolCommand
    if ($null -ne $ciTool) {
        $stagingRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("WinDefState-Wdac-{0}" -f ([guid]::NewGuid().ToString('N')))
        Ensure-Directory -Path $stagingRoot
        $usedDirectFileFallback = $false
        try {
            foreach ($file in $ciPolicyFiles) {
                $filePolicyId = Get-WdacPolicyIdFromFileName -FileName ([string]$file.FileName)
                if (Test-WdacSnapshotFileMatchesLiveFile -File $file -SnapshotPath $SnapshotPath) {
                    Write-Verbose ("WDAC policy file already matches snapshot on disk: {0}" -f ([string]$file.RelativePath))
                    continue
                }

                $tempPath = Join-Path $stagingRoot (Get-WdacCiToolStagingFileName -File $file)
                try {
                    Write-Verbose ("Restoring WDAC policy via CiTool from staged file {0}" -f $tempPath)
                    [System.IO.File]::WriteAllBytes($tempPath, (Get-WdacSnapshotFileBytes -File $file -SnapshotPath $SnapshotPath))
                    $ciToolOutput = (& $ciTool.Source -up $tempPath -json 2>&1 | Out-String).Trim()
                    if ($LASTEXITCODE -ne 0) {
                        $ciToolExitCode = $LASTEXITCODE
                        $message = "CiTool failed to restore WDAC policy from $tempPath (exit code $LASTEXITCODE)"
                        if (-not [string]::IsNullOrWhiteSpace($ciToolOutput)) {
                            $message = "$message. Output: $ciToolOutput"
                        }

                        $postFailureState = Get-WdacPolicyState
                        if (Test-WdacSnapshotFileMatchesLiveFile -File $file -SnapshotPath $SnapshotPath) {
                            Write-Warning ("CiTool reported failure while restoring WDAC policy {0}, but the on-disk policy file already matches the snapshot. Continuing." -f ([string]$file.FileName))
                            continue
                        }

                        if ((-not [string]::IsNullOrWhiteSpace($filePolicyId)) -and (Test-WdacPolicyPresent -State $postFailureState -PolicyId $filePolicyId)) {
                            Write-Warning ("CiTool reported failure while restoring WDAC policy {0}, but the policy is already present in the live CiTool state. Continuing." -f $filePolicyId)
                            continue
                        }

                        if ($ciToolExitCode -eq -2147024891) {
                            Write-Warning ("CiTool access denied while restoring WDAC policy {0}. Falling back to direct file restore; a reboot may be required before live WDAC state fully matches the snapshot." -f ([string]$file.FileName))
                            Write-WdacPolicyFiles -Files @($file) -SnapshotPath $SnapshotPath
                            if (Test-WdacSnapshotFileMatchesLiveFile -File $file -SnapshotPath $SnapshotPath) {
                                $usedDirectFileFallback = $true
                                continue
                            }
                        }

                        throw $message
                    }
                } finally {
                    if (Test-Path -LiteralPath $tempPath) {
                        Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
                    }
                }
            }

            if ($ciPolicyFiles.Count -gt 0 -and -not $usedDirectFileFallback) {
                $refreshOutput = (& $ciTool.Source -r 2>&1 | Out-String).Trim()
                if ($LASTEXITCODE -ne 0) {
                    $message = "CiTool could not refresh restored WDAC policy state (exit ${LASTEXITCODE})."
                    if (-not [string]::IsNullOrWhiteSpace($refreshOutput)) {
                        $message = "$message Output: $refreshOutput"
                    }
                    throw $message
                }
            }

            if ($singlePolicyFiles.Count -gt 0) {
                Write-WdacPolicyFiles -Files @($singlePolicyFiles) -SnapshotPath $SnapshotPath
            }

            if ($usedDirectFileFallback) {
                Write-Warning 'One or more WDAC policy files were restored directly to disk after CiTool access was denied. A reboot may be required before verification fully matches the snapshot.'
            }

            return
        } finally {
            if (Test-Path -LiteralPath $stagingRoot) {
                Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Write-Warning 'CiTool is unavailable. Restoring captured WDAC files directly without deleting unrecognized live policy files; a reboot may be required.'
    Write-WdacPolicyFiles -Files $files -SnapshotPath $SnapshotPath
}

function Initialize-AuditPolicyInterop {
    if ($null -ne ('WinDefState.Native.AuditPolicy' -as [type])) {
        return
    }

    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace WinDefState.Native
{
    public static class AuditPolicy
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct AuditPolicyInformation
        {
            public Guid AuditSubCategoryGuid;
            public uint AuditingInformation;
            public Guid AuditCategoryGuid;
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool AuditQuerySystemPolicy(
            [In] Guid[] subCategoryGuids,
            uint policyCount,
            out IntPtr auditPolicy);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool AuditSetSystemPolicy(
            [In] AuditPolicyInformation[] auditPolicy,
            uint policyCount);

        [DllImport("advapi32.dll")]
        private static extern void AuditFree(IntPtr buffer);

        public static uint Query(Guid subCategoryGuid)
        {
            IntPtr buffer;
            if (!AuditQuerySystemPolicy(new[] { subCategoryGuid }, 1, out buffer))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }

            try
            {
                AuditPolicyInformation policy = (AuditPolicyInformation)Marshal.PtrToStructure(
                    buffer,
                    typeof(AuditPolicyInformation));
                return policy.AuditingInformation;
            }
            finally
            {
                AuditFree(buffer);
            }
        }

        public static void Set(Guid subCategoryGuid, bool success, bool failure)
        {
            uint options = (success ? 1u : 0u) | (failure ? 2u : 0u);
            AuditPolicyInformation[] policy = new[]
            {
                new AuditPolicyInformation
                {
                    AuditSubCategoryGuid = subCategoryGuid,
                    AuditingInformation = options,
                    AuditCategoryGuid = Guid.Empty
                }
            };

            if (!AuditSetSystemPolicy(policy, 1))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
        }
    }
}
'@ -ErrorAction Stop -Verbose:$false
}

function Resolve-AuditSubcategoryGuid {
    param(
        [Parameter(Mandatory)] [string]$Subcategory,
        [string]$SubcategoryGuid
    )

    $candidate = $SubcategoryGuid
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        $candidate = switch ($Subcategory) {
            'Process Creation' { '{0CCE922B-69AE-11D9-BED3-505054503030}' }
            default { $null }
        }
    }

    $guid = [guid]::Empty
    if ([string]::IsNullOrWhiteSpace($candidate) -or -not [guid]::TryParse($candidate, [ref]$guid)) {
        throw "No valid audit subcategory GUID is registered for '$Subcategory'."
    }

    $guid
}

function Get-AuditPolicyState {
    param(
        [Parameter(Mandatory)] [string]$Subcategory,
        [string]$SubcategoryGuid
    )

    try {
        Initialize-AuditPolicyInterop
        $guid = Resolve-AuditSubcategoryGuid -Subcategory $Subcategory -SubcategoryGuid $SubcategoryGuid
        $options = [WinDefState.Native.AuditPolicy]::Query($guid)
        [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $true
            Error            = $null
            Success          = ($options -band 1) -ne 0
            Failure          = ($options -band 2) -ne 0
        }
    } catch {
        [PSCustomObject]@{
            CommandAvailable = $null -ne ('WinDefState.Native.AuditPolicy' -as [type])
            Captured         = $false
            Error            = $_.Exception.Message
            Success          = $false
            Failure          = $false
        }
    }
}

function Set-AuditPolicyState {
    param(
        [Parameter(Mandatory)] [string]$Subcategory,
        [string]$SubcategoryGuid,
        [Parameter(Mandatory)] [bool]$Success,
        [Parameter(Mandatory)] [bool]$Failure
    )

    Initialize-AuditPolicyInterop
    $guid = Resolve-AuditSubcategoryGuid -Subcategory $Subcategory -SubcategoryGuid $SubcategoryGuid
    [WinDefState.Native.AuditPolicy]::Set($guid, $Success, $Failure)
}

#endregion

#region Firewall and service mutation providers

function Convert-ServiceStartModeToScValue {
    param([Parameter(Mandatory)] [string]$StartMode)

    switch ($StartMode) {
        'Auto' { 'auto' }
        'Automatic' { 'auto' }
        'Manual' { 'demand' }
        'Demand' { 'demand' }
        'Disabled' { 'disabled' }
        default { 'demand' }
    }
}

function Set-ServiceStartModeValue {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [string]$StartModeValue
    )

    if (-not (Test-CommandAvailable -Name 'sc.exe')) {
        throw 'sc.exe was not found.'
    }

    $output = & sc.exe config $Name "start= $StartModeValue" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "sc.exe could not configure service '$Name' (exit ${LASTEXITCODE}): $($output -join ' ')"
    }
}

function Set-ServiceRunningState {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [bool]$Running,
        [int]$TimeoutSeconds = 20
    )

    $service = Get-Service -Name $Name -ErrorAction Stop
    $targetStatus = if ($Running) { [System.ServiceProcess.ServiceControllerStatus]::Running } else { [System.ServiceProcess.ServiceControllerStatus]::Stopped }
    if ($service.Status -ne $targetStatus) {
        if ($Running) {
            Start-Service -Name $Name -ErrorAction Stop
        } else {
            Stop-Service -Name $Name -Force -ErrorAction Stop
        }
        $service = Get-Service -Name $Name -ErrorAction Stop
        $service.WaitForStatus($targetStatus, [TimeSpan]::FromSeconds($TimeoutSeconds))
    }
}

function ConvertTo-FirewallProfileEnabledValue {
    param([AllowNull()] [object]$Value)

    switch ([string]$Value) {
        'True' { 'True' }
        'False' { 'False' }
        'NotConfigured' { 'NotConfigured' }
        '1' { 'True' }
        '0' { 'False' }
        default {
            if ($Value -is [bool]) {
                if ($Value) { 'True' } else { 'False' }
            } else {
                [string]$Value
            }
        }
    }
}

function ConvertTo-FirewallProfileActionValue {
    param([AllowNull()] [object]$Value)

    switch ([string]$Value) {
        'Allow' { 'Allow' }
        'Block' { 'Block' }
        'NotConfigured' { 'NotConfigured' }
        default { [string]$Value }
    }
}

function ConvertTo-FirewallProfileTriStateValue {
    param([AllowNull()] [object]$Value)

    switch ([string]$Value) {
        'True' { 'True' }
        'False' { 'False' }
        'NotConfigured' { 'NotConfigured' }
        '1' { 'True' }
        '0' { 'False' }
        default {
            if ($Value -is [bool]) {
                if ($Value) { 'True' } else { 'False' }
            } else {
                [string]$Value
            }
        }
    }
}

function New-FirewallProfileCaptureIssue {
    param(
        [string]$ProfileName,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        Profile = $ProfileName
        Message = $Message
    }
}

function Normalize-FirewallProfileEntry {
    param([AllowNull()] [object]$ProfileState)

    if ($null -eq $ProfileState) {
        return $null
    }

    [PSCustomObject]@{
        Profile                         = if ($ProfileState.PSObject.Properties['Profile']) { [string]$ProfileState.Profile } elseif ($ProfileState.PSObject.Properties['Name']) { [string]$ProfileState.Name } else { $null }
        Enabled                         = if ($ProfileState.PSObject.Properties['Enabled']) { ConvertTo-FirewallProfileEnabledValue -Value $ProfileState.Enabled } else { $null }
        DefaultInboundAction            = if ($ProfileState.PSObject.Properties['DefaultInboundAction']) { ConvertTo-FirewallProfileActionValue -Value $ProfileState.DefaultInboundAction } else { $null }
        DefaultOutboundAction           = if ($ProfileState.PSObject.Properties['DefaultOutboundAction']) { ConvertTo-FirewallProfileActionValue -Value $ProfileState.DefaultOutboundAction } else { $null }
        AllowUnicastResponseToMulticast = if ($ProfileState.PSObject.Properties['AllowUnicastResponseToMulticast']) { ConvertTo-FirewallProfileTriStateValue -Value $ProfileState.AllowUnicastResponseToMulticast } else { $null }
        NotifyOnListen                  = if ($ProfileState.PSObject.Properties['NotifyOnListen']) { ConvertTo-FirewallProfileTriStateValue -Value $ProfileState.NotifyOnListen } else { $null }
        LogAllowed                      = if ($ProfileState.PSObject.Properties['LogAllowed']) { ConvertTo-FirewallProfileTriStateValue -Value $ProfileState.LogAllowed } else { $null }
        LogBlocked                      = if ($ProfileState.PSObject.Properties['LogBlocked']) { ConvertTo-FirewallProfileTriStateValue -Value $ProfileState.LogBlocked } else { $null }
        LogIgnored                      = if ($ProfileState.PSObject.Properties['LogIgnored']) { ConvertTo-FirewallProfileTriStateValue -Value $ProfileState.LogIgnored } else { $null }
        LogMaxSizeKilobytes             = if ($ProfileState.PSObject.Properties['LogMaxSizeKilobytes'] -and $null -ne $ProfileState.LogMaxSizeKilobytes) { [uint64]$ProfileState.LogMaxSizeKilobytes } else { $null }
        LogFileName                     = if ($ProfileState.PSObject.Properties['LogFileName'] -and -not [string]::IsNullOrWhiteSpace([string]$ProfileState.LogFileName)) { [string]$ProfileState.LogFileName } else { $null }
    }
}

function Normalize-FirewallProfileState {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            CaptureIssues    = @()
            Profiles         = @()
        }
    }

    $profiles = @(if ($State.PSObject.Properties['Profiles']) { $State.Profiles } else { $State })
    $captureIssues = @()
    if ($State.PSObject.Properties['CaptureIssues']) {
        $captureIssues = @(
            foreach ($issue in @($State.CaptureIssues)) {
                if ($null -eq $issue) {
                    continue
                }

                [PSCustomObject]@{
                    Profile = if ($issue.PSObject.Properties['Profile']) { [string]$issue.Profile } else { $null }
                    Message = if ($issue.PSObject.Properties['Message']) { [string]$issue.Message } else { $null }
                }
            }
        )
    }

    [PSCustomObject]@{
        CommandAvailable = if ($State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
        CaptureIssues    = @($captureIssues)
        Profiles         = @(
            foreach ($firewallProfile in @($profiles)) {
                Normalize-FirewallProfileEntry -ProfileState $firewallProfile
            }
        )
    }
}

function Test-FirewallProfileStateHasExtendedFields {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $false
    }

    if ($State.PSObject.Properties['CommandAvailable'] -or $State.PSObject.Properties['CaptureIssues']) {
        return $true
    }

    $profiles = @(if ($State.PSObject.Properties['Profiles']) { $State.Profiles } else { $State })
    foreach ($firewallProfile in @($profiles)) {
        if ($null -eq $firewallProfile) {
            continue
        }

        foreach ($propertyName in @('DefaultInboundAction', 'DefaultOutboundAction', 'AllowUnicastResponseToMulticast', 'NotifyOnListen', 'LogAllowed', 'LogBlocked', 'LogIgnored', 'LogMaxSizeKilobytes', 'LogFileName')) {
            if ($firewallProfile.PSObject.Properties[$propertyName]) {
                return $true
            }
        }
    }

    $false
}

function Test-FirewallProfileStateCapturedExactly {
    param([AllowNull()] [object]$State)

    $normalized = Normalize-FirewallProfileState -State $State
    $normalized.CommandAvailable -and (@($normalized.CaptureIssues).Count -eq 0) -and (@($normalized.Profiles).Count -gt 0)
}

function Get-FirewallProfileStates {
    param([string[]]$Profiles)

    if (-not (Test-CommandAvailable -Name 'Get-NetFirewallProfile')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            CaptureIssues    = @()
            Profiles         = @()
        }
    }

    $requestedProfiles = @($Profiles | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
    $profilesByName = @{}
    foreach ($firewallProfile in @(Get-NetFirewallProfile -Profile $requestedProfiles -ErrorAction SilentlyContinue)) {
        $profileName = if ($firewallProfile.PSObject.Properties['Name']) { [string]$firewallProfile.Name } elseif ($firewallProfile.PSObject.Properties['Profile']) { [string]$firewallProfile.Profile } else { $null }
        if (-not [string]::IsNullOrWhiteSpace($profileName)) {
            $profilesByName[$profileName] = $firewallProfile
        }
    }

    $captureIssues = [System.Collections.Generic.List[object]]::new()
    $states = foreach ($profileName in $requestedProfiles) {
        $firewallProfile = if ($profilesByName.ContainsKey($profileName)) { $profilesByName[$profileName] } else { $null }
        if ($null -eq $firewallProfile) {
            $captureIssues.Add((New-FirewallProfileCaptureIssue -ProfileName $profileName -Message 'Get-NetFirewallProfile did not return this profile.')) | Out-Null
            continue
        }

        Normalize-FirewallProfileEntry -ProfileState $firewallProfile
    }

    [PSCustomObject]@{
        CommandAvailable = $true
        CaptureIssues    = @($captureIssues)
        Profiles         = @($states)
    }
}

function Set-FirewallProfileStateExact {
    param([Parameter(Mandatory)] [object]$ProfileState)

    if (-not (Test-CommandAvailable -Name 'Set-NetFirewallProfile')) {
        return
    }

    $normalizedProfile = Normalize-FirewallProfileEntry -ProfileState $ProfileState
    $profileName = [string]$normalizedProfile.Profile
    if ([string]::IsNullOrWhiteSpace($profileName)) {
        return
    }

    $params = @{
        Profile = $profileName
    }

    if ($null -ne $normalizedProfile.Enabled) {
        $params['Enabled'] = $normalizedProfile.Enabled
    }

    foreach ($propertyName in @('DefaultInboundAction', 'DefaultOutboundAction', 'AllowUnicastResponseToMulticast', 'NotifyOnListen', 'LogAllowed', 'LogBlocked', 'LogIgnored')) {
        $value = $normalizedProfile.$propertyName
        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
            $params[$propertyName] = [string]$value
        }
    }

    if ($null -ne $normalizedProfile.LogMaxSizeKilobytes) {
        $params['LogMaxSizeKilobytes'] = [uint64]$normalizedProfile.LogMaxSizeKilobytes
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$normalizedProfile.LogFileName)) {
        $params['LogFileName'] = [string]$normalizedProfile.LogFileName
    }

    Set-NetFirewallProfile @params
}

function Set-Permissive-FirewallProfiles {
    param([Parameter(Mandatory)] [object]$Definition)

    if (-not (Test-CommandAvailable -Name 'Set-NetFirewallProfile')) {
        return
    }

    $profileNames = @($Definition.Profiles)
    if ($profileNames.Count -eq 0) {
        return
    }

    Set-NetFirewallProfile -Profile $profileNames `
        -Enabled (ConvertTo-FirewallProfileEnabledValue -Value $Definition.PermissiveValue) `
        -DefaultInboundAction $Definition.PermissiveDefaultInboundAction `
        -DefaultOutboundAction $Definition.PermissiveDefaultOutboundAction `
        -AllowUnicastResponseToMulticast (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveAllowUnicastResponseToMulticast) `
        -NotifyOnListen (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveNotifyOnListen) `
        -LogAllowed (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveLogAllowed) `
        -LogBlocked (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveLogBlocked) `
        -LogIgnored (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveLogIgnored)
}

function Restore-FirewallProfiles {
    param([AllowNull()] [object]$State)

    if (-not (Test-FirewallProfileStateCapturedExactly -State $State)) {
        return
    }

    foreach ($firewallProfile in @((Normalize-FirewallProfileState -State $State).Profiles)) {
        Set-FirewallProfileStateExact -ProfileState $firewallProfile
    }
}

function ConvertTo-ComparableFirewallProfileState {
    param(
        [AllowNull()] [object]$State,
        [AllowNull()] [object]$ReferenceState
    )

    $normalizedState = Normalize-FirewallProfileState -State $State
    $comparisonReference = if ($null -ne $ReferenceState) { $ReferenceState } else { $State }
    $useExtendedFields = Test-FirewallProfileStateHasExtendedFields -State $comparisonReference

    if (-not $useExtendedFields) {
        return @(
            foreach ($firewallProfile in @($normalizedState.Profiles | Sort-Object -Property Profile)) {
                [PSCustomObject]@{
                    Profile = [string]$firewallProfile.Profile
                    Enabled = if ($null -ne $firewallProfile.Enabled) { [string]$firewallProfile.Enabled } else { $null }
                }
            }
        )
    }

    [ordered]@{
        CommandAvailable = $normalizedState.CommandAvailable
        CaptureIssues    = @(
            foreach ($issue in @($normalizedState.CaptureIssues | Sort-Object -Property Profile, Message)) {
                [PSCustomObject]@{
                    Profile = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Profile)) { [string]$issue.Profile } else { $null }
                    Message = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Message)) { [string]$issue.Message } else { $null }
                }
            }
        )
        Profiles         = @(
            foreach ($firewallProfile in @($normalizedState.Profiles | Sort-Object -Property Profile)) {
                [PSCustomObject]@{
                    Profile                         = [string]$firewallProfile.Profile
                    Enabled                         = if ($null -ne $firewallProfile.Enabled) { [string]$firewallProfile.Enabled } else { $null }
                    DefaultInboundAction            = if ($null -ne $firewallProfile.DefaultInboundAction) { [string]$firewallProfile.DefaultInboundAction } else { $null }
                    DefaultOutboundAction           = if ($null -ne $firewallProfile.DefaultOutboundAction) { [string]$firewallProfile.DefaultOutboundAction } else { $null }
                    AllowUnicastResponseToMulticast = if ($null -ne $firewallProfile.AllowUnicastResponseToMulticast) { [string]$firewallProfile.AllowUnicastResponseToMulticast } else { $null }
                    NotifyOnListen                  = if ($null -ne $firewallProfile.NotifyOnListen) { [string]$firewallProfile.NotifyOnListen } else { $null }
                    LogAllowed                      = if ($null -ne $firewallProfile.LogAllowed) { [string]$firewallProfile.LogAllowed } else { $null }
                    LogBlocked                      = if ($null -ne $firewallProfile.LogBlocked) { [string]$firewallProfile.LogBlocked } else { $null }
                    LogIgnored                      = if ($null -ne $firewallProfile.LogIgnored) { [string]$firewallProfile.LogIgnored } else { $null }
                    LogMaxSizeKilobytes             = $firewallProfile.LogMaxSizeKilobytes
                    LogFileName                     = if ($null -ne $firewallProfile.LogFileName) { [string]$firewallProfile.LogFileName } else { $null }
                }
            }
        )
    }
}

function New-FirewallRuleCaptureIssue {
    param(
        [string]$Group,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        Group   = $Group
        Message = $Message
    }
}

function Normalize-FirewallRuleEntry {
    param([AllowNull()] [object]$Rule)

    if ($null -eq $Rule) {
        return $null
    }

    [PSCustomObject]@{
        Name        = if ($Rule.PSObject.Properties['Name']) { [string]$Rule.Name } else { $null }
        DisplayName = if ($Rule.PSObject.Properties['DisplayName']) { [string]$Rule.DisplayName } else { $null }
        Group       = if ($Rule.PSObject.Properties['Group']) { [string]$Rule.Group } elseif ($Rule.PSObject.Properties['RuleGroup']) { [string]$Rule.RuleGroup } else { $null }
        Enabled     = if ($Rule.PSObject.Properties['Enabled']) { ConvertTo-FirewallProfileEnabledValue -Value $Rule.Enabled } else { $null }
        Direction   = if ($Rule.PSObject.Properties['Direction']) { [string]$Rule.Direction } else { $null }
        Action      = if ($Rule.PSObject.Properties['Action']) { [string]$Rule.Action } else { $null }
        Profile     = if ($Rule.PSObject.Properties['Profile']) { [string]$Rule.Profile } else { $null }
    }
}

function Normalize-FirewallRuleState {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Group            = $null
            CaptureIssues    = @()
            Rules            = @()
        }
    }

    $rules = @(if ($State.PSObject.Properties['Rules']) { $State.Rules } else { $State })
    $captureIssues = @()
    if ($State.PSObject.Properties['CaptureIssues']) {
        $captureIssues = @(
            foreach ($issue in @($State.CaptureIssues)) {
                if ($null -eq $issue) {
                    continue
                }

                [PSCustomObject]@{
                    Group   = if ($issue.PSObject.Properties['Group']) { [string]$issue.Group } else { $null }
                    Message = if ($issue.PSObject.Properties['Message']) { [string]$issue.Message } else { $null }
                }
            }
        )
    }

    [PSCustomObject]@{
        CommandAvailable = if ($State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
        Group            = if ($State.PSObject.Properties['Group']) { [string]$State.Group } else { $null }
        CaptureIssues    = @($captureIssues)
        Rules            = @(
            foreach ($rule in @($rules)) {
                Normalize-FirewallRuleEntry -Rule $rule
            }
        )
    }
}

function Test-FirewallRuleStateCapturedExactly {
    param([AllowNull()] [object]$State)

    $normalized = Normalize-FirewallRuleState -State $State
    if (-not $normalized.CommandAvailable -or @($normalized.CaptureIssues).Count -ne 0) {
        return $false
    }

    $seenNames = @{}
    foreach ($rule in @($normalized.Rules)) {
        if ($null -eq $rule -or [string]::IsNullOrWhiteSpace([string]$rule.Name)) {
            return $false
        }

        $enabled = [string](ConvertTo-FirewallProfileEnabledValue -Value $rule.Enabled)
        if ($enabled -notin @('True', 'False') -or $seenNames.ContainsKey([string]$rule.Name)) {
            return $false
        }
        $seenNames[[string]$rule.Name] = $true
    }

    return $true
}

function Get-FirewallRuleGroupState {
    param([Parameter(Mandatory)] [string]$Group)

    if (-not (Test-CommandAvailable -Name 'Get-NetFirewallRule')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Group            = $Group
            CaptureIssues    = @()
            Rules            = @()
        }
    }

    try {
        $rules = @(
            foreach ($rule in @(Get-NetFirewallRule -Group $Group -ErrorAction Stop)) {
                Normalize-FirewallRuleEntry -Rule ([PSCustomObject]@{
                    Name        = $rule.Name
                    DisplayName = $rule.DisplayName
                    Group       = $rule.Group
                    Enabled     = $rule.Enabled
                    Direction   = $rule.Direction
                    Action      = $rule.Action
                    Profile     = $rule.Profile
                })
            }
        )

        return [PSCustomObject]@{
            CommandAvailable = $true
            Group            = $Group
            CaptureIssues    = @()
            Rules            = @($rules)
        }
    } catch {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Group            = $Group
            CaptureIssues    = @(
                New-FirewallRuleCaptureIssue -Group $Group -Message $_.Exception.Message
            )
            Rules            = @()
        }
    }
}

function Set-Permissive-FirewallRules {
    param([Parameter(Mandatory)] [object]$Definition)

    if (-not (Test-CommandAvailable -Name 'Set-NetFirewallRule')) {
        return
    }

    Set-NetFirewallRule -Group $Definition.Group -Enabled (ConvertTo-FirewallProfileEnabledValue -Value $Definition.PermissiveEnabled) -ErrorAction Stop | Out-Null
}

function Restore-FirewallRules {
    param([AllowNull()] [object]$State)

    if (-not (Test-FirewallRuleStateCapturedExactly -State $State)) {
        return
    }

    if (-not (Test-CommandAvailable -Name 'Set-NetFirewallRule')) {
        return
    }

    $rules = @((Normalize-FirewallRuleState -State $State).Rules)
    foreach ($enabled in @('True', 'False')) {
        $names = @(
            $rules |
                Where-Object { [string](ConvertTo-FirewallProfileEnabledValue -Value $_.Enabled) -eq $enabled } |
                ForEach-Object { [string]$_.Name }
        )
        if ($names.Count -gt 0) {
            Set-NetFirewallRule -Name $names -Enabled $enabled -ErrorAction Stop | Out-Null
        }
    }
}

function ConvertTo-ComparableFirewallRuleState {
    param([AllowNull()] [object]$State)

    $normalizedState = Normalize-FirewallRuleState -State $State
    [ordered]@{
        CommandAvailable = $normalizedState.CommandAvailable
        Group            = if (-not [string]::IsNullOrWhiteSpace([string]$normalizedState.Group)) { [string]$normalizedState.Group } else { $null }
        CaptureIssues    = @(
            foreach ($issue in @($normalizedState.CaptureIssues | Sort-Object -Property Group, Message)) {
                [PSCustomObject]@{
                    Group   = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Group)) { [string]$issue.Group } else { $null }
                    Message = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Message)) { [string]$issue.Message } else { $null }
                }
            }
        )
        Rules            = @(
            foreach ($rule in @($normalizedState.Rules | Sort-Object -Property Name, DisplayName, Direction, Action, Profile)) {
                [PSCustomObject]@{
                    Name        = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Name)) { [string]$rule.Name } else { $null }
                    DisplayName = if (-not [string]::IsNullOrWhiteSpace([string]$rule.DisplayName)) { [string]$rule.DisplayName } else { $null }
                    Group       = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Group)) { [string]$rule.Group } else { $null }
                    Enabled     = if ($null -ne $rule.Enabled) { [string]$rule.Enabled } else { $null }
                    Direction   = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Direction)) { [string]$rule.Direction } else { $null }
                    Action      = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Action)) { [string]$rule.Action } else { $null }
                    Profile     = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Profile)) { [string]$rule.Profile } else { $null }
                }
            }
        )
    }
}

#endregion

#region Canonicalization, reporting, and verification

function ConvertTo-CanonicalValue {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if (
        $Value -is [string] -or
        $Value -is [char] -or
        $Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [sbyte] -or
        $Value -is [int16] -or
        $Value -is [uint16] -or
        $Value -is [int32] -or
        $Value -is [uint32] -or
        $Value -is [int64] -or
        $Value -is [uint64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]
    ) {
        return $Value
    }

    if ($Value -is [datetime] -or $Value -is [guid] -or $Value -is [version]) {
        return [string]$Value
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $ordered = [ordered]@{}
        foreach ($key in @($Value.Keys | ForEach-Object { [string]$_ } | Sort-Object)) {
            $ordered[$key] = ConvertTo-CanonicalValue -Value $Value[$key]
        }

        return [PSCustomObject]$ordered
    }

    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @(
            foreach ($item in @($Value)) {
                ConvertTo-CanonicalValue -Value $item
            }
        )

        $sortedItems = @(
            foreach ($item in $items) {
                [PSCustomObject]@{
                    SortKey = ConvertTo-Json -InputObject $item -Depth 12 -Compress
                    Value   = $item
                }
            }
        ) | Sort-Object -Property SortKey

        return @($sortedItems | ForEach-Object { $_.Value })
    }

    $properties = @($Value.PSObject.Properties | Where-Object { $_.MemberType -match 'Property' } | Sort-Object -Property Name)
    if ($properties.Count -eq 0) {
        return [string]$Value
    }

    $ordered = [ordered]@{}
    foreach ($property in $properties) {
        $ordered[$property.Name] = ConvertTo-CanonicalValue -Value $property.Value
    }

    [PSCustomObject]$ordered
}

function Get-CanonicalJson {
    param([AllowNull()] [object]$Value)

    ConvertTo-Json -InputObject (ConvertTo-CanonicalValue -Value $Value) -Depth 12 -Compress
}

function ConvertTo-DisplayString {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return '<null>'
    }

    if ($Value -is [bool]) {
        return $Value.ToString()
    }

    if (
        $Value -is [string] -or
        $Value -is [char] -or
        $Value -is [byte] -or
        $Value -is [sbyte] -or
        $Value -is [int16] -or
        $Value -is [uint16] -or
        $Value -is [int32] -or
        $Value -is [uint32] -or
        $Value -is [int64] -or
        $Value -is [uint64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]
    ) {
        return [string]$Value
    }

    ConvertTo-Json -InputObject (ConvertTo-CanonicalValue -Value $Value) -Depth 12 -Compress
}

function Add-ReportKeyValueLine {
    param(
        [Parameter(Mandatory)] [System.Collections.Generic.List[string]]$Lines,
        [Parameter(Mandatory)] [string]$Label,
        [AllowNull()] [object]$Value,
        [int]$Indent = 2
    )

    $Lines.Add(('{0}{1}: {2}' -f (' ' * $Indent), $Label, (ConvertTo-DisplayString -Value $Value)))
}

function Add-ReportJsonBlock {
    param(
        [Parameter(Mandatory)] [System.Collections.Generic.List[string]]$Lines,
        [Parameter(Mandatory)] [string]$Label,
        [AllowNull()] [object]$Value,
        [int]$Indent = 2
    )

    $Lines.Add(('{0}{1}:' -f (' ' * $Indent), $Label))

    $json = if ($null -eq $Value) {
        'null'
    } else {
        ConvertTo-Json -InputObject (ConvertTo-CanonicalValue -Value $Value) -Depth 12
    }

    foreach ($line in @($json -split "`r?`n")) {
        $Lines.Add(('{0}{1}' -f (' ' * ($Indent + 2)), $line))
    }
}

function Add-SnapshotEntryReportLines {
    param(
        [Parameter(Mandatory)] [System.Collections.Generic.List[string]]$Lines,
        [Parameter(Mandatory)] [object]$Entry,
        [string]$SnapshotPath
    )

    $Lines.Add("[$($Entry.Id)] $($Entry.Type)")
    Add-ReportKeyValueLine -Lines $Lines -Label 'Requires reboot' -Value $Entry.RequiresReboot

    switch ($Entry.Type) {
        'RegistryValue' {
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Path' -Value $Entry.Path
            Add-ReportKeyValueLine -Lines $Lines -Label 'Name' -Value $Entry.Name
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Exists' -Value $Entry.Exists
            Add-ReportKeyValueLine -Lines $Lines -Label 'Value kind' -Value $Entry.ValueKind
            Add-ReportKeyValueLine -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
            if (-not $captured -and $Entry.PSObject.Properties['CaptureError']) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
        }
        'RegistryKeyFlat' {
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Path' -Value $Entry.Path
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Exists' -Value $Entry.Exists
            if (-not $captured -and $Entry.PSObject.Properties['CaptureError']) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
            Add-ReportJsonBlock -Lines $Lines -Label 'Values' -Value @($Entry.CurrentValue)
        }
        'MpPreferenceValue' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $null -ne $Entry.RestoreValue }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Property' -Value $Entry.Property
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
            Add-ReportKeyValueLine -Lines $Lines -Label 'Restore value' -Value $Entry.RestoreValue
            if (-not ($commandAvailable -and $captured) -and $Entry.PSObject.Properties['CaptureError']) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
        }
        'MpPreferenceList' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            $items = @(Normalize-MpPreferenceListItems -Value $Entry.CurrentValue)

            Add-ReportKeyValueLine -Lines $Lines -Label 'Property' -Value $Entry.Property
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Item count' -Value @($items).Count
            if ($Entry.PSObject.Properties['CaptureError'] -and -not [string]::IsNullOrWhiteSpace([string]$Entry.CaptureError)) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
            Add-ReportJsonBlock -Lines $Lines -Label 'Items' -Value @($items)
        }
        'DefenderRuntimeStatus' {
            $state = if ($null -ne $Entry.CurrentValue) { $Entry.CurrentValue } else { [PSCustomObject]@{} }
            $commandAvailable = if ($state.PSObject.Properties['CommandAvailable']) { [bool]$state.CommandAvailable } else { $true }
            $captured = if ($state.PSObject.Properties['Captured']) { [bool]$state.Captured } else { $true }
            $amRunningMode = if ($state.PSObject.Properties['AMRunningMode']) { [string]$state.AMRunningMode } else { $null }
            $realTimeProtectionEnabled = if ($state.PSObject.Properties['RealTimeProtectionEnabled']) { ConvertTo-NullableBoolean -Value $state.RealTimeProtectionEnabled } else { $null }
            $antivirusEnabled = if ($state.PSObject.Properties['AntivirusEnabled']) { ConvertTo-NullableBoolean -Value $state.AntivirusEnabled } else { $null }
            $isTamperProtected = if ($state.PSObject.Properties['IsTamperProtected']) { ConvertTo-NullableBoolean -Value $state.IsTamperProtected } else { $null }

            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Running mode' -Value $amRunningMode
            Add-ReportKeyValueLine -Lines $Lines -Label 'Real-time protection enabled' -Value $realTimeProtectionEnabled
            Add-ReportKeyValueLine -Lines $Lines -Label 'Antivirus enabled' -Value $antivirusEnabled
            Add-ReportKeyValueLine -Lines $Lines -Label 'Tamper protection' -Value $isTamperProtected
            if ($state.PSObject.Properties['Error'] -and -not [string]::IsNullOrWhiteSpace([string]$state.Error)) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $state.Error
            }
            if ($captured -and ($isTamperProtected -eq $true -or $amRunningMode -match 'Passive')) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'Tamper protection can cause local Defender preference changes to be ignored or later reverted, and passive mode means Defender Antivirus is not the primary enforcement engine. Interpret Defender restore mismatches in that context.'
            }
        }
        'BitLockerVolumes' {
            $state = Normalize-BitLockerState -State $Entry.CurrentValue
            $timedOutMountPoints = @($state.TimedOutMountPoints)
            $captureIssues = @($state.CaptureIssues)
            $capturedVolumes = @($state.Volumes)
            $baselineComplete = Test-BitLockerStateCapturedExactly -State $state

            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Timed out mount point count' -Value @($timedOutMountPoints).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value @($captureIssues).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($baselineComplete) { 'Complete' } else { 'Partial / incomplete' })
            if (-not $baselineComplete) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'BitLocker round-trip requires an exact mounted-volume baseline. Permissive, restore, and verification skip this setting if mount points time out or richer BitLocker fields cannot be captured exactly.'
            }
            if (@($timedOutMountPoints).Count -gt 0) {
                foreach ($mountPoint in $timedOutMountPoints) {
                    $Lines.Add(('  - Timed out mount point: {0}' -f $mountPoint))
                }
            }
            foreach ($issue in @($captureIssues | Sort-Object -Property MountPoint, Message)) {
                $issueMountPoint = if (-not [string]::IsNullOrWhiteSpace([string]$issue.MountPoint)) { [string]$issue.MountPoint } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | MountPoint={0} | {1}' -f $issueMountPoint, $issue.Message))
            }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Captured volume count' -Value @($capturedVolumes).Count
            foreach ($volume in @($capturedVolumes | Sort-Object -Property MountPoint)) {
                $mountPoint = if (-not [string]::IsNullOrWhiteSpace([string]$volume.MountPoint)) { [string]$volume.MountPoint } else { '<unknown>' }
                $protectionMode = if (-not [string]::IsNullOrWhiteSpace([string]$volume.ProtectionMode)) { [string]$volume.ProtectionMode } else { '<unknown>' }
                $volumeStatus = if (-not [string]::IsNullOrWhiteSpace([string]$volume.VolumeStatus)) { [string]$volume.VolumeStatus } else { '<unknown>' }
                $lockStatus = if (-not [string]::IsNullOrWhiteSpace([string]$volume.LockStatus)) { [string]$volume.LockStatus } else { '<unknown>' }
                $encryptionMethod = if (-not [string]::IsNullOrWhiteSpace([string]$volume.EncryptionMethod)) { [string]$volume.EncryptionMethod } else { '<unknown>' }
                $encryptionSummary = if ($null -ne $volume.EncryptionPercentage) {
                    ('{0}% ({1})' -f ([int]$volume.EncryptionPercentage), $encryptionMethod)
                } else {
                    $encryptionMethod
                }
                $protectorCount = if ($null -ne $volume.KeyProtectorCount) { [int]$volume.KeyProtectorCount } else { '<unknown>' }
                $protectorSummary = Get-BitLockerProtectorTypeSummary -Volume $volume
                $autoUnlockSummary = if (Test-BitLockerAutoUnlockSupportedVolume -Volume $volume) {
                    if ($null -ne $volume.AutoUnlockEnabled) {
                        ConvertTo-DisplayString -Value $volume.AutoUnlockEnabled
                    } else {
                        '<unknown>'
                    }
                } else {
                    'n/a'
                }

                $Lines.Add(('  - {0} | Mode={1} | Protection={2} | Status={3} | Lock={4} | Encryption={5} | Protectors={6} [{7}] | AutoUnlock={8}' -f $mountPoint, $protectionMode, $volume.ProtectionStatus, $volumeStatus, $lockStatus, $encryptionSummary, $protectorCount, $protectorSummary, $autoUnlockSummary))
            }
        }
        'ExploitProtectionPolicy' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $Entry.CurrentValue.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'XML captured' -Value (-not [string]::IsNullOrWhiteSpace((Get-ExploitProtectionPolicyXml -State $Entry.CurrentValue -SnapshotPath $SnapshotPath)))
        }
        'WdacPolicies' {
            $captured = Test-WdacPolicyStateCapturedExactly -State $Entry.CurrentValue
            $captureIssues = @(
                if ($Entry.CurrentValue.PSObject.Properties['CaptureIssues']) {
                    $Entry.CurrentValue.CaptureIssues
                }
            )
            Add-ReportKeyValueLine -Lines $Lines -Label 'CiTool available' -Value $Entry.CurrentValue.CiToolAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value $captureIssues.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Policy count' -Value @($Entry.CurrentValue.Policies).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Policy file count' -Value @($Entry.CurrentValue.Files).Count
            foreach ($issue in $captureIssues) {
                $Lines.Add(('  - Capture issue | {0}' -f $issue))
            }
            $policyRows = @(Get-WdacPolicyReportRows -State $Entry.CurrentValue)
            $platformManagedPolicies = @($policyRows | Where-Object { $_.IsPlatformManaged })
            $activePolicies = @($policyRows | Where-Object { $_.IsEnforced -eq $true })
            $onDiskOnlyPolicies = @($policyRows | Where-Object { $_.HasFileOnDisk -eq $true -and $_.IsEnforced -eq $false })
            $activeOnlyPolicies = @($policyRows | Where-Object { $_.HasFileOnDisk -eq $false -and $_.IsEnforced -eq $true })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Platform-managed policy count' -Value @($platformManagedPolicies).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Active policy count' -Value @($activePolicies).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'On-disk only / pending reboot count' -Value @($onDiskOnlyPolicies).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Active-only policy count' -Value @($activeOnlyPolicies).Count
            if (@($platformManagedPolicies).Count -gt 0 -or @($onDiskOnlyPolicies).Count -gt 0) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'WDAC reporting distinguishes active enforcement from on-disk file presence. On-disk-only policies usually indicate pending-reboot state, and platform-managed policies can require file-based handling or a reboot before live WDAC state fully matches the snapshot.'
            }
            foreach ($policyRow in $policyRows) {
                if ($policyRow.IsPlatformManaged) {
                    $Lines.Add(('  - {0} | {1} | {2} | Enforced={3} | FileOnDisk={4} | {5} [platform-managed]' -f $policyRow.PolicyId, $policyRow.FriendlyName, $policyRow.Presence, (ConvertTo-DisplayString -Value $policyRow.IsEnforced), (ConvertTo-DisplayString -Value $policyRow.HasFileOnDisk), $policyRow.Classification))
                } else {
                    $Lines.Add(('  - {0} | {1} | {2} | Enforced={3} | FileOnDisk={4}' -f $policyRow.PolicyId, $policyRow.FriendlyName, $policyRow.Presence, (ConvertTo-DisplayString -Value $policyRow.IsEnforced), (ConvertTo-DisplayString -Value $policyRow.HasFileOnDisk)))
                }
            }
        }
        'AsrRules' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Configured rule count' -Value @($Entry.CurrentValue).Count
            $invalidEntries = @(Get-AsrInvalidEntriesFromEntry -Entry $Entry)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Invalid capture entry count' -Value $invalidEntries.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured -and $invalidEntries.Count -eq 0) { 'Complete' } else { 'Partial / incomplete' })
            if (-not ($commandAvailable -and $captured) -and $Entry.PSObject.Properties['CaptureError']) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
            if (-not ($commandAvailable -and $captured) -or $invalidEntries.Count -gt 0) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'This snapshot did not capture a reliable ASR baseline. Permissive, restore, and verification will skip ASR changes to avoid unsafe round-trip behavior.'
            }
            foreach ($rule in @($Entry.CurrentValue)) {
                $Lines.Add(('  - {0} | {1} | {2}' -f $rule.Id, $rule.Name, $rule.ActionLabel))
            }
            foreach ($rule in $invalidEntries) {
                $displayId = if ([string]::IsNullOrWhiteSpace([string]$rule.Id)) { '<blank>' } else { [string]$rule.Id }
                $displayAction = if ([string]::IsNullOrWhiteSpace([string]$rule.Action)) {
                    '<blank>'
                } elseif ([string]::IsNullOrWhiteSpace([string]$rule.ActionLabel)) {
                    [string]$rule.Action
                } else {
                    [string]$rule.ActionLabel
                }
                $Lines.Add(('  - Incomplete capture entry | RuleId={0} | Action={1}' -f $displayId, $displayAction))
            }
        }
        'PowerShellModuleLogging' {
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Base path' -Value $Entry.BasePath
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Exists' -Value $Entry.Exists
            if (-not $captured -and $Entry.PSObject.Properties['CaptureError']) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
            Add-ReportJsonBlock -Lines $Lines -Label 'Base values' -Value @($Entry.CurrentValue.BaseValues)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Module names key exists' -Value $Entry.CurrentValue.ModuleNamesExists
            Add-ReportJsonBlock -Lines $Lines -Label 'Module names values' -Value @($Entry.CurrentValue.ModuleNamesValues)
        }
        'AppLockerPolicy' {
            $state = $Entry.CurrentValue
            $commandAvailable = if ($state.PSObject.Properties['CommandAvailable']) { [bool]$state.CommandAvailable } else { $true }
            $localCaptured = if ($state.PSObject.Properties['LocalCaptured']) { [bool]$state.LocalCaptured } else { $false }
            $effectiveCaptured = if ($state.PSObject.Properties['EffectiveCaptured']) { [bool]$state.EffectiveCaptured } else { $false }
            $localMatchesEffective = if ($state.PSObject.Properties['LocalMatchesEffective']) { [bool]$state.LocalMatchesEffective } else { $false }
            $captureIssues = @(if ($state.PSObject.Properties['CaptureIssues']) { $state.CaptureIssues })
            $collectionSummaries = @(if ($state.PSObject.Properties['CollectionSummaries']) {
                $state.CollectionSummaries
            } else {
                Get-AppLockerCollectionSummaries -Xml (Get-AppLockerPolicyXml -State $state -PolicyScope Effective -SnapshotPath $SnapshotPath)
            })
            $baselineComplete = Test-AppLockerPolicyCapturedExactly -State $state -SnapshotPath $SnapshotPath

            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Local policy captured' -Value $localCaptured
            Add-ReportKeyValueLine -Lines $Lines -Label 'Effective policy captured' -Value $effectiveCaptured
            Add-ReportKeyValueLine -Lines $Lines -Label 'Local matches effective' -Value $localMatchesEffective
            Add-ReportKeyValueLine -Lines $Lines -Label 'Collection count' -Value @($collectionSummaries).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value @($captureIssues).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($baselineComplete) { 'Complete' } else { 'Partial / incomplete' })
            if (-not $baselineComplete) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'AppLocker round-trip is only treated as exact when the local and effective AppLocker policies match. If Group Policy or another higher-precedence source changes the effective policy, permissive, restore, and verification skip this entry rather than guessing.'
            }

            foreach ($issue in @($captureIssues | Sort-Object -Property Scope, Message)) {
                $scope = if ($issue.PSObject.Properties['Scope'] -and -not [string]::IsNullOrWhiteSpace([string]$issue.Scope)) { [string]$issue.Scope } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | Scope={0} | {1}' -f $scope, $issue.Message))
            }

            foreach ($collection in @($collectionSummaries | Sort-Object -Property Type)) {
                $Lines.Add(('  - {0} | Enforcement={1} | Rules={2} | Services={3} | SystemApps={4}' -f (ConvertTo-DisplayString -Value $collection.Type), (ConvertTo-DisplayString -Value $collection.EnforcementMode), (ConvertTo-DisplayString -Value $collection.RuleCount), (ConvertTo-DisplayString -Value $collection.ServicesEnforcement), (ConvertTo-DisplayString -Value $collection.SystemAppsAllow)))
            }
        }
        'FirewallProfiles' {
            $state = Normalize-FirewallProfileState -State $Entry.CurrentValue
            $baselineComplete = Test-FirewallProfileStateCapturedExactly -State $state
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Captured profile count' -Value (@($state.Profiles).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value (@($state.CaptureIssues).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($baselineComplete) { 'Complete' } else { 'Partial / incomplete' })
            if (-not $baselineComplete) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'Firewall round-trip requires an exact per-profile baseline. Permissive, restore, and verification skip this setting if a profile cannot be captured exactly.'
            }

            foreach ($issue in @($state.CaptureIssues | Sort-Object -Property Profile, Message)) {
                $issueProfile = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Profile)) { [string]$issue.Profile } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | Profile={0} | {1}' -f $issueProfile, $issue.Message))
            }

            foreach ($firewallProfile in @($state.Profiles | Sort-Object -Property Profile)) {
                $Lines.Add(('  - {0} | Enabled={1} | Inbound={2} | Outbound={3} | Notify={4} | Unicast={5} | LogAllowed={6} | LogBlocked={7} | LogIgnored={8} | LogMaxKB={9} | LogFile={10}' -f $firewallProfile.Profile, (ConvertTo-DisplayString -Value $firewallProfile.Enabled), (ConvertTo-DisplayString -Value $firewallProfile.DefaultInboundAction), (ConvertTo-DisplayString -Value $firewallProfile.DefaultOutboundAction), (ConvertTo-DisplayString -Value $firewallProfile.NotifyOnListen), (ConvertTo-DisplayString -Value $firewallProfile.AllowUnicastResponseToMulticast), (ConvertTo-DisplayString -Value $firewallProfile.LogAllowed), (ConvertTo-DisplayString -Value $firewallProfile.LogBlocked), (ConvertTo-DisplayString -Value $firewallProfile.LogIgnored), (ConvertTo-DisplayString -Value $firewallProfile.LogMaxSizeKilobytes), (ConvertTo-DisplayString -Value $firewallProfile.LogFileName)))
            }
        }
        'FirewallRules' {
            $state = Normalize-FirewallRuleState -State $Entry.CurrentValue
            $baselineComplete = Test-FirewallRuleStateCapturedExactly -State $state
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Group' -Value $state.Group
            Add-ReportKeyValueLine -Lines $Lines -Label 'Rule count' -Value (@($state.Rules).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value (@($state.CaptureIssues).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($baselineComplete) { 'Complete' } else { 'Partial / incomplete' })
            if (-not $baselineComplete) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'Firewall rule group round-trip requires an exact baseline. Permissive, restore, and verification skip this setting if the rule group cannot be captured exactly.'
            }

            foreach ($issue in @($state.CaptureIssues | Sort-Object -Property Group, Message)) {
                $issueGroup = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Group)) { [string]$issue.Group } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | Group={0} | {1}' -f $issueGroup, $issue.Message))
            }

            foreach ($rule in @($state.Rules | Sort-Object -Property Name, DisplayName)) {
                $displayName = if (-not [string]::IsNullOrWhiteSpace([string]$rule.DisplayName)) { [string]$rule.DisplayName } else { '<unnamed>' }
                $ruleName = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Name)) { [string]$rule.Name } else { '<unknown>' }
                $Lines.Add(('  - {0} | Name={1} | Enabled={2} | Direction={3} | Action={4} | Profile={5}' -f $displayName, $ruleName, (ConvertTo-DisplayString -Value $rule.Enabled), (ConvertTo-DisplayString -Value $rule.Direction), (ConvertTo-DisplayString -Value $rule.Action), (ConvertTo-DisplayString -Value $rule.Profile)))
            }
        }
        'NetBiosAdapters' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            if (-not ($commandAvailable -and $captured) -and $Entry.PSObject.Properties['CaptureError']) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
            foreach ($adapter in @($Entry.CurrentValue)) {
                $Lines.Add(('  - Index {0} | {1} | TcpipNetbiosOptions={2}' -f $adapter.Index, $adapter.Description, $adapter.TcpipNetbiosOptions))
            }
        }
        'LoadedUserRegistryValues' {
            $state = Normalize-UserRegistryValueState -State $Entry.CurrentValue
            Add-ReportKeyValueLine -Lines $Lines -Label 'Captured profile count' -Value (@($state.Entries | Group-Object -Property Sid).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value (@($state.CaptureIssues).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if (@($state.CaptureIssues).Count -eq 0) { 'Complete' } else { 'Partial / incomplete' })
            if (@($state.CaptureIssues).Count -gt 0) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'One or more user profile hives could not be accessed. Permissive, restore, and verification will skip this setting to avoid partial round-trip behavior.'
            }

            foreach ($value in @($state.Entries)) {
                $displayValue = if ($value.Exists) { ConvertTo-DisplayString -Value $value.CurrentValue } else { '<absent>' }
                $Lines.Add(('  - {0} | {1} | {2} | {3}' -f $value.Sid, $value.RelativePath, $value.Name, $displayValue))
            }

            foreach ($issue in @($state.CaptureIssues)) {
                $profilePath = if ($null -ne $issue.PSObject.Properties['ProfilePath'] -and -not [string]::IsNullOrWhiteSpace([string]$issue.ProfilePath)) { [string]$issue.ProfilePath } else { '<unknown>' }
                $hivePath = if ($null -ne $issue.PSObject.Properties['HivePath'] -and -not [string]::IsNullOrWhiteSpace([string]$issue.HivePath)) { [string]$issue.HivePath } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | SID={0} | Profile={1} | Hive={2} | {3}' -f $issue.Sid, $profilePath, $hivePath, $issue.Message))
            }
        }
        'MachineEnvironmentValue' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Name' -Value $Entry.Name
            Add-ReportKeyValueLine -Lines $Lines -Label 'Exists' -Value $Entry.Exists
            Add-ReportKeyValueLine -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
        }
        'ServiceConfig' {
            $serviceState = if ($null -ne $Entry.CurrentValue) { $Entry.CurrentValue } else { [PSCustomObject]@{} }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else {
                $null -ne $Entry.CurrentValue -and
                $serviceState.PSObject.Properties['StartMode'] -and
                $serviceState.PSObject.Properties['State'] -and
                -not [string]::IsNullOrWhiteSpace([string]$serviceState.StartMode) -and
                -not [string]::IsNullOrWhiteSpace([string]$serviceState.State)
            }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Name' -Value $Entry.Name
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Start mode' -Value $(if ($serviceState.PSObject.Properties['StartMode']) { $serviceState.StartMode } else { $null })
            Add-ReportKeyValueLine -Lines $Lines -Label 'State' -Value $(if ($serviceState.PSObject.Properties['State']) { $serviceState.State } else { $null })
            if (-not $captured -and $Entry.PSObject.Properties['CaptureError']) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
        }
        'LocalUser' {
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else {
                -not [string]::IsNullOrWhiteSpace([string]$Entry.Sid) -and $null -ne $Entry.CurrentValue
            }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Name' -Value $Entry.Name
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'SID' -Value $Entry.Sid
            Add-ReportKeyValueLine -Lines $Lines -Label 'Enabled' -Value $Entry.CurrentValue
            if (-not $captured -and $Entry.PSObject.Properties['CaptureError']) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
        }
        'WsManValue' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Path' -Value $Entry.Path
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
            if ($Entry.PSObject.Properties['CaptureError'] -and -not [string]::IsNullOrWhiteSpace([string]$Entry.CaptureError)) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
        }
        'WinRmListeners' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            $listeners = @($Entry.CurrentValue)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Listener count' -Value @($listeners).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            if ($Entry.PSObject.Properties['CaptureError'] -and -not [string]::IsNullOrWhiteSpace([string]$Entry.CaptureError)) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }

            foreach ($listener in $listeners) {
                $Lines.Add(('  - {0} {1} | Port={2} | Enabled={3} | Hostname={4} | URLPrefix={5} | Cert={6}' -f $listener.Transport, $listener.Address, $listener.Port, $listener.Enabled, (ConvertTo-DisplayString -Value $listener.Hostname), (ConvertTo-DisplayString -Value $listener.URLPrefix), (ConvertTo-DisplayString -Value $listener.CertificateThumbprint)))
            }
        }
        'AuditPolicy' {
            $auditState = if ($null -ne $Entry.CurrentValue) { $Entry.CurrentValue } else { [PSCustomObject]@{} }
            $commandAvailable = if ($auditState.PSObject.Properties['CommandAvailable']) { [bool]$auditState.CommandAvailable } else { $true }
            $captured = if ($auditState.PSObject.Properties['Captured']) { [bool]$auditState.Captured } else {
                [bool]($auditState.PSObject.Properties['Success'] -and $auditState.PSObject.Properties['Failure'])
            }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Subcategory' -Value $Entry.Subcategory
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Success' -Value $(if ($auditState.PSObject.Properties['Success']) { $auditState.Success } else { $null })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Failure' -Value $(if ($auditState.PSObject.Properties['Failure']) { $auditState.Failure } else { $null })
            if (-not ($commandAvailable -and $captured) -and $auditState.PSObject.Properties['Error']) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $auditState.Error
            }
        }
        'SmbClientConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Timed out' -Value $state.TimedOut
            Add-ReportKeyValueLine -Lines $Lines -Label 'Require security signature' -Value $state.RequireSecuritySignature
        }
        'SmbServerConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Timed out' -Value $state.TimedOut
            Add-ReportKeyValueLine -Lines $Lines -Label 'Require security signature' -Value $state.RequireSecuritySignature
        }
        default {
            Add-ReportJsonBlock -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
        }
    }
}

function Add-CapturePerformanceReportLines {
    param(
        [Parameter(Mandatory)] [System.Collections.Generic.List[string]]$Lines,
        [AllowNull()] [object]$Metrics,
        [int]$Top = 5,
        [ValidateSet('Capture', 'Mutation')] [string]$Activity = 'Capture'
    )

    if ($null -eq $Metrics) {
        return
    }

    $durationMs = if ($Metrics.PSObject.Properties['DurationMs']) { [double]$Metrics.DurationMs } else { 0 }
    $providerQueryCount = if ($Metrics.PSObject.Properties['ProviderQueryCount']) { [int]$Metrics.ProviderQueryCount } else { 0 }
    $cacheHitCount = if ($Metrics.PSObject.Properties['CacheHitCount']) { [int]$Metrics.CacheHitCount } else { 0 }
    $settings = @(if ($Metrics.PSObject.Properties['Settings']) { $Metrics.Settings })
    $providerQueries = @(if ($Metrics.PSObject.Properties['ProviderQueries']) { $Metrics.ProviderQueries })

    $Lines.Add(('{0} duration: {1:N2} seconds' -f $Activity, ($durationMs / 1000)))
    $providerLabel = if ($Activity -eq 'Capture') { 'Provider queries' } else { 'Provider setup queries' }
    $cacheLabel = if ($Activity -eq 'Capture') { 'Shared provider cache hits' } else { 'Shared provider setup cache hits' }
    $Lines.Add(('{0}: {1}' -f $providerLabel, $providerQueryCount))
    $Lines.Add(('{0}: {1}' -f $cacheLabel, $cacheHitCount))

    $slowSettings = @($settings | Sort-Object -Property DurationMs -Descending | Select-Object -First $Top)
    if ($slowSettings.Count -gt 0) {
        $Lines.Add($(if ($Activity -eq 'Capture') { 'Slowest settings:' } else { 'Slowest mutations:' }))
        foreach ($timing in $slowSettings) {
            $Lines.Add(('  - {0}: {1:N1} ms' -f ([string]$timing.Id), ([double]$timing.DurationMs)))
        }
    }

    $slowQueries = @($providerQueries | Sort-Object -Property DurationMs -Descending | Select-Object -First $Top)
    if ($slowQueries.Count -gt 0) {
        $Lines.Add($(if ($Activity -eq 'Capture') { 'Slowest provider queries:' } else { 'Slowest provider setup queries:' }))
        foreach ($timing in $slowQueries) {
            $Lines.Add(('  - {0}: {1:N1} ms' -f ([string]$timing.Key), ([double]$timing.DurationMs)))
        }
    }
}

function Add-RuntimeInfoReportLines {
    param(
        [Parameter(Mandatory)] [System.Collections.Generic.List[string]]$Lines,
        [AllowNull()] [object]$RuntimeInfo,
        [Parameter(Mandatory)] [string]$Prefix
    )

    if ($null -eq $RuntimeInfo) {
        $Lines.Add(("{0} script SHA-256: <not recorded>" -f $Prefix))
        return
    }

    $scriptFileName = if ($RuntimeInfo.PSObject.Properties['ScriptFileName']) { [string]$RuntimeInfo.ScriptFileName } else { '<unknown>' }
    $scriptSha256 = if ($RuntimeInfo.PSObject.Properties['ScriptSha256'] -and -not [string]::IsNullOrWhiteSpace([string]$RuntimeInfo.ScriptSha256)) { [string]$RuntimeInfo.ScriptSha256 } else { '<unavailable>' }
    $powerShellVersion = if ($RuntimeInfo.PSObject.Properties['PowerShellVersion']) { [string]$RuntimeInfo.PowerShellVersion } else { '<unknown>' }
    $powerShellEdition = if ($RuntimeInfo.PSObject.Properties['PowerShellEdition']) { [string]$RuntimeInfo.PowerShellEdition } else { '<unknown>' }
    $architecture = if ($RuntimeInfo.PSObject.Properties['ProcessArchitecture'] -and -not [string]::IsNullOrWhiteSpace([string]$RuntimeInfo.ProcessArchitecture)) { [string]$RuntimeInfo.ProcessArchitecture } else { '<unknown>' }

    $Lines.Add(("{0} script: {1}" -f $Prefix, $scriptFileName))
    $Lines.Add(("{0} script SHA-256: {1}" -f $Prefix, $scriptSha256))
    $Lines.Add(("{0} runtime: PowerShell {1} ({2}, {3})" -f $Prefix, $powerShellVersion, $powerShellEdition, $architecture))
}

function Get-SnapshotReportLines {
    param(
        [Parameter(Mandatory)] [object]$Snapshot,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $settings = @($Snapshot.Settings)
    $incompleteBaselineCount = @($settings | Where-Object { -not (Test-SnapshotEntryCapturedExactly -Entry $_ -SnapshotPath $SnapshotPath) }).Count
    $platformManagedWdacCount = 0
    $wdacEntry = @($settings | Where-Object { [string]$_.Type -eq 'WdacPolicies' } | Select-Object -First 1)
    if ($wdacEntry.Count -gt 0) {
        $platformManagedWdacCount = @(Get-WdacPlatformManagedPolicies -State $wdacEntry[0].CurrentValue).Count
    }

    $lines.Add('WinDefState Snapshot Report')
    $lines.Add(('Snapshot JSON: {0}' -f $SnapshotPath))
    $producer = if ($Snapshot.PSObject.Properties['Producer']) { $Snapshot.Producer } else { $null }
    Add-RuntimeInfoReportLines -Lines $lines -RuntimeInfo $producer -Prefix 'Producer'
    $lines.Add(('ComputerName: {0}' -f $Snapshot.ComputerName))
    $lines.Add(('CapturedAtUtc: {0}' -f $Snapshot.CapturedAtUtc))
    $lines.Add(('Settings captured: {0}' -f $settings.Count))
    $lines.Add(('Reboot-required settings: {0}' -f (@($settings | Where-Object { $_.RequiresReboot }).Count)))
    $lines.Add(('Incomplete-baseline settings: {0}' -f $incompleteBaselineCount))
    $lines.Add(('Platform-managed WDAC policies: {0}' -f $platformManagedWdacCount))
    if ($Snapshot.PSObject.Properties['CaptureScope'] -and $null -ne $Snapshot.CaptureScope) {
        $scopeLabel = if ($Snapshot.CaptureScope.PSObject.Properties['IsFiltered'] -and [bool]$Snapshot.CaptureScope.IsFiltered) { 'Filtered' } else { 'All settings' }
        $lines.Add(('Capture scope: {0}' -f $scopeLabel))
        $includedIds = @(if ($Snapshot.CaptureScope.PSObject.Properties['IncludeId']) { $Snapshot.CaptureScope.IncludeId })
        $excludedIds = @(if ($Snapshot.CaptureScope.PSObject.Properties['ExcludeId']) { $Snapshot.CaptureScope.ExcludeId })
        if ($includedIds.Count -gt 0) {
            $lines.Add(('Included ID filters: {0}' -f ($includedIds -join ', ')))
        }
        if ($excludedIds.Count -gt 0) {
            $lines.Add(('Excluded ID filters: {0}' -f ($excludedIds -join ', ')))
        }
    }
    if ($Snapshot.PSObject.Properties['CaptureMetrics']) {
        Add-CapturePerformanceReportLines -Lines $lines -Metrics $Snapshot.CaptureMetrics
    }

    foreach ($entry in $settings) {
        $lines.Add(' ')
        Add-SnapshotEntryReportLines -Lines $lines -Entry $entry -SnapshotPath $SnapshotPath
    }

    [string[]]$lines
}

function Show-ReportLines {
    param([AllowEmptyString()] [string[]]$Lines)

    foreach ($line in @($Lines)) {
        Write-Host $line
    }
}

function Get-ReportSummaryLines {
    param([AllowEmptyString()] [string[]]$Lines)

    $summary = [System.Collections.Generic.List[string]]::new()
    foreach ($line in @($Lines)) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            break
        }
        $summary.Add($line)
    }

    [string[]]$summary
}

function ConvertTo-ComparableSnapshotEntry {
    param(
        [Parameter(Mandatory)] [object]$Entry,
        [string]$SnapshotPath,
        [AllowNull()] [object]$ReferenceEntry
    )

    switch ($Entry.Type) {
        'RegistryValue' {
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Path           = $Entry.Path
                Name           = $Entry.Name
                ValueKind      = $Entry.ValueKind
                Captured       = $captured
                Exists         = [bool]$Entry.Exists
                CurrentValue   = $Entry.CurrentValue
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'RegistryKeyFlat' {
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Path           = $Entry.Path
                Captured       = $captured
                Exists         = $Entry.Exists
                CurrentValue   = @(
                    foreach ($value in @($Entry.CurrentValue)) {
                        [PSCustomObject]@{
                            Name      = $value.Name
                            ValueKind = $value.ValueKind
                            Value     = $value.Value
                        }
                    }
                )
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'MpPreferenceValue' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $null -ne $Entry.RestoreValue }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Property       = $Entry.Property
                CommandAvailable = $commandAvailable
                Captured       = $captured
                RestoreValue   = $Entry.RestoreValue
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'MpPreferenceList' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id               = $Entry.Id
                Type             = $Entry.Type
                Property         = $Entry.Property
                CommandAvailable = $commandAvailable
                Captured         = $captured
                CurrentValue     = @(Normalize-MpPreferenceListItems -Value $Entry.CurrentValue)
                RequiresReboot   = $Entry.RequiresReboot
            })
        }
        'DefenderRuntimeStatus' {
            $state = if ($null -ne $Entry.CurrentValue) { $Entry.CurrentValue } else { [PSCustomObject]@{} }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable          = if ($state.PSObject.Properties['CommandAvailable']) { [bool]$state.CommandAvailable } else { $true }
                    Captured                  = if ($state.PSObject.Properties['Captured']) { [bool]$state.Captured } else { $true }
                    AMRunningMode             = if ($state.PSObject.Properties['AMRunningMode']) { [string]$state.AMRunningMode } else { $null }
                    RealTimeProtectionEnabled = if ($state.PSObject.Properties['RealTimeProtectionEnabled']) { ConvertTo-NullableBoolean -Value $state.RealTimeProtectionEnabled } else { $null }
                    AntivirusEnabled          = if ($state.PSObject.Properties['AntivirusEnabled']) { ConvertTo-NullableBoolean -Value $state.AntivirusEnabled } else { $null }
                    IsTamperProtected         = if ($state.PSObject.Properties['IsTamperProtected']) { ConvertTo-NullableBoolean -Value $state.IsTamperProtected } else { $null }
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'BitLockerVolumes' {
            $referenceState = if ($null -ne $ReferenceEntry -and [string]$ReferenceEntry.Type -eq 'BitLockerVolumes') { $ReferenceEntry.CurrentValue } else { $null }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = ConvertTo-ComparableBitLockerState -State $Entry.CurrentValue -ReferenceState $referenceState
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'AppLockerPolicy' {
            $effectiveXml = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $Entry.CurrentValue -PolicyScope Effective -SnapshotPath $SnapshotPath)
            $localXml = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $Entry.CurrentValue -PolicyScope Local -SnapshotPath $SnapshotPath)
            $collectionSummaries = if ($Entry.CurrentValue.PSObject.Properties['CollectionSummaries']) { @($Entry.CurrentValue.CollectionSummaries) } else { @(Get-AppLockerCollectionSummaries -Xml $effectiveXml) }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable      = if ($Entry.CurrentValue.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CurrentValue.CommandAvailable } else { $true }
                    LocalCaptured         = if ($Entry.CurrentValue.PSObject.Properties['LocalCaptured']) { [bool]$Entry.CurrentValue.LocalCaptured } else { $false }
                    EffectiveCaptured     = if ($Entry.CurrentValue.PSObject.Properties['EffectiveCaptured']) { [bool]$Entry.CurrentValue.EffectiveCaptured } else { $false }
                    LocalMatchesEffective = if ($Entry.CurrentValue.PSObject.Properties['LocalMatchesEffective']) { [bool]$Entry.CurrentValue.LocalMatchesEffective } else { $false }
                    CaptureIssues         = @(
                        foreach ($issue in @($Entry.CurrentValue.CaptureIssues | Sort-Object -Property Scope, Message)) {
                            [PSCustomObject]@{
                                Scope   = if ($null -ne $issue.PSObject.Properties['Scope']) { [string]$issue.Scope } else { $null }
                                Message = if ($null -ne $issue.PSObject.Properties['Message']) { [string]$issue.Message } else { $null }
                            }
                        }
                    )
                    CollectionSummaries   = @(
                        foreach ($collection in @($collectionSummaries | Sort-Object -Property Type)) {
                            [PSCustomObject]@{
                                Type                = if ($null -ne $collection.PSObject.Properties['Type']) { [string]$collection.Type } else { $null }
                                EnforcementMode     = if ($null -ne $collection.PSObject.Properties['EnforcementMode']) { [string]$collection.EnforcementMode } else { $null }
                                RuleCount           = if ($null -ne $collection.PSObject.Properties['RuleCount']) { [int]$collection.RuleCount } else { 0 }
                                ServicesEnforcement = if ($null -ne $collection.PSObject.Properties['ServicesEnforcement']) { [string]$collection.ServicesEnforcement } else { $null }
                                SystemAppsAllow     = if ($null -ne $collection.PSObject.Properties['SystemAppsAllow']) { [string]$collection.SystemAppsAllow } else { $null }
                            }
                        }
                    )
                    EffectiveXml          = $effectiveXml
                    LocalXml              = $localXml
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'ExploitProtectionPolicy' {
            $xml = Get-ExploitProtectionPolicyXml -State $Entry.CurrentValue -SnapshotPath $SnapshotPath
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable = $Entry.CurrentValue.CommandAvailable
                    Xml              = Normalize-ExploitProtectionXml -Xml $xml
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'WdacPolicies' {
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = ConvertTo-ComparableWdacState -State $Entry.CurrentValue
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'AsrRules' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            $invalidEntries = @(Get-AsrInvalidEntriesFromEntry -Entry $Entry)
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CommandAvailable = $commandAvailable
                Captured       = $captured
                CurrentValue   = @(
                    foreach ($rule in @($Entry.CurrentValue)) {
                        [PSCustomObject]@{
                            Id          = [string]$rule.Id
                            ActionLabel = if ($null -ne $rule.PSObject.Properties['ActionLabel']) { [string]$rule.ActionLabel } else { Get-AsrActionLabel -Action $rule.Action }
                        }
                    }
                )
                InvalidEntries = @(
                    foreach ($rule in $invalidEntries) {
                        [PSCustomObject]@{
                            Id          = [string]$rule.Id
                            ActionLabel = if ($null -ne $rule.PSObject.Properties['ActionLabel']) { [string]$rule.ActionLabel } else { Get-AsrActionLabel -Action $rule.Action }
                        }
                    }
                )
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'PowerShellModuleLogging' {
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                BasePath       = $Entry.BasePath
                Captured       = $captured
                Exists         = $Entry.Exists
                CurrentValue   = [ordered]@{
                    BaseValues        = @(
                        foreach ($value in @($Entry.CurrentValue.BaseValues)) {
                            [PSCustomObject]@{
                                Name      = $value.Name
                                ValueKind = $value.ValueKind
                                Value     = $value.Value
                            }
                        }
                    )
                    ModuleNamesExists = $Entry.CurrentValue.ModuleNamesExists
                    ModuleNamesValues = @(
                        foreach ($value in @($Entry.CurrentValue.ModuleNamesValues)) {
                            [PSCustomObject]@{
                                Name      = $value.Name
                                ValueKind = $value.ValueKind
                                Value     = $value.Value
                            }
                        }
                    )
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'FirewallProfiles' {
            $referenceState = if ($null -ne $ReferenceEntry -and [string]$ReferenceEntry.Type -eq 'FirewallProfiles') { $ReferenceEntry.CurrentValue } else { $null }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = ConvertTo-ComparableFirewallProfileState -State $Entry.CurrentValue -ReferenceState $referenceState
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'FirewallRules' {
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = ConvertTo-ComparableFirewallRuleState -State $Entry.CurrentValue
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'NetBiosAdapters' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CommandAvailable = $commandAvailable
                Captured       = $captured
                CurrentValue   = @(
                    foreach ($adapter in @($Entry.CurrentValue)) {
                        [PSCustomObject]@{
                            Index               = [int]$adapter.Index
                            TcpipNetbiosOptions = [int]$adapter.TcpipNetbiosOptions
                        }
                    }
                )
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'LoadedUserRegistryValues' {
            $state = Normalize-UserRegistryValueState -State $Entry.CurrentValue
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    Entries       = @(
                        foreach ($value in @($state.Entries | Sort-Object -Property Sid, RelativePath, Name)) {
                            [PSCustomObject]@{
                                Sid          = $value.Sid
                                RelativePath = $value.RelativePath
                                Name         = $value.Name
                                Exists       = $value.Exists
                                CurrentValue = $value.CurrentValue
                                ValueKind    = $value.ValueKind
                            }
                        }
                    )
                    CaptureIssues = @(
                        foreach ($issue in @($state.CaptureIssues | Sort-Object -Property Sid)) {
                            [string]$issue.Sid
                        }
                    )
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'MachineEnvironmentValue' {
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Name           = $Entry.Name
                Exists         = [bool]$Entry.Exists
                CurrentValue   = if ($null -ne $Entry.CurrentValue) { [string]$Entry.CurrentValue } else { $null }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'ServiceConfig' {
            $serviceState = if ($null -ne $Entry.CurrentValue) { $Entry.CurrentValue } else { [PSCustomObject]@{} }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else {
                $null -ne $Entry.CurrentValue -and
                $serviceState.PSObject.Properties['StartMode'] -and
                $serviceState.PSObject.Properties['State'] -and
                -not [string]::IsNullOrWhiteSpace([string]$serviceState.StartMode) -and
                -not [string]::IsNullOrWhiteSpace([string]$serviceState.State)
            }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Name           = $Entry.Name
                Captured       = $captured
                CurrentValue   = [ordered]@{
                    StartMode = if ($serviceState.PSObject.Properties['StartMode'] -and $null -ne $serviceState.StartMode) { [string]$serviceState.StartMode } else { $null }
                    State     = if ($serviceState.PSObject.Properties['State'] -and $null -ne $serviceState.State) { [string]$serviceState.State } else { $null }
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'LocalUser' {
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else {
                -not [string]::IsNullOrWhiteSpace([string]$Entry.Sid) -and $null -ne $Entry.CurrentValue
            }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Name           = if ($null -ne $Entry.Name) { [string]$Entry.Name } else { $null }
                Sid            = if ($null -ne $Entry.Sid) { [string]$Entry.Sid } else { $null }
                Rid            = if ($null -ne $Entry.Rid) { [int]$Entry.Rid } else { $null }
                Captured       = $captured
                CurrentValue   = if ($null -ne $Entry.CurrentValue) { [bool]$Entry.CurrentValue } else { $null }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'WsManValue' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Path           = $Entry.Path
                CommandAvailable = $commandAvailable
                Captured       = $captured
                CurrentValue   = $Entry.CurrentValue
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'WinRmListeners' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CommandAvailable = $commandAvailable
                Captured       = $captured
                CurrentValue   = @(
                    foreach ($listener in @($Entry.CurrentValue | Sort-Object -Property Address, Transport, Port, Hostname, URLPrefix, CertificateThumbprint)) {
                        [PSCustomObject]@{
                            Address               = [string]$listener.Address
                            Transport             = [string]$listener.Transport
                            Port                  = if ($null -ne $listener.PSObject.Properties['Port'] -and $null -ne $listener.Port) { [int]$listener.Port } else { $null }
                            Hostname              = if ($null -ne $listener.PSObject.Properties['Hostname']) { [string]$listener.Hostname } else { $null }
                            Enabled               = if ($null -ne $listener.PSObject.Properties['Enabled']) { ConvertTo-NullableBoolean -Value $listener.Enabled } else { $null }
                            URLPrefix             = if ($null -ne $listener.PSObject.Properties['URLPrefix']) { [string]$listener.URLPrefix } else { $null }
                            CertificateThumbprint = if ($null -ne $listener.PSObject.Properties['CertificateThumbprint']) { [string]$listener.CertificateThumbprint } else { $null }
                        }
                    }
                )
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'AuditPolicy' {
            $auditState = if ($null -ne $Entry.CurrentValue) { $Entry.CurrentValue } else { [PSCustomObject]@{} }
            $commandAvailable = if ($auditState.PSObject.Properties['CommandAvailable']) { [bool]$auditState.CommandAvailable } else { $true }
            $captured = if ($auditState.PSObject.Properties['Captured']) { [bool]$auditState.Captured } else {
                [bool]($auditState.PSObject.Properties['Success'] -and $auditState.PSObject.Properties['Failure'])
            }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Subcategory    = $Entry.Subcategory
                CurrentValue   = [ordered]@{
                    CommandAvailable = $commandAvailable
                    Captured         = $captured
                    Success          = if ($auditState.PSObject.Properties['Success']) { [bool]$auditState.Success } else { $null }
                    Failure          = if ($auditState.PSObject.Properties['Failure']) { [bool]$auditState.Failure } else { $null }
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'SmbClientConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable         = $state.CommandAvailable
                    TimedOut                 = $state.TimedOut
                    RequireSecuritySignature = $state.RequireSecuritySignature
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'SmbServerConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable         = $state.CommandAvailable
                    TimedOut                 = $state.TimedOut
                    RequireSecuritySignature = $state.RequireSecuritySignature
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        default {
            return ConvertTo-CanonicalValue -Value $Entry
        }
    }
}

function Test-SnapshotEntryCapturedExactly {
    param(
        [Parameter(Mandatory)] [object]$Entry,
        [string]$SnapshotPath
    )

    switch ($Entry.Type) {
        'RegistryValue' {
            return $(if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true })
        }
        'RegistryKeyFlat' {
            return $(if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true })
        }
        'PowerShellModuleLogging' {
            return $(if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true })
        }
        'MachineEnvironmentValue' {
            return $true
        }
        'ServiceConfig' {
            if ($Entry.PSObject.Properties['Captured']) {
                return [bool]$Entry.Captured
            }
            $serviceState = if ($null -ne $Entry.CurrentValue) { $Entry.CurrentValue } else { [PSCustomObject]@{} }
            return (
                $serviceState.PSObject.Properties['StartMode'] -and
                $serviceState.PSObject.Properties['State'] -and
                -not [string]::IsNullOrWhiteSpace([string]$serviceState.StartMode) -and
                -not [string]::IsNullOrWhiteSpace([string]$serviceState.State)
            )
        }
        'LocalUser' {
            if ($Entry.PSObject.Properties['Captured']) {
                return [bool]$Entry.Captured
            }
            return (-not [string]::IsNullOrWhiteSpace([string]$Entry.Sid) -and $null -ne $Entry.CurrentValue)
        }
        'NetBiosAdapters' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ($commandAvailable -and $captured)
        }
        'AuditPolicy' {
            if ($null -eq $Entry.CurrentValue) {
                return $false
            }
            $commandAvailable = if ($Entry.CurrentValue.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CurrentValue.CommandAvailable } else { $true }
            $captured = if ($Entry.CurrentValue.PSObject.Properties['Captured']) { [bool]$Entry.CurrentValue.Captured } else {
                $null -ne $Entry.CurrentValue.PSObject.Properties['Success'] -and
                $null -ne $Entry.CurrentValue.PSObject.Properties['Failure']
            }
            return ($commandAvailable -and $captured)
        }
        'MpPreferenceValue' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $null -ne $Entry.RestoreValue }
            return ($commandAvailable -and $captured)
        }
        'MpPreferenceList' {
            return (Test-MpPreferenceListCapturedExactly -Entry $Entry)
        }
        'DefenderRuntimeStatus' {
            return (Test-DefenderRuntimeStatusCapturedExactly -State $Entry.CurrentValue)
        }
        'BitLockerVolumes' {
            return (Test-BitLockerStateCapturedExactly -State $Entry.CurrentValue)
        }
        'AppLockerPolicy' {
            return (Test-AppLockerPolicyCapturedExactly -State $Entry.CurrentValue -SnapshotPath $SnapshotPath)
        }
        'ExploitProtectionPolicy' {
            return (Test-ExploitProtectionPolicyCapturedExactly -State $Entry.CurrentValue -SnapshotPath $SnapshotPath)
        }
        'WdacPolicies' {
            return (Test-WdacPolicyStateCapturedExactly -State $Entry.CurrentValue)
        }
        'AsrRules' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            $invalidEntries = @(Get-AsrInvalidEntriesFromEntry -Entry $Entry)
            return ($commandAvailable -and $captured -and $invalidEntries.Count -eq 0)
        }
        'FirewallProfiles' {
            return (Test-FirewallProfileStateCapturedExactly -State $Entry.CurrentValue)
        }
        'FirewallRules' {
            return (Test-FirewallRuleStateCapturedExactly -State $Entry.CurrentValue)
        }
        'SmbClientConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            return $state.CommandAvailable -and -not $state.TimedOut -and $null -ne $state.RequireSecuritySignature
        }
        'SmbServerConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            return $state.CommandAvailable -and -not $state.TimedOut -and $null -ne $state.RequireSecuritySignature
        }
        'LoadedUserRegistryValues' {
            $state = Normalize-UserRegistryValueState -State $Entry.CurrentValue
            return (@($state.CaptureIssues).Count -eq 0)
        }
        'WsManValue' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ($commandAvailable -and $captured)
        }
        'WinRmListeners' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ($commandAvailable -and $captured)
        }
        default {
            return $false
        }
    }
}

function Test-DefenseSnapshot {
    param(
        [Parameter(Mandatory)] [object]$Snapshot,
        [string]$SnapshotPath,
        [AllowNull()] [string[]]$IncludeId,
        [AllowNull()] [string[]]$ExcludeId
    )

    $definitions = Get-DefenseDefinitionMap

    $results = @()
    $entriesToVerify = @(
        $Snapshot.Settings | Where-Object {
            Test-SettingIdIncluded -Id ([string]$_.Id) -IncludeId $IncludeId -ExcludeId $ExcludeId
        }
    )
    $captureSession = New-CaptureSession -Phase Verification
    $captureSession.UserRegistryDefinitionsRemaining = @(
        foreach ($entry in $entriesToVerify) {
            $entryId = [string]$entry.Id
            if (
                [string]$entry.Type -eq 'LoadedUserRegistryValues' -and
                $definitions.ContainsKey($entryId) -and
                (Test-SnapshotEntryCapturedExactly -Entry $entry -SnapshotPath $SnapshotPath)
            ) {
                $entry
            }
        }
    ).Count
    $captureSession.ServiceNames = @(
        foreach ($entry in $entriesToVerify) {
            $entryId = [string]$entry.Id
            if (
                [string]$entry.Type -eq 'ServiceConfig' -and
                $definitions.ContainsKey($entryId) -and
                (Test-SnapshotEntryCapturedExactly -Entry $entry -SnapshotPath $SnapshotPath)
            ) {
                [string]$definitions[$entryId].Name
            }
        }
    )
    $captureMetrics = $null
    try {
        for ($entryIndex = 0; $entryIndex -lt $entriesToVerify.Count; $entryIndex++) {
            $entry = $entriesToVerify[$entryIndex]
            $entryId = [string]$entry.Id
            Write-OperationProgress -Phase 'Verify' -Current ($entryIndex + 1) -Total $entriesToVerify.Count -Id $entryId

            $baselineComplete = Test-SnapshotEntryCapturedExactly -Entry $entry -SnapshotPath $SnapshotPath
            if (-not $baselineComplete) {
                $results += [PSCustomObject]@{
                    Id             = $entryId
                    Type           = $entry.Type
                    RequiresReboot = $entry.RequiresReboot
                    Matches        = $false
                    Skipped        = $true
                    SkipCategory   = 'IncompleteBaseline'
                    Reason         = 'Verification skipped because the snapshot baseline was incomplete.'
                    Expected       = $null
                    Actual         = $null
                }
                continue
            }

            if (-not $definitions.ContainsKey($entryId)) {
                $results += [PSCustomObject]@{
                    Id             = $entryId
                    Type           = $entry.Type
                    RequiresReboot = $entry.RequiresReboot
                    Matches        = $false
                    Skipped        = $false
                    Reason         = 'This snapshot entry no longer has a matching definition in the current script.'
                    Expected       = $null
                    Actual         = $null
                }
                continue
            }

            if (-not (Test-DefinitionHasRestoreAction -Definition $definitions[$entryId])) {
                $results += [PSCustomObject]@{
                    Id             = $entryId
                    Type           = $entry.Type
                    RequiresReboot = $entry.RequiresReboot
                    Matches        = $true
                    Skipped        = $true
                    SkipCategory   = 'InventoryOnly'
                    Reason         = 'Verification skipped because this entry is inventory-only and has no restore action.'
                    Expected       = $null
                    Actual         = $null
                }
                continue
            }

            $expectedComparable = ConvertTo-ComparableSnapshotEntry -Entry $entry -SnapshotPath $SnapshotPath
            try {
                $liveEntry = Invoke-TimedDefinitionCapture -Definition $definitions[$entryId] -CaptureSession $captureSession
                $actualComparable = ConvertTo-ComparableSnapshotEntry -Entry $liveEntry -ReferenceEntry $entry
                $isMatch = (Get-CanonicalJson -Value $expectedComparable) -eq (Get-CanonicalJson -Value $actualComparable)

                $results += [PSCustomObject]@{
                    Id             = $entryId
                    Type           = $entry.Type
                    RequiresReboot = $entry.RequiresReboot
                    Matches        = $isMatch
                    Skipped        = $false
                    Reason         = if ($isMatch) { $null } else { 'Live state does not match the snapshot entry.' }
                    Expected       = $expectedComparable
                    Actual         = $actualComparable
                }
            } catch {
                $results += [PSCustomObject]@{
                    Id             = $entryId
                    Type           = $entry.Type
                    RequiresReboot = $entry.RequiresReboot
                    Matches        = $false
                    Skipped        = $false
                    Reason         = $_.Exception.Message
                    Expected       = $expectedComparable
                    Actual         = $null
                }
            }
        }
    } finally {
        $captureMetrics = Complete-CaptureSession -Session $captureSession
    }

    [PSCustomObject]@{
        Tool          = 'WinDefState'
        BaselineProducer = if ($Snapshot.PSObject.Properties['Producer']) { $Snapshot.Producer } else { $null }
        Verifier      = Get-WinDefStateRuntimeInfo
        ComputerName  = $env:COMPUTERNAME
        VerifiedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        MatchedCount  = @($results | Where-Object { $_.Matches -and -not $_.Skipped }).Count
        SkippedCount  = @($results | Where-Object { $_.Skipped }).Count
        IncompleteCount = @($results | Where-Object { $_.Skipped -and $_.PSObject.Properties['SkipCategory'] -and [string]$_.SkipCategory -eq 'IncompleteBaseline' }).Count
        InventoryCount = @($results | Where-Object { $_.Skipped -and $_.PSObject.Properties['SkipCategory'] -and [string]$_.SkipCategory -eq 'InventoryOnly' }).Count
        MismatchCount = @($results | Where-Object { -not $_.Matches -and -not $_.Skipped }).Count
        CaptureMetrics = $captureMetrics
        Results       = @($results)
    }
}

function Get-VerificationReportLines {
    param(
        [Parameter(Mandatory)] [object]$Verification,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $mismatches = @($Verification.Results | Where-Object { -not $_.Matches -and -not $_.Skipped })
    $skipped = @($Verification.Results | Where-Object { $_.Skipped })

    $lines.Add('WinDefState Restore Verification')
    $lines.Add(('Snapshot JSON: {0}' -f $SnapshotPath))
    $baselineProducer = if ($Verification.PSObject.Properties['BaselineProducer']) { $Verification.BaselineProducer } else { $null }
    $verifier = if ($Verification.PSObject.Properties['Verifier']) { $Verification.Verifier } else { $null }
    Add-RuntimeInfoReportLines -Lines $lines -RuntimeInfo $baselineProducer -Prefix 'Baseline producer'
    Add-RuntimeInfoReportLines -Lines $lines -RuntimeInfo $verifier -Prefix 'Verifier'
    $lines.Add(('ComputerName: {0}' -f $Verification.ComputerName))
    $lines.Add(('VerifiedAtUtc: {0}' -f $Verification.VerifiedAtUtc))
    $lines.Add(('Matched settings: {0}' -f $Verification.MatchedCount))
    $lines.Add(('Skipped settings: {0}' -f $Verification.SkippedCount))
    $incompleteCount = if ($Verification.PSObject.Properties['IncompleteCount']) { [int]$Verification.IncompleteCount } else { [int]$Verification.SkippedCount }
    $inventoryCount = if ($Verification.PSObject.Properties['InventoryCount']) { [int]$Verification.InventoryCount } else { 0 }
    $lines.Add(('Incomplete-baseline settings: {0}' -f $incompleteCount))
    $lines.Add(('Inventory-only settings: {0}' -f $inventoryCount))
    $lines.Add(('Mismatched settings: {0}' -f $Verification.MismatchCount))
    if ($Verification.PSObject.Properties['RestoreCheckpointSummary']) {
        $checkpoint = $Verification.RestoreCheckpointSummary
        $lines.Add(('Restore attempt: {0}' -f $checkpoint.AttemptNumber))
        $lines.Add(('Restore mutation settings requested: {0}' -f $checkpoint.RequestedMutationCount))
        $lines.Add(('Checkpointed settings revalidated: {0}' -f $checkpoint.PreviouslyCompletedCount))
        $lines.Add(('Checkpointed settings skipped as matching: {0}' -f $checkpoint.RevalidatedAndSkippedCount))
        $lines.Add(('Checkpointed settings scheduled again: {0}' -f $checkpoint.ReappliedCheckpointCount))
        $lines.Add(('Restore mutation settings scheduled this attempt: {0}' -f $checkpoint.ScheduledMutationCount))
    }
    if ($Verification.PSObject.Properties['CaptureMetrics']) {
        Add-CapturePerformanceReportLines -Lines $lines -Metrics $Verification.CaptureMetrics
    }
    if ($Verification.PSObject.Properties['MutationMetrics']) {
        Add-CapturePerformanceReportLines -Lines $lines -Metrics $Verification.MutationMetrics -Activity Mutation
    }

    if ($mismatches.Count -eq 0) {
        $lines.Add(' ')
        if ($skipped.Count -eq 0) {
            $lines.Add('All captured settings match the requested snapshot.')
        } elseif ($incompleteCount -gt 0) {
            $lines.Add('All restorable settings with complete baselines match the requested snapshot.')
        } else {
            $lines.Add('All restorable settings match the requested snapshot; inventory-only entries were not treated as restore targets.')
        }
    }

    foreach ($skippedResult in $skipped) {
        $lines.Add(' ')
        $lines.Add("[$($skippedResult.Id)] $($skippedResult.Type)")
        Add-ReportKeyValueLine -Lines $lines -Label 'Requires reboot' -Value $skippedResult.RequiresReboot
        if ($skippedResult.PSObject.Properties['SkipCategory']) {
            Add-ReportKeyValueLine -Lines $lines -Label 'Skip category' -Value $skippedResult.SkipCategory
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$skippedResult.Reason)) {
            Add-ReportKeyValueLine -Lines $lines -Label 'Reason' -Value $skippedResult.Reason
        }
    }

    if ($mismatches.Count -eq 0) {
        return [string[]]$lines
    }

    foreach ($mismatch in $mismatches) {
        $lines.Add(' ')
        $lines.Add("[$($mismatch.Id)] $($mismatch.Type)")
        Add-ReportKeyValueLine -Lines $lines -Label 'Requires reboot' -Value $mismatch.RequiresReboot
        if (-not [string]::IsNullOrWhiteSpace([string]$mismatch.Reason)) {
            Add-ReportKeyValueLine -Lines $lines -Label 'Reason' -Value $mismatch.Reason
        }
        Add-ReportJsonBlock -Lines $lines -Label 'Expected' -Value $mismatch.Expected
        Add-ReportJsonBlock -Lines $lines -Label 'Actual' -Value $mismatch.Actual
    }

    [string[]]$lines
}

function Get-PermissiveVerificationReportLines {
    param(
        [Parameter(Mandatory)] [object]$Verification,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('WinDefState Permissive Verification')
    $lines.Add(('Baseline snapshot: {0}' -f $SnapshotPath))
    $baselineProducer = if ($Verification.PSObject.Properties['BaselineProducer']) { $Verification.BaselineProducer } else { $null }
    $verifier = if ($Verification.PSObject.Properties['Verifier']) { $Verification.Verifier } else { $null }
    Add-RuntimeInfoReportLines -Lines $lines -RuntimeInfo $baselineProducer -Prefix 'Baseline producer'
    Add-RuntimeInfoReportLines -Lines $lines -RuntimeInfo $verifier -Prefix 'Verifier'
    $lines.Add(('ComputerName: {0}' -f $Verification.ComputerName))
    $lines.Add(('VerifiedAtUtc: {0}' -f $Verification.VerifiedAtUtc))
    $lines.Add(('Verified settings: {0}' -f $Verification.VerifiedCount))
    $lines.Add(('Configured, pending reboot: {0}' -f $Verification.PendingRebootCount))
    $lines.Add(('Mismatched settings: {0}' -f $Verification.MismatchCount))
    if ($Verification.PSObject.Properties['CaptureMetrics']) {
        Add-CapturePerformanceReportLines -Lines $lines -Metrics $Verification.CaptureMetrics
    }
    if ($Verification.PSObject.Properties['MutationMetrics']) {
        Add-CapturePerformanceReportLines -Lines $lines -Metrics $Verification.MutationMetrics -Activity Mutation
    }

    foreach ($result in @($Verification.Results)) {
        $lines.Add(' ')
        $lines.Add("[$($result.Id)] $($result.Type)")
        Add-ReportKeyValueLine -Lines $lines -Label 'Status' -Value $result.Status
        Add-ReportKeyValueLine -Lines $lines -Label 'Configured state changed' -Value $result.Changed
        Add-ReportKeyValueLine -Lines $lines -Label 'Requires reboot' -Value $result.RequiresReboot
        if (-not [string]::IsNullOrWhiteSpace([string]$result.Reason)) {
            Add-ReportKeyValueLine -Lines $lines -Label 'Reason' -Value $result.Reason
        }
        if ([string]$result.Status -ne 'Verified') {
            Add-ReportJsonBlock -Lines $lines -Label 'Expected permissive state' -Value $result.Expected
            Add-ReportJsonBlock -Lines $lines -Label 'Observed state' -Value $result.Actual
        }
    }

    [string[]]$lines
}

function Get-WdacVerificationReportLines {
    param(
        [Parameter(Mandatory)] [object]$Verification,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $wdacResults = @($Verification.Results | Where-Object { [string]$_.Type -eq 'WdacPolicies' -or [string]$_.Id -eq 'wdac.policies' })

    $lines.Add('WinDefState WDAC Restore Verification')
    $lines.Add(('Snapshot JSON: {0}' -f $SnapshotPath))
    $baselineProducer = if ($Verification.PSObject.Properties['BaselineProducer']) { $Verification.BaselineProducer } else { $null }
    $verifier = if ($Verification.PSObject.Properties['Verifier']) { $Verification.Verifier } else { $null }
    Add-RuntimeInfoReportLines -Lines $lines -RuntimeInfo $baselineProducer -Prefix 'Baseline producer'
    Add-RuntimeInfoReportLines -Lines $lines -RuntimeInfo $verifier -Prefix 'Verifier'
    $lines.Add(('ComputerName: {0}' -f $Verification.ComputerName))
    $lines.Add(('VerifiedAtUtc: {0}' -f $Verification.VerifiedAtUtc))
    $lines.Add(('WDAC result count: {0}' -f $wdacResults.Count))

    if ($wdacResults.Count -eq 0) {
        $lines.Add(' ')
        $lines.Add('No WDAC verification entry was present in this verification run.')
        return [string[]]$lines
    }

    foreach ($result in $wdacResults) {
        $lines.Add(' ')
        $lines.Add("[$($result.Id)] $($result.Type)")
        Add-ReportKeyValueLine -Lines $lines -Label 'Requires reboot' -Value $result.RequiresReboot
        Add-ReportKeyValueLine -Lines $lines -Label 'Matches snapshot' -Value $result.Matches
        Add-ReportKeyValueLine -Lines $lines -Label 'Skipped' -Value $result.Skipped
        if (-not [string]::IsNullOrWhiteSpace([string]$result.Reason)) {
            Add-ReportKeyValueLine -Lines $lines -Label 'Reason' -Value $result.Reason
        }
        Add-ReportJsonBlock -Lines $lines -Label 'Expected' -Value $result.Expected
        Add-ReportJsonBlock -Lines $lines -Label 'Actual' -Value $result.Actual
    }

    [string[]]$lines
}

function Get-RestoreCompletionMessage {
    param(
        [Parameter(Mandatory)] [string]$SnapshotPath,
        [Parameter(Mandatory)] [object]$Verification
    )

    $incompleteCount = if ($Verification.PSObject.Properties['IncompleteCount']) { [int]$Verification.IncompleteCount } else { [int]$Verification.SkippedCount }
    $inventoryCount = if ($Verification.PSObject.Properties['InventoryCount']) { [int]$Verification.InventoryCount } else { 0 }
    if ($incompleteCount -gt 0) {
        return "Restore completed and verified all settings with complete baselines from snapshot: $SnapshotPath"
    }
    if ($inventoryCount -gt 0) {
        return "Restore completed and verified all restorable settings from snapshot: $SnapshotPath"
    }

    "Restore completed and verified from snapshot: $SnapshotPath"
}

function Persist-SnapshotExternalAssets {
    param(
        [Parameter(Mandatory)] [object]$Snapshot,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $SnapshotPath
    $utf8Encoding = [System.Text.UTF8Encoding]::new($false)

    foreach ($entry in @($Snapshot.Settings)) {
        if ([string]$entry.Type -eq 'AppLockerPolicy' -and $null -ne $entry.CurrentValue) {
            $localXml = Get-AppLockerPolicyXml -State $entry.CurrentValue -PolicyScope Local
            $effectiveXml = Get-AppLockerPolicyXml -State $entry.CurrentValue -PolicyScope Effective
            $localAssetRelativePath = if (-not [string]::IsNullOrWhiteSpace($localXml)) { Join-Path 'applocker' 'local-policy.xml' } else { $null }
            $effectiveAssetRelativePath = if (-not [string]::IsNullOrWhiteSpace($effectiveXml)) { Join-Path 'applocker' 'effective-policy.xml' } else { $null }
            $localAssetSha256 = $null
            $effectiveAssetSha256 = $null

            if (-not [string]::IsNullOrWhiteSpace($localXml)) {
                Write-TextAtomic -Path (Join-Path $assetRoot $localAssetRelativePath) -Content $localXml
                $localAssetSha256 = Get-Sha256HashFromBytes -Content ($utf8Encoding.GetBytes($localXml))
            }

            if (-not [string]::IsNullOrWhiteSpace($effectiveXml)) {
                Write-TextAtomic -Path (Join-Path $assetRoot $effectiveAssetRelativePath) -Content $effectiveXml
                $effectiveAssetSha256 = Get-Sha256HashFromBytes -Content ($utf8Encoding.GetBytes($effectiveXml))
            }

            $entry.CurrentValue = [PSCustomObject]@{
                CommandAvailable               = $entry.CurrentValue.CommandAvailable
                LocalCaptured                  = $entry.CurrentValue.LocalCaptured
                EffectiveCaptured              = $entry.CurrentValue.EffectiveCaptured
                LocalMatchesEffective          = $entry.CurrentValue.LocalMatchesEffective
                CaptureIssues                  = @($entry.CurrentValue.CaptureIssues)
                CollectionSummaries            = @($entry.CurrentValue.CollectionSummaries)
                LocalSnapshotAssetRelativePath = $localAssetRelativePath
                LocalSnapshotAssetSha256       = $localAssetSha256
                EffectiveSnapshotAssetRelativePath = $effectiveAssetRelativePath
                EffectiveSnapshotAssetSha256   = $effectiveAssetSha256
            }
            continue
        }

        if ([string]$entry.Type -eq 'ExploitProtectionPolicy' -and $null -ne $entry.CurrentValue) {
            $xml = Get-ExploitProtectionPolicyXml -State $entry.CurrentValue
            $assetRelativePath = Join-Path 'exploit-protection' 'policy.xml'
            $assetSha256 = $null
            if (-not [string]::IsNullOrWhiteSpace($xml)) {
                $assetPath = Join-Path $assetRoot $assetRelativePath
                Write-TextAtomic -Path $assetPath -Content $xml
                $assetSha256 = Get-Sha256HashFromBytes -Content ($utf8Encoding.GetBytes($xml))
            }

            $entry.CurrentValue = [PSCustomObject]@{
                CommandAvailable         = $entry.CurrentValue.CommandAvailable
                SnapshotAssetRelativePath = if (-not [string]::IsNullOrWhiteSpace($xml)) { $assetRelativePath } else { $null }
                SnapshotAssetSha256       = $assetSha256
            }
            continue
        }

        if ([string]$entry.Type -ne 'WdacPolicies' -or $null -eq $entry.CurrentValue) {
            continue
        }

        $rewrittenFiles = @()
        foreach ($file in @($entry.CurrentValue.Files)) {
            $relativePath = [string]$file.RelativePath
            $assetRelativePath = Join-Path 'wdac' $relativePath
            $assetPath = Join-Path $assetRoot $assetRelativePath
            $assetSha256 = if ($file.PSObject.Properties['Sha256']) { [string]$file.Sha256 } else { $null }

            if ($file.PSObject.Properties['Base64'] -and -not [string]::IsNullOrWhiteSpace([string]$file.Base64)) {
                $assetBytes = [Convert]::FromBase64String([string]$file.Base64)
                Write-BytesAtomic -Path $assetPath -Content $assetBytes
                $assetSha256 = Get-Sha256HashFromBytes -Content $assetBytes
            }

            $rewrittenFiles += [PSCustomObject]@{
                RelativePath              = $relativePath
                FileName                  = [string]$file.FileName
                Sha256                    = $assetSha256
                SnapshotAssetRelativePath = $assetRelativePath
            }
        }

        $entry.CurrentValue = [PSCustomObject]@{
            CiToolAvailable = $entry.CurrentValue.CiToolAvailable
            Captured        = if ($entry.CurrentValue.PSObject.Properties['Captured']) { [bool]$entry.CurrentValue.Captured } else { $true }
            CaptureIssues   = @(
                if ($entry.CurrentValue.PSObject.Properties['CaptureIssues']) {
                    $entry.CurrentValue.CaptureIssues
                }
            )
            Policies        = @($entry.CurrentValue.Policies)
            Files           = @($rewrittenFiles)
        }
    }
}

function Initialize-SnapshotAssetCache {
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Entries,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    foreach ($entry in @($Entries)) {
        if ($null -eq $entry -or -not $entry.PSObject.Properties['Type']) {
            continue
        }

        switch ([string]$entry.Type) {
            'AppLockerPolicy' {
                if (
                    $null -ne $entry.CurrentValue -and
                    (Test-AppLockerPolicyCapturedExactly -State $entry.CurrentValue -SnapshotPath $SnapshotPath)
                ) {
                    $null = Get-AppLockerPolicyXml -State $entry.CurrentValue -PolicyScope Local -SnapshotPath $SnapshotPath
                    $null = Get-AppLockerPolicyXml -State $entry.CurrentValue -PolicyScope Effective -SnapshotPath $SnapshotPath
                }
            }
            'ExploitProtectionPolicy' {
                if (
                    $null -ne $entry.CurrentValue -and
                    (Test-ExploitProtectionPolicyCapturedExactly -State $entry.CurrentValue -SnapshotPath $SnapshotPath)
                ) {
                    $null = Get-ExploitProtectionPolicyXml -State $entry.CurrentValue -SnapshotPath $SnapshotPath
                }
            }
            'WdacPolicies' {
                if (
                    $null -ne $entry.CurrentValue -and
                    (Test-WdacPolicyStateCapturedExactly -State $entry.CurrentValue)
                ) {
                    foreach ($file in @($entry.CurrentValue.Files)) {
                        $null = Get-WdacSnapshotFileBytes -File $file -SnapshotPath $SnapshotPath
                    }
                }
            }
        }
    }
}

#endregion

#region Definition catalog and lifecycle dispatch

function Get-DefenseDefinitions {
    if ($null -ne $script:WinDefStateDefinitionCache) {
        return $script:WinDefStateDefinitionCache
    }

    $officeMacroApps = @('access', 'excel', 'powerpoint', 'project', 'publisher', 'visio', 'word')
    $officeMacroItems = foreach ($app in $officeMacroApps) {
        [PSCustomObject]@{
            RelativePath    = "Software\Policies\Microsoft\Office\16.0\$app\Security"
            Name            = 'BlockContentExecutionFromInternet'
            ValueKind       = 'DWord'
            PermissiveExists = $true
            PermissiveValue = 0
        }
    }

    $script:WinDefStateDefinitionCache = @(
        [PSCustomObject]@{ Id = 'defender.runtime_status'; Type = 'DefenderRuntimeStatus'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_realtime_monitoring'; Type = 'MpPreferenceValue'; Property = 'DisableRealtimeMonitoring'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_behavior_monitoring'; Type = 'MpPreferenceValue'; Property = 'DisableBehaviorMonitoring'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.maps_reporting'; Type = 'MpPreferenceValue'; Property = 'MAPSReporting'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Basic'; '2' = 'Advanced' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.submit_samples_consent'; Type = 'MpPreferenceValue'; Property = 'SubmitSamplesConsent'; PermissiveValue = 'NeverSend'; ValueMap = @{ '0' = 'AlwaysPrompt'; '1' = 'SendSafeSamples'; '2' = 'NeverSend'; '3' = 'SendAllSamples' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.pua_protection'; Type = 'MpPreferenceValue'; Property = 'PUAProtection'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_script_scanning'; Type = 'MpPreferenceValue'; Property = 'DisableScriptScanning'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_ioav_protection'; Type = 'MpPreferenceValue'; Property = 'DisableIOAVProtection'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_intrusion_prevention_system'; Type = 'MpPreferenceValue'; Property = 'DisableIntrusionPreventionSystem'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.enable_network_protection'; Type = 'MpPreferenceValue'; Property = 'EnableNetworkProtection'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.enable_controlled_folder_access'; Type = 'MpPreferenceValue'; Property = 'EnableControlledFolderAccess'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.exclusion_paths'; Type = 'MpPreferenceList'; Property = 'ExclusionPath'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.exclusion_processes'; Type = 'MpPreferenceList'; Property = 'ExclusionProcess'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.exclusion_extensions'; Type = 'MpPreferenceList'; Property = 'ExclusionExtension'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.cfa_allowed_applications'; Type = 'MpPreferenceList'; Property = 'ControlledFolderAccessAllowedApplications'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.cfa_protected_folders'; Type = 'MpPreferenceList'; Property = 'ControlledFolderAccessProtectedFolders'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.asr_rules'; Type = 'AsrRules'; RequiresReboot = $false }

        [PSCustomObject]@{
            Id                                     = 'firewall.profiles'
            Type                                   = 'FirewallProfiles'
            Profiles                               = @('Domain', 'Private', 'Public')
            PermissiveValue                        = $false
            PermissiveDefaultInboundAction         = 'Allow'
            PermissiveDefaultOutboundAction        = 'Allow'
            PermissiveAllowUnicastResponseToMulticast = $true
            PermissiveNotifyOnListen               = $false
            PermissiveLogAllowed                   = $false
            PermissiveLogBlocked                   = $false
            PermissiveLogIgnored                   = $false
            RequiresReboot                         = $false
        }

        [PSCustomObject]@{ Id = 'powershell.lockdown'; Type = 'MachineEnvironmentValue'; Name = '__PSLockdownPolicy'; PermissiveExists = $true; PermissiveValue = '0'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'powershell.script_block_logging'; Type = 'RegistryKeyFlat'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'powershell.module_logging'; Type = 'PowerShellModuleLogging'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'powershell.transcription'; Type = 'RegistryKeyFlat'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\Transcription'; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'applocker.policy'; Type = 'AppLockerPolicy'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'applocker.service'; Type = 'ServiceConfig'; Name = 'AppIDSvc'; PermissiveStartup = 'demand'; PermissiveState = 'Stopped'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'print.spooler_service'; Type = 'ServiceConfig'; Name = 'Spooler'; PermissiveStartup = 'auto'; PermissiveState = 'Running'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'localuser.administrator'; Type = 'LocalUser'; Rid = 500; Name = 'Built-in local administrator'; PermissiveValue = $true; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'accounts.limit_blank_password_use'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'LimitBlankPasswordUse'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'logon.cached_domain_logons'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'; Name = 'CachedLogonsCount'; ValueKind = 'String'; PermissiveExists = $true; PermissiveValue = '50'; RequiresReboot = $true }

        [PSCustomObject]@{ Id = 'uac.prompt_on_secure_desktop'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'PromptOnSecureDesktop'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'uac.enable_lua'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'EnableLUA'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'uac.consent_prompt_behavior_admin'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'ConsentPromptBehaviorAdmin'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'uac.local_account_token_filter_policy'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'LocalAccountTokenFilterPolicy'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'rdp.allow_connections'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'; Name = 'fDenyTSConnections'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'; Name = 'UserAuthentication'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'rdp.security_layer'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'; Name = 'SecurityLayer'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'rdp.min_encryption_level'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'; Name = 'MinEncryptionLevel'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'rdp.listener_enabled'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'; Name = 'fEnableWinStation'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'rdp.allow_clipboard_redirection'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'; Name = 'fDisableClip'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'rdp.allow_drive_redirection'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'; Name = 'fDisableCdm'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'rdp.firewall_rules'; Type = 'FirewallRules'; Group = '@FirewallAPI.dll,-28752'; PermissiveEnabled = $true; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'wsh.enabled'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows Script Host\Settings'; Name = 'Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'smartscreen.enable'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Name = 'EnableSmartScreen'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'sehop.disable_exception_chain_validation'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\kernel'; Name = 'DisableExceptionChainValidation'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'exploit_protection.policy'; Type = 'ExploitProtectionPolicy'; RequiresReboot = $true }

        [PSCustomObject]@{ Id = 'lsa.run_as_ppl'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'RunAsPPL'; ValueKind = 'DWord'; PermissiveExists = $false; PermissiveValue = $null; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'lsa.no_lm_hash'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'NoLMHash'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'lsa.lsa_cfg_flags'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'LsaCfgFlags'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'rdp.disable_restricted_admin'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'DisableRestrictedAdmin'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }

        [PSCustomObject]@{ Id = 'deviceguard.enable_vbs'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard'; Name = 'EnableVirtualizationBasedSecurity'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'deviceguard.require_platform_security_features'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard'; Name = 'RequirePlatformSecurityFeatures'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'hvci.enabled'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity'; Name = 'Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'wdigest.use_logon_credential'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest'; Name = 'UseLogonCredential'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'ntlm.lm_compatibility_level'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'LmCompatibilityLevel'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'ntlm.minimum_client_session_security'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0'; Name = 'NTLMMinClientSec'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'ntlm.minimum_server_session_security'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\MSV1_0'; Name = 'NTLMMinServerSec'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'ldap.client_signing'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\LDAP'; Name = 'LDAPClientIntegrity'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'network.netbios_adapters'; Type = 'NetBiosAdapters'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'network.restrict_anonymous'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'RestrictAnonymous'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'network.restrict_anonymous_sam'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'RestrictAnonymousSAM'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'network.everyone_includes_anonymous'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'EveryoneIncludesAnonymous'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'wpad.disable_wpad'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp'; Name = 'DisableWpad'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'wpad.user_auto_detect'; Type = 'LoadedUserRegistryValues'; Items = @([PSCustomObject]@{ RelativePath = 'Software\Microsoft\Windows\CurrentVersion\Internet Settings'; Name = 'AutoDetect'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1 }); RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'office.block_macros_from_internet'; Type = 'LoadedUserRegistryValues'; Items = @($officeMacroItems); RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'llmnr.enable_multicast'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient'; Name = 'EnableMulticast'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'mdns.enable'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters'; Name = 'EnableMDNS'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'telemetry.allow'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'; Name = 'AllowTelemetry'; ValueKind = 'DWord'; PermissiveExists = $false; PermissiveValue = $null; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'audit.process_cmdline'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit'; Name = 'ProcessCreationIncludeCmdLine_Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'audit.process_creation'; Type = 'AuditPolicy'; Subcategory = 'Process Creation'; SubcategoryGuid = '{0CCE922B-69AE-11D9-BED3-505054503030}'; PermissiveSuccess = $false; PermissiveFailure = $false; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'print.register_spooler_remote_rpc_endpoint'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers'; Name = 'RegisterSpoolerRemoteRpcEndPoint'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'winrm.service'; Type = 'ServiceConfig'; Name = 'WinRM'; PermissiveStartup = 'auto'; PermissiveState = 'Running'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.allow_unencrypted'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\AllowUnencrypted'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.allow_unencrypted'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\AllowUnencrypted'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.basic'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Basic'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.basic'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Basic'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.digest'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Digest'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.certificate'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Certificate'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.certificate'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Certificate'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.kerberos'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Kerberos'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.kerberos'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Kerberos'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.negotiate'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Negotiate'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.negotiate'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Negotiate'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.credssp'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\CredSSP'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.credssp'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\CredSSP'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.cbt_hardening_level'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\CbtHardeningLevel'; PermissiveValue = 'Relaxed'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.ipv4_filter'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\IPv4Filter'; PermissiveValue = '*'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.ipv6_filter'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\IPv6Filter'; PermissiveValue = '*'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.trusted_hosts'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\TrustedHosts'; PermissiveValue = '*'; RequiresReboot = $false }
        [PSCustomObject]@{
            Id             = 'winrm.listeners'
            Type           = 'WinRmListeners'
            PermissiveValue = @(
                [PSCustomObject]@{
                    Address               = '*'
                    Transport             = 'HTTP'
                    Port                  = 5985
                    Hostname              = ''
                    Enabled               = $true
                    URLPrefix             = 'wsman'
                    CertificateThumbprint = ''
                }
            )
            RequiresReboot = $false
        }

        [PSCustomObject]@{ Id = 'smb.client.require_security_signature'; Type = 'SmbClientConfig'; PermissiveValue = $false; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'smb.server.require_security_signature'; Type = 'SmbServerConfig'; PermissiveValue = $false; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'smb.client.allow_insecure_guest_auth'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation'; Name = 'AllowInsecureGuestAuth'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'smb.client.require_encryption'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\LanmanWorkstation'; Name = 'RequireEncryption'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'bitlocker.volumes'; Type = 'BitLockerVolumes'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'wdac.policies'; Type = 'WdacPolicies'; RequiresReboot = $true }
    )

    $script:WinDefStateDefinitionCache
}

function Get-DefenseDefinitionMap {
    if ($null -ne $script:WinDefStateDefinitionMapCache) {
        return $script:WinDefStateDefinitionMapCache
    }

    $definitionMap = @{}
    foreach ($definition in Get-DefenseDefinitions) {
        $definitionMap[[string]$definition.Id] = $definition
    }
    $script:WinDefStateDefinitionMapCache = $definitionMap
    $script:WinDefStateDefinitionMapCache
}

function Capture-Definition {
    param(
        [Parameter(Mandatory)] [object]$Definition,
        [AllowNull()] [object]$CaptureSession
    )

    switch ($Definition.Type) {
        'RegistryValue' {
            $state = Get-RegistryValueCaptureState -Path $Definition.Path -Name $Definition.Name -DefaultValueKind $Definition.ValueKind -CaptureSession $CaptureSession

            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'RegistryValue'
                Path           = $Definition.Path
                Name           = $Definition.Name
                ValueKind      = $state.ValueKind
                Captured       = $state.Captured
                CaptureError   = $state.Error
                Exists         = $state.Exists
                CurrentValue   = $state.CurrentValue
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'RegistryKeyFlat' {
            return Capture-RegistryKeyFlatState -Id $Definition.Id -Path $Definition.Path -RequiresReboot $Definition.RequiresReboot
        }
        'MpPreferenceValue' {
            $state = Get-MpPreferencePropertyState -Property $Definition.Property -CaptureSession $CaptureSession

            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'MpPreferenceValue'
                Property       = $Definition.Property
                CommandAvailable = $state.CommandAvailable
                Captured       = $state.Captured
                CaptureError   = $state.Error
                CurrentValue   = $state.Value
                RestoreValue   = Resolve-MpPreferenceValue -Definition $Definition -Value $state.Value
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'MpPreferenceList' {
            $state = Get-MpPreferenceListState -Property $Definition.Property -CaptureSession $CaptureSession

            return [PSCustomObject]@{
                Id               = $Definition.Id
                Type             = 'MpPreferenceList'
                Property         = $Definition.Property
                CurrentValue     = @($state.Items)
                CommandAvailable = $state.CommandAvailable
                Captured         = $state.Captured
                CaptureError     = $state.Error
                RequiresReboot   = $Definition.RequiresReboot
            }
        }
        'DefenderRuntimeStatus' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'DefenderRuntimeStatus'
                CurrentValue   = Get-CaptureSessionValue -Session $CaptureSession -Key 'defender.runtime' -Factory { Get-DefenderRuntimeStatus }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'BitLockerVolumes' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'BitLockerVolumes'
                CurrentValue   = Get-CaptureSessionValue -Session $CaptureSession -Key 'bitlocker.volumes' -Factory { Get-BitLockerVolumeStates }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'ExploitProtectionPolicy' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'ExploitProtectionPolicy'
                CurrentValue   = Get-CaptureSessionValue -Session $CaptureSession -Key 'exploit-protection.policy' -Factory { Get-ExploitProtectionPolicyState }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'WdacPolicies' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'WdacPolicies'
                CurrentValue   = Get-CaptureSessionValue -Session $CaptureSession -Key 'wdac.policies' -Factory { Get-WdacPolicyState }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'AsrRules' {
            $asrState = Get-AsrRuleCaptureState -CaptureSession $CaptureSession
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'AsrRules'
                CommandAvailable = $asrState.CommandAvailable
                Captured       = $asrState.Captured
                CaptureError   = $asrState.Error
                CurrentValue   = @($asrState.Rules)
                InvalidEntries = @($asrState.InvalidEntries)
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'PowerShellModuleLogging' {
            return Capture-PowerShellModuleLoggingState
        }
        'AppLockerPolicy' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'AppLockerPolicy'
                CurrentValue   = Get-CaptureSessionValue -Session $CaptureSession -Key 'applocker.policy' -Factory { Get-AppLockerPolicyState }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'FirewallProfiles' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'FirewallProfiles'
                CurrentValue   = Get-CaptureSessionValue -Session $CaptureSession -Key 'firewall.profiles' -Factory { Get-FirewallProfileStates -Profiles @($Definition.Profiles) }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'FirewallRules' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'FirewallRules'
                CurrentValue   = Get-CaptureSessionValue -Session $CaptureSession -Key ("firewall.rules:{0}" -f ([string]$Definition.Group)) -Factory { Get-FirewallRuleGroupState -Group ([string]$Definition.Group) }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'NetBiosAdapters' {
            $state = Get-CaptureSessionValue -Session $CaptureSession -Key 'network.netbios-adapters' -Factory { Get-NetBiosAdapterCaptureState }
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'NetBiosAdapters'
                CommandAvailable = $state.CommandAvailable
                Captured       = $state.Captured
                CaptureError   = $state.Error
                CurrentValue   = @($state.Adapters)
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'LoadedUserRegistryValues' {
            $currentValue = Get-CaptureSessionValue -Session $CaptureSession -Key ("user.registry.values:{0}" -f ([string]$Definition.Id)) -Factory { Get-LoadedUserRegistryValueStates -Items @($Definition.Items) -CaptureSession $CaptureSession }
            Complete-UserRegistrySessionDefinition -Session $CaptureSession

            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'LoadedUserRegistryValues'
                CurrentValue   = $currentValue
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'MachineEnvironmentValue' {
            $value = [Environment]::GetEnvironmentVariable($Definition.Name, 'Machine')
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'MachineEnvironmentValue'
                Name           = $Definition.Name
                Exists         = $null -ne $value
                CurrentValue   = $value
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'ServiceConfig' {
            $service = Get-ServiceCaptureState -Name ([string]$Definition.Name) -CaptureSession $CaptureSession
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'ServiceConfig'
                Name           = $Definition.Name
                Captured       = $null -ne $service
                CaptureError   = if ($null -eq $service) { "Service '$($Definition.Name)' was not returned by Win32_Service." } else { $null }
                CurrentValue   = [PSCustomObject]@{
                    StartMode = if ($null -ne $service) { $service.StartMode } else { $null }
                    State     = if ($null -ne $service) { $service.State } else { $null }
                }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'LocalUser' {
            $user = Get-CaptureSessionValue -Session $CaptureSession -Key ("local-user:{0}" -f ([string]$Definition.Rid)) -Factory {
                Resolve-LocalUserTarget -Reference $Definition
            }
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'LocalUser'
                Name           = if ($null -ne $user) { $user.Name } else { $Definition.Name }
                Sid            = if ($null -ne $user -and $null -ne $user.SID) { $user.SID.Value } else { $null }
                Rid            = if ($Definition.PSObject.Properties['Rid']) { $Definition.Rid } else { $null }
                Captured       = $null -ne $user -and $null -ne $user.SID -and $null -ne $user.Enabled
                CaptureError   = if ($null -eq $user -or $null -eq $user.SID) { "Local user RID $($Definition.Rid) could not be resolved." } else { $null }
                CurrentValue   = if ($null -ne $user) { $user.Enabled } else { $null }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'WsManValue' {
            $state = Get-WsManConfigValueState -Path $Definition.Path -CaptureSession $CaptureSession
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'WsManValue'
                Path           = $Definition.Path
                CurrentValue   = $state.Value
                CommandAvailable = $state.CommandAvailable
                Captured       = $state.Captured
                CaptureError   = $state.Error
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'WinRmListeners' {
            $state = Get-WinRmListenerStates -CaptureSession $CaptureSession
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'WinRmListeners'
                CurrentValue   = @($state.Listeners)
                CommandAvailable = $state.CommandAvailable
                Captured       = $state.Captured
                CaptureError   = $state.Error
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'AuditPolicy' {
            $state = Get-CaptureSessionValue -Session $CaptureSession -Key ("audit:{0}" -f ([string]$Definition.SubcategoryGuid)) -Factory {
                Get-AuditPolicyState -Subcategory $Definition.Subcategory -SubcategoryGuid $Definition.SubcategoryGuid
            }
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'AuditPolicy'
                Subcategory    = $Definition.Subcategory
                SubcategoryGuid = $Definition.SubcategoryGuid
                CurrentValue   = $state
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'SmbClientConfig' {
            $smbStates = Get-CaptureSessionValue -Session $CaptureSession -Key 'smb.configurations' -Factory { Get-SmbConfigurationStates }
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'SmbClientConfig'
                CurrentValue   = $smbStates.Client
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'SmbServerConfig' {
            $smbStates = Get-CaptureSessionValue -Session $CaptureSession -Key 'smb.configurations' -Factory { Get-SmbConfigurationStates }
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'SmbServerConfig'
                CurrentValue   = $smbStates.Server
                RequiresReboot = $Definition.RequiresReboot
            }
        }
    }
}

function New-PermissiveTargetDescriptor {
    param(
        [Parameter(Mandatory)] [string]$Summary,
        [ValidateSet('Exact', 'BaselineDependent')] [string]$Mode = 'Exact'
    )

    [PSCustomObject]@{
        Mode    = $Mode
        Summary = $Summary
    }
}

function Get-DefinitionPermissiveTargetDescriptor {
    param(
        [Parameter(Mandatory)] [object]$Definition,
        [AllowNull()] [object]$Entry
    )

    if (-not (Test-DefinitionHasPermissiveAction -Definition $Definition)) {
        return $null
    }

    $type = [string]$Definition.Type
    switch ($type) {
        'RegistryValue' {
            $summary = if ([bool]$Definition.PermissiveExists) {
                "Set $($Definition.Name) to $(ConvertTo-DisplayString -Value $Definition.PermissiveValue) ($($Definition.ValueKind))."
            } else {
                "Remove the $($Definition.Name) registry value."
            }
            return (New-PermissiveTargetDescriptor -Summary $summary)
        }
        'RegistryKeyFlat' {
            return (New-PermissiveTargetDescriptor -Summary 'Remove this policy registry key and its flat values.')
        }
        'PowerShellModuleLogging' {
            return (New-PermissiveTargetDescriptor -Summary 'Remove the PowerShell module-logging policy key and module-name filters.')
        }
        'MachineEnvironmentValue' {
            $summary = if ([bool]$Definition.PermissiveExists) {
                "Set the machine environment value to $(ConvertTo-DisplayString -Value $Definition.PermissiveValue)."
            } else {
                'Remove the machine environment value.'
            }
            return (New-PermissiveTargetDescriptor -Summary $summary)
        }
        'MpPreferenceValue' {
            $resolvedValue = Resolve-MpPreferenceValue -Definition $Definition -Value $Definition.PermissiveValue
            return (New-PermissiveTargetDescriptor -Summary ("Set Defender {0} to {1}." -f $Definition.Property, (ConvertTo-DisplayString -Value $resolvedValue)))
        }
        'MpPreferenceList' {
            $itemCount = @($Definition.PermissiveValue).Count
            return (New-PermissiveTargetDescriptor -Summary ("Replace Defender {0} with an exact {1}-item list." -f $Definition.Property, $itemCount))
        }
        'AsrRules' {
            return (New-PermissiveTargetDescriptor -Summary 'Remove every completely captured configured ASR rule.' -Mode BaselineDependent)
        }
        'ServiceConfig' {
            $startMode = ConvertTo-PermissiveServiceStartMode -Value $Definition.PermissiveStartup
            return (New-PermissiveTargetDescriptor -Summary ("Set service startup to {0} and running state to {1}." -f $startMode, $Definition.PermissiveState))
        }
        'LocalUser' {
            $enabledText = if ([bool]$Definition.PermissiveValue) { 'enabled' } else { 'disabled' }
            return (New-PermissiveTargetDescriptor -Summary ("Set the resolved local account to {0}." -f $enabledText) -Mode BaselineDependent)
        }
        'NetBiosAdapters' {
            $adapterCount = if ($null -ne $Entry -and $Entry.PSObject.Properties['CurrentValue']) { @($Entry.CurrentValue).Count } else { 0 }
            return (New-PermissiveTargetDescriptor -Summary ("Enable NetBIOS over TCP/IP on {0} captured IP-enabled adapter(s)." -f $adapterCount) -Mode BaselineDependent)
        }
        'AuditPolicy' {
            return (New-PermissiveTargetDescriptor -Summary ("Set audit success={0} and failure={1}." -f ([bool]$Definition.PermissiveSuccess), ([bool]$Definition.PermissiveFailure)))
        }
        'WsManValue' {
            return (New-PermissiveTargetDescriptor -Summary ("Set the WSMan value to {0}." -f (ConvertTo-DisplayString -Value $Definition.PermissiveValue)))
        }
        'WinRmListeners' {
            return (New-PermissiveTargetDescriptor -Summary 'Replace listeners with one enabled HTTP listener on all addresses at port 5985.')
        }
        'SmbClientConfig' {
            return (New-PermissiveTargetDescriptor -Summary ("Set SMB client RequireSecuritySignature to {0}." -f ([bool]$Definition.PermissiveValue)))
        }
        'SmbServerConfig' {
            return (New-PermissiveTargetDescriptor -Summary ("Set SMB server RequireSecuritySignature to {0}." -f ([bool]$Definition.PermissiveValue)))
        }
        'LoadedUserRegistryValues' {
            $state = if ($null -ne $Entry -and $Entry.PSObject.Properties['CurrentValue']) { Normalize-UserRegistryValueState -State $Entry.CurrentValue } else { $null }
            $profileCount = if ($null -ne $state) { @($state.Entries | ForEach-Object { [string]$_.Sid } | Sort-Object -Unique).Count } else { 0 }
            return (New-PermissiveTargetDescriptor -Summary ("Apply {0} exact user-scoped registry target(s) across {1} captured profile(s)." -f @($Definition.Items).Count, $profileCount) -Mode BaselineDependent)
        }
        'FirewallProfiles' {
            return (New-PermissiveTargetDescriptor -Summary 'Disable Domain, Private, and Public profiles; allow default inbound/outbound traffic and disable profile logging/notifications.')
        }
        'FirewallRules' {
            $ruleCount = 0
            if ($null -ne $Entry -and $Entry.PSObject.Properties['CurrentValue'] -and $null -ne $Entry.CurrentValue) {
                $ruleState = Normalize-FirewallRuleState -State $Entry.CurrentValue
                $ruleCount = @($ruleState.Rules).Count
            }
            return (New-PermissiveTargetDescriptor -Summary ("Enable {0} captured rule(s) in the configured firewall group." -f $ruleCount) -Mode BaselineDependent)
        }
        'BitLockerVolumes' {
            $volumeCount = if ($null -ne $Entry -and $Entry.PSObject.Properties['CurrentValue'] -and $null -ne $Entry.CurrentValue -and $Entry.CurrentValue.PSObject.Properties['Volumes']) { @($Entry.CurrentValue.Volumes).Count } else { 0 }
            return (New-PermissiveTargetDescriptor -Summary ("Suspend protection on eligible captured volumes and enable supported data-volume auto-unlock across {0} volume(s)." -f $volumeCount) -Mode BaselineDependent)
        }
        'AppLockerPolicy' {
            return (New-PermissiveTargetDescriptor -Summary 'Replace the local AppLocker policy with an empty policy; effective-policy precedence remains verified separately.')
        }
        'ExploitProtectionPolicy' {
            return (New-PermissiveTargetDescriptor -Summary 'Apply the bundled permissive exploit-protection system policy.')
        }
        'WdacPolicies' {
            return (New-PermissiveTargetDescriptor -Summary 'Remove identified non-platform WDAC policies while preserving platform-managed and unclassified policy state.' -Mode BaselineDependent)
        }
        default {
            return (New-PermissiveTargetDescriptor -Summary ("Apply the registered permissive action for provider type {0}." -f $type) -Mode BaselineDependent)
        }
    }
}

function Invoke-TimedDefinitionCapture {
    param(
        [Parameter(Mandatory)] [object]$Definition,
        [AllowNull()] [object]$CaptureSession
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $errorMessage = $null
    try {
        $entry = Capture-Definition -Definition $Definition -CaptureSession $CaptureSession
        $entry | Add-Member -NotePropertyName Capabilities -NotePropertyValue (Get-DefinitionCapabilities -Definition $Definition) -Force
        $entry | Add-Member -NotePropertyName PermissiveTarget -NotePropertyValue (Get-DefinitionPermissiveTargetDescriptor -Definition $Definition -Entry $entry) -Force
        return $entry
    } catch {
        $errorMessage = $_.Exception.Message
        throw
    } finally {
        $stopwatch.Stop()
        Add-CaptureSessionSettingTiming `
            -Session $CaptureSession `
            -Id ([string]$Definition.Id) `
            -Type ([string]$Definition.Type) `
            -DurationMs $stopwatch.Elapsed.TotalMilliseconds `
            -Succeeded ([string]::IsNullOrWhiteSpace($errorMessage)) `
            -ErrorMessage $errorMessage
        Write-Verbose ("Captured {0} in {1:N1} ms" -f ([string]$Definition.Id), $stopwatch.Elapsed.TotalMilliseconds)
    }
}

function Invoke-TimedSettingOperation {
    param(
        [Parameter(Mandatory)] [ValidateSet('Permissive', 'Restore')] [string]$Phase,
        [Parameter(Mandatory)] [string]$Id,
        [Parameter(Mandatory)] [string]$Type,
        [Parameter(Mandatory)] [object]$Session,
        [Parameter(Mandatory)] [scriptblock]$Action
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $errorMessage = $null
    try {
        & $Action
    } catch {
        $errorMessage = $_.Exception.Message
        throw
    } finally {
        $stopwatch.Stop()
        Add-CaptureSessionSettingTiming `
            -Session $Session `
            -Id $Id `
            -Type $Type `
            -DurationMs $stopwatch.Elapsed.TotalMilliseconds `
            -Succeeded ([string]::IsNullOrWhiteSpace($errorMessage)) `
            -ErrorMessage $errorMessage
        Write-Verbose ("{0} {1} in {2:N1} ms" -f $Phase, $Id, $stopwatch.Elapsed.TotalMilliseconds)
    }
}

function Test-DefinitionHasPermissiveAction {
    param([Parameter(Mandatory)] [object]$Definition)

    switch ([string]$Definition.Type) {
        'DefenderRuntimeStatus' { return $false }
        'MpPreferenceList' { return ($null -ne $Definition.PSObject.Properties['PermissiveValue']) }
        default { return $true }
    }
}

function Test-DefinitionHasRestoreAction {
    param([Parameter(Mandatory)] [object]$Definition)

    switch ([string]$Definition.Type) {
        'RegistryValue' { return $true }
        'RegistryKeyFlat' { return $true }
        'MpPreferenceValue' { return $true }
        'MpPreferenceList' { return $true }
        'DefenderRuntimeStatus' { return $false }
        'BitLockerVolumes' { return $true }
        'ExploitProtectionPolicy' { return $true }
        'WdacPolicies' { return $true }
        'AsrRules' { return $true }
        'PowerShellModuleLogging' { return $true }
        'AppLockerPolicy' { return $true }
        'FirewallProfiles' { return $true }
        'FirewallRules' { return $true }
        'NetBiosAdapters' { return $true }
        'LoadedUserRegistryValues' { return $true }
        'MachineEnvironmentValue' { return $true }
        'ServiceConfig' { return $true }
        'LocalUser' { return $true }
        'WsManValue' { return $true }
        'WinRmListeners' { return $true }
        'AuditPolicy' { return $true }
        'SmbClientConfig' { return $true }
        'SmbServerConfig' { return $true }
        default { return $false }
    }
}

function Get-DefinitionCapabilities {
    param([Parameter(Mandatory)] [object]$Definition)

    $permissive = Test-DefinitionHasPermissiveAction -Definition $Definition
    $restore = Test-DefinitionHasRestoreAction -Definition $Definition
    [PSCustomObject]@{
        Permissive    = $permissive
        Restore       = $restore
        InventoryOnly = -not $permissive -and -not $restore
    }
}

function Test-CanonicalStateEqual {
    param(
        [AllowNull()] [object]$Expected,
        [AllowNull()] [object]$Actual
    )

    (Get-CanonicalJson -Value (ConvertTo-CanonicalValue -Value $Expected)) -eq
        (Get-CanonicalJson -Value (ConvertTo-CanonicalValue -Value $Actual))
}

function New-PermissiveTargetEvaluation {
    param(
        [Parameter(Mandatory)] [bool]$IsMatch,
        [AllowNull()] [object]$Expected,
        [AllowNull()] [object]$Actual,
        [string]$Reason,
        [bool]$PendingReboot = $false,
        [AllowNull()] [Nullable[bool]]$Changed
    )

    [PSCustomObject]@{
        Matches       = $IsMatch
        Expected      = $Expected
        Actual        = $Actual
        Reason        = $Reason
        PendingReboot = $PendingReboot
        Changed       = $Changed
    }
}

function ConvertTo-PermissiveServiceStartMode {
    param([AllowNull()] [object]$Value)

    switch (([string]$Value).Trim().ToLowerInvariant()) {
        'auto' { return 'Auto' }
        'automatic' { return 'Auto' }
        'delayed-auto' { return 'Auto' }
        'demand' { return 'Manual' }
        'manual' { return 'Manual' }
        'disabled' { return 'Disabled' }
        default { return [string]$Value }
    }
}

function ConvertTo-PermissiveWsManValue {
    param([AllowNull()] [object]$Value)

    $converted = ConvertFrom-WsManTextValue -Value $Value
    if ($converted -is [string]) {
        return ([string]$converted).Trim().ToLowerInvariant()
    }

    $converted
}

function Get-PermissiveExploitProtectionEvaluation {
    param([Parameter(Mandatory)] [object]$LiveEntry)

    $xml = Get-ExploitProtectionPolicyXml -State $LiveEntry.CurrentValue
    $requirements = @(
        [PSCustomObject]@{ Node = 'DEP'; Attribute = 'Enable' }
        [PSCustomObject]@{ Node = 'DEP'; Attribute = 'EmulateAtlThunks' }
        [PSCustomObject]@{ Node = 'ControlFlowGuard'; Attribute = 'Enable' }
        [PSCustomObject]@{ Node = 'ASLR'; Attribute = 'ForceRelocateImages' }
        [PSCustomObject]@{ Node = 'ASLR'; Attribute = 'BottomUp' }
        [PSCustomObject]@{ Node = 'ASLR'; Attribute = 'HighEntropy' }
        [PSCustomObject]@{ Node = 'SEHOP'; Attribute = 'Enable' }
    )
    $expected = @(
        foreach ($requirement in $requirements) {
            [PSCustomObject]@{
                Setting = "$($requirement.Node).$($requirement.Attribute)"
                Value   = $false
            }
        }
    )

    try {
        $document = New-Object System.Xml.XmlDocument
        $document.PreserveWhitespace = $false
        $document.LoadXml($xml)
    } catch {
        return New-PermissiveTargetEvaluation -IsMatch $false -Expected $expected -Actual $xml -Reason "Exploit protection returned invalid XML: $($_.Exception.Message)"
    }

    $actual = @(
        foreach ($requirement in $requirements) {
            $node = $document.SelectSingleNode("/MitigationPolicy/SystemConfig/$($requirement.Node)")
            $rawValue = if ($null -ne $node -and $node.Attributes[$requirement.Attribute]) {
                [string]$node.Attributes[$requirement.Attribute].Value
            } else {
                $null
            }
            [PSCustomObject]@{
                Setting = "$($requirement.Node).$($requirement.Attribute)"
                Value   = ConvertTo-NullableBoolean -Value $rawValue
            }
        }
    )
    $isMatch = @($actual | Where-Object { $null -eq $_.Value -or [bool]$_.Value }).Count -eq 0
    $reason = if ($isMatch) { $null } else { 'One or more system exploit mitigations were not reported as explicitly disabled.' }
    New-PermissiveTargetEvaluation -IsMatch $isMatch -Expected $expected -Actual $actual -Reason $reason
}

function Get-PermissiveBitLockerEvaluation {
    param(
        [Parameter(Mandatory)] [object]$LiveEntry,
        [AllowNull()] [object]$BaselineEntry
    )

    if ($null -eq $BaselineEntry) {
        return New-PermissiveTargetEvaluation -IsMatch $false -Expected $null -Actual $LiveEntry.CurrentValue -Reason 'The BitLocker permissive target requires the captured baseline.'
    }

    $baseline = Normalize-BitLockerState -State $BaselineEntry.CurrentValue
    $live = Normalize-BitLockerState -State $LiveEntry.CurrentValue
    $liveByMountPoint = @{}
    foreach ($volume in @($live.Volumes)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$volume.MountPoint)) {
            $liveByMountPoint[[string]$volume.MountPoint] = $volume
        }
    }

    $expected = [System.Collections.Generic.List[object]]::new()
    $actual = [System.Collections.Generic.List[object]]::new()
    $isMatch = $true
    foreach ($baselineVolume in @($baseline.Volumes)) {
        $mountPoint = [string]$baselineVolume.MountPoint
        $targetProtection = if (Test-BitLockerProtectionEnabled -Value $baselineVolume.ProtectionStatus) {
            'Off'
        } else {
            ConvertTo-BitLockerProtectionStatusLabel -Value $baselineVolume.ProtectionStatus
        }
        $targetAutoUnlock = $baselineVolume.AutoUnlockEnabled
        if (
            (Test-BitLockerAutoUnlockSupportedVolume -Volume $baselineVolume) -and
            $baselineVolume.AutoUnlockEnabled -eq $false -and
            [string]$baselineVolume.ProtectionMode -ne 'Decrypted'
        ) {
            $targetAutoUnlock = $true
        }

        $expected.Add([PSCustomObject]@{
            MountPoint        = $mountPoint
            ProtectionStatus  = $targetProtection
            AutoUnlockEnabled = $targetAutoUnlock
        }) | Out-Null

        $liveVolume = if ($liveByMountPoint.ContainsKey($mountPoint)) { $liveByMountPoint[$mountPoint] } else { $null }
        $actualProtection = if ($null -ne $liveVolume) { ConvertTo-BitLockerProtectionStatusLabel -Value $liveVolume.ProtectionStatus } else { $null }
        $actualAutoUnlock = if ($null -ne $liveVolume) { $liveVolume.AutoUnlockEnabled } else { $null }
        $actual.Add([PSCustomObject]@{
            MountPoint        = $mountPoint
            ProtectionStatus  = $actualProtection
            AutoUnlockEnabled = $actualAutoUnlock
        }) | Out-Null

        if ($null -eq $liveVolume -or $targetProtection -ne $actualProtection) {
            $isMatch = $false
        }
        if ($null -ne $targetAutoUnlock -and $targetAutoUnlock -ne $actualAutoUnlock) {
            $isMatch = $false
        }
    }

    $reason = if ($isMatch) { $null } else { 'One or more mounted BitLocker volumes did not reach the requested suspended/auto-unlock posture.' }
    New-PermissiveTargetEvaluation -IsMatch $isMatch -Expected @($expected) -Actual @($actual) -Reason $reason
}

function Get-PermissiveWdacEvaluation {
    param(
        [Parameter(Mandatory)] [object]$LiveEntry,
        [AllowNull()] [object]$BaselineEntry
    )

    $state = $LiveEntry.CurrentValue
    $rows = @(Get-WdacPolicyReportRows -State $state)
    $platformPolicyIds = @(
        $rows |
            Where-Object { $_.IsPlatformManaged -and -not [string]::IsNullOrWhiteSpace([string]$_.PolicyId) } |
            ForEach-Object { Get-WdacNormalizedPolicyId -Value $_.PolicyId }
    )
    $nonPlatformOnDiskPolicies = @(
        $rows | Where-Object { -not $_.IsPlatformManaged -and $_.HasFileOnDisk -eq $true }
    )
    $activeOnlyPolicies = @(
        $rows | Where-Object { -not $_.IsPlatformManaged -and $_.IsEnforced -eq $true -and $_.HasFileOnDisk -ne $true }
    )
    $nonPlatformFiles = @(
        foreach ($file in @($state.Files)) {
            $filePolicyId = Get-WdacPolicyIdFromFileName -FileName ([string]$file.FileName)
            if (-not [string]::IsNullOrWhiteSpace($filePolicyId) -and $filePolicyId -in $platformPolicyIds) {
                continue
            }

            [PSCustomObject]@{
                RelativePath = [string]$file.RelativePath
                FileName     = [string]$file.FileName
            }
        }
    )

    $expected = [PSCustomObject]@{
        NonPlatformOnDiskPolicies = @()
        NonPlatformPolicyFiles    = @()
    }
    $actual = [PSCustomObject]@{
        NonPlatformOnDiskPolicies = @($nonPlatformOnDiskPolicies | Select-Object PolicyId, FriendlyName, Presence)
        NonPlatformPolicyFiles    = @($nonPlatformFiles)
        ActiveOnlyPolicies        = @($activeOnlyPolicies | Select-Object PolicyId, FriendlyName, Presence)
        PlatformManagedPolicies   = @($rows | Where-Object { $_.IsPlatformManaged } | Select-Object PolicyId, FriendlyName, Classification, Presence)
    }
    $isMatch = $nonPlatformOnDiskPolicies.Count -eq 0 -and $nonPlatformFiles.Count -eq 0
    $pendingReboot = $isMatch -and $activeOnlyPolicies.Count -gt 0
    $changed = $true
    if ($null -ne $BaselineEntry -and $null -ne $BaselineEntry.CurrentValue) {
        $baselineRows = @(Get-WdacPolicyReportRows -State $BaselineEntry.CurrentValue)
        $baselinePlatformPolicyIds = @(
            $baselineRows |
                Where-Object { $_.IsPlatformManaged -and -not [string]::IsNullOrWhiteSpace([string]$_.PolicyId) } |
                ForEach-Object { Get-WdacNormalizedPolicyId -Value $_.PolicyId }
        )
        $baselineRemovalPolicies = @($baselineRows | Where-Object { -not $_.IsPlatformManaged -and ($_.HasFileOnDisk -eq $true -or $_.IsEnforced -eq $true) })
        $baselineRemovalFiles = @(
            foreach ($file in @($BaselineEntry.CurrentValue.Files)) {
                $filePolicyId = Get-WdacPolicyIdFromFileName -FileName ([string]$file.FileName)
                if (-not [string]::IsNullOrWhiteSpace($filePolicyId) -and $filePolicyId -in $baselinePlatformPolicyIds) {
                    continue
                }
                $file
            }
        )
        $changed = $baselineRemovalPolicies.Count -gt 0 -or $baselineRemovalFiles.Count -gt 0
    }
    $reason = if (-not $isMatch) {
        'One or more removable WDAC policies or policy files remain on disk.'
    } elseif ($pendingReboot) {
        'Removable WDAC policy files are gone, but one or more policies remain active until reboot.'
    } else {
        $null
    }
    New-PermissiveTargetEvaluation -IsMatch $isMatch -Expected $expected -Actual $actual -Reason $reason -PendingReboot $pendingReboot -Changed $changed
}

function Test-PermissiveDefinitionState {
    param(
        [Parameter(Mandatory)] [object]$Definition,
        [Parameter(Mandatory)] [object]$LiveEntry,
        [AllowNull()] [object]$BaselineEntry,
        [string]$SnapshotPath
    )

    if (-not (Test-DefinitionHasPermissiveAction -Definition $Definition)) {
        return [PSCustomObject]@{
            Id             = [string]$Definition.Id
            Type           = [string]$Definition.Type
            RequiresReboot = [bool]$Definition.RequiresReboot
            Status         = 'NotApplicable'
            Matches        = $true
            Changed        = $false
            Reason         = 'This definition is capture-only.'
            Expected       = $null
            Actual         = $null
        }
    }

    if (-not (Test-SnapshotEntryCapturedExactly -Entry $LiveEntry)) {
        $captureError = if ($LiveEntry.PSObject.Properties['CaptureError']) { [string]$LiveEntry.CaptureError } else { $null }
        $reason = 'The post-apply provider state could not be captured exactly.'
        if (-not [string]::IsNullOrWhiteSpace($captureError)) {
            $reason = "$reason $captureError"
        }
        return [PSCustomObject]@{
            Id             = [string]$Definition.Id
            Type           = [string]$Definition.Type
            RequiresReboot = [bool]$Definition.RequiresReboot
            Status         = 'Mismatch'
            Matches        = $false
            Changed        = $false
            Reason         = $reason
            Expected       = $null
            Actual         = $LiveEntry
        }
    }

    $evaluation = $null
    switch ([string]$Definition.Type) {
        'RegistryValue' {
            $expected = [PSCustomObject]@{
                Exists       = [bool]$Definition.PermissiveExists
                ValueKind    = [string]$Definition.ValueKind
                CurrentValue = if ($Definition.PermissiveExists) { $Definition.PermissiveValue } else { $null }
            }
            $actual = [PSCustomObject]@{
                Exists       = [bool]$LiveEntry.Exists
                ValueKind    = [string]$LiveEntry.ValueKind
                CurrentValue = if ($LiveEntry.Exists) { $LiveEntry.CurrentValue } else { $null }
            }
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'The registry value does not match its permissive target.'
            break
        }
        'RegistryKeyFlat' {
            $expected = [PSCustomObject]@{ Exists = $false }
            $actual = [PSCustomObject]@{ Exists = [bool]$LiveEntry.Exists }
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (-not [bool]$LiveEntry.Exists) -Expected $expected -Actual $actual -Reason 'The policy registry key still exists.'
            break
        }
        'PowerShellModuleLogging' {
            $expected = [PSCustomObject]@{ Exists = $false }
            $actual = [PSCustomObject]@{ Exists = [bool]$LiveEntry.Exists }
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (-not [bool]$LiveEntry.Exists) -Expected $expected -Actual $actual -Reason 'The PowerShell module logging policy key still exists.'
            break
        }
        'MachineEnvironmentValue' {
            $expected = [PSCustomObject]@{
                Exists       = [bool]$Definition.PermissiveExists
                CurrentValue = if ($Definition.PermissiveExists) { [string]$Definition.PermissiveValue } else { $null }
            }
            $actual = [PSCustomObject]@{
                Exists       = [bool]$LiveEntry.Exists
                CurrentValue = if ($LiveEntry.Exists) { [string]$LiveEntry.CurrentValue } else { $null }
            }
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'The machine environment value does not match its permissive target.'
            break
        }
        'MpPreferenceValue' {
            $expectedValue = Resolve-MpPreferenceValue -Definition $Definition -Value $Definition.PermissiveValue
            $actualValue = $LiveEntry.RestoreValue
            $isMatch = if ($expectedValue -is [string] -and $actualValue -is [string]) {
                [string]::Equals([string]$expectedValue, [string]$actualValue, [System.StringComparison]::OrdinalIgnoreCase)
            } else {
                Test-CanonicalStateEqual -Expected $expectedValue -Actual $actualValue
            }
            $evaluation = New-PermissiveTargetEvaluation -IsMatch $isMatch -Expected $expectedValue -Actual $actualValue -Reason 'The Defender preference does not match its permissive target. Policy precedence or tamper protection may have rejected it.'
            break
        }
        'MpPreferenceList' {
            $expected = @(Normalize-MpPreferenceListItems -Value $Definition.PermissiveValue)
            $actual = @(Normalize-MpPreferenceListItems -Value $LiveEntry.CurrentValue)
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'The Defender list preference does not match its permissive target.'
            break
        }
        'DefenderRuntimeStatus' {
            $evaluation = New-PermissiveTargetEvaluation -IsMatch $true -Expected $null -Actual $LiveEntry.CurrentValue -Reason $null
            break
        }
        'AsrRules' {
            $actual = @($LiveEntry.CurrentValue)
            $isMatch = $actual.Count -eq 0 -and @(Get-AsrInvalidEntriesFromEntry -Entry $LiveEntry).Count -eq 0
            $evaluation = New-PermissiveTargetEvaluation -IsMatch $isMatch -Expected @() -Actual $actual -Reason 'One or more configured ASR rules remain.'
            break
        }
        'ServiceConfig' {
            $expected = [PSCustomObject]@{
                StartMode = ConvertTo-PermissiveServiceStartMode -Value $Definition.PermissiveStartup
                State     = [string]$Definition.PermissiveState
            }
            $actual = [PSCustomObject]@{
                StartMode = ConvertTo-PermissiveServiceStartMode -Value $LiveEntry.CurrentValue.StartMode
                State     = [string]$LiveEntry.CurrentValue.State
            }
            $isMatch = [string]::Equals($expected.StartMode, $actual.StartMode, [System.StringComparison]::OrdinalIgnoreCase) -and [string]::Equals($expected.State, $actual.State, [System.StringComparison]::OrdinalIgnoreCase)
            $evaluation = New-PermissiveTargetEvaluation -IsMatch $isMatch -Expected $expected -Actual $actual -Reason 'The service startup or running state does not match its permissive target.'
            break
        }
        'LocalUser' {
            $expected = [bool]$Definition.PermissiveValue
            $actual = [bool]$LiveEntry.CurrentValue
            $evaluation = New-PermissiveTargetEvaluation -IsMatch ($expected -eq $actual) -Expected $expected -Actual $actual -Reason 'The local user enabled state does not match its permissive target.'
            break
        }
        'NetBiosAdapters' {
            $actual = @(
                foreach ($adapter in @($LiveEntry.CurrentValue | Sort-Object -Property Index)) {
                    [PSCustomObject]@{ Index = [int]$adapter.Index; TcpipNetbiosOptions = [int]$adapter.TcpipNetbiosOptions }
                }
            )
            $expected = @(
                foreach ($adapter in $actual) {
                    [PSCustomObject]@{ Index = [int]$adapter.Index; TcpipNetbiosOptions = 1 }
                }
            )
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'One or more IP-enabled adapters do not have NetBIOS enabled.'
            break
        }
        'AuditPolicy' {
            $expected = [PSCustomObject]@{ Success = [bool]$Definition.PermissiveSuccess; Failure = [bool]$Definition.PermissiveFailure }
            $actual = [PSCustomObject]@{ Success = [bool]$LiveEntry.CurrentValue.Success; Failure = [bool]$LiveEntry.CurrentValue.Failure }
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'The audit policy does not match its permissive target.'
            break
        }
        'WsManValue' {
            $expected = ConvertTo-PermissiveWsManValue -Value $Definition.PermissiveValue
            $actual = ConvertTo-PermissiveWsManValue -Value $LiveEntry.CurrentValue
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'The WSMan value does not match its permissive target.'
            break
        }
        'WinRmListeners' {
            $expectedEntry = [PSCustomObject]@{
                Id = $Definition.Id; Type = $Definition.Type; CommandAvailable = $true; Captured = $true
                CurrentValue = @($Definition.PermissiveValue); RequiresReboot = $Definition.RequiresReboot
            }
            $expected = ConvertTo-ComparableSnapshotEntry -Entry $expectedEntry
            $actual = ConvertTo-ComparableSnapshotEntry -Entry $LiveEntry -ReferenceEntry $expectedEntry
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'The WinRM listener inventory does not match its permissive target.'
            break
        }
        'SmbClientConfig' {
            $state = Normalize-SmbConfigState -Value $LiveEntry.CurrentValue
            $expected = [bool]$Definition.PermissiveValue
            $actual = $state.RequireSecuritySignature
            $evaluation = New-PermissiveTargetEvaluation -IsMatch ($null -ne $actual -and $expected -eq [bool]$actual) -Expected $expected -Actual $actual -Reason 'The SMB client signing requirement does not match its permissive target.'
            break
        }
        'SmbServerConfig' {
            $state = Normalize-SmbConfigState -Value $LiveEntry.CurrentValue
            $expected = [bool]$Definition.PermissiveValue
            $actual = $state.RequireSecuritySignature
            $evaluation = New-PermissiveTargetEvaluation -IsMatch ($null -ne $actual -and $expected -eq [bool]$actual) -Expected $expected -Actual $actual -Reason 'The SMB server signing requirement does not match its permissive target.'
            break
        }
        'LoadedUserRegistryValues' {
            $state = Normalize-UserRegistryValueState -State $LiveEntry.CurrentValue
            $actual = @(
                foreach ($value in @($state.Entries | Sort-Object -Property Sid, RelativePath, Name)) {
                    [PSCustomObject]@{
                        Sid          = [string]$value.Sid
                        RelativePath = [string]$value.RelativePath
                        Name         = [string]$value.Name
                        Exists       = [bool]$value.Exists
                        CurrentValue = if ($value.Exists) { $value.CurrentValue } else { $null }
                        ValueKind    = [string]$value.ValueKind
                    }
                }
            )
            $baselineState = if ($null -ne $BaselineEntry) { Normalize-UserRegistryValueState -State $BaselineEntry.CurrentValue } else { $null }
            $actualSids = @($actual | ForEach-Object { [string]$_.Sid })
            $baselineSids = if ($null -ne $baselineState) { @($baselineState.Entries | ForEach-Object { [string]$_.Sid }) } else { @() }
            $targetSids = @(
                @(@($actualSids) + @($baselineSids)) |
                    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                    Sort-Object -Unique
            )
            $expected = @(@(
                foreach ($sid in $targetSids) {
                    foreach ($item in @($Definition.Items)) {
                    [PSCustomObject]@{
                            Sid          = [string]$sid
                            RelativePath = [string]$item.RelativePath
                            Name         = [string]$item.Name
                            Exists       = [bool]$item.PermissiveExists
                            CurrentValue = if ($item.PermissiveExists) { $item.PermissiveValue } else { $null }
                            ValueKind    = [string]$item.ValueKind
                        }
                    }
                }
            ) | Sort-Object -Property Sid, RelativePath, Name)
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'One or more user-scoped registry values do not match their permissive targets.'
            break
        }
        'FirewallProfiles' {
            $state = Normalize-FirewallProfileState -State $LiveEntry.CurrentValue
            $actual = @(
                foreach ($firewallProfile in @($state.Profiles | Sort-Object -Property Profile)) {
                    [PSCustomObject]@{
                        Profile                         = [string]$firewallProfile.Profile
                        Enabled                         = [string]$firewallProfile.Enabled
                        DefaultInboundAction            = [string]$firewallProfile.DefaultInboundAction
                        DefaultOutboundAction           = [string]$firewallProfile.DefaultOutboundAction
                        AllowUnicastResponseToMulticast = [string]$firewallProfile.AllowUnicastResponseToMulticast
                        NotifyOnListen                  = [string]$firewallProfile.NotifyOnListen
                        LogAllowed                      = [string]$firewallProfile.LogAllowed
                        LogBlocked                      = [string]$firewallProfile.LogBlocked
                        LogIgnored                      = [string]$firewallProfile.LogIgnored
                    }
                }
            )
            $expected = @(
                foreach ($profileName in @($Definition.Profiles | Sort-Object)) {
                    [PSCustomObject]@{
                        Profile                         = [string]$profileName
                        Enabled                         = ConvertTo-FirewallProfileEnabledValue -Value $Definition.PermissiveValue
                        DefaultInboundAction            = ConvertTo-FirewallProfileActionValue -Value $Definition.PermissiveDefaultInboundAction
                        DefaultOutboundAction           = ConvertTo-FirewallProfileActionValue -Value $Definition.PermissiveDefaultOutboundAction
                        AllowUnicastResponseToMulticast = ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveAllowUnicastResponseToMulticast
                        NotifyOnListen                  = ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveNotifyOnListen
                        LogAllowed                      = ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveLogAllowed
                        LogBlocked                      = ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveLogBlocked
                        LogIgnored                      = ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveLogIgnored
                    }
                }
            )
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'One or more firewall profiles do not match their permissive targets.'
            break
        }
        'FirewallRules' {
            $state = Normalize-FirewallRuleState -State $LiveEntry.CurrentValue
            $actual = @(
                foreach ($rule in @($state.Rules | Sort-Object -Property Name)) {
                    [PSCustomObject]@{ Name = [string]$rule.Name; Enabled = [string]$rule.Enabled }
                }
            )
            $expected = @(
                foreach ($rule in $actual) {
                    [PSCustomObject]@{ Name = $rule.Name; Enabled = ConvertTo-FirewallProfileEnabledValue -Value $Definition.PermissiveEnabled }
                }
            )
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'One or more firewall rules do not match their permissive enabled state.'
            break
        }
        'BitLockerVolumes' {
            $evaluation = Get-PermissiveBitLockerEvaluation -LiveEntry $LiveEntry -BaselineEntry $BaselineEntry
            break
        }
        'AppLockerPolicy' {
            $emptyXml = Normalize-AppLockerXml -Xml (Get-EmptyAppLockerPolicyXml)
            $expected = [PSCustomObject]@{ LocalXml = $emptyXml; EffectiveXml = $emptyXml }
            $actual = [PSCustomObject]@{
                LocalXml     = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $LiveEntry.CurrentValue -PolicyScope Local)
                EffectiveXml = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $LiveEntry.CurrentValue -PolicyScope Effective)
            }
            $evaluation = New-PermissiveTargetEvaluation -IsMatch (Test-CanonicalStateEqual -Expected $expected -Actual $actual) -Expected $expected -Actual $actual -Reason 'The local/effective AppLocker policy is not empty.'
            break
        }
        'ExploitProtectionPolicy' {
            $evaluation = Get-PermissiveExploitProtectionEvaluation -LiveEntry $LiveEntry
            break
        }
        'WdacPolicies' {
            $evaluation = Get-PermissiveWdacEvaluation -LiveEntry $LiveEntry -BaselineEntry $BaselineEntry
            break
        }
        default {
            $evaluation = New-PermissiveTargetEvaluation -IsMatch $false -Expected $null -Actual $LiveEntry -Reason "No permissive verification contract exists for type '$($Definition.Type)'."
            break
        }
    }

    $changed = $true
    if ($null -ne $BaselineEntry) {
        try {
            $baselineComparable = ConvertTo-ComparableSnapshotEntry -Entry $BaselineEntry -SnapshotPath $SnapshotPath
            $liveComparable = ConvertTo-ComparableSnapshotEntry -Entry $LiveEntry -ReferenceEntry $BaselineEntry
            $changed = -not (Test-CanonicalStateEqual -Expected $baselineComparable -Actual $liveComparable)
        } catch {
            $changed = $true
        }
    }
    if ($evaluation.PSObject.Properties['Changed'] -and $null -ne $evaluation.Changed) {
        $changed = [bool]$evaluation.Changed
    }

    $pendingReboot = [bool]$evaluation.PendingReboot -or ([bool]$evaluation.Matches -and [bool]$Definition.RequiresReboot -and $changed)
    $status = if (-not $evaluation.Matches) { 'Mismatch' } elseif ($pendingReboot) { 'ConfiguredPendingReboot' } else { 'Verified' }
    $reason = [string]$evaluation.Reason
    if ($evaluation.Matches -and -not $pendingReboot) {
        $reason = $null
    } elseif ($evaluation.Matches -and $pendingReboot -and [string]::IsNullOrWhiteSpace($reason)) {
        $reason = 'The configured target was verified, but activation requires a reboot.'
    }

    [PSCustomObject]@{
        Id             = [string]$Definition.Id
        Type           = [string]$Definition.Type
        RequiresReboot = [bool]$Definition.RequiresReboot
        Status         = $status
        Matches        = [bool]$evaluation.Matches
        Changed        = $changed
        Reason         = $reason
        Expected       = $evaluation.Expected
        Actual         = $evaluation.Actual
    }
}

function Test-DefensePermissiveState {
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Definitions,
        [Parameter(Mandatory)] [hashtable]$BaselineEntriesById,
        [string]$SnapshotPath
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $captureSession = New-CaptureSession -Phase PermissiveVerification
    $captureSession.UserRegistryDefinitionsRemaining = @($Definitions | Where-Object { [string]$_.Type -eq 'LoadedUserRegistryValues' }).Count
    $captureSession.ServiceNames = @($Definitions | Where-Object { [string]$_.Type -eq 'ServiceConfig' } | ForEach-Object { [string]$_.Name })
    $captureMetrics = $null

    try {
        for ($i = 0; $i -lt $Definitions.Count; $i++) {
            $definition = $Definitions[$i]
            $id = [string]$definition.Id
            Write-OperationProgress -Phase 'Verify permissive' -Current ($i + 1) -Total $Definitions.Count -Id $id
            $baselineEntry = if ($BaselineEntriesById.ContainsKey($id)) { $BaselineEntriesById[$id] } else { $null }
            try {
                $liveEntry = Invoke-TimedDefinitionCapture -Definition $definition -CaptureSession $captureSession
                $results.Add((Test-PermissiveDefinitionState -Definition $definition -LiveEntry $liveEntry -BaselineEntry $baselineEntry -SnapshotPath $SnapshotPath)) | Out-Null
            } catch {
                $results.Add([PSCustomObject]@{
                    Id             = $id
                    Type           = [string]$definition.Type
                    RequiresReboot = [bool]$definition.RequiresReboot
                    Status         = 'Mismatch'
                    Matches        = $false
                    Changed        = $false
                    Reason         = $_.Exception.Message
                    Expected       = $null
                    Actual         = $null
                }) | Out-Null
            }
        }
    } finally {
        $captureMetrics = Complete-CaptureSession -Session $captureSession
    }

    [PSCustomObject]@{
        Tool                = 'WinDefState'
        Verifier            = Get-WinDefStateRuntimeInfo
        ComputerName        = $env:COMPUTERNAME
        VerifiedAtUtc       = (Get-Date).ToUniversalTime().ToString('o')
        VerifiedCount       = @($results | Where-Object { [string]$_.Status -eq 'Verified' }).Count
        PendingRebootCount  = @($results | Where-Object { [string]$_.Status -eq 'ConfiguredPendingReboot' }).Count
        MismatchCount       = @($results | Where-Object { [string]$_.Status -eq 'Mismatch' }).Count
        CaptureMetrics      = $captureMetrics
        Results             = @($results)
    }
}

function Get-MutationWorkItems {
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Items)

    $workItems = [System.Collections.Generic.List[object]]::new()
    $batchedItemsByKey = @{}
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        $type = [string]$item.Type
        $id = [string]$item.Id
        $canBatch = if ($item.PSObject.Properties['CanBatch']) { [bool]$item.CanBatch } else { $true }
        $batchKey = $null
        $batchKind = $null
        $operationId = $id
        $operationType = $type

        if ($canBatch -and $type -eq 'MpPreferenceValue') {
            $batchKey = 'defender.preferences'
            $batchKind = 'DefenderPreferenceValues'
            $operationId = 'defender.preferences'
            $operationType = 'MpPreferenceValueBatch'
        } elseif ($canBatch -and $type -eq 'WsManValue') {
            $subject = if ($item.PSObject.Properties['Definition'] -and $null -ne $item.Definition) { $item.Definition } else { $item.Entry }
            $target = Resolve-WsManConfigTarget -Path ([string]$subject.Path)
            $resourceUri = [string]$target.ResourceUri
            $batchKey = "winrm.resource:$($resourceUri.ToLowerInvariant())"
            $batchKind = 'WsManValues'
            $operationId = "winrm.resource:$resourceUri"
            $operationType = 'WsManValueBatch'
        }

        if ($null -eq $batchKey) {
            $singleItems = [System.Collections.Generic.List[object]]::new()
            $singleItems.Add($item) | Out-Null
            $workItems.Add([PSCustomObject]@{
                Id        = $operationId
                Type      = $operationType
                BatchKind = $null
                Items     = $singleItems
            }) | Out-Null
            continue
        }

        if (-not $batchedItemsByKey.ContainsKey($batchKey)) {
            $batchItems = [System.Collections.Generic.List[object]]::new()
            $workItem = [PSCustomObject]@{
                Id        = $operationId
                Type      = $operationType
                BatchKind = $batchKind
                Items     = $batchItems
            }
            $batchedItemsByKey[$batchKey] = $workItem
            $workItems.Add($workItem) | Out-Null
        }
        $batchedItemsByKey[$batchKey].Items.Add($item) | Out-Null
    }

    @($workItems)
}

function Invoke-PermissiveMutationWorkItem {
    param(
        [Parameter(Mandatory)] [object]$WorkItem,
        [string]$SnapshotPath,
        [AllowNull()] [object]$CaptureSession
    )

    switch ([string]$WorkItem.BatchKind) {
        'DefenderPreferenceValues' {
            $values = @(
                foreach ($item in @($WorkItem.Items)) {
                    [PSCustomObject]@{ Property = [string]$item.Definition.Property; Value = $item.Definition.PermissiveValue }
                }
            )
            Set-MpPreferencePropertyValues -Items $values
        }
        'WsManValues' {
            $values = @(
                foreach ($item in @($WorkItem.Items)) {
                    [PSCustomObject]@{ Path = [string]$item.Definition.Path; Value = $item.Definition.PermissiveValue }
                }
            )
            Set-WsManConfigValues -Items $values -CaptureSession $CaptureSession
        }
        default {
            $item = @($WorkItem.Items)[0]
            Apply-PermissiveDefinition -Definition $item.Definition -Entry $item.Entry -SnapshotPath $SnapshotPath -CaptureSession $CaptureSession
        }
    }
}

function Invoke-RestoreMutationWorkItem {
    param(
        [Parameter(Mandatory)] [object]$WorkItem,
        [string]$SnapshotPath,
        [AllowNull()] [object]$CaptureSession
    )

    switch ([string]$WorkItem.BatchKind) {
        'DefenderPreferenceValues' {
            $values = @(
                foreach ($item in @($WorkItem.Items)) {
                    [PSCustomObject]@{ Property = [string]$item.Entry.Property; Value = $item.Entry.RestoreValue }
                }
            )
            Set-MpPreferencePropertyValues -Items $values
        }
        'WsManValues' {
            $values = @(
                foreach ($item in @($WorkItem.Items)) {
                    [PSCustomObject]@{ Path = [string]$item.Entry.Path; Value = $item.Entry.CurrentValue }
                }
            )
            Set-WsManConfigValues -Items $values -CaptureSession $CaptureSession
        }
        default {
            $item = @($WorkItem.Items)[0]
            Restore-SnapshotEntry -Entry $item.Entry -SnapshotPath $SnapshotPath -CaptureSession $CaptureSession
        }
    }
}

function Complete-MutationWorkItem {
    param(
        [AllowNull()] [object]$CaptureSession,
        [Parameter(Mandatory)] [object]$WorkItem
    )

    foreach ($item in @($WorkItem.Items)) {
        Complete-MutationSessionDefinition -Session $CaptureSession -Type ([string]$item.Type)
    }
}

function Apply-PermissiveDefinition {
    param(
        [Parameter(Mandatory)] [object]$Definition,
        [AllowNull()] [object]$Entry,
        [string]$SnapshotPath,
        [AllowNull()] [object]$CaptureSession
    )

    switch ($Definition.Type) {
        'RegistryValue' {
            Ensure-RegistryPath -Path $Definition.Path
            if ($Definition.PermissiveExists) {
                New-ItemProperty -Path $Definition.Path -Name $Definition.Name -PropertyType $Definition.ValueKind -Value $Definition.PermissiveValue -Force | Out-Null
            } else {
                Remove-ItemProperty -Path $Definition.Path -Name $Definition.Name -ErrorAction SilentlyContinue
            }
        }
        'RegistryKeyFlat' {
            Remove-RegistryKeyIfExists -Path $Definition.Path
        }
        'MpPreferenceValue' {
            Set-MpPreferencePropertyValue -Property $Definition.Property -Value $Definition.PermissiveValue
        }
        'MpPreferenceList' {
            if ($Definition.PSObject.Properties['PermissiveValue']) {
                Set-MpPreferenceListValue -Property $Definition.Property -DesiredItems @($Definition.PermissiveValue) -CaptureSession $CaptureSession
            }
        }
        'DefenderRuntimeStatus' {
        }
        'BitLockerVolumes' {
            $bitLockerState = if ($null -ne $Entry) { $Entry.CurrentValue } else { $null }
            Set-Permissive-BitLockerVolumes -State $bitLockerState
        }
        'ExploitProtectionPolicy' {
            Set-Permissive-ExploitProtection
        }
        'WdacPolicies' {
            Remove-WdacPolicies -State (Get-WdacPolicyState)
        }
        'AsrRules' {
            Disable-ConfiguredAsrRules -CaptureSession $CaptureSession
        }
        'PowerShellModuleLogging' {
            Set-Permissive-PowerShellModuleLogging
        }
        'AppLockerPolicy' {
            $appLockerState = if ($null -ne $Entry) { $Entry.CurrentValue } else { $null }
            Set-Permissive-AppLockerPolicy -State $appLockerState -SnapshotPath $SnapshotPath
        }
        'FirewallProfiles' {
            Set-Permissive-FirewallProfiles -Definition $Definition
        }
        'FirewallRules' {
            Set-Permissive-FirewallRules -Definition $Definition
        }
        'NetBiosAdapters' {
            Set-Permissive-NetBiosAdapters -Adapters @($Entry.CurrentValue)
        }
        'LoadedUserRegistryValues' {
            Set-Permissive-LoadedUserRegistryValues -Items @($Definition.Items) -CaptureSession $CaptureSession
        }
        'MachineEnvironmentValue' {
            if ($Definition.PermissiveExists) {
                [Environment]::SetEnvironmentVariable($Definition.Name, $Definition.PermissiveValue, 'Machine')
            } else {
                [Environment]::SetEnvironmentVariable($Definition.Name, $null, 'Machine')
            }
        }
        'ServiceConfig' {
            Set-ServiceStartModeValue -Name $Definition.Name -StartModeValue $Definition.PermissiveStartup
            Set-ServiceRunningState -Name $Definition.Name -Running ([string]$Definition.PermissiveState -eq 'Running')
        }
        'LocalUser' {
            Set-LocalUserEnabledState -Reference $Definition -Enabled ([bool]$Definition.PermissiveValue)
        }
        'WsManValue' {
            Set-WsManConfigValue -Path $Definition.Path -Value $Definition.PermissiveValue -CaptureSession $CaptureSession
        }
        'WinRmListeners' {
            Set-Permissive-WinRmListeners -Listeners @($Definition.PermissiveValue) -CaptureSession $CaptureSession
        }
        'AuditPolicy' {
            Set-AuditPolicyState -Subcategory $Definition.Subcategory -SubcategoryGuid $Definition.SubcategoryGuid -Success $Definition.PermissiveSuccess -Failure $Definition.PermissiveFailure
        }
        'SmbClientConfig' {
            if (Test-CommandAvailable -Name 'Set-SmbClientConfiguration') {
                Set-SmbClientConfiguration -RequireSecuritySignature $Definition.PermissiveValue -Confirm:$false | Out-Null
            }
        }
        'SmbServerConfig' {
            if (Test-CommandAvailable -Name 'Set-SmbServerConfiguration') {
                Set-SmbServerConfiguration -RequireSecuritySignature $Definition.PermissiveValue -Confirm:$false | Out-Null
            }
        }
    }
}

function Restore-SnapshotEntry {
    param(
        [Parameter(Mandatory)] [object]$Entry,
        [string]$SnapshotPath,
        [AllowNull()] [object]$CaptureSession
    )

    if (-not (Test-SnapshotEntryCapturedExactly -Entry $Entry -SnapshotPath $SnapshotPath)) {
        Write-Warning "Skipping restore for $($Entry.Id) because the snapshot baseline was incomplete."
        return
    }

    switch ($Entry.Type) {
        'RegistryValue' {
            Ensure-RegistryPath -Path $Entry.Path
            if ($Entry.Exists) {
                New-ItemProperty -Path $Entry.Path -Name $Entry.Name -PropertyType $Entry.ValueKind -Value $Entry.CurrentValue -Force | Out-Null
            } else {
                Remove-ItemProperty -Path $Entry.Path -Name $Entry.Name -ErrorAction SilentlyContinue
            }
        }
        'RegistryKeyFlat' {
            Restore-RegistryKeyFlatState -Entry $Entry
        }
        'MpPreferenceValue' {
            Set-MpPreferencePropertyValue -Property $Entry.Property -Value $Entry.RestoreValue
        }
        'MpPreferenceList' {
            Set-MpPreferenceListValue -Property $Entry.Property -DesiredItems @($Entry.CurrentValue) -CaptureSession $CaptureSession
        }
        'DefenderRuntimeStatus' {
        }
        'BitLockerVolumes' {
            Restore-BitLockerVolumes -State $Entry.CurrentValue
        }
        'ExploitProtectionPolicy' {
            # Reset lingering system mitigation state first because PolicyFilePath import
            # does not reliably clear permissive-era SystemConfig entries on its own.
            Reset-ExploitProtectionSystemConfig
            Apply-ExploitProtectionPolicyXml -Xml (Get-ExploitProtectionPolicyXml -State $Entry.CurrentValue -SnapshotPath $SnapshotPath)
        }
        'WdacPolicies' {
            Restore-WdacPolicies -State $Entry.CurrentValue -SnapshotPath $SnapshotPath
        }
        'AsrRules' {
            Restore-AsrRules -Rules @($Entry.CurrentValue) -CaptureSession $CaptureSession
        }
        'PowerShellModuleLogging' {
            Restore-PowerShellModuleLogging -Entry $Entry
        }
        'AppLockerPolicy' {
            Restore-AppLockerPolicy -State $Entry.CurrentValue -SnapshotPath $SnapshotPath
        }
        'FirewallProfiles' {
            Restore-FirewallProfiles -State $Entry.CurrentValue
        }
        'FirewallRules' {
            Restore-FirewallRules -State $Entry.CurrentValue
        }
        'NetBiosAdapters' {
            Restore-NetBiosAdapters -Adapters @($Entry.CurrentValue)
        }
        'LoadedUserRegistryValues' {
            $state = Normalize-UserRegistryValueState -State $Entry.CurrentValue
            Restore-LoadedUserRegistryValues -Entries @($state.Entries) -CaptureSession $CaptureSession
        }
        'MachineEnvironmentValue' {
            if ($Entry.Exists) {
                [Environment]::SetEnvironmentVariable($Entry.Name, [string]$Entry.CurrentValue, 'Machine')
            } else {
                [Environment]::SetEnvironmentVariable($Entry.Name, $null, 'Machine')
            }
        }
        'ServiceConfig' {
            $startMode = Convert-ServiceStartModeToScValue -StartMode ([string]$Entry.CurrentValue.StartMode)
            Set-ServiceStartModeValue -Name $Entry.Name -StartModeValue $startMode
            Set-ServiceRunningState -Name $Entry.Name -Running ([string]$Entry.CurrentValue.State -eq 'Running')
        }
        'LocalUser' {
            if ($null -eq $Entry.CurrentValue) {
                return
            }

            Set-LocalUserEnabledState -Reference $Entry -Enabled ([bool]$Entry.CurrentValue)
        }
        'WsManValue' {
            Set-WsManConfigValue -Path $Entry.Path -Value $Entry.CurrentValue -CaptureSession $CaptureSession
        }
        'WinRmListeners' {
            Restore-WinRmListeners -Listeners @($Entry.CurrentValue) -CaptureSession $CaptureSession
        }
        'AuditPolicy' {
            $subcategoryGuid = if ($Entry.PSObject.Properties['SubcategoryGuid']) { [string]$Entry.SubcategoryGuid } else { $null }
            Set-AuditPolicyState -Subcategory $Entry.Subcategory -SubcategoryGuid $subcategoryGuid -Success ([bool]$Entry.CurrentValue.Success) -Failure ([bool]$Entry.CurrentValue.Failure)
        }
        'SmbClientConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            if (Test-CommandAvailable -Name 'Set-SmbClientConfiguration') {
                Set-SmbClientConfiguration -RequireSecuritySignature ([bool]$state.RequireSecuritySignature) -Confirm:$false | Out-Null
            }
        }
        'SmbServerConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            if (Test-CommandAvailable -Name 'Set-SmbServerConfiguration') {
                Set-SmbServerConfiguration -RequireSecuritySignature ([bool]$state.RequireSecuritySignature) -Confirm:$false | Out-Null
            }
        }
    }
}

#endregion

#region Public orchestration

function Export-DefenseSnapshot {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [AllowNull()] [string[]]$IncludeId,
        [AllowNull()] [string[]]$ExcludeId,
        [AllowNull()] [string]$CancellationPath
    )

    $fullPath = [IO.Path]::GetFullPath($Path)
    $definitions = @(Get-SelectedDefenseDefinitions -IncludeId $IncludeId -ExcludeId $ExcludeId)
    $settings = [System.Collections.Generic.List[object]]::new()
    $captureSession = New-CaptureSession -Phase Snapshot
    $captureSession.UserRegistryDefinitionsRemaining = @($definitions | Where-Object { [string]$_.Type -eq 'LoadedUserRegistryValues' }).Count
    $captureSession.ServiceNames = @($definitions | Where-Object { [string]$_.Type -eq 'ServiceConfig' } | ForEach-Object { [string]$_.Name })
    $captureMetrics = $null

    try {
        for ($i = 0; $i -lt $definitions.Count; $i++) {
            Assert-OperationNotCancelled -Path $CancellationPath -Stage 'read-only capture'
            $definition = $definitions[$i]
            Write-OperationProgress -Phase 'Capture' -Current ($i + 1) -Total $definitions.Count -Id ([string]$definition.Id)
            $settings.Add((Invoke-TimedDefinitionCapture -Definition $definition -CaptureSession $captureSession))
        }
    } finally {
        $captureMetrics = Complete-CaptureSession -Session $captureSession
    }

    $snapshot = [PSCustomObject]@{
        SchemaVersion = 2
        Tool          = 'WinDefState'
        Producer      = Get-WinDefStateRuntimeInfo
        ComputerName  = $env:COMPUTERNAME
        CapturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        CaptureMetrics = $captureMetrics
        CaptureScope  = [PSCustomObject]@{
            IsFiltered = Test-IdFilterActive -IncludeId $IncludeId -ExcludeId $ExcludeId
            IncludeId  = @(Get-NormalizedIdFilter -Ids $IncludeId)
            ExcludeId  = @(Get-NormalizedIdFilter -Ids $ExcludeId)
        }
        Settings      = @($settings)
    }

    Write-Verbose ("Capture completed in {0:N1} ms with {1} provider query/queries and {2} cache hit(s)." -f $captureMetrics.DurationMs, $captureMetrics.ProviderQueryCount, $captureMetrics.CacheHitCount)
    Assert-OperationNotCancelled -Path $CancellationPath -Stage 'before snapshot persistence'

    Write-OperationProgress -Phase 'Persist' -Current 1 -Total 4 -Id 'Snapshot sidecar assets'
    Persist-SnapshotExternalAssets -Snapshot $snapshot -SnapshotPath $fullPath

    Write-OperationProgress -Phase 'Persist' -Current 2 -Total 4 -Id 'Snapshot JSON'
    Write-SnapshotJsonAtomic -Path $fullPath -Snapshot $snapshot

    $reportPath = Get-SnapshotReportPath -SnapshotPath $fullPath
    try {
        Write-OperationProgress -Phase 'Persist' -Current 3 -Total 4 -Id 'Snapshot report'
        $reportLines = Get-SnapshotReportLines -Snapshot $snapshot -SnapshotPath $fullPath
        Write-OperationProgress -Phase 'Persist' -Current 4 -Total 4 -Id 'Snapshot report file'
        Write-TextAtomic -Path $reportPath -Content ($reportLines -join [Environment]::NewLine)
    } catch {
        throw "Snapshot JSON was saved successfully to '$fullPath', but its human-readable report could not be created. No defense setting was changed by this capture step. $($_.Exception.Message)"
    }

    Write-OperationResult -Name 'SnapshotPath' -Value $fullPath

    [PSCustomObject]@{
        JsonPath    = $fullPath
        ReportPath  = $reportPath
        Snapshot    = $snapshot
        ReportLines = @($reportLines)
    }
}

function Set-DefensePermissive {
    param(
        [string]$Path,
        [AllowNull()] [string[]]$IncludeId,
        [AllowNull()] [string[]]$ExcludeId,
        [AllowNull()] [string]$CancellationPath,
        [AllowNull()] [string]$MutationApprovalPath
    )

    $definitions = @(Get-SelectedDefenseDefinitions -IncludeId $IncludeId -ExcludeId $ExcludeId)
    $mutationDefinitions = @($definitions | Where-Object { Test-DefinitionHasPermissiveAction -Definition $_ })
    if ($mutationDefinitions.Count -eq 0) {
        throw 'The selected settings are capture-only and have no permissive action.'
    }

    $null = Invoke-WinDefStatePreflight -Action Permissive -Items $definitions -Root $StateRoot -SnapshotPath $Path
    Assert-NoActiveOperation -Root $StateRoot

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Get-DefaultSnapshotPath -Root $StateRoot
    }

    $export = Export-DefenseSnapshot -Path $Path -IncludeId $IncludeId -ExcludeId $ExcludeId -CancellationPath $CancellationPath
    Assert-OperationNotCancelled -Path $CancellationPath -Stage 'before permissive mutation'
    Wait-MutationApproval -Path $MutationApprovalPath -Action Permissive -SnapshotPath $export.JsonPath -TrustedRoot $StateRoot -CancellationPath $CancellationPath
    Assert-OperationNotCancelled -Path $CancellationPath -Stage 'after pre-change review'
    $operationState = Write-OperationState -Root $StateRoot -SnapshotPath $export.JsonPath -Mode 'Permissive' -IncludeId $IncludeId -ExcludeId $ExcludeId

    $snapshotEntriesById = @{}
    foreach ($entry in @($export.Snapshot.Settings)) {
        $snapshotEntriesById[[string]$entry.Id] = $entry
    }

    $completedApplyCount = 0
    $skippedIncompleteCount = 0
    $appliedDefinitions = [System.Collections.Generic.List[object]]::new()
    $mutationSession = New-CaptureSession -Phase Permissive
    $mutationMetrics = $null
    try {
        $mutationItems = [System.Collections.Generic.List[object]]::new()
        foreach ($definition in $mutationDefinitions) {
            $entry = if ($snapshotEntriesById.ContainsKey([string]$definition.Id)) { $snapshotEntriesById[[string]$definition.Id] } else { $null }
            if ($null -eq $entry) {
                throw "The persisted baseline is missing setting '$($definition.Id)'. No mutation was attempted for that setting."
            }
            if (-not (Test-SnapshotEntryCapturedExactly -Entry $entry -SnapshotPath $export.JsonPath)) {
                Write-Warning "Skipping permissive change for $($definition.Id) because the baseline capture was incomplete."
                $skippedIncompleteCount++
                continue
            }

            $mutationItems.Add([PSCustomObject]@{
                Id         = [string]$definition.Id
                Type       = [string]$definition.Type
                Definition = $definition
                Entry      = $entry
                CanBatch   = $true
            }) | Out-Null
        }

        $mutationSession.UserRegistryDefinitionsRemaining = @($mutationItems | Where-Object { [string]$_.Type -eq 'LoadedUserRegistryValues' }).Count
        $mutationSession.WinRmMutationDefinitionsRemaining = @($mutationItems | Where-Object { [string]$_.Type -in @('WsManValue', 'WinRmListeners') }).Count
        $workItems = @(Get-MutationWorkItems -Items @($mutationItems))
        $processedSettingCount = 0
        foreach ($workItem in $workItems) {
            $itemCount = @($workItem.Items).Count
            $progressId = if ($itemCount -gt 1) { "$($workItem.Id) ($itemCount settings)" } else { [string]$workItem.Id }
            Write-OperationProgress -Phase 'Permissive' -Current ($processedSettingCount + $itemCount) -Total $mutationItems.Count -Id $progressId
            try {
                Invoke-TimedSettingOperation -Phase Permissive -Id ([string]$workItem.Id) -Type ([string]$workItem.Type) -Session $mutationSession -Action {
                    Invoke-PermissiveMutationWorkItem -WorkItem $workItem -SnapshotPath $export.JsonPath -CaptureSession $mutationSession
                }
            } finally {
                Complete-MutationWorkItem -CaptureSession $mutationSession -WorkItem $workItem
            }
            foreach ($item in @($workItem.Items)) {
                $appliedDefinitions.Add($item.Definition) | Out-Null
            }
            $processedSettingCount += $itemCount
            $completedApplyCount += $itemCount
        }
    } catch {
        $applyError = $_
        try {
            $operationState = Update-OperationStateStatus -Root $StateRoot -Operation $operationState -Status 'ApplyFailed'
        } catch {
            Write-Warning "The permissive operation failed and its journal status could not be updated: $($_.Exception.Message)"
        }
        throw "Permissive apply failed after $completedApplyCount setting(s) completed. The original snapshot and current-operation.json were preserved for restore. $($applyError.Exception.Message)"
    } finally {
        $mutationMetrics = Complete-CaptureSession -Session $mutationSession
    }
    Write-Verbose ("Permissive mutation phase completed in {0:N1} ms with {1} shared provider setup query/queries and {2} cache hit(s)." -f $mutationMetrics.DurationMs, $mutationMetrics.ProviderQueryCount, $mutationMetrics.CacheHitCount)

    $verification = $null
    $verificationPath = $null
    try {
        $verification = Test-DefensePermissiveState -Definitions @($appliedDefinitions) -BaselineEntriesById $snapshotEntriesById -SnapshotPath $export.JsonPath
        $baselineProducer = if ($export.Snapshot.PSObject.Properties['Producer']) { $export.Snapshot.Producer } else { $null }
        $verification | Add-Member -NotePropertyName BaselineProducer -NotePropertyValue $baselineProducer -Force
        $verification | Add-Member -NotePropertyName MutationMetrics -NotePropertyValue $mutationMetrics -Force
        $verificationPath = Get-PermissiveVerificationReportPath -Root $StateRoot -SnapshotPath $export.JsonPath
        $verificationLines = Get-PermissiveVerificationReportLines -Verification $verification -SnapshotPath $export.JsonPath
        Write-TextAtomic -Path $verificationPath -Content ($verificationLines -join [Environment]::NewLine)
    } catch {
        $verificationError = $_
        try {
            $operationState = Update-OperationStateStatus -Root $StateRoot -Operation $operationState -Status 'ApplyVerificationFailed'
        } catch {
            Write-Warning "Permissive verification failed and its journal status could not be updated: $($_.Exception.Message)"
        }
        throw "Permissive commands completed, but post-apply verification could not finish. The original snapshot and current-operation.json were preserved for restore. $($verificationError.Exception.Message)"
    }

    if ($verification.MismatchCount -gt 0) {
        $operationState = Update-OperationStateStatus -Root $StateRoot -Operation $operationState -Status 'ApplyVerificationFailed' -PermissiveVerification $verification -PermissiveVerificationReportPath $verificationPath
        $mismatchIds = @($verification.Results | Where-Object { [string]$_.Status -eq 'Mismatch' } | ForEach-Object { [string]$_.Id })
        Write-Warning "Permissive verification found $($verification.MismatchCount) setting(s) that did not reach the requested target."
        Write-Warning ("Mismatched IDs: {0}" -f ($mismatchIds -join ', '))
        Write-Warning "Permissive verification report saved to: $verificationPath"
        throw 'Permissive verification failed. The baseline and current-operation.json were left in place so Restore can return the host to its original state.'
    }

    $operationStatus = if ($verification.PendingRebootCount -gt 0) { 'AppliedPendingReboot' } else { 'AppliedVerified' }
    $operationState = Update-OperationStateStatus -Root $StateRoot -Operation $operationState -Status $operationStatus -PermissiveVerification $verification -PermissiveVerificationReportPath $verificationPath

    Write-Host ("Permissive mode completed: {0} verified immediately, {1} configured and pending reboot, {2} skipped because capture was incomplete, {3} capture-only." -f $verification.VerifiedCount, $verification.PendingRebootCount, $skippedIncompleteCount, ($definitions.Count - $mutationDefinitions.Count))
    Write-Host "Snapshot JSON saved to: $($export.JsonPath)"
    Write-Host "Snapshot report saved to: $($export.ReportPath)"
    Write-Host "Permissive verification report saved to: $verificationPath"
    if ($verification.PendingRebootCount -gt 0) {
        $pendingIds = @($verification.Results | Where-Object { [string]$_.Status -eq 'ConfiguredPendingReboot' } | ForEach-Object { [string]$_.Id })
        Write-Warning ("A reboot is required before these configured changes are fully active: {0}" -f ($pendingIds -join ', '))
    }
}

function Get-OrderedRestoreEntries {
    param([Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$Entries)

    $deferredWinRmServiceEntries = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in @($Entries)) {
        $isWinRmService = (
            [string]$entry.Type -eq 'ServiceConfig' -and
            $entry.PSObject.Properties['Name'] -and
            [string]::Equals([string]$entry.Name, 'WinRM', [System.StringComparison]::OrdinalIgnoreCase)
        )
        if ($isWinRmService) {
            $deferredWinRmServiceEntries.Add($entry) | Out-Null
        } else {
            $entry
        }
    }

    foreach ($entry in $deferredWinRmServiceEntries) {
        $entry
    }
}

function Get-RestoreCheckpointResumeState {
    param(
        [Parameter(Mandatory)] [object]$Operation,
        [Parameter(Mandatory)] [object]$Snapshot,
        [Parameter(Mandatory)] [string]$SnapshotPath,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [object[]]$MutationEntries
    )

    $emptyResult = [PSCustomObject]@{
        CandidateIds  = @()
        VerifiedIds   = @()
        RetryIds      = @()
        Verification  = $null
    }
    if (
        -not $Operation.PSObject.Properties['RestoreCheckpoint'] -or
        $null -eq $Operation.RestoreCheckpoint -or
        -not $Operation.RestoreCheckpoint.PSObject.Properties['CompletedIds']
    ) {
        return $emptyResult
    }

    $requestedIds = @(Get-NormalizedRestoreCheckpointIds -Id @($MutationEntries | ForEach-Object { [string]$_.Id }))
    $requestedSet = @{}
    foreach ($requestedId in $requestedIds) {
        $requestedSet[([string]$requestedId).ToLowerInvariant()] = $true
    }
    $candidateIds = @(
        Get-NormalizedRestoreCheckpointIds -Id @($Operation.RestoreCheckpoint.CompletedIds) |
            Where-Object { $requestedSet.ContainsKey(([string]$_).ToLowerInvariant()) }
    )
    if ($candidateIds.Count -eq 0) {
        return $emptyResult
    }

    $verification = Test-DefenseSnapshot -Snapshot $Snapshot -SnapshotPath $SnapshotPath -IncludeId $candidateIds
    $verifiedIds = @(
        $verification.Results |
            Where-Object { $_.Matches -and -not $_.Skipped } |
            ForEach-Object { [string]$_.Id }
    )
    $verifiedSet = @{}
    foreach ($verifiedId in $verifiedIds) {
        $verifiedSet[([string]$verifiedId).ToLowerInvariant()] = $true
    }
    $retryIds = @($candidateIds | Where-Object { -not $verifiedSet.ContainsKey(([string]$_).ToLowerInvariant()) })
    [PSCustomObject]@{
        CandidateIds = @($candidateIds)
        VerifiedIds  = @($verifiedIds)
        RetryIds     = @($retryIds)
        Verification = $verification
    }
}

function Restore-DefenseSnapshot {
    param(
        [string]$Path,
        [AllowNull()] [string[]]$IncludeId,
        [AllowNull()] [string[]]$ExcludeId,
        [switch]$AllowDifferentComputer,
        [AllowNull()] [string]$CancellationPath,
        [AllowNull()] [string]$MutationApprovalPath
    )

    $operation = Get-OperationState -Root $StateRoot
    if ([string]::IsNullOrWhiteSpace($Path)) {
        if ($null -eq $operation) {
            $null = Invoke-WinDefStatePreflight -Action Restore -Items @() -Root $StateRoot
            throw "No active permissive operation was found. A successful restore clears current-operation.json. To restore a saved snapshot explicitly, rerun Restore with -SnapshotPath 'C:\path\to\snapshot.json'."
        }

        $Path = [string]$operation.SnapshotPath
    }

    $fullPath = [IO.Path]::GetFullPath($Path)
    $operationTargetsSnapshot = Test-OperationTargetsSnapshot -Operation $operation -SnapshotPath $fullPath
    if ($operationTargetsSnapshot) {
        Assert-OperationSnapshotIntegrity -Operation $operation
    }
    $snapshot = Read-JsonFile -Path $fullPath
    Assert-ValidDefenseSnapshot -Snapshot $snapshot -AllowDifferentComputer:$AllowDifferentComputer
    $allEntries = @($snapshot.Settings)
    Assert-ValidSettingIdFilter -AvailableId @($allEntries.Id) -IncludeId $IncludeId -ExcludeId $ExcludeId
    Assert-OperationNotCancelled -Path $CancellationPath -Stage 'before restore mutation'
    $entries = @(
        Get-OrderedRestoreEntries -Entries @(
            $allEntries | Where-Object { Test-SettingIdIncluded -Id ([string]$_.Id) -IncludeId $IncludeId -ExcludeId $ExcludeId }
        )
    )
    $isFilteredRestore = Test-IdFilterActive -IncludeId $IncludeId -ExcludeId $ExcludeId
    $includedIds = @(Get-NormalizedIdFilter -Ids $IncludeId)
    $excludedIds = @(Get-NormalizedIdFilter -Ids $ExcludeId)
    if ($includedIds.Count -eq $allEntries.Count -and $excludedIds.Count -eq 0 -and $entries.Count -eq $allEntries.Count) {
        $isFilteredRestore = $false
    }
    $null = Invoke-WinDefStatePreflight -Action Restore -Items $entries -Root $StateRoot -SnapshotPath $fullPath
    Initialize-SnapshotAssetCache -Entries $entries -SnapshotPath $fullPath
    Write-OperationResult -Name 'SnapshotPath' -Value $fullPath
    Wait-MutationApproval -Path $MutationApprovalPath -Action Restore -SnapshotPath $fullPath -TrustedRoot $StateRoot -CancellationPath $CancellationPath
    Assert-OperationNotCancelled -Path $CancellationPath -Stage 'after pre-change review'
    if ($operationTargetsSnapshot) {
        $operation = Update-OperationStateStatus -Root $StateRoot -Operation $operation -Status 'Restoring'
    }
    $mutationEntries = @($entries | Where-Object { Test-DefinitionHasRestoreAction -Definition $_ })
    $requestedMutationCount = $mutationEntries.Count
    $restoreAttemptNumber = 0
    $resumeState = [PSCustomObject]@{ CandidateIds = @(); VerifiedIds = @(); RetryIds = @(); Verification = $null }
    $resumedSettingCount = 0
    if ($operationTargetsSnapshot) {
        $requestedMutationIds = @($mutationEntries | ForEach-Object { [string]$_.Id })
        $operation = Initialize-RestoreCheckpoint -Root $StateRoot -Operation $operation -RequestedIds $requestedMutationIds
        $restoreAttemptNumber = [int]$operation.RestoreCheckpoint.AttemptNumber
        try {
            $resumeState = Get-RestoreCheckpointResumeState -Operation $operation -Snapshot $snapshot -SnapshotPath $fullPath -MutationEntries $mutationEntries
        } catch {
            Write-Warning "Previously completed restore settings could not be revalidated and will be retried: $($_.Exception.Message)"
            $resumeState = [PSCustomObject]@{ CandidateIds = @(); VerifiedIds = @(); RetryIds = @(); Verification = $null }
        }
        $operation = Set-RestoreCheckpointVerifiedIds -Root $StateRoot -Operation $operation -VerifiedIds @($resumeState.VerifiedIds)
        $verifiedCheckpointSet = @{}
        foreach ($verifiedId in @($resumeState.VerifiedIds)) {
            $verifiedCheckpointSet[([string]$verifiedId).ToLowerInvariant()] = $true
        }
        $mutationEntries = @(
            $mutationEntries | Where-Object {
                -not $verifiedCheckpointSet.ContainsKey(([string]$_.Id).ToLowerInvariant())
            }
        )
        $resumedSettingCount = @($resumeState.VerifiedIds).Count
        if ($resumedSettingCount -gt 0) {
            Write-Host ("Restore resume: {0} previously completed setting(s) still match the baseline and will not be reapplied." -f $resumedSettingCount)
        }
        if (@($resumeState.RetryIds).Count -gt 0) {
            Write-Verbose ("Restore resume will reapply checkpointed setting(s) that no longer match: {0}" -f (@($resumeState.RetryIds) -join ', '))
        }
    }
    $restoreSession = New-CaptureSession -Phase Restore
    $restoreSession.UserRegistryDefinitionsRemaining = @($mutationEntries | Where-Object { [string]$_.Type -eq 'LoadedUserRegistryValues' }).Count
    $restoreSession.WinRmMutationDefinitionsRemaining = @($mutationEntries | Where-Object { [string]$_.Type -in @('WsManValue', 'WinRmListeners') }).Count
    $restoreMetrics = $null
    try {
        $restoreItems = @(
            foreach ($entry in $mutationEntries) {
                $type = [string]$entry.Type
                $canBatch = if ($type -in @('MpPreferenceValue', 'WsManValue')) {
                    Test-SnapshotEntryCapturedExactly -Entry $entry -SnapshotPath $fullPath
                } else {
                    $false
                }
                [PSCustomObject]@{
                    Id       = [string]$entry.Id
                    Type     = $type
                    Entry    = $entry
                    CanBatch = $canBatch
                }
            }
        )
        $workItems = @(Get-MutationWorkItems -Items $restoreItems)
        $processedSettingCount = 0
        foreach ($workItem in $workItems) {
            $itemCount = @($workItem.Items).Count
            $workItemSettingIds = @($workItem.Items | ForEach-Object { [string]$_.Id })
            $progressId = if ($itemCount -gt 1) { "$($workItem.Id) ($itemCount settings)" } else { [string]$workItem.Id }
            Write-OperationProgress -Phase 'Restore' -Current ($processedSettingCount + $itemCount) -Total $mutationEntries.Count -Id $progressId
            if ($operationTargetsSnapshot) {
                $operation = Start-RestoreCheckpointWorkItem -Root $StateRoot -Operation $operation -WorkItemId ([string]$workItem.Id) -SettingIds $workItemSettingIds
            }
            try {
                Invoke-TimedSettingOperation -Phase Restore -Id ([string]$workItem.Id) -Type ([string]$workItem.Type) -Session $restoreSession -Action {
                    Invoke-RestoreMutationWorkItem -WorkItem $workItem -SnapshotPath $fullPath -CaptureSession $restoreSession
                }
            } finally {
                Complete-MutationWorkItem -CaptureSession $restoreSession -WorkItem $workItem
            }
            if ($operationTargetsSnapshot) {
                $operation = Complete-RestoreCheckpointWorkItem -Root $StateRoot -Operation $operation -SettingIds $workItemSettingIds
            }
            $processedSettingCount += $itemCount
        }
    } catch {
        $restoreError = $_
        if ($operationTargetsSnapshot) {
            try {
                $operation = Set-RestoreCheckpointFailure -Root $StateRoot -Operation $operation -Message $restoreError.Exception.Message
            } catch {
                Write-Warning "Restore failed and its per-setting checkpoint could not be updated: $($_.Exception.Message)"
            }
            try {
                $operation = Update-OperationStateStatus -Root $StateRoot -Operation $operation -Status 'RestoreFailed'
            } catch {
                Write-Warning "Restore failed and the operation journal status could not be updated: $($_.Exception.Message)"
            }
        }
        throw "Restore failed while applying snapshot '$fullPath'. current-operation.json was preserved when this snapshot owns the active operation. $($restoreError.Exception.Message)"
    } finally {
        $restoreMetrics = Complete-CaptureSession -Session $restoreSession
    }
    Write-Verbose ("Restore mutation phase completed in {0:N1} ms with {1} shared provider setup query/queries and {2} cache hit(s)." -f $restoreMetrics.DurationMs, $restoreMetrics.ProviderQueryCount, $restoreMetrics.CacheHitCount)

    $verification = $null
    $verificationPath = $null
    $wdacVerificationPath = $null
    try {
        $verification = Test-DefenseSnapshot -Snapshot $snapshot -SnapshotPath $fullPath -IncludeId $IncludeId -ExcludeId $ExcludeId
        $verification | Add-Member -NotePropertyName MutationMetrics -NotePropertyValue $restoreMetrics -Force
        if ($operationTargetsSnapshot) {
            $verification | Add-Member -NotePropertyName RestoreCheckpointSummary -NotePropertyValue ([PSCustomObject]@{
                AttemptNumber               = $restoreAttemptNumber
                RequestedMutationCount      = $requestedMutationCount
                PreviouslyCompletedCount    = @($resumeState.CandidateIds).Count
                RevalidatedAndSkippedCount  = $resumedSettingCount
                ReappliedCheckpointCount    = @($resumeState.RetryIds).Count
                ScheduledMutationCount      = $mutationEntries.Count
            }) -Force
        }
        $verificationPath = Get-VerificationReportPath -Root $StateRoot -SnapshotPath $fullPath
        $verificationLines = Get-VerificationReportLines -Verification $verification -SnapshotPath $fullPath
        $wdacVerificationPath = Get-WdacVerificationReportPath -VerificationPath $verificationPath
        $wdacVerificationLines = Get-WdacVerificationReportLines -Verification $verification -SnapshotPath $fullPath
        Write-TextAtomic -Path $verificationPath -Content ($verificationLines -join [Environment]::NewLine)
        Write-TextAtomic -Path $wdacVerificationPath -Content ($wdacVerificationLines -join [Environment]::NewLine)
    } catch {
        $verificationError = $_
        if ($operationTargetsSnapshot) {
            try {
                $operation = Update-OperationStateStatus -Root $StateRoot -Operation $operation -Status 'RestoreVerificationFailed'
            } catch {
                Write-Warning "Restore verification failed and the operation journal status could not be updated: $($_.Exception.Message)"
            }
        }
        throw "Restore commands completed, but post-restore verification or report persistence could not finish. current-operation.json was preserved when this snapshot owns the active operation. $($verificationError.Exception.Message)"
    }

    if ($verification.MismatchCount -gt 0) {
        $mismatchIds = @($verification.Results | Where-Object { -not $_.Matches -and -not $_.Skipped } | ForEach-Object { $_.Id })
        Write-Warning "Restore verification found $($verification.MismatchCount) mismatched setting(s)."
        Write-Warning "Verification report saved to: $verificationPath"
        Write-Warning "WDAC verification report saved to: $wdacVerificationPath"
        Write-Warning ("Mismatched IDs: {0}" -f ($mismatchIds -join ', '))
        if ($operationTargetsSnapshot) {
            $operation = Update-OperationStateStatus -Root $StateRoot -Operation $operation -Status 'RestoreVerificationFailed'
            throw 'Restore verification failed. current-operation.json was left in place so you can retry the same snapshot.'
        }
        throw 'Restore verification failed. Review the saved report and retry the same snapshot explicitly.'
    }

    if ($isFilteredRestore) {
        if ($operationTargetsSnapshot) {
            $operation = Update-OperationStateStatus -Root $StateRoot -Operation $operation -Status 'PartiallyRestored'
            Write-Warning 'Filtered restore completed. current-operation.json was left in place because only selected settings were restored.'
        }
    } elseif ($operationTargetsSnapshot) {
        Clear-OperationState -Root $StateRoot
    }
    Write-Host (Get-RestoreCompletionMessage -SnapshotPath $fullPath -Verification $verification)
    Write-Host "Verification report saved to: $verificationPath"
    Write-Host "WDAC verification report saved to: $wdacVerificationPath"
    $incompleteResults = @(
        $verification.Results | Where-Object {
            $_.Skipped -and (
                -not $_.PSObject.Properties['SkipCategory'] -or
                [string]$_.SkipCategory -eq 'IncompleteBaseline'
            )
        }
    )
    if ($incompleteResults.Count -gt 0) {
        $skippedIds = @($incompleteResults | ForEach-Object { $_.Id })
        Write-Warning ("Verification skipped {0} setting(s) because the snapshot baseline was incomplete: {1}" -f $incompleteResults.Count, ($skippedIds -join ', '))
    }
}

#endregion

#region Script entry point

if ($script:IsDotSourced) {
    return
}

if ([string]::IsNullOrWhiteSpace($Command)) {
    throw 'Specify -Command Snapshot, Permissive, or Restore.'
}

$IncludeId = @(Merge-SettingIdFilter -Id $IncludeId -Category $IncludeCategory)
$ExcludeId = @(Merge-SettingIdFilter -Id $ExcludeId -Category $ExcludeCategory)

switch ($Command) {
    'Snapshot' {
        $snapshotTarget = if ([string]::IsNullOrWhiteSpace($SnapshotPath)) {
            Join-Path (Join-Path $StateRoot 'snapshots') '<automatic timestamp>.json'
        } else {
            $SnapshotPath
        }
        if ($PSCmdlet.ShouldProcess($snapshotTarget, 'Capture and persist a defense-state snapshot')) {
            Invoke-ProtectedWinDefStateOperation {
                if ([string]::IsNullOrWhiteSpace($SnapshotPath)) {
                    $SnapshotPath = Get-DefaultSnapshotPath -Root $StateRoot
                }
                $preflightDefinitions = @(Get-SelectedDefenseDefinitions -IncludeId $IncludeId -ExcludeId $ExcludeId)
                $null = Invoke-WinDefStatePreflight -Action Snapshot -Items $preflightDefinitions -Root $StateRoot -SnapshotPath $SnapshotPath
                $export = Export-DefenseSnapshot -Path $SnapshotPath -IncludeId $IncludeId -ExcludeId $ExcludeId -CancellationPath $CancellationPath
                Write-Host "Snapshot JSON saved to: $($export.JsonPath)"
                Write-Host "Snapshot report saved to: $($export.ReportPath)"
                if ($ConsoleReport -ne 'None') {
                    Write-Host ''
                    if ($ConsoleReport -eq 'Full') {
                        Show-ReportLines -Lines $export.ReportLines
                    } else {
                        Show-ReportLines -Lines (Get-ReportSummaryLines -Lines $export.ReportLines)
                    }
                }
            }
        }
    }
    'Permissive' {
        $computerTarget = if (-not [string]::IsNullOrWhiteSpace([string]$env:COMPUTERNAME)) { [string]$env:COMPUTERNAME } else { [Environment]::MachineName }
        if ($PSCmdlet.ShouldProcess($computerTarget, 'Capture a baseline and apply the supported permissive defense posture')) {
            Invoke-ProtectedWinDefStateOperation {
                Set-DefensePermissive -Path $SnapshotPath -IncludeId $IncludeId -ExcludeId $ExcludeId -CancellationPath $CancellationPath -MutationApprovalPath $MutationApprovalPath
            }
        }
    }
    'Restore' {
        $restoreTarget = if ([string]::IsNullOrWhiteSpace($SnapshotPath)) {
            Join-Path $StateRoot 'current-operation.json'
        } else {
            $SnapshotPath
        }
        if ($PSCmdlet.ShouldProcess($restoreTarget, 'Restore and verify defense state from the saved snapshot')) {
            Invoke-ProtectedWinDefStateOperation {
                Restore-DefenseSnapshot -Path $SnapshotPath -IncludeId $IncludeId -ExcludeId $ExcludeId -AllowDifferentComputer:$AllowDifferentComputer -CancellationPath $CancellationPath -MutationApprovalPath $MutationApprovalPath
            }
        }
    }
}

#endregion

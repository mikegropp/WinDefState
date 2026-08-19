[CmdletBinding()]
param(
    [string]$StateRoot,
    [switch]$ValidateOnly
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
    $StateRoot = if (-not [string]::IsNullOrWhiteSpace([string]$env:ProgramData)) {
        Join-Path $env:ProgramData 'WinDefState'
    } else {
        Join-Path $scriptRoot 'state'
    }
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

$isAdministrator = Test-Administrator
$requiresRelaunch = (-not $ValidateOnly -and -not $isAdministrator) -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA'
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
    if ($ValidateOnly) {
        $arguments += '-ValidateOnly'
    }

    if ($ValidateOnly) {
        $validationProcess = Start-Process -FilePath 'powershell.exe' -ArgumentList ($arguments -join ' ') -Wait -PassThru
        if ($validationProcess.ExitCode -ne 0) {
            throw "WPF validation failed in the STA child process with exit code $($validationProcess.ExitCode)."
        }
    } elseif ($isAdministrator) {
        Start-Process -FilePath 'powershell.exe' -ArgumentList ($arguments -join ' ')
    } else {
        Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList ($arguments -join ' ')
    }
    return
}

Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase
if ($null -eq ('WinDefState.Gui.ProcessOutputCollector' -as [type])) {
    Add-Type -TypeDefinition @'
using System.Collections.Concurrent;
using System.Diagnostics;

namespace WinDefState.Gui
{
    public sealed class ProcessOutputCollector
    {
        private readonly ConcurrentQueue<string> lines = new ConcurrentQueue<string>();

        public void Attach(Process process)
        {
            process.OutputDataReceived += OnOutputDataReceived;
            process.ErrorDataReceived += OnErrorDataReceived;
        }

        public bool TryDequeue(out string line)
        {
            return lines.TryDequeue(out line);
        }

        private void OnOutputDataReceived(object sender, DataReceivedEventArgs args)
        {
            if (args.Data != null)
            {
                lines.Enqueue(args.Data);
            }
        }

        private void OnErrorDataReceived(object sender, DataReceivedEventArgs args)
        {
            if (args.Data != null)
            {
                lines.Enqueue("ERROR: " + args.Data);
            }
        }
    }
}
'@
}

$script:EnginePath = $enginePath
$script:StateRoot = [IO.Path]::GetFullPath($StateRoot)
$script:SnapshotPath = $null
$script:Snapshot = $null
$script:ComparisonSnapshotPath = $null
$script:ComparisonSnapshot = $null
$script:SnapshotCache = @{}
$script:Rows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:ActiveOperation = $null
$script:OperationTimer = $null
$script:CurrentPhase = $null
$script:SettingsView = $null
$script:RowFilter = $null

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

    if ($Entry.PSObject.Properties['Captured'] -and -not [bool]$Entry.Captured) {
        $captureError = if ($Entry.PSObject.Properties['CaptureError']) { [string]$Entry.CaptureError } else { $null }
        if (-not [string]::IsNullOrWhiteSpace($captureError)) {
            return (ConvertTo-ShortText -Value ("Incomplete: {0}" -f $captureError))
        }
        return '<incomplete>'
    }

    $currentValue = if ($Entry.PSObject.Properties['CurrentValue']) { $Entry.CurrentValue } else { $null }
    switch ([string]$Entry.Type) {
        'RegistryValue' {
            if ($Entry.PSObject.Properties['Exists'] -and [bool]$Entry.Exists) { return (ConvertTo-ShortText -Value $currentValue) }
            return '<absent>'
        }
        'MpPreferenceValue' {
            return (ConvertTo-ShortText -Value $currentValue)
        }
        'MpPreferenceList' {
            return ('{0} item(s)' -f (Get-ItemCount -Value $currentValue))
        }
        'ServiceConfig' {
            if ($null -eq $currentValue) { return '<not captured>' }
            $startMode = if ($currentValue.PSObject.Properties['StartMode']) { [string]$currentValue.StartMode } else { '<unknown>' }
            $state = if ($currentValue.PSObject.Properties['State']) { [string]$currentValue.State } else { '<unknown>' }
            return ('{0} / {1}' -f $startMode, $state)
        }
        'FirewallProfiles' {
            $profiles = if ($null -ne $currentValue -and $currentValue.PSObject.Properties['Profiles']) { $currentValue.Profiles } else { @() }
            return ('{0} profile(s)' -f (Get-ItemCount -Value $profiles))
        }
        'FirewallRules' {
            $rules = if ($null -ne $currentValue -and $currentValue.PSObject.Properties['Rules']) { $currentValue.Rules } else { @() }
            return ('{0} rule(s)' -f (Get-ItemCount -Value $rules))
        }
        'LoadedUserRegistryValues' {
            $entries = if ($null -ne $currentValue -and $currentValue.PSObject.Properties['Entries']) { $currentValue.Entries } else { @() }
            $issues = if ($null -ne $currentValue -and $currentValue.PSObject.Properties['CaptureIssues']) { $currentValue.CaptureIssues } else { @() }
            return ('{0} value(s), {1} issue(s)' -f (Get-ItemCount -Value $entries), (Get-ItemCount -Value $issues))
        }
        'WdacPolicies' {
            $policies = if ($null -ne $currentValue -and $currentValue.PSObject.Properties['Policies']) { $currentValue.Policies } else { @() }
            $files = if ($null -ne $currentValue -and $currentValue.PSObject.Properties['Files']) { $currentValue.Files } else { @() }
            $issues = if ($null -ne $currentValue -and $currentValue.PSObject.Properties['CaptureIssues']) { $currentValue.CaptureIssues } else { @() }
            return ('{0} policy item(s), {1} file(s), {2} issue(s)' -f (Get-ItemCount -Value $policies), (Get-ItemCount -Value $files), (Get-ItemCount -Value $issues))
        }
        'BitLockerVolumes' {
            $volumes = if ($null -ne $currentValue -and $currentValue.PSObject.Properties['Volumes']) { $currentValue.Volumes } else { @() }
            return ('{0} volume(s)' -f (Get-ItemCount -Value $volumes))
        }
        default {
            return (ConvertTo-ShortText -Value $currentValue)
        }
    }
}

function Get-EntryCapabilities {
    param([Parameter(Mandatory)] [object]$Entry)

    if ($Entry.PSObject.Properties['Capabilities'] -and $null -ne $Entry.Capabilities) {
        $capabilities = $Entry.Capabilities
        $permissive = $capabilities.PSObject.Properties['Permissive'] -and [bool]$capabilities.Permissive
        $restore = $capabilities.PSObject.Properties['Restore'] -and [bool]$capabilities.Restore
        $inventoryOnly = if ($capabilities.PSObject.Properties['InventoryOnly']) {
            [bool]$capabilities.InventoryOnly
        } else {
            -not $permissive -and -not $restore
        }
        return [PSCustomObject]@{
            Permissive    = $permissive
            Restore       = $restore
            InventoryOnly = $inventoryOnly
        }
    }

    $type = [string]$Entry.Type
    $legacyPermissive = $type -notin @('DefenderRuntimeStatus', 'MpPreferenceList')
    $legacyRestore = $type -ne 'DefenderRuntimeStatus'
    [PSCustomObject]@{
        Permissive    = $legacyPermissive
        Restore       = $legacyRestore
        InventoryOnly = -not $legacyPermissive -and -not $legacyRestore
    }
}

function Get-EntryBadges {
    param([Parameter(Mandatory)] [object]$Entry)

    $badges = New-Object System.Collections.Generic.List[string]
    $type = [string]$Entry.Type
    $capabilities = Get-EntryCapabilities -Entry $Entry
    if ($Entry.PSObject.Properties['RequiresReboot'] -and [bool]$Entry.RequiresReboot) {
        $badges.Add('Requires reboot')
    }

    if ($capabilities.InventoryOnly) {
        $badges.Add('Inventory only')
    } elseif (-not $capabilities.Permissive -and $capabilities.Restore) {
        $badges.Add('Restore only')
    }

    $structuredTypes = @(
        'ServiceConfig', 'AuditPolicy', 'DefenderRuntimeStatus', 'BitLockerVolumes',
        'AppLockerPolicy', 'ExploitProtectionPolicy', 'WdacPolicies', 'FirewallProfiles',
        'FirewallRules', 'SmbClientConfig', 'SmbServerConfig', 'LoadedUserRegistryValues'
    )
    if ($structuredTypes -contains $type -and (
        -not $Entry.PSObject.Properties['CurrentValue'] -or $null -eq $Entry.CurrentValue
    )) {
        $badges.Add('Partial')
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
        if ($state.PSObject.Properties['Captured'] -and -not [bool]$state.Captured) {
            $badges.Add('Partial')
        }
        if ($state.PSObject.Properties['CommandAvailable'] -and -not [bool]$state.CommandAvailable) {
            $badges.Add('Provider missing')
        }
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
                if ($null -eq $policy) {
                    continue
                }
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

function ConvertTo-GuiCanonicalValue {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return $null
    }
    if (
        $Value -is [string] -or $Value -is [char] -or $Value -is [bool] -or
        $Value -is [byte] -or $Value -is [sbyte] -or $Value -is [int16] -or
        $Value -is [uint16] -or $Value -is [int32] -or $Value -is [uint32] -or
        $Value -is [int64] -or $Value -is [uint64] -or $Value -is [single] -or
        $Value -is [double] -or $Value -is [decimal]
    ) {
        return $Value
    }
    if ($Value -is [datetime] -or $Value -is [datetimeoffset] -or $Value -is [guid]) {
        return [string]$Value
    }
    if ($Value -is [System.Collections.IDictionary]) {
        $dictionaryResult = [ordered]@{}
        foreach ($key in @($Value.Keys | ForEach-Object { [string]$_ } | Sort-Object)) {
            $dictionaryResult[$key] = ConvertTo-GuiCanonicalValue -Value $Value[$key]
        }
        return $dictionaryResult
    }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        return @(
            foreach ($item in $Value) {
                ConvertTo-GuiCanonicalValue -Value $item
            }
        )
    }

    $objectResult = [ordered]@{}
    foreach ($property in @($Value.PSObject.Properties | Where-Object { $_.MemberType -in @('NoteProperty', 'Property') } | Sort-Object -Property Name)) {
        $objectResult[[string]$property.Name] = ConvertTo-GuiCanonicalValue -Value $property.Value
    }
    $objectResult
}

function Get-GuiEntryFingerprint {
    param([Parameter(Mandatory)] [object]$Entry)

    $state = [ordered]@{}
    foreach ($property in @($Entry.PSObject.Properties | Sort-Object -Property Name)) {
        if ([string]$property.Name -in @('Capabilities', 'PermissiveTarget')) {
            continue
        }
        $state[[string]$property.Name] = ConvertTo-GuiCanonicalValue -Value $property.Value
    }
    ConvertTo-Json -InputObject $state -Depth 12 -Compress
}

function Get-EntryPermissiveTargetSummary {
    param([Parameter(Mandatory)] [object]$Entry)

    if (
        $Entry.PSObject.Properties['PermissiveTarget'] -and
        $null -ne $Entry.PermissiveTarget -and
        $Entry.PermissiveTarget.PSObject.Properties['Summary'] -and
        -not [string]::IsNullOrWhiteSpace([string]$Entry.PermissiveTarget.Summary)
    ) {
        return [string]$Entry.PermissiveTarget.Summary
    }

    switch ([string]$Entry.Type) {
        'RegistryValue' { 'Apply the engine-defined registry permissive value.' }
        'RegistryKeyFlat' { 'Remove this policy registry key.' }
        'PowerShellModuleLogging' { 'Remove the module-logging policy.' }
        'MpPreferenceValue' { 'Apply the engine-defined Defender preference.' }
        'AsrRules' { 'Remove completely captured configured ASR rules.' }
        'ServiceConfig' { 'Apply the engine-defined startup and running state.' }
        'LoadedUserRegistryValues' { 'Apply exact user-scoped targets to captured profiles.' }
        'FirewallProfiles' { 'Apply the permissive firewall profile posture.' }
        'FirewallRules' { 'Enable the captured rules in this firewall group.' }
        'BitLockerVolumes' { 'Suspend eligible protection and enable supported data-volume auto-unlock.' }
        'AppLockerPolicy' { 'Replace the local AppLocker policy with an empty policy.' }
        'ExploitProtectionPolicy' { 'Apply the bundled permissive exploit-protection policy.' }
        'WdacPolicies' { 'Remove identified non-platform WDAC policies.' }
        default { 'Apply the registered engine-defined permissive action.' }
    }
}

function New-SnapshotRow {
    param([Parameter(Mandatory)] [object]$Entry)

    $type = [string]$Entry.Type
    $capabilities = Get-EntryCapabilities -Entry $Entry
    $badges = Get-EntryBadges -Entry $Entry
    $isIncomplete = $badges -match '(^|, )(Partial|Provider missing)(, |$)'
    $canRun = -not $isIncomplete -and ($capabilities.Permissive -or $capabilities.Restore)
    $action = if ($isIncomplete) {
        'Unavailable'
    } elseif ($capabilities.Permissive) {
        'Permissive target'
    } elseif ($capabilities.Restore) {
        'Restore captured'
    } elseif ($capabilities.InventoryOnly) {
        'Inventory only'
    } else {
        'Unavailable'
    }
    $actionOptions = if (-not $canRun) {
        @($action)
    } elseif ($capabilities.Permissive -and $capabilities.Restore) {
        @('Permissive target', 'Restore captured')
    } elseif ($capabilities.Permissive) {
        @('Permissive target')
    } else {
        @('Restore captured')
    }

    [PSCustomObject]@{
        Selected      = $false
        CanRun        = $canRun
        SupportsPermissive = [bool]$capabilities.Permissive
        SupportsRestore = [bool]$capabilities.Restore
        Category      = Get-EntryCategory -Entry $Entry
        Id            = [string]$Entry.Id
        Type          = $type
        Reboot        = if ($Entry.PSObject.Properties['RequiresReboot'] -and [bool]$Entry.RequiresReboot) { 'Yes' } else { 'No' }
        Badges        = $badges
        Current       = Get-EntryCurrentSummary -Entry $Entry
        Compare       = ''
        Difference    = ''
        PermissiveTarget = Get-EntryPermissiveTargetSummary -Entry $Entry
        RestoreTarget = Get-EntryCurrentSummary -Entry $Entry
        Action        = $action
        ActionOptions = @($actionOptions)
        SourceEntry   = $Entry
        ComparisonEntry = $null
    }
}

function Test-SnapshotRowVisible {
    param(
        [Parameter(Mandatory)] [object]$Row,
        [AllowNull()] [string]$SearchText,
        [AllowNull()] [string]$Category,
        [bool]$ChangedOnly = $false
    )

    $normalizedCategory = if ($null -eq $Category) { '' } else { $Category.Trim() }
    if (
        -not [string]::IsNullOrWhiteSpace($normalizedCategory) -and
        $normalizedCategory -ne 'All categories' -and
        -not [string]::Equals([string]$Row.Category, $normalizedCategory, [System.StringComparison]::OrdinalIgnoreCase)
    ) {
        return $false
    }

    $difference = if ($Row.PSObject.Properties['Difference']) { [string]$Row.Difference } else { '' }
    if ($ChangedOnly -and $difference -notin @('Changed', 'Added', 'Removed')) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($SearchText)) {
        return $true
    }

    $haystack = @(
        [string]$Row.Category,
        [string]$Row.Id,
        [string]$Row.Type,
        [string]$Row.Reboot,
        [string]$Row.Badges,
        [string]$Row.Current,
        $(if ($Row.PSObject.Properties['Compare']) { [string]$Row.Compare } else { '' }),
        $difference,
        $(if ($Row.PSObject.Properties['PermissiveTarget']) { [string]$Row.PermissiveTarget } else { '' }),
        [string]$Row.Action
    ) -join ' '
    foreach ($token in @($SearchText.Trim() -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if ($haystack.IndexOf($token, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) {
            return $false
        }
    }

    $true
}

function Get-SnapshotRowSummary {
    param(
        [AllowNull()] [object[]]$Rows,
        [AllowNull()] [object[]]$VisibleRows
    )

    $allRows = @($Rows)
    $visible = @($VisibleRows)
    [PSCustomObject]@{
        Total     = (Get-ItemCount -Value $allRows)
        Visible   = (Get-ItemCount -Value $visible)
        Runnable  = (Get-ItemCount -Value @($visible | Where-Object { [bool]$_.CanRun }))
        Reboot    = (Get-ItemCount -Value @($visible | Where-Object { [string]$_.Reboot -eq 'Yes' }))
        Selected  = (Get-ItemCount -Value @($allRows | Where-Object { [bool]$_.Selected }))
        Changed   = (Get-ItemCount -Value @($allRows | Where-Object { $_.PSObject.Properties['Difference'] -and [string]$_.Difference -eq 'Changed' }))
        Added     = (Get-ItemCount -Value @($allRows | Where-Object { $_.PSObject.Properties['Difference'] -and [string]$_.Difference -eq 'Added' }))
        Removed   = (Get-ItemCount -Value @($allRows | Where-Object { $_.PSObject.Properties['Difference'] -and [string]$_.Difference -eq 'Removed' }))
    }
}

function Read-GuiSnapshot {
    param([Parameter(Mandatory)] [string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        throw "Snapshot was not found: $fullPath"
    }

    $file = Get-Item -LiteralPath $fullPath
    $cacheKey = $fullPath.ToLowerInvariant()
    if ($script:SnapshotCache.ContainsKey($cacheKey)) {
        $cached = $script:SnapshotCache[$cacheKey]
        if ([long]$cached.Length -eq [long]$file.Length -and [long]$cached.LastWriteUtcTicks -eq [long]$file.LastWriteTimeUtc.Ticks) {
            return $cached.Snapshot
        }
    }

    $snapshot = Get-Content -LiteralPath $fullPath -Raw | ConvertFrom-Json
    if ($null -eq $snapshot -or -not $snapshot.PSObject.Properties['Settings']) {
        throw "Snapshot JSON does not contain a Settings collection: $fullPath"
    }
    $script:SnapshotCache[$cacheKey] = [PSCustomObject]@{
        Length            = [long]$file.Length
        LastWriteUtcTicks = [long]$file.LastWriteTimeUtc.Ticks
        Snapshot          = $snapshot
    }
    $snapshot
}

function ConvertTo-GuiFileSize {
    param([Parameter(Mandatory)] [long]$Length)

    if ($Length -ge 1MB) {
        return ('{0:N1} MB' -f ($Length / 1MB))
    }
    if ($Length -ge 1KB) {
        return ('{0:N0} KB' -f ($Length / 1KB))
    }
    "${Length} B"
}

function Get-SnapshotHistoryItems {
    $snapshotDir = Join-Path $script:StateRoot 'snapshots'
    if (-not (Test-Path -LiteralPath $snapshotDir)) {
        return @()
    }

    @(
        foreach ($file in @(Get-ChildItem -LiteralPath $snapshotDir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object -Property LastWriteTime -Descending)) {
            try {
                $capturedText = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                $scopeText = 'full'
                $settingCount = $null
                $cacheKey = $file.FullName.ToLowerInvariant()
                $cachedSnapshot = $null
                if ($script:SnapshotCache.ContainsKey($cacheKey)) {
                    $cachedRecord = $script:SnapshotCache[$cacheKey]
                    if ([long]$cachedRecord.Length -eq [long]$file.Length -and [long]$cachedRecord.LastWriteUtcTicks -eq [long]$file.LastWriteTimeUtc.Ticks) {
                        $cachedSnapshot = $cachedRecord.Snapshot
                    } else {
                        $script:SnapshotCache.Remove($cacheKey)
                    }
                }
                if ($null -ne $cachedSnapshot) {
                    if ($cachedSnapshot.PSObject.Properties['CapturedAtUtc'] -and -not [string]::IsNullOrWhiteSpace([string]$cachedSnapshot.CapturedAtUtc)) {
                        try {
                            $capturedText = ([DateTimeOffset]::Parse([string]$cachedSnapshot.CapturedAtUtc)).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
                        } catch {
                            Write-Verbose ("Could not parse cached snapshot capture time '{0}': {1}" -f $cachedSnapshot.CapturedAtUtc, $_.Exception.Message)
                        }
                    }
                    $settingCount = Get-ItemCount -Value @($cachedSnapshot.Settings)
                    if (
                        $cachedSnapshot.PSObject.Properties['CaptureScope'] -and
                        $null -ne $cachedSnapshot.CaptureScope -and
                        $cachedSnapshot.CaptureScope.PSObject.Properties['IsFiltered'] -and
                        [bool]$cachedSnapshot.CaptureScope.IsFiltered
                    ) {
                        $scopeText = 'scoped'
                    }
                } else {
                    # Snapshot headers are line-oriented; avoid parsing every full history document at startup.
                    $headerText = @(Get-Content -LiteralPath $file.FullName -TotalCount 40 -ErrorAction Stop) -join [Environment]::NewLine
                    if ($headerText -notmatch '(?m)^\s*\{' -or $headerText -notmatch '"Settings"\s*:') {
                        throw 'The bounded snapshot header does not contain a JSON object and Settings collection.'
                    }
                    if ($headerText -match '"CapturedAtUtc"\s*:\s*"([^"]+)"') {
                        try {
                            $capturedText = ([DateTimeOffset]::Parse([string]$matches[1])).ToLocalTime().ToString('yyyy-MM-dd HH:mm:ss')
                        } catch {
                            Write-Verbose ("Could not parse snapshot capture time '{0}': {1}" -f $matches[1], $_.Exception.Message)
                        }
                    }
                    if ($headerText -match '"CaptureScope"\s*:\s*\{[^\r\n]*"IsFiltered"\s*:\s*true') {
                        $scopeText = 'scoped'
                    }
                }
                $sizeText = ConvertTo-GuiFileSize -Length ([long]$file.Length)
                $detailText = if ($null -ne $settingCount) { "$settingCount settings" } else { $sizeText }
                [PSCustomObject]@{
                    Path        = $file.FullName
                    Name        = $file.BaseName
                    Captured    = $capturedText
                    SettingCount = $settingCount
                    Scope       = $scopeText
                    DisplayName = $file.BaseName
                    DisplayMeta = ("{0} / {1} / {2}" -f $capturedText, $detailText, $scopeText)
                    IsReadable  = $true
                }
            } catch {
                [PSCustomObject]@{
                    Path        = $file.FullName
                    Name        = $file.BaseName
                    Captured    = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                    SettingCount = 0
                    Scope       = 'unreadable'
                    DisplayName = $file.BaseName
                    DisplayMeta = ("Unreadable snapshot: {0}" -f $_.Exception.Message)
                    IsReadable  = $false
                }
            }
        }
    )
}

function Rebuild-SnapshotRows {
    $script:Rows.Clear()
    if ($null -eq $script:Snapshot) {
        return
    }

    $comparisonById = @{}
    if ($null -ne $script:ComparisonSnapshot) {
        foreach ($entry in @($script:ComparisonSnapshot.Settings)) {
            $comparisonById[[string]$entry.Id] = $entry
        }
    }

    $loadedIds = @{}
    foreach ($entry in @($script:Snapshot.Settings)) {
        $id = [string]$entry.Id
        $loadedIds[$id] = $true
        $row = New-SnapshotRow -Entry $entry
        if ($null -ne $script:ComparisonSnapshot) {
            if ($comparisonById.ContainsKey($id)) {
                $comparisonEntry = $comparisonById[$id]
                $row.ComparisonEntry = $comparisonEntry
                $row.Compare = Get-EntryCurrentSummary -Entry $comparisonEntry
                $row.Difference = if ((Get-GuiEntryFingerprint -Entry $entry) -eq (Get-GuiEntryFingerprint -Entry $comparisonEntry)) { 'Unchanged' } else { 'Changed' }
            } else {
                $row.Compare = '<not present>'
                $row.Difference = 'Added'
            }
        }
        $script:Rows.Add($row) | Out-Null
    }

    if ($null -ne $script:ComparisonSnapshot) {
        foreach ($comparisonEntry in @($script:ComparisonSnapshot.Settings | Sort-Object -Property Id)) {
            $id = [string]$comparisonEntry.Id
            if ($loadedIds.ContainsKey($id)) {
                continue
            }
            $row = New-SnapshotRow -Entry $comparisonEntry
            $row.Current = '<not present>'
            $row.Compare = Get-EntryCurrentSummary -Entry $comparisonEntry
            $row.Difference = 'Removed'
            $row.Selected = $false
            $row.CanRun = $false
            $row.SupportsPermissive = $false
            $row.SupportsRestore = $false
            $row.Action = 'Unavailable'
            $row.ActionOptions = @('Unavailable')
            $row.Badges = if ([string]::IsNullOrWhiteSpace([string]$row.Badges)) { 'Comparison only' } else { "$($row.Badges), Comparison only" }
            $row.SourceEntry = $null
            $row.ComparisonEntry = $comparisonEntry
            $script:Rows.Add($row) | Out-Null
        }
    }
}

function Set-GuiComparisonState {
    param([Parameter(Mandatory)] [bool]$Active)

    $DifferenceColumn.Visibility = if ($Active) { 'Visible' } else { 'Collapsed' }
    $CompareColumn.Visibility = if ($Active) { 'Visible' } else { 'Collapsed' }
    $isIdle = $null -eq $script:ActiveOperation
    $ChangedOnlyCheckBox.IsEnabled = $Active -and $isIdle
    $ClearCompareButton.IsEnabled = $Active -and $isIdle
    if (-not $Active) {
        $ChangedOnlyCheckBox.IsChecked = $false
        $ComparisonStateText.Text = 'Choose a snapshot to compare against the loaded baseline.'
    } else {
        $ComparisonStateText.Text = "Compared with $([IO.Path]::GetFileName($script:ComparisonSnapshotPath))"
    }
}

function Refresh-SnapshotHistory {
    $selectedPath = if ($null -ne $HistoryCombo.SelectedItem) { [string]$HistoryCombo.SelectedItem.Path } else { $null }
    $history = @(Get-SnapshotHistoryItems)
    $HistoryCombo.ItemsSource = $history
    $HistoryCountText.Text = [string]$history.Count

    $selection = $null
    if (-not [string]::IsNullOrWhiteSpace($selectedPath)) {
        $selection = @($history | Where-Object { [string]::Equals([string]$_.Path, $selectedPath, [System.StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1)
        if ($selection.Count -gt 0) {
            $selection = $selection[0]
        } else {
            $selection = $null
        }
    }
    if ($null -eq $selection) {
        $selection = @($history | Where-Object {
            [bool]$_.IsReadable -and (
                [string]::IsNullOrWhiteSpace($script:SnapshotPath) -or
                -not [string]::Equals([string]$_.Path, $script:SnapshotPath, [System.StringComparison]::OrdinalIgnoreCase)
            )
        } | Select-Object -First 1)
        if ($selection.Count -gt 0) {
            $selection = $selection[0]
        } else {
            $selection = $null
        }
    }
    $HistoryCombo.SelectedItem = $selection
}

function Clear-SnapshotComparison {
    $script:ComparisonSnapshotPath = $null
    $script:ComparisonSnapshot = $null
    Rebuild-SnapshotRows
    Set-GuiComparisonState -Active $false
    Update-GuiCategoryFilter
    Refresh-GuiFilter
}

function Compare-Snapshot {
    param([Parameter(Mandatory)] [string]$Path)

    if ($null -eq $script:Snapshot -or [string]::IsNullOrWhiteSpace($script:SnapshotPath)) {
        throw 'Load a primary snapshot before starting a comparison.'
    }
    $fullPath = [IO.Path]::GetFullPath($Path)
    if ([string]::Equals($fullPath, $script:SnapshotPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Choose a different snapshot to compare with the loaded baseline.'
    }

    $script:ComparisonSnapshot = Read-GuiSnapshot -Path $fullPath
    $script:ComparisonSnapshotPath = $fullPath
    Rebuild-SnapshotRows
    Set-GuiComparisonState -Active $true
    Update-GuiCategoryFilter
    Refresh-GuiFilter
    Set-GuiStatus ("Comparison ready: {0} vs {1}" -f [IO.Path]::GetFileName($script:SnapshotPath), [IO.Path]::GetFileName($fullPath))
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

function Get-CurrentOperationSnapshotPath {
    $operationPath = Join-Path $script:StateRoot 'current-operation.json'
    if (-not (Test-Path -LiteralPath $operationPath)) {
        return $null
    }

    try {
        $operation = Get-Content -LiteralPath $operationPath -Raw | ConvertFrom-Json
        if ($operation.PSObject.Properties['SnapshotPath'] -and -not [string]::IsNullOrWhiteSpace([string]$operation.SnapshotPath)) {
            return [IO.Path]::GetFullPath([string]$operation.SnapshotPath)
        }
    } catch {
        Write-Verbose ("Failed to read current operation file {0}: {1}" -f $operationPath, $_.Exception.Message)
    }

    $null
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
    if ($LogBox.Text.Length -gt 600000) {
        $trimAt = $LogBox.Text.IndexOf([Environment]::NewLine, 100000)
        if ($trimAt -gt 0) {
            $LogBox.Text = $LogBox.Text.Substring($trimAt + [Environment]::NewLine.Length)
        }
    }
    $LogBox.ScrollToEnd()
}

function Get-VisibleSnapshotRows {
    if ($null -eq $script:SettingsView) {
        return @()
    }

    @($script:SettingsView | ForEach-Object { $_ })
}

function Update-GuiSummary {
    $summary = Get-SnapshotRowSummary -Rows @($script:Rows) -VisibleRows @(Get-VisibleSnapshotRows)
    $VisibleCountText.Text = if ($summary.Total -gt 0) { '{0} / {1}' -f $summary.Visible, $summary.Total } else { '0' }
    $RunnableCountText.Text = [string]$summary.Runnable
    $RebootCountText.Text = [string]$summary.Reboot
    $SelectedCountText.Text = [string]$summary.Selected
    $ChangedCountText.Text = [string]$summary.Changed
    $AddedCountText.Text = [string]$summary.Added
    $RemovedCountText.Text = [string]$summary.Removed
}

function Refresh-GuiFilter {
    if ($null -ne $script:SettingsView) {
        $script:SettingsView.Refresh()
    }
    Update-GuiSummary
}

function Update-GuiCategoryFilter {
    $previousCategory = if ($null -eq $CategoryFilter.SelectedItem) { 'All categories' } else { [string]$CategoryFilter.SelectedItem }
    $categories = @('All categories') + @($script:Rows | ForEach-Object { [string]$_.Category } | Sort-Object -Unique)
    $CategoryFilter.ItemsSource = @($categories)
    if (@($categories) -contains $previousCategory) {
        $CategoryFilter.SelectedItem = $previousCategory
    } else {
        $CategoryFilter.SelectedIndex = 0
    }
}

function Load-Snapshot {
    param([Parameter(Mandatory)] [string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    $snapshot = Read-GuiSnapshot -Path $fullPath

    $script:SnapshotPath = $fullPath
    $script:Snapshot = $snapshot
    $script:ComparisonSnapshotPath = $null
    $script:ComparisonSnapshot = $null
    Rebuild-SnapshotRows
    Set-GuiComparisonState -Active $false
    $SnapshotNameText.Text = [IO.Path]::GetFileName($fullPath)
    $SnapshotPathText.Text = $fullPath
    $SnapshotPathText.ToolTip = $fullPath
    $snapshotMetadata = New-Object System.Collections.Generic.List[string]
    $snapshotMetadata.Add(('{0} settings' -f (Get-ItemCount -Value $script:Rows)))
    $status = "Loaded {0} setting(s) from {1}" -f (Get-ItemCount -Value $script:Rows), [IO.Path]::GetFileName($fullPath)
    if ($snapshot.PSObject.Properties['Producer'] -and $null -ne $snapshot.Producer -and $snapshot.Producer.PSObject.Properties['ScriptSha256'] -and -not [string]::IsNullOrWhiteSpace([string]$snapshot.Producer.ScriptSha256)) {
        $shortHash = ([string]$snapshot.Producer.ScriptSha256).Substring(0, [Math]::Min(12, ([string]$snapshot.Producer.ScriptSha256).Length))
        $status += " | Build $shortHash"
        $snapshotMetadata.Add("build $shortHash")
    }
    if ($snapshot.PSObject.Properties['CaptureMetrics'] -and $null -ne $snapshot.CaptureMetrics) {
        $durationMs = if ($snapshot.CaptureMetrics.PSObject.Properties['DurationMs']) { [double]$snapshot.CaptureMetrics.DurationMs } else { 0 }
        $cacheHits = if ($snapshot.CaptureMetrics.PSObject.Properties['CacheHitCount']) { [int]$snapshot.CaptureMetrics.CacheHitCount } else { 0 }
        $status += " | Capture {0:N2}s | {1} shared read(s) reused" -f ($durationMs / 1000), $cacheHits
        $snapshotMetadata.Add(('capture {0:N2}s' -f ($durationMs / 1000)))
        $snapshotMetadata.Add(('{0} shared reads reused' -f $cacheHits))
    }
    $SnapshotMetaText.Text = @($snapshotMetadata) -join ' / '
    Update-GuiCategoryFilter
    Refresh-GuiFilter
    Refresh-SnapshotHistory
    Set-GuiStatus $status
}

function Get-SelectedRows {
    @($script:Rows | Where-Object { $_.Selected })
}

function Get-GuiOperationResultValue {
    param(
        [AllowNull()] [object]$Operation,
        [Parameter(Mandatory)] [string]$Name
    )

    if (
        $null -eq $Operation -or
        -not $Operation.PSObject.Properties['Results'] -or
        $null -eq $Operation.Results -or
        -not $Operation.Results.ContainsKey($Name)
    ) {
        return $null
    }

    [string]$Operation.Results[$Name]
}

function Load-CompletedOperationSnapshot {
    param([Parameter(Mandatory)] [object]$Operation)

    $snapshotPath = Get-GuiOperationResultValue -Operation $Operation -Name 'SnapshotPath'
    if ([string]::IsNullOrWhiteSpace($snapshotPath)) {
        $snapshotPath = Get-LatestSnapshotPath
    }
    if ([string]::IsNullOrWhiteSpace($snapshotPath)) {
        throw 'The operation completed but did not report a snapshot path.'
    }

    Load-Snapshot -Path $snapshotPath
}

function ConvertTo-WindowsCommandLineArgument {
    param([AllowNull()] [string]$Argument)

    if ($null -eq $Argument -or $Argument.Length -eq 0) {
        return '""'
    }

    if ($Argument -notmatch '[\s"]') {
        return $Argument
    }

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append([char]34)
    $backslashCount = 0

    foreach ($character in $Argument.ToCharArray()) {
        if ($character -eq [char]92) {
            $backslashCount++
            continue
        }

        if ($character -eq [char]34) {
            if ($backslashCount -gt 0) {
                [void]$builder.Append([char]92, ($backslashCount * 2))
            }
            [void]$builder.Append([char]92)
            [void]$builder.Append([char]34)
            $backslashCount = 0
            continue
        }

        if ($backslashCount -gt 0) {
            [void]$builder.Append([char]92, $backslashCount)
            $backslashCount = 0
        }
        [void]$builder.Append($character)
    }

    if ($backslashCount -gt 0) {
        [void]$builder.Append([char]92, ($backslashCount * 2))
    }
    [void]$builder.Append([char]34)
    $builder.ToString()
}

function Get-WinDefStateArguments {
    param(
        [Parameter(Mandatory)] [ValidateSet('Snapshot', 'Permissive', 'Restore')] [string]$Command,
        [string]$SnapshotPath,
        [string[]]$IncludeId,
        [string]$CancellationPath,
        [string]$MutationApprovalPath
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
        '-EmitProgress'
        '-ConsoleReport'
        'None'
    )

    if (-not [string]::IsNullOrWhiteSpace($SnapshotPath)) {
        $arguments += @('-SnapshotPath', $SnapshotPath)
    }

    if ($null -ne $IncludeId -and $IncludeId.Length -gt 0) {
        $arguments += @('-IncludeId', (@($IncludeId) -join ','))
    }

    if (-not [string]::IsNullOrWhiteSpace($CancellationPath)) {
        $arguments += @('-CancellationPath', $CancellationPath)
    }

    if (-not [string]::IsNullOrWhiteSpace($MutationApprovalPath)) {
        $arguments += @('-MutationApprovalPath', $MutationApprovalPath)
    }

    @($arguments)
}

function Write-GuiMarkerAtomic {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Content
    )

    $fullPath = [IO.Path]::GetFullPath($Path)
    $tempPath = "$fullPath.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        [IO.File]::WriteAllText($tempPath, $Content, [Text.UTF8Encoding]::new($false))
        Move-Item -LiteralPath $tempPath -Destination $fullPath -Force
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Show-GuiError {
    param([Parameter(Mandatory)] [string]$Message)

    Set-GuiStatus $Message
    Add-GuiLog ("ERROR: {0}" -f $Message)
    [System.Windows.MessageBox]::Show($Message, 'WinDefState', 'OK', 'Error') | Out-Null
}

function Set-GuiOperationState {
    param([Parameter(Mandatory)] [bool]$Running)

    foreach ($control in @(
        $SnapshotButton,
        $LoadLatestButton,
        $BrowseButton,
        $SelectAllButton,
        $ClearButton,
        $RunSelectedButton,
        $SearchBox,
        $CategoryFilter,
        $ClearFilterButton,
        $HistoryCombo,
        $HistoryRefreshButton,
        $CompareButton,
        $ClearCompareButton,
        $ChangedOnlyCheckBox
    )) {
        $control.IsEnabled = -not $Running
    }

    $journalPath = Join-Path $script:StateRoot 'current-operation.json'
    $journalExists = Test-Path -LiteralPath $journalPath
    $restoreSnapshotPath = if ($journalExists) { Get-CurrentOperationSnapshotPath } else { $null }
    $PermissiveButton.IsEnabled = -not $Running -and -not $journalExists
    $RestoreButton.IsEnabled = -not $Running -and -not [string]::IsNullOrWhiteSpace($restoreSnapshotPath)
    $ChangedOnlyCheckBox.IsEnabled = -not $Running -and $null -ne $script:ComparisonSnapshot
    $ClearCompareButton.IsEnabled = -not $Running -and $null -ne $script:ComparisonSnapshot
    $CompareButton.IsEnabled = -not $Running -and $null -ne $HistoryCombo.SelectedItem -and $null -ne $script:Snapshot
    $PermissiveButton.ToolTip = if ($journalExists) { 'Restore the active baseline before starting another permissive operation.' } else { 'Capture a fresh baseline and apply the permissive posture.' }
    $RestoreButton.ToolTip = if ($journalExists -and [string]::IsNullOrWhiteSpace($restoreSnapshotPath)) { 'The operation journal exists but cannot be read. Inspect the state folder.' } else { 'Restore and verify the active operation baseline.' }
    $OperationStateText.Text = if ($Running) { 'OPERATION RUNNING' } else { 'READY' }
    $JournalStateText.Text = if ($journalExists) { 'RESTORE POINT ACTIVE' } else { 'NO ACTIVE CHANGE' }

    $SettingsGrid.IsEnabled = -not $Running
    $CancelButton.Visibility = if ($Running) { 'Visible' } else { 'Collapsed' }
    $CancelButton.IsEnabled = (
        $Running -and
        $null -ne $script:ActiveOperation -and
        [bool]$script:ActiveOperation.CanCancel -and
        -not [bool]$script:ActiveOperation.CancellationRequested
    )
    $OperationProgress.Visibility = if ($Running) { 'Visible' } else { 'Collapsed' }
    $OperationProgress.IsIndeterminate = $Running
    if (-not $Running) {
        $OperationProgress.Value = 0
    }
}

function Set-GuiCancellationAvailability {
    param([AllowNull()] [string]$Phase)

    if ($null -eq $script:ActiveOperation) {
        return
    }

    $script:CurrentPhase = $Phase
    $command = [string]$script:ActiveOperation.Command
    $canCancel = switch ($command) {
        'Snapshot' { [string]::IsNullOrWhiteSpace($Phase) -or $Phase -eq 'Capture' }
        'Permissive' { [string]::IsNullOrWhiteSpace($Phase) -or $Phase -in @('Capture', 'Persist') }
        default { $false }
    }
    if ([bool]$script:ActiveOperation.CancellationRequested) {
        $canCancel = $false
    }

    $script:ActiveOperation.CanCancel = $canCancel
    $CancelButton.IsEnabled = $canCancel
    $CancelButton.ToolTip = if ($canCancel) {
        'Request cancellation at the next read-only boundary. The process will not be killed during mutation.'
    } elseif ($command -eq 'Restore' -or $Phase -in @('Permissive', 'Restore', 'Verify', 'Verify permissive')) {
        'Cancellation is unavailable after mutation begins because restore safety takes priority.'
    } else {
        'Cancellation has already been requested or the operation is completing.'
    }
}

function Update-GuiProgressFromLine {
    param([Parameter(Mandatory)] [string]$Line)

    if ($Line -match '^WDS_REVIEW\|(Permissive|Restore)\|(.*)$') {
        $reviewCommand = [string]$matches[1]
        $reviewSnapshotPath = [string]$matches[2]
        Set-GuiCancellationAvailability -Phase 'Review'
        Add-GuiLog ("Pre-change review ready: {0}" -f $reviewSnapshotPath)
        Set-GuiStatus 'Waiting for pre-change approval. No live setting has changed.'
        try {
            Refresh-SnapshotHistory
            $approved = Show-MutationPreview -Command $reviewCommand -SnapshotPath $reviewSnapshotPath -IncludeId @($script:ActiveOperation.IncludeId)
            if ($approved) {
                Write-GuiMarkerAtomic -Path ([string]$script:ActiveOperation.MutationApprovalPath) -Content 'APPROVE'
                Add-GuiLog 'Pre-change review approved. Mutation may now begin.'
                Set-GuiStatus 'Approved. Starting the reviewed mutation...'
            } else {
                Write-GuiMarkerAtomic -Path ([string]$script:ActiveOperation.MutationApprovalPath) -Content 'CANCEL'
                $script:ActiveOperation.CancellationRequested = $true
                $script:ActiveOperation.CanCancel = $false
                Add-GuiLog 'Pre-change review cancelled. Waiting for the engine to exit before mutation.'
                Set-GuiStatus 'Cancelled in pre-change review. No live setting has changed.'
            }
        } catch {
            $reviewError = $_
            try {
                Write-GuiMarkerAtomic -Path ([string]$script:ActiveOperation.MutationApprovalPath) -Content 'CANCEL'
            } catch {
                Add-GuiLog ("ERROR: The review failed and its protected cancellation decision could not be written: {0}" -f $_.Exception.Message)
            }
            $script:ActiveOperation.CancellationRequested = $true
            $script:ActiveOperation.CanCancel = $false
            Show-GuiError -Message ("Pre-change review failed safely before mutation: {0}" -f $reviewError.Exception.Message)
        }
        return $true
    }

    if ($Line -match '^WDS_APPROVED\|(Permissive|Restore)$') {
        Add-GuiLog ("Engine accepted the {0} approval marker." -f [string]$matches[1])
        return $true
    }

    if ($Line -match '^WDS_RESULT\|([^|]+)\|(.*)$') {
        $name = [string]$matches[1]
        $value = [string]$matches[2]
        if ($null -ne $script:ActiveOperation) {
            $script:ActiveOperation.Results[$name] = $value
        }
        Add-GuiLog ("Result: {0} = {1}" -f $name, $value)
        return $true
    }

    if ($Line -match '^WDS_CANCELLED\|(.*)$') {
        $stage = [string]$matches[1]
        if ($null -ne $script:ActiveOperation) {
            $script:ActiveOperation.CancellationObserved = $true
            $script:ActiveOperation.CanCancel = $false
        }
        $CancelButton.IsEnabled = $false
        $message = "Cancellation accepted at safe boundary: $stage"
        Set-GuiStatus $message
        Add-GuiLog $message
        return $true
    }

    if ($Line -match '^WDS_PROGRESS\|([^|]+)\|(\d+)\|(\d+)\|(.*)$') {
        $phase = [string]$matches[1]
        $current = [int]$matches[2]
        $total = [int]$matches[3]
        $id = [string]$matches[4]
        Set-GuiCancellationAvailability -Phase $phase
        $displayLine = if ([string]::IsNullOrWhiteSpace($id)) {
            "{0} {1}/{2}" -f $phase, $current, $total
        } else {
            "{0} {1}/{2}: {3}" -f $phase, $current, $total, $id
        }

        $OperationProgress.IsIndeterminate = $false
        $OperationProgress.Maximum = [Math]::Max(1, $total)
        $OperationProgress.Value = [Math]::Min($current, $OperationProgress.Maximum)
        Set-GuiStatus $displayLine
        Add-GuiLog $displayLine
        return $true
    }

    $displayLine = $Line -replace '^(VERBOSE|WARNING):\s*', ''
    if ($displayLine -match '\[post\s+(\d+)/(\d+)\]') {
        $OperationProgress.IsIndeterminate = $false
        $OperationProgress.Maximum = [double]$matches[2]
        $OperationProgress.Value = [double]$matches[1]
        Set-GuiStatus $displayLine
        return $true
    }

    if ($displayLine -match '\[(\d+)/(\d+)\]') {
        $OperationProgress.IsIndeterminate = $false
        $OperationProgress.Maximum = [double]$matches[2]
        $OperationProgress.Value = [double]$matches[1]
        Set-GuiStatus $displayLine
        return $true
    }

    $false
}

function Drain-WinDefStateOutput {
    param([Parameter(Mandatory)] [object]$Operation)

    $line = $null
    while ($Operation.Collector.TryDequeue([ref]$line)) {
        if (-not (Update-GuiProgressFromLine -Line $line)) {
            Add-GuiLog $line
        }
        $line = $null
    }
}

function Update-WinDefStateOperation {
    if ($null -eq $script:ActiveOperation) {
        return
    }

    $operation = $script:ActiveOperation
    Drain-WinDefStateOutput -Operation $operation
    if (-not $operation.Process.HasExited) {
        return
    }

    $operation.Process.WaitForExit()
    Drain-WinDefStateOutput -Operation $operation
    $exitCode = $operation.Process.ExitCode
    $completion = $operation.OnCompleted
    $name = [string]$operation.Name
    $cancellationObserved = [bool]$operation.CancellationObserved
    $cancellationPath = [string]$operation.CancellationPath
    $mutationApprovalPath = [string]$operation.MutationApprovalPath

    $script:OperationTimer.Stop()
    $operation.Process.Dispose()
    $script:ActiveOperation = $null
    $script:CurrentPhase = $null
    if (-not [string]::IsNullOrWhiteSpace($cancellationPath) -and (Test-Path -LiteralPath $cancellationPath)) {
        Remove-Item -LiteralPath $cancellationPath -Force -ErrorAction SilentlyContinue
    }
    if (-not [string]::IsNullOrWhiteSpace($mutationApprovalPath) -and (Test-Path -LiteralPath $mutationApprovalPath)) {
        Remove-Item -LiteralPath $mutationApprovalPath -Force -ErrorAction SilentlyContinue
    }
    Set-GuiOperationState -Running $false

    if ($exitCode -ne 0) {
        if ($cancellationObserved) {
            Add-GuiLog ("Cancelled safely: {0}" -f $name)
            Set-GuiStatus ("Cancelled safely: {0}. No defense setting was changed by the cancelled operation." -f $name)
            return
        }
        Show-GuiError -Message ("{0} failed with exit code {1}. Review the log for the provider error." -f $name, $exitCode)
        return
    }

    Add-GuiLog ("Completed: {0}" -f $name)
    Set-GuiStatus ("Completed: {0}" -f $name)
    if ($null -ne $completion) {
        try {
            & $completion $operation
        } catch {
            Show-GuiError -Message $_.Exception.Message
        }
    }
}

function Start-WinDefStateOperation {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [ValidateSet('Snapshot', 'Permissive', 'Restore')] [string]$Command,
        [string]$SnapshotPath,
        [string[]]$IncludeId,
        [scriptblock]$OnCompleted
    )

    if ($null -ne $script:ActiveOperation) {
        throw 'Another WinDefState operation is already running.'
    }

    $cancellationPath = Join-Path ([IO.Path]::GetTempPath()) ("WinDefState-Gui-{0}.cancel" -f [guid]::NewGuid().ToString('N'))
    $mutationApprovalPath = if ($Command -eq 'Snapshot') { $null } else { Join-Path $script:StateRoot (".review-{0}.approve" -f [guid]::NewGuid().ToString('N')) }
    $arguments = @(Get-WinDefStateArguments -Command $Command -SnapshotPath $SnapshotPath -IncludeId $IncludeId -CancellationPath $cancellationPath -MutationApprovalPath $mutationApprovalPath)
    $commandLine = @($arguments | ForEach-Object { ConvertTo-WindowsCommandLineArgument -Argument ([string]$_) }) -join ' '
    Add-GuiLog ("> powershell.exe {0}" -f $commandLine)

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = 'powershell.exe'
    $startInfo.Arguments = $commandLine
    $startInfo.WorkingDirectory = $scriptRoot
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $collector = New-Object WinDefState.Gui.ProcessOutputCollector
    $collector.Attach($process)

    try {
        if (-not $process.Start()) {
            throw 'powershell.exe did not start.'
        }
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
    } catch {
        $process.Dispose()
        throw
    }

    $script:ActiveOperation = [PSCustomObject]@{
        Name        = $Name
        Command     = $Command
        SnapshotPath = $SnapshotPath
        IncludeId    = @($IncludeId)
        Process     = $process
        Collector   = $collector
        OnCompleted = $OnCompleted
        CancellationPath = $cancellationPath
        MutationApprovalPath = $mutationApprovalPath
        CancellationRequested = $false
        CancellationObserved = $false
        Results     = @{}
        CanCancel   = $Command -ne 'Restore'
    }
    $script:CurrentPhase = $null

    if ($null -eq $script:OperationTimer) {
        $script:OperationTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:OperationTimer.Interval = [TimeSpan]::FromMilliseconds(150)
        $script:OperationTimer.Add_Tick({ Update-WinDefStateOperation })
    }

    Set-GuiOperationState -Running $true
    Set-GuiCancellationAvailability -Phase $null
    Set-GuiStatus $Name
    $script:OperationTimer.Start()
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

function Get-MutationPreviewRows {
    param(
        [Parameter(Mandatory)] [object]$Snapshot,
        [Parameter(Mandatory)] [ValidateSet('Permissive', 'Restore')] [string]$Command,
        [AllowNull()] [string[]]$IncludeId
    )

    $includedIds = @{}
    foreach ($value in @($IncludeId)) {
        foreach ($id in @(([string]$value) -split ',')) {
            $trimmedId = $id.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmedId)) {
                $includedIds[$trimmedId] = $true
            }
        }
    }

    @(
        foreach ($entry in @($Snapshot.Settings)) {
            $id = [string]$entry.Id
            if ($includedIds.Count -gt 0 -and -not $includedIds.ContainsKey($id)) {
                continue
            }

            $row = New-SnapshotRow -Entry $entry
            $isRunnable = if ($Command -eq 'Permissive') {
                [bool]$row.CanRun -and [bool]$row.SupportsPermissive
            } else {
                [bool]$row.CanRun -and [bool]$row.SupportsRestore
            }
            if (-not $isRunnable) {
                continue
            }

            $highImpactTypes = @('BitLockerVolumes', 'AppLockerPolicy', 'ExploitProtectionPolicy', 'WdacPolicies')
            $impact = if ([string]$row.Type -in $highImpactTypes) {
                'HIGH IMPACT'
            } elseif ([string]$row.Reboot -eq 'Yes') {
                'REBOOT'
            } else {
                'IMMEDIATE'
            }
            $target = if ($Command -eq 'Permissive') { [string]$row.PermissiveTarget } else { [string]$row.RestoreTarget }
            $before = if ($Command -eq 'Permissive') { [string]$row.Current } else { 'Live state will be verified after restore.' }
            $recovery = if ($Command -eq 'Permissive') { [string]$row.RestoreTarget } else { [string]$row.RestoreTarget }

            [PSCustomObject]@{
                Category = [string]$row.Category
                Id       = $id
                Before   = $before
                Target   = $target
                Recovery = $recovery
                Reboot   = [string]$row.Reboot
                Impact   = $impact
            }
        }
    )
}

$script:MutationPreviewXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Review WinDefState changes" Width="1120" Height="720" MinWidth="920" MinHeight="580"
        WindowStartupLocation="CenterOwner" ResizeMode="CanResize" Background="#F3F6F4"
        FontFamily="Segoe UI Variable Text" FontSize="12" UseLayoutRounding="True" SnapsToDevicePixels="True">
  <Window.Resources>
    <SolidColorBrush x:Key="PreviewInk" Color="#112D34"/>
    <SolidColorBrush x:Key="PreviewMuted" Color="#65777D"/>
    <SolidColorBrush x:Key="PreviewLine" Color="#D8E2DE"/>
    <Style TargetType="{x:Type Button}">
      <Setter Property="Padding" Value="16,8"/><Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/><Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="Foreground" Value="{StaticResource PreviewInk}"/><Setter Property="BorderBrush" Value="#B9C9C4"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type Button}">
            <Border x:Name="Chrome" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Chrome" Property="Opacity" Value="0.9"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Chrome" Property="Opacity" Value="0.78"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="{x:Type DataGridColumnHeader}">
      <Setter Property="Background" Value="#EDF3F0"/><Setter Property="Foreground" Value="#4E6369"/>
      <Setter Property="FontFamily" Value="Bahnschrift SemiCondensed"/><Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="10"/><Setter Property="Padding" Value="10,0"/>
      <Setter Property="BorderBrush" Value="#D8E2DE"/><Setter Property="BorderThickness" Value="0,0,0,1"/>
    </Style>
    <Style TargetType="{x:Type DataGridRow}">
      <Setter Property="Background" Value="#FFFFFF"/><Setter Property="Foreground" Value="{StaticResource PreviewInk}"/>
      <Setter Property="Height" Value="48"/><Setter Property="BorderBrush" Value="#E7EEEB"/><Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Style.Triggers>
        <Trigger Property="AlternationIndex" Value="1"><Setter Property="Background" Value="#FAFCFB"/></Trigger>
        <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#EEF6F3"/></Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="{x:Type DataGridCell}">
      <Setter Property="Padding" Value="10,0"/><Setter Property="BorderThickness" Value="0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
  </Window.Resources>
  <Grid>
    <Grid.RowDefinitions><RowDefinition Height="112"/><RowDefinition Height="96"/><RowDefinition Height="*"/><RowDefinition Height="72"/></Grid.RowDefinitions>
    <Border Grid.Row="0">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,0"><GradientStop Color="#102F37" Offset="0"/><GradientStop Color="#12685D" Offset="1"/></LinearGradientBrush>
      </Border.Background>
      <Grid Margin="24,0">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="PRE-CHANGE REVIEW" Foreground="#90B9B0" FontFamily="Bahnschrift SemiCondensed" FontSize="10" FontWeight="SemiBold"/>
          <TextBlock x:Name="PreviewActionText" Text="Review planned changes" Foreground="#FFFFFF" FontFamily="Bahnschrift SemiCondensed" FontSize="25" FontWeight="SemiBold" Margin="0,3,0,0"/>
          <TextBlock x:Name="PreviewSubtitleText" Foreground="#C9DFDA" Margin="0,3,0,0"/>
        </StackPanel>
        <Border Grid.Column="1" Background="#1FFFFFFF" BorderBrush="#3AFFFFFF" BorderThickness="1" CornerRadius="12" Padding="13,8" VerticalAlignment="Center">
          <StackPanel><TextBlock Text="BASELINE SAVED" Foreground="#92CDBE" FontFamily="Bahnschrift SemiCondensed" FontSize="9" FontWeight="SemiBold"/>
            <TextBlock x:Name="PreviewSnapshotText" Foreground="#FFFFFF" MaxWidth="360" TextTrimming="CharacterEllipsis" Margin="0,2,0,0"/></StackPanel>
        </Border>
      </Grid>
    </Border>
    <Grid Grid.Row="1" Margin="22,14,22,12">
      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="132"/><ColumnDefinition Width="132"/><ColumnDefinition Width="132"/></Grid.ColumnDefinitions>
      <Border Background="#FFF4E8" BorderBrush="#E8BE8E" BorderThickness="1" CornerRadius="10" Padding="14,9" Margin="0,0,12,0">
        <StackPanel><TextBlock Text="NO LIVE SETTING HAS CHANGED YET" Foreground="#A4521F" FontFamily="Bahnschrift SemiCondensed" FontWeight="SemiBold" FontSize="10"/>
          <TextBlock Text="Review the exact persisted baseline and planned targets. Closing this window cancels safely before mutation." Foreground="#6D4C37" Margin="0,3,0,0" TextWrapping="Wrap"/></StackPanel>
      </Border>
      <Border Grid.Column="1" Background="#FFFFFF" BorderBrush="{StaticResource PreviewLine}" BorderThickness="1" CornerRadius="10" Padding="12,8" Margin="0,0,8,0">
        <StackPanel><TextBlock Text="RUNNABLE" Foreground="{StaticResource PreviewMuted}" FontFamily="Bahnschrift SemiCondensed" FontSize="9" FontWeight="SemiBold"/><TextBlock x:Name="PreviewCountText" Text="0" Foreground="#08786A" FontFamily="Bahnschrift SemiCondensed" FontSize="21" FontWeight="SemiBold"/></StackPanel>
      </Border>
      <Border Grid.Column="2" Background="#FFFFFF" BorderBrush="{StaticResource PreviewLine}" BorderThickness="1" CornerRadius="10" Padding="12,8" Margin="0,0,8,0">
        <StackPanel><TextBlock Text="REBOOT" Foreground="{StaticResource PreviewMuted}" FontFamily="Bahnschrift SemiCondensed" FontSize="9" FontWeight="SemiBold"/><TextBlock x:Name="PreviewRebootText" Text="0" Foreground="#B85D24" FontFamily="Bahnschrift SemiCondensed" FontSize="21" FontWeight="SemiBold"/></StackPanel>
      </Border>
      <Border Grid.Column="3" Background="#FFFFFF" BorderBrush="{StaticResource PreviewLine}" BorderThickness="1" CornerRadius="10" Padding="12,8">
        <StackPanel><TextBlock Text="HIGH IMPACT" Foreground="{StaticResource PreviewMuted}" FontFamily="Bahnschrift SemiCondensed" FontSize="9" FontWeight="SemiBold"/><TextBlock x:Name="PreviewImpactText" Text="0" Foreground="#9E3D2D" FontFamily="Bahnschrift SemiCondensed" FontSize="21" FontWeight="SemiBold"/></StackPanel>
      </Border>
    </Grid>
    <Border Grid.Row="2" Margin="22,0" Background="#FFFFFF" BorderBrush="{StaticResource PreviewLine}" BorderThickness="1" CornerRadius="12" ClipToBounds="True">
      <DataGrid x:Name="PreviewGrid" AutomationProperties.Name="Planned setting changes" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="True" HeadersVisibility="Column"
                GridLinesVisibility="None" RowHeaderWidth="0" ColumnHeaderHeight="36" AlternationCount="2"
                EnableRowVirtualization="True" EnableColumnVirtualization="True" VirtualizingPanel.IsVirtualizing="True"
                VirtualizingPanel.VirtualizationMode="Recycling" ScrollViewer.CanContentScroll="True">
        <DataGrid.Columns>
          <DataGridTextColumn Header="CATEGORY" Width="95" Binding="{Binding Category}"/>
          <DataGridTextColumn Header="SETTING ID" Width="245" Binding="{Binding Id}">
            <DataGridTextColumn.ElementStyle><Style TargetType="{x:Type TextBlock}"><Setter Property="FontFamily" Value="Cascadia Mono"/><Setter Property="FontSize" Value="10"/><Setter Property="TextTrimming" Value="CharacterEllipsis"/></Style></DataGridTextColumn.ElementStyle>
          </DataGridTextColumn>
          <DataGridTextColumn Header="BASELINE / SOURCE" Width="190" Binding="{Binding Before}">
            <DataGridTextColumn.ElementStyle><Style TargetType="{x:Type TextBlock}"><Setter Property="TextTrimming" Value="CharacterEllipsis"/><Setter Property="ToolTip" Value="{Binding Before}"/></Style></DataGridTextColumn.ElementStyle>
          </DataGridTextColumn>
          <DataGridTextColumn Header="PLANNED TARGET" Width="*" MinWidth="260" Binding="{Binding Target}">
            <DataGridTextColumn.ElementStyle><Style TargetType="{x:Type TextBlock}"><Setter Property="TextTrimming" Value="CharacterEllipsis"/><Setter Property="ToolTip" Value="{Binding Target}"/></Style></DataGridTextColumn.ElementStyle>
          </DataGridTextColumn>
          <DataGridTextColumn Header="RESTORE VALUE" Width="190" Binding="{Binding Recovery}">
            <DataGridTextColumn.ElementStyle><Style TargetType="{x:Type TextBlock}"><Setter Property="TextTrimming" Value="CharacterEllipsis"/><Setter Property="ToolTip" Value="{Binding Recovery}"/></Style></DataGridTextColumn.ElementStyle>
          </DataGridTextColumn>
          <DataGridTextColumn Header="IMPACT" Width="92" Binding="{Binding Impact}">
            <DataGridTextColumn.ElementStyle>
              <Style TargetType="{x:Type TextBlock}"><Setter Property="FontFamily" Value="Bahnschrift SemiCondensed"/><Setter Property="FontSize" Value="9"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Foreground" Value="#557069"/>
                <Style.Triggers><DataTrigger Binding="{Binding Impact}" Value="REBOOT"><Setter Property="Foreground" Value="#B85D24"/></DataTrigger><DataTrigger Binding="{Binding Impact}" Value="HIGH IMPACT"><Setter Property="Foreground" Value="#9E3D2D"/></DataTrigger></Style.Triggers>
              </Style>
            </DataGridTextColumn.ElementStyle>
          </DataGridTextColumn>
        </DataGrid.Columns>
      </DataGrid>
    </Border>
    <Border Grid.Row="3" Background="#E8EFEC" BorderBrush="#CCD9D5" BorderThickness="0,1,0,0">
      <Grid Margin="22,0"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <TextBlock Text="Approval applies only to this persisted snapshot and this selected scope." Foreground="#60757B" VerticalAlignment="Center"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
          <Button x:Name="PreviewCancelButton" Content="Cancel safely" AutomationProperties.Name="Cancel planned changes"
                  AutomationProperties.HelpText="Close this review without changing live settings." MinWidth="112" Margin="0,0,10,0"/>
          <Button x:Name="PreviewConfirmButton" Content="Approve changes" AutomationProperties.Name="Approve planned changes"
                  AutomationProperties.HelpText="Apply the exact changes listed in this review." MinWidth="145" Background="#08786A" Foreground="#FFFFFF" BorderBrush="#08786A"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

function New-MutationPreviewWindow {
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$script:MutationPreviewXaml)
    $previewWindow = [Windows.Markup.XamlReader]::Load($reader)
    $controls = [ordered]@{}
    foreach ($name in @(
        'PreviewActionText', 'PreviewSubtitleText', 'PreviewSnapshotText', 'PreviewCountText',
        'PreviewRebootText', 'PreviewImpactText', 'PreviewGrid', 'PreviewConfirmButton', 'PreviewCancelButton'
    )) {
        $controls[$name] = $previewWindow.FindName($name)
    }
    $missing = @($controls.GetEnumerator() | Where-Object { $null -eq $_.Value } | ForEach-Object { [string]$_.Key })
    if ($missing.Count -gt 0) {
        throw "The mutation-preview visual tree is missing required control(s): $($missing -join ', ')"
    }

    [PSCustomObject]@{ Window = $previewWindow; Controls = $controls }
}

function Show-MutationPreview {
    param(
        [Parameter(Mandatory)] [ValidateSet('Permissive', 'Restore')] [string]$Command,
        [Parameter(Mandatory)] [string]$SnapshotPath,
        [AllowNull()] [string[]]$IncludeId
    )

    $snapshot = Read-GuiSnapshot -Path $SnapshotPath
    $rows = @(Get-MutationPreviewRows -Snapshot $snapshot -Command $Command -IncludeId $IncludeId)
    if ($rows.Count -eq 0) {
        throw 'The persisted baseline contains no complete runnable setting in the requested scope.'
    }

    $preview = New-MutationPreviewWindow
    $previewWindow = $preview.Window
    $controls = $preview.Controls
    $previewWindow.Owner = $Window
    $controls.PreviewGrid.ItemsSource = $rows
    $controls.PreviewCountText.Text = [string]$rows.Count
    $controls.PreviewRebootText.Text = [string]@($rows | Where-Object { [string]$_.Reboot -eq 'Yes' }).Count
    $controls.PreviewImpactText.Text = [string]@($rows | Where-Object { [string]$_.Impact -eq 'HIGH IMPACT' }).Count
    $controls.PreviewSnapshotText.Text = [IO.Path]::GetFileName($SnapshotPath)
    $controls.PreviewSnapshotText.ToolTip = $SnapshotPath
    if ($Command -eq 'Permissive') {
        $controls.PreviewActionText.Text = 'Approve snapshot + permissive'
        $controls.PreviewSubtitleText.Text = 'The fresh baseline is durable. Approval starts the first live mutation.'
        $controls.PreviewConfirmButton.Content = 'Approve permissive'
    } else {
        $controls.PreviewActionText.Text = 'Approve baseline restore'
        $controls.PreviewSubtitleText.Text = 'The snapshot and sidecars passed preflight validation. Approval starts restore.'
        $controls.PreviewConfirmButton.Content = 'Approve restore'
    }

    $controls.PreviewConfirmButton.Add_Click({ $previewWindow.DialogResult = $true })
    $controls.PreviewCancelButton.Add_Click({ $previewWindow.DialogResult = $false })
    [bool]($previewWindow.ShowDialog() -eq $true)
}

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="WinDefState Control Room" Width="1420" Height="940" MinWidth="1120" MinHeight="760"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI Variable Text" FontSize="12"
        Background="#F2F5F3" UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display">
  <Window.Resources>
    <SolidColorBrush x:Key="InkBrush" Color="#112D34"/>
    <SolidColorBrush x:Key="MutedBrush" Color="#607279"/>
    <SolidColorBrush x:Key="LineBrush" Color="#D8E2DE"/>
    <SolidColorBrush x:Key="SurfaceBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="TealBrush" Color="#08786A"/>
    <SolidColorBrush x:Key="AmberBrush" Color="#B85D24"/>
    <SolidColorBrush x:Key="PaleTealBrush" Color="#E2F0EC"/>
    <SolidColorBrush x:Key="PaleAmberBrush" Color="#F8EBDD"/>

    <Style x:Key="EyebrowTextStyle" TargetType="{x:Type TextBlock}">
      <Setter Property="FontFamily" Value="Bahnschrift SemiCondensed"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
    </Style>
    <Style x:Key="SectionTitleStyle" TargetType="{x:Type TextBlock}">
      <Setter Property="FontFamily" Value="Bahnschrift SemiCondensed"/>
      <Setter Property="FontSize" Value="18"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
    </Style>
    <Style x:Key="CardStyle" TargetType="{x:Type Border}">
      <Setter Property="Background" Value="{StaticResource SurfaceBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource LineBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="14"/>
      <Setter Property="Effect">
        <Setter.Value>
          <DropShadowEffect Color="#1D3840" BlurRadius="18" ShadowDepth="3" Opacity="0.09" RenderingBias="Performance"/>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="SecondaryButtonStyle" TargetType="{x:Type Button}">
      <Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
      <Setter Property="BorderBrush" Value="#BCCBC6"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="13,7"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type Button}">
            <Border x:Name="Chrome" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Chrome" Property="BorderBrush" Value="#6E9187"/>
                <Setter TargetName="Chrome" Property="Opacity" Value="0.94"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="Chrome" Property="Opacity" Value="0.82"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="Chrome" Property="BorderBrush" Value="{StaticResource TealBrush}"/>
                <Setter TargetName="Chrome" Property="BorderThickness" Value="2"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="Chrome" Property="Opacity" Value="0.38"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="{x:Type Button}" BasedOn="{StaticResource SecondaryButtonStyle}"/>
    <Style x:Key="RunbookButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource SecondaryButtonStyle}">
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#22FFFFFF"/>
      <Setter Property="Height" Value="76"/>
      <Setter Property="Padding" Value="16,12"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
    </Style>
    <Style x:Key="FilledButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource SecondaryButtonStyle}">
      <Setter Property="Background" Value="{StaticResource TealBrush}"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="{StaticResource TealBrush}"/>
    </Style>
    <Style x:Key="CancelButtonStyle" TargetType="{x:Type Button}" BasedOn="{StaticResource SecondaryButtonStyle}">
      <Setter Property="Background" Value="#FFF8F4"/>
      <Setter Property="Foreground" Value="#9C3F18"/>
      <Setter Property="BorderBrush" Value="#DDA783"/>
    </Style>
    <Style x:Key="InputTextBoxStyle" TargetType="{x:Type TextBox}">
      <Setter Property="Height" Value="36"/>
      <Setter Property="Padding" Value="11,7"/>
      <Setter Property="Background" Value="#FBFCFB"/>
      <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
      <Setter Property="BorderBrush" Value="#BCCBC6"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type TextBox}">
            <Border x:Name="InputChrome" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="InputChrome" Property="BorderBrush" Value="{StaticResource TealBrush}"/>
                <Setter TargetName="InputChrome" Property="BorderThickness" Value="2"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="InputChrome" Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="FilterComboStyle" TargetType="{x:Type ComboBox}">
      <Setter Property="Height" Value="36"/>
      <Setter Property="Padding" Value="8,5"/>
      <Setter Property="Background" Value="#FBFCFB"/>
      <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
      <Setter Property="BorderBrush" Value="#BCCBC6"/>
    </Style>
    <Style TargetType="{x:Type CheckBox}">
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="{x:Type CheckBox}">
            <Grid Width="18" Height="18">
              <Border x:Name="CheckChrome" Background="#FFFFFF" BorderBrush="#9FB1AB" BorderThickness="1.5" CornerRadius="4"/>
              <Path x:Name="CheckMark" Data="M 3,8 L 7,12 L 15,4" Stroke="#FFFFFF" StrokeThickness="2"
                    StrokeStartLineCap="Round" StrokeEndLineCap="Round" Visibility="Collapsed"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="CheckChrome" Property="Background" Value="{StaticResource TealBrush}"/>
                <Setter TargetName="CheckChrome" Property="BorderBrush" Value="{StaticResource TealBrush}"/>
                <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="CheckChrome" Property="BorderBrush" Value="{StaticResource TealBrush}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="CheckChrome" Property="Opacity" Value="0.35"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="{x:Type DataGridColumnHeader}">
      <Setter Property="Background" Value="#EDF3F0"/>
      <Setter Property="Foreground" Value="#455C63"/>
      <Setter Property="FontFamily" Value="Bahnschrift SemiCondensed"/>
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="10,0"/>
      <Setter Property="BorderBrush" Value="#D8E2DE"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
    </Style>
    <Style TargetType="{x:Type DataGridRow}">
      <Setter Property="Background" Value="#FFFFFF"/>
      <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
      <Setter Property="BorderBrush" Value="#E8EEEB"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="Height" Value="44"/>
      <Style.Triggers>
        <Trigger Property="AlternationIndex" Value="1">
          <Setter Property="Background" Value="#FAFCFB"/>
        </Trigger>
        <DataTrigger Binding="{Binding Difference}" Value="Changed">
          <Setter Property="Background" Value="#FFF8EC"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Difference}" Value="Added">
          <Setter Property="Background" Value="#EDF8F2"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Difference}" Value="Removed">
          <Setter Property="Background" Value="#FFF2EF"/>
          <Setter Property="Foreground" Value="#735E58"/>
        </DataTrigger>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#EEF6F3"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#DCECE7"/>
          <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style TargetType="{x:Type DataGridCell}">
      <Setter Property="Padding" Value="10,0"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
  </Window.Resources>

  <Grid x:Name="RootShell" Opacity="0">
    <Grid.Background>
      <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
        <GradientStop Color="#F7F5EF" Offset="0"/>
        <GradientStop Color="#EEF4F1" Offset="1"/>
      </LinearGradientBrush>
    </Grid.Background>
    <Grid.RowDefinitions>
      <RowDefinition Height="106"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="38"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0">
      <Border.Background>
        <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
          <GradientStop Color="#102F37" Offset="0"/>
          <GradientStop Color="#174A4C" Offset="0.62"/>
          <GradientStop Color="#12685D" Offset="1"/>
        </LinearGradientBrush>
      </Border.Background>
      <Grid ClipToBounds="True">
        <Ellipse Width="430" Height="430" Fill="#10FFFFFF" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,-250,-70,0"/>
        <Ellipse Width="230" Height="230" Fill="#0BFFFFFF" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,40,250,0"/>
        <Grid Margin="28,0">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="WINDEFSTATE" FontFamily="Bahnschrift SemiCondensed" FontSize="26" FontWeight="SemiBold"
                       Foreground="#FFFFFF"/>
            <TextBlock Text="Snapshot precisely. Test safely. Restore exactly." Margin="1,3,0,0"
                       Foreground="#CDE2DD" FontSize="12.5"/>
          </StackPanel>
          <StackPanel Grid.Column="1" VerticalAlignment="Center" HorizontalAlignment="Right">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
              <Border Background="#1FFFFFFF" BorderBrush="#38FFFFFF" BorderThickness="1" CornerRadius="12" Padding="11,5">
                <StackPanel Orientation="Horizontal">
                  <Ellipse Width="7" Height="7" Fill="#71E2BA" Margin="0,0,7,0" VerticalAlignment="Center"/>
                  <TextBlock x:Name="OperationStateText" Text="READY" Foreground="#FFFFFF" FontFamily="Bahnschrift SemiCondensed"
                             FontWeight="SemiBold" FontSize="10"/>
                </StackPanel>
              </Border>
              <Border Background="#1FFFFFFF" BorderBrush="#38FFFFFF" BorderThickness="1" CornerRadius="12" Padding="11,5" Margin="8,0,0,0">
                <TextBlock x:Name="JournalStateText" Text="NO ACTIVE CHANGE" Foreground="#DCEBE7" FontFamily="Bahnschrift SemiCondensed"
                           FontWeight="SemiBold" FontSize="10"/>
              </Border>
            </StackPanel>
            <TextBlock Text="Elevated Windows PowerShell control surface" Foreground="#AFCBC5" FontSize="10.5"
                       HorizontalAlignment="Right" Margin="0,7,0,0"/>
          </StackPanel>
        </Grid>
      </Grid>
    </Border>

    <Border Grid.Row="1" Style="{StaticResource CardStyle}" Margin="24,18,24,14" Padding="20,17">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel>
            <TextBlock Text="PRIMARY RUNBOOK" Style="{StaticResource EyebrowTextStyle}"/>
            <TextBlock Text="Three deliberate operations" Style="{StaticResource SectionTitleStyle}" Margin="0,2,0,0"/>
          </StackPanel>
          <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
            <Button x:Name="CancelButton" Content="Cancel safely" AutomationProperties.Name="Cancel current operation safely"
                    AutomationProperties.HelpText="Request cancellation at the next safe checkpoint." Style="{StaticResource CancelButtonStyle}"
                    Margin="0,0,12,0" Visibility="Collapsed" IsEnabled="False"/>
            <ProgressBar x:Name="OperationProgress" Width="260" Height="7" Foreground="{StaticResource TealBrush}"
                         VerticalAlignment="Center" Visibility="Collapsed" IsIndeterminate="True"/>
          </StackPanel>
        </Grid>
        <UniformGrid Grid.Row="1" Columns="3" Margin="0,14,0,0">
          <Button x:Name="SnapshotButton" AutomationProperties.Name="Snapshot only"
                  AutomationProperties.HelpText="Capture a read-only baseline and report without changing defenses."
                  Style="{StaticResource RunbookButtonStyle}" Background="#173B43" Margin="0,0,8,0"
                  ToolTip="Capture a read-only baseline and report without changing defenses.">
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="43"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <Border Width="34" Height="34" CornerRadius="17" Background="#1FFFFFFF" VerticalAlignment="Center">
                <TextBlock Text="01" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Bahnschrift SemiCondensed" FontWeight="Bold"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="10,0,0,0">
                <TextBlock Text="Snapshot only" FontFamily="Bahnschrift SemiCondensed" FontSize="16" FontWeight="SemiBold"/>
                <TextBlock Text="Read-only baseline and report" Foreground="#C7DCDA" FontSize="10.5" Margin="0,2,0,0"/>
              </StackPanel>
            </Grid>
          </Button>
          <Button x:Name="PermissiveButton" AutomationProperties.Name="Snapshot and apply permissive settings"
                  AutomationProperties.HelpText="Capture a fresh restore point before applying every exact permissive target."
                  Style="{StaticResource RunbookButtonStyle}" Background="#B85D24" Margin="4,0"
                  ToolTip="Capture a fresh baseline, then apply every exact permissive target.">
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="43"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <Border Width="34" Height="34" CornerRadius="17" Background="#24FFFFFF" VerticalAlignment="Center">
                <TextBlock Text="02" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Bahnschrift SemiCondensed" FontWeight="Bold"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="10,0,0,0">
                <TextBlock Text="Snapshot + permissive" FontFamily="Bahnschrift SemiCondensed" FontSize="16" FontWeight="SemiBold"/>
                <TextBlock Text="Creates the restore point first" Foreground="#FFE0CC" FontSize="10.5" Margin="0,2,0,0"/>
              </StackPanel>
            </Grid>
          </Button>
          <Button x:Name="RestoreButton" AutomationProperties.Name="Restore baseline"
                  AutomationProperties.HelpText="Restore and verify the baseline in the active operation journal."
                  Style="{StaticResource RunbookButtonStyle}" Background="#08786A" Margin="8,0,0,0"
                  ToolTip="Restore and verify the baseline in the active operation journal.">
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="43"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <Border Width="34" Height="34" CornerRadius="17" Background="#24FFFFFF" VerticalAlignment="Center">
                <TextBlock Text="03" HorizontalAlignment="Center" VerticalAlignment="Center" FontFamily="Bahnschrift SemiCondensed" FontWeight="Bold"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="10,0,0,0">
                <TextBlock Text="Restore baseline" FontFamily="Bahnschrift SemiCondensed" FontSize="16" FontWeight="SemiBold"/>
                <TextBlock Text="Apply exact state and verify" Foreground="#C9E9E2" FontSize="10.5" Margin="0,2,0,0"/>
              </StackPanel>
            </Grid>
          </Button>
        </UniformGrid>
      </Grid>
    </Border>

    <Border Grid.Row="2" Style="{StaticResource CardStyle}" Margin="24,0,24,18" ClipToBounds="True">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="64"/>
          <RowDefinition Height="72"/>
          <RowDefinition Height="70"/>
          <RowDefinition Height="52"/>
          <RowDefinition Height="*"/>
          <RowDefinition Height="5"/>
          <RowDefinition Height="145"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" BorderBrush="{StaticResource LineBrush}" BorderThickness="0,0,0,1" Padding="18,10">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <StackPanel VerticalAlignment="Center">
              <TextBlock Text="SNAPSHOT EXPLORER" Style="{StaticResource EyebrowTextStyle}"/>
              <StackPanel Orientation="Horizontal" Margin="0,2,0,0">
                <TextBlock x:Name="SnapshotNameText" Text="No snapshot loaded" Foreground="{StaticResource InkBrush}" FontWeight="SemiBold"/>
                <TextBlock Text="  /  " Foreground="#9AACA6"/>
                <TextBlock x:Name="SnapshotMetaText" Text="Take a snapshot to populate the workspace" Foreground="{StaticResource MutedBrush}"/>
              </StackPanel>
              <TextBlock x:Name="SnapshotPathText" Foreground="#82938D" FontSize="9.5" TextTrimming="CharacterEllipsis" Margin="0,2,0,0"/>
            </StackPanel>
            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
              <Button x:Name="LoadLatestButton" Content="Latest" AutomationProperties.Name="Load latest snapshot" Margin="0,0,7,0"/>
              <Button x:Name="BrowseButton" Content="Load..." AutomationProperties.Name="Load snapshot from file" Margin="0,0,7,0"/>
              <Button x:Name="OpenReportButton" Content="Report" AutomationProperties.Name="Open snapshot report" Margin="0,0,7,0"/>
              <Button x:Name="OpenStateButton" Content="State folder" AutomationProperties.Name="Open protected state folder"/>
            </StackPanel>
          </Grid>
        </Border>

        <Border Grid.Row="1" Background="#F8FAF9" BorderBrush="{StaticResource LineBrush}" BorderThickness="0,0,0,1" Padding="18,10">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <StackPanel Orientation="Horizontal">
              <StackPanel Width="330">
                <TextBlock Text="SEARCH SETTINGS" Style="{StaticResource EyebrowTextStyle}" Margin="1,0,0,4"/>
                <TextBox x:Name="SearchBox" AutomationProperties.Name="Search settings"
                         AutomationProperties.HelpText="Search category, ID, type, badges, current value, or action."
                         Style="{StaticResource InputTextBoxStyle}" ToolTip="Search category, ID, type, badges, current value, or action."/>
              </StackPanel>
              <StackPanel Width="190" Margin="12,0,0,0">
                <TextBlock Text="CATEGORY" Style="{StaticResource EyebrowTextStyle}" Margin="1,0,0,4"/>
                <ComboBox x:Name="CategoryFilter" AutomationProperties.Name="Filter by category" Style="{StaticResource FilterComboStyle}"/>
              </StackPanel>
              <Button x:Name="ClearFilterButton" Content="Reset filter" Margin="9,20,0,0" Height="36"/>
            </StackPanel>
            <UniformGrid Grid.Column="1" Columns="4" Width="424" VerticalAlignment="Center">
              <Border Background="#FFFFFF" BorderBrush="{StaticResource LineBrush}" BorderThickness="1" CornerRadius="9" Padding="11,7" Margin="4,0">
                <StackPanel><TextBlock Text="VISIBLE" Style="{StaticResource EyebrowTextStyle}"/><TextBlock x:Name="VisibleCountText" Text="0" FontFamily="Bahnschrift SemiCondensed" FontSize="17" FontWeight="SemiBold" Foreground="{StaticResource InkBrush}"/></StackPanel>
              </Border>
              <Border Background="#FFFFFF" BorderBrush="{StaticResource LineBrush}" BorderThickness="1" CornerRadius="9" Padding="11,7" Margin="4,0">
                <StackPanel><TextBlock Text="RUNNABLE" Style="{StaticResource EyebrowTextStyle}"/><TextBlock x:Name="RunnableCountText" Text="0" FontFamily="Bahnschrift SemiCondensed" FontSize="17" FontWeight="SemiBold" Foreground="{StaticResource TealBrush}"/></StackPanel>
              </Border>
              <Border Background="#FFFFFF" BorderBrush="{StaticResource LineBrush}" BorderThickness="1" CornerRadius="9" Padding="11,7" Margin="4,0">
                <StackPanel><TextBlock Text="REBOOT" Style="{StaticResource EyebrowTextStyle}"/><TextBlock x:Name="RebootCountText" Text="0" FontFamily="Bahnschrift SemiCondensed" FontSize="17" FontWeight="SemiBold" Foreground="{StaticResource AmberBrush}"/></StackPanel>
              </Border>
              <Border Background="#FFFFFF" BorderBrush="{StaticResource LineBrush}" BorderThickness="1" CornerRadius="9" Padding="11,7" Margin="4,0">
                <StackPanel><TextBlock Text="SELECTED" Style="{StaticResource EyebrowTextStyle}"/><TextBlock x:Name="SelectedCountText" Text="0" FontFamily="Bahnschrift SemiCondensed" FontSize="17" FontWeight="SemiBold" Foreground="{StaticResource InkBrush}"/></StackPanel>
              </Border>
            </UniformGrid>
          </Grid>
        </Border>

        <Border Grid.Row="2" Background="#FCFDFD" BorderBrush="{StaticResource LineBrush}" BorderThickness="0,0,0,1" Padding="18,9">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <Grid>
              <Grid.ColumnDefinitions><ColumnDefinition Width="138"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
              <StackPanel VerticalAlignment="Center">
                <StackPanel Orientation="Horizontal"><TextBlock Text="SNAPSHOT HISTORY" Style="{StaticResource EyebrowTextStyle}"/><Border Background="#E2F0EC" CornerRadius="8" Padding="6,1" Margin="7,-1,0,0"><TextBlock x:Name="HistoryCountText" Text="0" Foreground="{StaticResource TealBrush}" FontFamily="Bahnschrift SemiCondensed" FontSize="9" FontWeight="SemiBold"/></Border></StackPanel>
                <TextBlock x:Name="ComparisonStateText" Text="Choose a snapshot to compare against the loaded baseline." Foreground="{StaticResource MutedBrush}" FontSize="9.5" TextTrimming="CharacterEllipsis" Margin="0,4,8,0"/>
              </StackPanel>
              <ComboBox x:Name="HistoryCombo" AutomationProperties.Name="Snapshot comparison history"
                        AutomationProperties.HelpText="Choose a snapshot to compare with the loaded baseline."
                        Grid.Column="1" Height="44" VerticalContentAlignment="Center" Padding="8,3" Background="#FFFFFF" BorderBrush="#BCCBC6">
                <ComboBox.ItemTemplate>
                  <DataTemplate><StackPanel><TextBlock Text="{Binding DisplayName}" Foreground="{StaticResource InkBrush}" FontWeight="SemiBold"/><TextBlock Text="{Binding DisplayMeta}" Foreground="{StaticResource MutedBrush}" FontSize="9.5"/></StackPanel></DataTemplate>
                </ComboBox.ItemTemplate>
              </ComboBox>
              <Button x:Name="HistoryRefreshButton" Grid.Column="2" Content="Refresh" Margin="8,4,0,4"/>
              <Button x:Name="CompareButton" Grid.Column="3" Content="Compare" Style="{StaticResource FilledButtonStyle}" Margin="7,4,0,4"/>
              <Button x:Name="ClearCompareButton" Grid.Column="4" Content="Clear" Margin="7,4,0,4"/>
            </Grid>
            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
              <StackPanel Orientation="Horizontal" Margin="0,0,15,0" VerticalAlignment="Center">
                <CheckBox x:Name="ChangedOnlyCheckBox" AutomationProperties.Name="Show changed settings only" IsEnabled="False" VerticalAlignment="Center"/>
                <TextBlock Text="Changed only" Foreground="{StaticResource MutedBrush}" Margin="7,0,0,0" VerticalAlignment="Center"/>
              </StackPanel>
              <Border Background="#FFF8EC" BorderBrush="#ECD4AA" BorderThickness="1" CornerRadius="9" Padding="10,6" Margin="0,0,6,0"><StackPanel Orientation="Horizontal"><TextBlock Text="CHANGED " Style="{StaticResource EyebrowTextStyle}"/><TextBlock x:Name="ChangedCountText" Text="0" Foreground="#A4521F" FontWeight="SemiBold"/></StackPanel></Border>
              <Border Background="#EDF8F2" BorderBrush="#C7E3D3" BorderThickness="1" CornerRadius="9" Padding="10,6" Margin="0,0,6,0"><StackPanel Orientation="Horizontal"><TextBlock Text="ADDED " Style="{StaticResource EyebrowTextStyle}"/><TextBlock x:Name="AddedCountText" Text="0" Foreground="#08786A" FontWeight="SemiBold"/></StackPanel></Border>
              <Border Background="#FFF2EF" BorderBrush="#E7C7BE" BorderThickness="1" CornerRadius="9" Padding="10,6"><StackPanel Orientation="Horizontal"><TextBlock Text="REMOVED " Style="{StaticResource EyebrowTextStyle}"/><TextBlock x:Name="RemovedCountText" Text="0" Foreground="#9E3D2D" FontWeight="SemiBold"/></StackPanel></Border>
            </StackPanel>
          </Grid>
        </Border>

        <Border Grid.Row="3" BorderBrush="{StaticResource LineBrush}" BorderThickness="0,0,0,1" Padding="18,8">
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <TextBlock Text="ADVANCED SELECTION" Style="{StaticResource EyebrowTextStyle}" VerticalAlignment="Center" Margin="0,0,14,0"/>
              <Button x:Name="SelectAllButton" Content="Select permissive" Margin="0,0,7,0"/>
              <Button x:Name="ClearButton" Content="Clear selection"/>
            </StackPanel>
            <Button x:Name="RunSelectedButton" Grid.Column="1" Content="Run selected" AutomationProperties.Name="Run selected setting actions"
                    AutomationProperties.HelpText="Review and run the actions selected in the settings grid." Style="{StaticResource FilledButtonStyle}" MinWidth="118"/>
          </Grid>
        </Border>

        <DataGrid x:Name="SettingsGrid" AutomationProperties.Name="Snapshot settings" Grid.Row="4"
                  AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                  IsReadOnly="False" HeadersVisibility="Column" GridLinesVisibility="None"
                  SelectionMode="Extended" SelectionUnit="FullRow" Background="#FFFFFF"
                  AlternationCount="2" RowHeaderWidth="0" ColumnHeaderHeight="36" FrozenColumnCount="3"
                  EnableRowVirtualization="True" EnableColumnVirtualization="True"
                  VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling"
                  ScrollViewer.CanContentScroll="True">
          <DataGrid.Columns>
            <DataGridTemplateColumn Header="USE" Width="48">
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <CheckBox AutomationProperties.Name="Select setting" AutomationProperties.HelpText="{Binding Id}" HorizontalAlignment="Center"
                            IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                            IsEnabled="{Binding CanRun}"/>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
            <DataGridTextColumn Header="CATEGORY" Width="110" IsReadOnly="True" Binding="{Binding Category}"/>
            <DataGridTextColumn Header="SETTING ID" Width="300" IsReadOnly="True" Binding="{Binding Id}">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="{x:Type TextBlock}"><Setter Property="FontFamily" Value="Cascadia Mono"/><Setter Property="FontSize" Value="10.5"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="TYPE" Width="165" IsReadOnly="True" Binding="{Binding Type}"/>
            <DataGridTextColumn Header="REBOOT" Width="82" IsReadOnly="True" Binding="{Binding Reboot}">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="{x:Type TextBlock}">
                  <Setter Property="Foreground" Value="#71827C"/><Setter Property="VerticalAlignment" Value="Center"/>
                  <Style.Triggers><DataTrigger Binding="{Binding Reboot}" Value="Yes"><Setter Property="Foreground" Value="{StaticResource AmberBrush}"/><Setter Property="FontWeight" Value="SemiBold"/></DataTrigger></Style.Triggers>
                </Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="BADGES" Width="190" IsReadOnly="True" Binding="{Binding Badges}">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="{x:Type TextBlock}"><Setter Property="Foreground" Value="#9A521F"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="TextTrimming" Value="CharacterEllipsis"/></Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="CHANGE" Width="90" IsReadOnly="True" Binding="{Binding Difference}">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="{x:Type TextBlock}"><Setter Property="Foreground" Value="#71827C"/><Setter Property="FontFamily" Value="Bahnschrift SemiCondensed"/><Setter Property="FontSize" Value="9.5"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="VerticalAlignment" Value="Center"/>
                  <Style.Triggers><DataTrigger Binding="{Binding Difference}" Value="Changed"><Setter Property="Foreground" Value="#A4521F"/></DataTrigger><DataTrigger Binding="{Binding Difference}" Value="Added"><Setter Property="Foreground" Value="#08786A"/></DataTrigger><DataTrigger Binding="{Binding Difference}" Value="Removed"><Setter Property="Foreground" Value="#9E3D2D"/></DataTrigger></Style.Triggers>
                </Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="LOADED" Width="*" MinWidth="210" IsReadOnly="True" Binding="{Binding Current}">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="{x:Type TextBlock}"><Setter Property="Foreground" Value="#41565C"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="TextTrimming" Value="CharacterEllipsis"/></Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTextColumn Header="COMPARE" Width="*" MinWidth="210" IsReadOnly="True" Binding="{Binding Compare}">
              <DataGridTextColumn.ElementStyle>
                <Style TargetType="{x:Type TextBlock}"><Setter Property="Foreground" Value="#6B5650"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="TextTrimming" Value="CharacterEllipsis"/></Style>
              </DataGridTextColumn.ElementStyle>
            </DataGridTextColumn>
            <DataGridTemplateColumn Header="ACTION" Width="164">
              <DataGridTemplateColumn.CellTemplate>
                <DataTemplate>
                  <ComboBox AutomationProperties.Name="Setting action" AutomationProperties.HelpText="{Binding Id}"
                            ItemsSource="{Binding ActionOptions}" SelectedItem="{Binding Action, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                            IsEnabled="{Binding CanRun}" MinWidth="140" Height="30" Padding="6,3"/>
                </DataTemplate>
              </DataGridTemplateColumn.CellTemplate>
            </DataGridTemplateColumn>
          </DataGrid.Columns>
        </DataGrid>

        <GridSplitter Grid.Row="5" Height="5" HorizontalAlignment="Stretch" Background="#D7E2DE" ResizeDirection="Rows" Cursor="SizeNS"/>

        <Border Grid.Row="6" Background="#0F1C23">
          <Grid>
            <Grid.RowDefinitions><RowDefinition Height="34"/><RowDefinition Height="*"/></Grid.RowDefinitions>
            <Border BorderBrush="#2A3C44" BorderThickness="0,0,0,1" Padding="14,0">
              <Grid>
                <TextBlock Text="OPERATION CONSOLE" Foreground="#8FB0A9" FontFamily="Bahnschrift SemiCondensed" FontSize="10" FontWeight="SemiBold" VerticalAlignment="Center"/>
                <Button x:Name="ClearLogButton" Content="Clear output" HorizontalAlignment="Right" VerticalAlignment="Center" Padding="9,3"
                        Background="#172831" Foreground="#C8D9D5" BorderBrush="#39505A" FontSize="10"/>
              </Grid>
            </Border>
            <TextBox x:Name="LogBox" Grid.Row="1" IsReadOnly="True" BorderThickness="0" Padding="14,9"
                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                     TextWrapping="NoWrap" FontFamily="Cascadia Mono" FontSize="10.5" Background="#0F1C23" Foreground="#D7E4E1"/>
          </Grid>
        </Border>
      </Grid>
    </Border>

    <Border Grid.Row="3" Background="#E8EFEC" BorderBrush="#CEDAD6" BorderThickness="0,1,0,0">
      <Grid Margin="24,0">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse Width="7" Height="7" Fill="{StaticResource TealBrush}" Margin="0,0,8,0"/>
          <TextBlock x:Name="StatusText" Text="Ready" Foreground="#3F575D" TextTrimming="CharacterEllipsis" VerticalAlignment="Center"/>
        </StackPanel>
        <TextBlock Grid.Column="1" Text="ADMINISTRATOR / PROTECTED STATE ROOT" Foreground="#758983" FontFamily="Bahnschrift SemiCondensed"
                   FontSize="9.5" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </Grid>
    </Border>
  </Grid>
  <Window.Triggers>
    <EventTrigger RoutedEvent="Window.Loaded">
      <BeginStoryboard>
        <Storyboard>
          <DoubleAnimation Storyboard.TargetName="RootShell" Storyboard.TargetProperty="Opacity" From="0" To="1" Duration="0:0:0.28"/>
        </Storyboard>
      </BeginStoryboard>
    </EventTrigger>
  </Window.Triggers>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

$SnapshotButton = $Window.FindName('SnapshotButton')
$PermissiveButton = $Window.FindName('PermissiveButton')
$RestoreButton = $Window.FindName('RestoreButton')
$LoadLatestButton = $Window.FindName('LoadLatestButton')
$BrowseButton = $Window.FindName('BrowseButton')
$SelectAllButton = $Window.FindName('SelectAllButton')
$ClearButton = $Window.FindName('ClearButton')
$RunSelectedButton = $Window.FindName('RunSelectedButton')
$OpenReportButton = $Window.FindName('OpenReportButton')
$OpenStateButton = $Window.FindName('OpenStateButton')
$CancelButton = $Window.FindName('CancelButton')
$SettingsGrid = $Window.FindName('SettingsGrid')
$SearchBox = $Window.FindName('SearchBox')
$CategoryFilter = $Window.FindName('CategoryFilter')
$ClearFilterButton = $Window.FindName('ClearFilterButton')
$HistoryCombo = $Window.FindName('HistoryCombo')
$HistoryRefreshButton = $Window.FindName('HistoryRefreshButton')
$CompareButton = $Window.FindName('CompareButton')
$ClearCompareButton = $Window.FindName('ClearCompareButton')
$ChangedOnlyCheckBox = $Window.FindName('ChangedOnlyCheckBox')
$HistoryCountText = $Window.FindName('HistoryCountText')
$ComparisonStateText = $Window.FindName('ComparisonStateText')
$ChangedCountText = $Window.FindName('ChangedCountText')
$AddedCountText = $Window.FindName('AddedCountText')
$RemovedCountText = $Window.FindName('RemovedCountText')
$SnapshotPathText = $Window.FindName('SnapshotPathText')
$SnapshotNameText = $Window.FindName('SnapshotNameText')
$SnapshotMetaText = $Window.FindName('SnapshotMetaText')
$LogBox = $Window.FindName('LogBox')
$ClearLogButton = $Window.FindName('ClearLogButton')
$StatusText = $Window.FindName('StatusText')
$OperationProgress = $Window.FindName('OperationProgress')
$OperationStateText = $Window.FindName('OperationStateText')
$JournalStateText = $Window.FindName('JournalStateText')
$VisibleCountText = $Window.FindName('VisibleCountText')
$RunnableCountText = $Window.FindName('RunnableCountText')
$RebootCountText = $Window.FindName('RebootCountText')
$SelectedCountText = $Window.FindName('SelectedCountText')
$RootShell = $Window.FindName('RootShell')
$DifferenceColumn = $SettingsGrid.Columns[6]
$CompareColumn = $SettingsGrid.Columns[8]

$requiredControls = [ordered]@{
    SnapshotButton       = $SnapshotButton
    PermissiveButton     = $PermissiveButton
    RestoreButton        = $RestoreButton
    LoadLatestButton     = $LoadLatestButton
    BrowseButton         = $BrowseButton
    SelectAllButton      = $SelectAllButton
    ClearButton          = $ClearButton
    RunSelectedButton    = $RunSelectedButton
    OpenReportButton     = $OpenReportButton
    OpenStateButton      = $OpenStateButton
    CancelButton         = $CancelButton
    SettingsGrid         = $SettingsGrid
    SearchBox            = $SearchBox
    CategoryFilter       = $CategoryFilter
    ClearFilterButton    = $ClearFilterButton
    HistoryCombo         = $HistoryCombo
    HistoryRefreshButton = $HistoryRefreshButton
    CompareButton        = $CompareButton
    ClearCompareButton   = $ClearCompareButton
    ChangedOnlyCheckBox  = $ChangedOnlyCheckBox
    HistoryCountText     = $HistoryCountText
    ComparisonStateText  = $ComparisonStateText
    ChangedCountText     = $ChangedCountText
    AddedCountText       = $AddedCountText
    RemovedCountText     = $RemovedCountText
    SnapshotPathText     = $SnapshotPathText
    SnapshotNameText     = $SnapshotNameText
    SnapshotMetaText     = $SnapshotMetaText
    LogBox               = $LogBox
    ClearLogButton       = $ClearLogButton
    StatusText           = $StatusText
    OperationProgress    = $OperationProgress
    OperationStateText   = $OperationStateText
    JournalStateText     = $JournalStateText
    VisibleCountText     = $VisibleCountText
    RunnableCountText    = $RunnableCountText
    RebootCountText      = $RebootCountText
    SelectedCountText    = $SelectedCountText
    RootShell            = $RootShell
}
$missingControls = @($requiredControls.GetEnumerator() | Where-Object { $null -eq $_.Value } | ForEach-Object { [string]$_.Key })
if ($missingControls.Count -gt 0) {
    throw "The WPF visual tree is missing required control(s): $($missingControls -join ', ')"
}
if ($ValidateOnly) {
    $previewValidation = New-MutationPreviewWindow
    $previewValidation.Window.Close()
    $Window.Close()
    Write-Host ("WinDefState WPF validation passed with {0} main-window controls and the mutation-preview visual tree." -f $requiredControls.Count)
    return
}

$script:SettingsView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:Rows)
$script:RowFilter = [System.Predicate[object]]{
    param($row)

    $selectedCategory = if ($null -eq $CategoryFilter.SelectedItem) { 'All categories' } else { [string]$CategoryFilter.SelectedItem }
    Test-SnapshotRowVisible -Row $row -SearchText $SearchBox.Text -Category $selectedCategory -ChangedOnly ([bool]$ChangedOnlyCheckBox.IsChecked)
}
$script:SettingsView.Filter = $script:RowFilter
$SettingsGrid.ItemsSource = $script:SettingsView
$CategoryFilter.ItemsSource = @('All categories')
$CategoryFilter.SelectedIndex = 0
$DifferenceColumn.Visibility = 'Collapsed'
$CompareColumn.Visibility = 'Collapsed'
Ensure-Directory -Path $script:StateRoot

$SearchBox.Add_TextChanged({ Refresh-GuiFilter })
$CategoryFilter.Add_SelectionChanged({ Refresh-GuiFilter })
$ChangedOnlyCheckBox.Add_Checked({ Refresh-GuiFilter })
$ChangedOnlyCheckBox.Add_Unchecked({ Refresh-GuiFilter })
$ClearFilterButton.Add_Click({
    $SearchBox.Clear()
    $CategoryFilter.SelectedIndex = 0
    Refresh-GuiFilter
})
$HistoryCombo.Add_SelectionChanged({
    $CompareButton.IsEnabled = (
        $null -eq $script:ActiveOperation -and
        $null -ne $script:Snapshot -and
        $null -ne $HistoryCombo.SelectedItem -and
        [bool]$HistoryCombo.SelectedItem.IsReadable
    )
})
$HistoryRefreshButton.Add_Click({
    Invoke-GuiOperation -Name 'Refreshing snapshot history...' -Operation {
        Refresh-SnapshotHistory
        Set-GuiOperationState -Running $false
        Set-GuiStatus ("Snapshot history refreshed: {0} item(s)." -f $HistoryCountText.Text)
    }
})
$CompareButton.Add_Click({
    Invoke-GuiOperation -Name 'Building snapshot comparison...' -Operation {
        if ($null -eq $HistoryCombo.SelectedItem -or -not [bool]$HistoryCombo.SelectedItem.IsReadable) {
            throw 'Choose a readable snapshot from history first.'
        }
        Compare-Snapshot -Path ([string]$HistoryCombo.SelectedItem.Path)
        Set-GuiOperationState -Running $false
    }
})
$ClearCompareButton.Add_Click({
    Invoke-GuiOperation -Name 'Clearing snapshot comparison...' -Operation {
        Clear-SnapshotComparison
        Set-GuiOperationState -Running $false
        Set-GuiStatus 'Snapshot comparison cleared.'
    }
})
$ClearLogButton.Add_Click({ $LogBox.Clear() })
$selectionSummaryHandler = [System.Windows.RoutedEventHandler]{
    param($selectionSender, $selectionEvent)

    $null = $Window.Dispatcher.BeginInvoke([Action]{ Update-GuiSummary }, [System.Windows.Threading.DispatcherPriority]::Background)
}
$SettingsGrid.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::CheckedEvent, $selectionSummaryHandler)
$SettingsGrid.AddHandler([System.Windows.Controls.Primitives.ToggleButton]::UncheckedEvent, $selectionSummaryHandler)

$CancelButton.Add_Click({
    try {
        if ($null -eq $script:ActiveOperation -or -not [bool]$script:ActiveOperation.CanCancel) {
            throw 'This operation can no longer be cancelled safely. Let it finish so the snapshot or restore journal remains consistent.'
        }

        $cancellationPath = [string]$script:ActiveOperation.CancellationPath
        [IO.File]::WriteAllText($cancellationPath, (Get-Date).ToUniversalTime().ToString('o'), [Text.UTF8Encoding]::new($false))
        $script:ActiveOperation.CancellationRequested = $true
        $script:ActiveOperation.CanCancel = $false
        $CancelButton.IsEnabled = $false
        $CancelButton.ToolTip = 'Cancellation requested. Waiting for the engine to acknowledge a safe boundary.'
        Add-GuiLog 'Cancellation requested. Waiting for a safe read-only boundary; the child process will not be killed.'
        Set-GuiStatus 'Cancellation requested. Waiting for a safe boundary...'
    } catch {
        Show-GuiError -Message $_.Exception.Message
    }
})

$SnapshotButton.Add_Click({
    try {
        Start-WinDefStateOperation -Name 'Capturing a new baseline' -Command Snapshot -OnCompleted {
            param($completedOperation)
            Load-CompletedOperationSnapshot -Operation $completedOperation
        }
    } catch {
        Show-GuiError -Message $_.Exception.Message
    }
})

$PermissiveButton.Add_Click({
    try {
        Start-WinDefStateOperation -Name 'Capturing baseline and applying permissive mode' -Command Permissive -OnCompleted {
            param($completedOperation)
            Load-CompletedOperationSnapshot -Operation $completedOperation
        }
    } catch {
        Show-GuiError -Message $_.Exception.Message
    }
})

$RestoreButton.Add_Click({
    $restoreSnapshotPath = Get-CurrentOperationSnapshotPath
    if ([string]::IsNullOrWhiteSpace($restoreSnapshotPath)) {
        Show-GuiError -Message 'No current operation journal was found. Load a snapshot and use Restore captured for an explicit restore.'
        return
    }

    try {
        Start-WinDefStateOperation -Name 'Restoring and verifying the saved baseline' -Command Restore -SnapshotPath $restoreSnapshotPath -OnCompleted {
            param($completedOperation)
            if (-not [string]::IsNullOrWhiteSpace([string]$completedOperation.SnapshotPath)) {
                Load-Snapshot -Path ([string]$completedOperation.SnapshotPath)
            }
        }
    } catch {
        Show-GuiError -Message $_.Exception.Message
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
        $supportsPermissive = [bool]$row.CanRun -and @($row.ActionOptions) -contains 'Permissive target'
        $row.Selected = $supportsPermissive
        if ($supportsPermissive) {
            $row.Action = 'Permissive target'
        }
    }
    $SettingsGrid.Items.Refresh()
    Update-GuiSummary
    Set-GuiStatus ("Selected {0} permissive setting(s)." -f (Get-ItemCount -Value (Get-SelectedRows)))
})

$ClearButton.Add_Click({
    foreach ($row in $script:Rows) {
        $row.Selected = $false
    }
    $SettingsGrid.Items.Refresh()
    Update-GuiSummary
    Set-GuiStatus 'Selection cleared.'
})

$RunSelectedButton.Add_Click({
    try {
        $selectedRows = @(Get-SelectedRows)
        if ((Get-ItemCount -Value $selectedRows) -eq 0) {
            throw 'Select at least one setting first.'
        }

        $unavailableRows = @($selectedRows | Where-Object { -not [bool]$_.CanRun })
        if ($unavailableRows.Count -gt 0) {
            throw ("Selected row(s) cannot be run because their baseline is incomplete or inventory-only: {0}" -f (@($unavailableRows.Id) -join ', '))
        }

        $actions = @($selectedRows | ForEach-Object { [string]$_.Action } | Sort-Object -Unique)
        if ($actions -contains 'Restore captured' -and (Get-ItemCount -Value $actions) -ne 1) {
            throw 'Restore captured cannot be mixed with permissive rows.'
        }

        $ids = @($selectedRows | ForEach-Object { [string]$_.Id })
        $selectedAction = if ($actions -contains 'Restore captured') {
            'Restore captured'
        } elseif ($actions -contains 'Permissive target') {
            'Permissive target'
        } else {
            throw 'The selected rows do not expose a runnable action.'
        }
        switch ($selectedAction) {
            'Permissive target' {
                Start-WinDefStateOperation -Name ("Applying permissive mode to {0} selected setting(s)" -f $ids.Count) -Command Permissive -IncludeId $ids -OnCompleted {
                    param($completedOperation)
                    Load-CompletedOperationSnapshot -Operation $completedOperation
                }
            }
            'Restore captured' {
                if ([string]::IsNullOrWhiteSpace($script:SnapshotPath)) {
                    throw 'Load a snapshot before restoring selected settings.'
                }
                Start-WinDefStateOperation -Name ("Restoring {0} selected setting(s)" -f $ids.Count) -Command Restore -SnapshotPath $script:SnapshotPath -IncludeId $ids -OnCompleted {
                    param($completedOperation)
                    Load-Snapshot -Path ([string]$completedOperation.SnapshotPath)
                }
            }
            default {
                throw "Unknown action: $selectedAction"
            }
        }
    } catch {
        Show-GuiError -Message $_.Exception.Message
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
    $SnapshotPathText.ToolTip = $script:StateRoot
    $SnapshotNameText.Text = 'No snapshot loaded'
    $SnapshotMetaText.Text = 'Take a snapshot to populate the workspace'
    Set-GuiComparisonState -Active $false
    Refresh-SnapshotHistory
    Update-GuiSummary
    Set-GuiStatus 'Ready. Take a snapshot to populate the grid.'
}
Set-GuiOperationState -Running $false

$Window.Add_Closing({
    param($windowSource, $closingEvent)

    if ($null -ne $script:ActiveOperation -and -not $script:ActiveOperation.Process.HasExited) {
        $closingEvent.Cancel = $true
        [System.Windows.MessageBox]::Show('A WinDefState operation is still running. Wait for it to finish before closing the window.', 'WinDefState', 'OK', 'Information') | Out-Null
    }
})

$Window.ShowDialog() | Out-Null

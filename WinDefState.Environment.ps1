#Requires -Version 5.1
<#!
.SYNOPSIS
Capture a read-only environment baseline before testing, or compare two baselines.
.DESCRIPTION
Inventory only: no remote scanning, policy changes, or restoration. A listener is
not proof of remote reachability. JSON baselines are independent of engine snapshots.
.EXAMPLE
.\WinDefState.Environment.ps1 -OutputPath .\before.json
.EXAMPLE
.\WinDefState.Environment.ps1 -BaselinePath .\before.json -CurrentPath .\after.json -OutputPath .\drift.json
#>
[CmdletBinding()]
param(
    [Alias('OutputPath')][string]$EnvironmentOutputPath,
    [Alias('Format')][ValidateSet('Json', 'Html')][string]$EnvironmentFormat = 'Json',
    [Alias('BaselinePath')][string]$EnvironmentBaselinePath,
    [Alias('CurrentPath')][string]$EnvironmentCurrentPath
)

. (Join-Path $PSScriptRoot 'WinDefState.Health.ps1')

function ConvertTo-EnvironmentValue {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('o') }
    if ($Value -is [enum]) { return [string]$Value }
    if ($Value -is [string] -or $Value.GetType().IsPrimitive -or $Value -is [decimal]) { return $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($key in @($Value.Keys | Sort-Object)) { $result[[string]$key] = ConvertTo-EnvironmentValue $Value[$key] }
        return [pscustomobject]$result
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $values = @(foreach ($item in $Value) { ConvertTo-EnvironmentValue $item })
        return ,$values
    }
    $result = [ordered]@{}
    foreach ($property in @($Value.PSObject.Properties | Sort-Object Name)) { $result[$property.Name] = ConvertTo-EnvironmentValue $property.Value }
    [pscustomobject]$result
}

function New-EnvironmentSection {
    param([string]$Id, [string]$Name, [string[]]$KeyProperties, [scriptblock]$Read,
        [string[]]$IgnoreForComparison = @())
    $timer = [Diagnostics.Stopwatch]::StartNew()
    try {
        $objects = @(& $Read)
        $items = @(foreach ($object in $objects) {
            $data = ConvertTo-EnvironmentValue $object
            $keyParts = @(foreach ($property in $KeyProperties) {
                $part = Get-HealthProperty $data $property
                if ($null -eq $part -or [string]::IsNullOrWhiteSpace([string]$part)) { throw "Missing identity property: $property" }
                [string]$part
            })
            # A JSON tuple avoids delimiter collisions in arbitrary names/paths.
            [pscustomobject]@{ Key = ConvertTo-Json -InputObject $keyParts -Compress; Data = $data }
        })
        $duplicate = @($items | Group-Object Key -CaseSensitive | Where-Object Count -gt 1)
        if ($duplicate.Count -gt 0) { throw 'Provider returned duplicate identities; this section cannot be compared reliably.' }
        [pscustomobject]@{ Id = $Id; Name = $Name; Status = 'Captured'; Error = ''; DurationMs = $timer.ElapsedMilliseconds; IgnoreForComparison = $IgnoreForComparison; Items = @($items | Sort-Object Key) }
    } catch {
        # An unreadable section must never look like an empty successful inventory.
        [pscustomobject]@{ Id = $Id; Name = $Name; Status = 'Unknown'; Error = $_.Exception.Message; DurationMs = $timer.ElapsedMilliseconds; IgnoreForComparison = $IgnoreForComparison; Items = @() }
    }
}

function Get-EnvironmentEndpoint {
    param($Endpoint, [string]$Protocol, [System.Collections.IDictionary]$Processes)
    $owner = $Processes[[string]$Endpoint.OwningProcess]
    $path = [string](Get-HealthProperty $owner ExecutablePath)
    $name = [string](Get-HealthProperty $owner Name)
    [pscustomobject]@{
        Protocol = $Protocol; LocalAddress = [string]$Endpoint.LocalAddress; LocalPort = [int]$Endpoint.LocalPort
        ProcessId = [int]$Endpoint.OwningProcess; ProcessName = $name; ExecutablePath = $path
        OwnerStatus = if ([string]::IsNullOrWhiteSpace($name)) { 'Unknown or exited' } elseif ([string]::IsNullOrWhiteSpace($path)) { 'Path unavailable' } else { 'Resolved' }
    }
}

function Get-WindowsEnvironmentBaseline {
    [CmdletBinding()]
    param()
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Environment capture requires Windows.' }
    if (-not [Environment]::Is64BitProcess -and [Environment]::Is64BitOperatingSystem) { throw 'Use native 64-bit Windows PowerShell for an environment baseline.' }
    $started = (Get-Date).ToUniversalTime().ToString('o')
    $health = Get-WindowsHealthReport
    $processes = @{}
    $processError = ''
    try {
        Get-CimInstance Win32_Process -OperationTimeoutSec 8 -ErrorAction Stop | ForEach-Object { $processes[[string]$_.ProcessId] = $_ }
    } catch { $processError = $_.Exception.Message }
    $sections = @(
        New-EnvironmentSection 'network.tcp' 'TCP listeners' @('Protocol', 'LocalAddress', 'LocalPort', 'ProcessId') {
            Get-NetTCPConnection -State Listen -ErrorAction Stop | ForEach-Object { Get-EnvironmentEndpoint $_ TCP $processes }
        } @('ProcessId')
        New-EnvironmentSection 'network.udp' 'UDP bound endpoints' @('Protocol', 'LocalAddress', 'LocalPort', 'ProcessId') {
            Get-NetUDPEndpoint -ErrorAction Stop | ForEach-Object { Get-EnvironmentEndpoint $_ UDP $processes }
        } @('ProcessId')
        New-EnvironmentSection 'firewall.profiles' 'Effective firewall profiles' @('Name') {
            Get-NetFirewallProfile -PolicyStore ActiveStore -ErrorAction Stop | Select-Object Name, Enabled, DefaultInboundAction, DefaultOutboundAction, AllowInboundRules, AllowLocalFirewallRules, AllowLocalIPsecRules, NotifyOnListen, LogAllowed, LogBlocked, LogFileName, LogMaxSizeKilobytes
        }
        New-EnvironmentSection 'firewall.rules' 'Effective firewall rules' @('Name') {
            Get-NetFirewallRule -PolicyStore ActiveStore -TracePolicyStore -ErrorAction Stop | Select-Object Name, DisplayName, Description, Enabled, Direction, Action, Profile, Group, EdgeTraversalPolicy, EnforcementStatus, PolicyStoreSource, PolicyStoreSourceType, PrimaryStatus, Status
        }
        # Filter records retain their provider identities; do not infer rule joins
        # from display names or silently discard filters when a rule read fails.
        New-EnvironmentSection 'firewall.ports' 'Firewall port filters' @('InstanceID') {
            Get-NetFirewallPortFilter -PolicyStore ActiveStore -ErrorAction Stop | Select-Object InstanceID, Protocol, LocalPort, RemotePort, IcmpType, DynamicTarget
        }
        New-EnvironmentSection 'firewall.addresses' 'Firewall address filters' @('InstanceID') {
            Get-NetFirewallAddressFilter -PolicyStore ActiveStore -ErrorAction Stop | Select-Object InstanceID, LocalAddress, RemoteAddress
        }
        New-EnvironmentSection 'firewall.applications' 'Firewall application filters' @('InstanceID') {
            Get-NetFirewallApplicationFilter -PolicyStore ActiveStore -ErrorAction Stop | Select-Object InstanceID, Program, Package
        }
        New-EnvironmentSection 'firewall.services' 'Firewall service filters' @('InstanceID') {
            Get-NetFirewallServiceFilter -PolicyStore ActiveStore -ErrorAction Stop | Select-Object InstanceID, Service
        }
        New-EnvironmentSection 'firewall.interfaces' 'Firewall interface filters' @('InstanceID') {
            Get-NetFirewallInterfaceFilter -PolicyStore ActiveStore -ErrorAction Stop | Select-Object InstanceID, InterfaceAlias
        }
        New-EnvironmentSection 'firewall.interfaceTypes' 'Firewall interface type filters' @('InstanceID') {
            Get-NetFirewallInterfaceTypeFilter -PolicyStore ActiveStore -ErrorAction Stop | Select-Object InstanceID, InterfaceType
        }
        New-EnvironmentSection 'firewall.security' 'Firewall security filters' @('InstanceID') {
            Get-NetFirewallSecurityFilter -PolicyStore ActiveStore -ErrorAction Stop | Select-Object InstanceID, Authentication, Encryption, OverrideBlockRules, LocalUser, RemoteUser, RemoteMachine
        }
        New-EnvironmentSection 'network.adapters' 'Network adapters' @('InterfaceGuid') {
            Get-NetAdapter -IncludeHidden -ErrorAction Stop | Select-Object InterfaceGuid, Name, InterfaceDescription, Status, MacAddress, LinkSpeed, DriverVersion
        }
        New-EnvironmentSection 'network.addresses' 'IP addresses' @('InterfaceIndex', 'AddressFamily', 'IPAddress', 'PrefixLength') {
            Get-NetIPAddress -ErrorAction Stop | Select-Object InterfaceIndex, InterfaceAlias, AddressFamily, IPAddress, PrefixLength, PrefixOrigin, SuffixOrigin, AddressState, SkipAsSource
        }
        New-EnvironmentSection 'network.dns' 'DNS server configuration' @('InterfaceIndex', 'AddressFamily') {
            Get-DnsClientServerAddress -ErrorAction Stop | Select-Object InterfaceIndex, InterfaceAlias, AddressFamily, ServerAddresses
        }
        New-EnvironmentSection 'network.routes' 'Active routes' @('InterfaceIndex', 'DestinationPrefix', 'NextHop') {
            Get-NetRoute -PolicyStore ActiveStore -ErrorAction Stop | Select-Object InterfaceIndex, InterfaceAlias, DestinationPrefix, NextHop, RouteMetric, Protocol, Publish
        }
        New-EnvironmentSection 'network.profiles' 'Network connection profiles' @('InterfaceIndex', 'Name') {
            Get-NetConnectionProfile -ErrorAction Stop | Select-Object InterfaceIndex, InterfaceAlias, Name, NetworkCategory, IPv4Connectivity, IPv6Connectivity
        }
        New-EnvironmentSection 'system.services' 'Services and startup configuration' @('Name') {
            Get-CimInstance Win32_Service -OperationTimeoutSec 8 -ErrorAction Stop | Select-Object Name, DisplayName, State, StartMode, StartName, PathName, DelayedAutoStart
        }
        New-EnvironmentSection 'system.tasks' 'Scheduled tasks' @('TaskPath', 'TaskName') {
            Get-ScheduledTask -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{ TaskPath = $_.TaskPath; TaskName = $_.TaskName; State = [string]$_.State
                    UserId = $_.Principal.UserId; RunLevel = [string]$_.Principal.RunLevel
                    Actions = @($_.Actions | Select-Object Execute, WorkingDirectory, ClassId)
                    Enabled = $_.Settings.Enabled }
            }
        }
        New-EnvironmentSection 'system.software' 'Machine installed software' @('RegistryPath') {
            foreach ($root in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
                if (Test-Path -LiteralPath $root -ErrorAction Stop) {
                    Get-ChildItem -LiteralPath $root -ErrorAction Stop | Get-ItemProperty -ErrorAction Stop | ForEach-Object {
                        if (-not [string]::IsNullOrWhiteSpace([string](Get-HealthProperty $_ DisplayName))) {
                            [pscustomobject]@{ RegistryPath = $_.PSPath; Name = $_.DisplayName; Version = Get-HealthProperty $_ DisplayVersion; Publisher = Get-HealthProperty $_ Publisher; InstallDate = Get-HealthProperty $_ InstallDate }
                        }
                    }
                }
            }
        }
        New-EnvironmentSection 'system.hotfixes' 'Reported Windows hotfixes' @('HotFixID') {
            Get-CimInstance Win32_QuickFixEngineering -OperationTimeoutSec 8 -ErrorAction Stop | Select-Object HotFixID, Description, InstalledOn
        }
    )
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    [pscustomobject]@{
        SchemaVersion = 1; ReportType = 'WinDefState.Environment'; ComputerName = $env:COMPUTERNAME
        StartedAtUtc = $started; CapturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Elevated = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        Health = $health; ProcessInventoryError = $processError
        Sections = $sections
        Limitations = @(
            'Local TCP listeners and UDP bindings are not proof of external reachability. No network scan was performed.'
            'Capture is sequential, not an atomic image. Processes and endpoints can change during capture.'
            'Inventory only; no automatic restore. Keep a VM checkpoint or disk backup for full recovery.'
            'Software inventory covers machine uninstall registry entries, not per-user or Store applications. Hotfix inventory is not complete patch compliance.'
            'Scheduled-task inventory omits arguments and triggers. No passwords, recovery keys or process command lines are collected intentionally.'
        )
    }
}

function Assert-EnvironmentBaseline {
    param([Parameter(Mandatory)]$Baseline)
    if ((Get-HealthProperty $Baseline ReportType) -ne 'WinDefState.Environment' -or (Get-HealthProperty $Baseline SchemaVersion) -ne 1 -or $null -eq $Baseline.PSObject.Properties['Sections'] -or $null -eq $Baseline.Sections -or [string]::IsNullOrWhiteSpace([string](Get-HealthProperty $Baseline ComputerName))) {
        throw 'Not a supported WinDefState environment baseline (schema 1).'
    }
    $sectionIds = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    foreach ($section in $Baseline.Sections) {
        if ([string]::IsNullOrWhiteSpace([string]$section.Id) -or -not $sectionIds.Add([string]$section.Id) -or $section.Status -notin @('Captured', 'Unknown') -or $null -eq $section.Items) { throw 'Invalid or duplicate environment section.' }
        $keys = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
        foreach ($item in $section.Items) {
            if ([string]::IsNullOrWhiteSpace([string]$item.Key) -or -not $keys.Add([string]$item.Key) -or $null -eq $item.Data) { throw "Invalid or duplicate inventory identity in $($section.Id)." }
        }
    }
}

function Read-EnvironmentBaseline {
    param([Parameter(Mandatory)][string]$Path)
    $file = Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($file.Length -gt 64MB) { throw 'Baseline exceeds the 64 MB reader limit.' }
    $baseline = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    Assert-EnvironmentBaseline $baseline
    $baseline
}

function Get-EnvironmentComparableItems {
    param($Section)
    # Compare endpoint multisets by address/port, excluding PID churn. Keep
    # ownership in the value so ownership changes are changed bindings.
    $items = @(foreach ($item in $Section.Items) {
        $data = [ordered]@{}
        foreach ($property in $item.Data.PSObject.Properties) {
            if ($Section.Id -in @('network.tcp', 'network.udp') -and $property.Name -eq 'ProcessId') { continue }
            $data[$property.Name] = $property.Value
        }
        $key = $item.Key
        if ($Section.Id -in @('network.tcp', 'network.udp')) {
            $key = ConvertTo-Json -InputObject @($data.Protocol, $data.LocalAddress, $data.LocalPort) -Compress
        }
        [pscustomobject]@{ Key = $key; Data = ConvertTo-EnvironmentValue $data }
    })
    $map = [System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
    foreach ($group in @($items | Group-Object Key -CaseSensitive)) {
        $map[$group.Name] = @($group.Group.Data | Sort-Object { ConvertTo-Json -InputObject $_ -Depth 12 -Compress })
    }
    return ,$map
}

function Compare-EnvironmentBaseline {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Before, [Parameter(Mandatory)]$After)
    Assert-EnvironmentBaseline $Before
    Assert-EnvironmentBaseline $After
    if ($Before.ComputerName -ne $After.ComputerName) { throw 'Baselines belong to different computer names; comparison was refused.' }
    $changes = New-Object 'System.Collections.Generic.List[object]'
    $ids = @(@($Before.Sections.Id) + @($After.Sections.Id) | Sort-Object -Unique)
    foreach ($id in $ids) {
        $old = @($Before.Sections | Where-Object Id -eq $id)
        $new = @($After.Sections | Where-Object Id -eq $id)
        if ($old.Count -ne 1 -or $new.Count -ne 1 -or $old[0].Status -ne 'Captured' -or $new[0].Status -ne 'Captured') {
            $changes.Add([pscustomobject]@{ Section = $id; Change = 'Unknown'; Key = ''; Before = $null; After = $null; Detail = 'Section missing or unreadable in at least one capture; no additions or removals inferred.' })
            continue
        }
        $oldMap = Get-EnvironmentComparableItems $old[0]
        $newMap = Get-EnvironmentComparableItems $new[0]
        $keys = @(@($oldMap.Keys) + @($newMap.Keys) | Sort-Object -Unique -CaseSensitive)
        foreach ($key in $keys) {
            $kind = if (-not $oldMap.ContainsKey($key)) { 'Added' } elseif (-not $newMap.ContainsKey($key)) { 'Removed' } elseif ((ConvertTo-Json -InputObject $oldMap[$key] -Depth 14 -Compress) -cne (ConvertTo-Json -InputObject $newMap[$key] -Depth 14 -Compress)) { 'Changed' } else { '' }
            $detail = ''
            if ($id -in @('network.tcp', 'network.udp') -and $oldMap.ContainsKey($key) -and $newMap.ContainsKey($key)) {
                $unresolved = @(@($oldMap[$key]) + @($newMap[$key]) | Where-Object {
                    [string]::IsNullOrWhiteSpace([string](Get-HealthProperty $_ ProcessName)) -or [string]::IsNullOrWhiteSpace([string](Get-HealthProperty $_ ExecutablePath))
                })
                if ($unresolved.Count -gt 0) { $kind = 'Unknown'; $detail = 'Binding observed in both captures, but ownership is not fully resolved. Review both values; an ownership change cannot be established.' }
            }
            if ($kind) { $changes.Add([pscustomobject]@{ Section = $id; Change = $kind; Key = $key; Before = $oldMap[$key]; After = $newMap[$key]; Detail = $detail }) }
        }
    }
    [pscustomobject]@{
        SchemaVersion = 1; ReportType = 'WinDefState.EnvironmentDiff'; ComputerName = $Before.ComputerName
        BeforeCapturedAtUtc = $Before.CapturedAtUtc; AfterCapturedAtUtc = $After.CapturedAtUtc
        Changes = @($changes.ToArray())
        Note = 'Inventory comparison; PID changes are ignored for endpoints. Unknown sections are not unchanged sections. Health checks are retained in each baseline and are not included in this inventory diff.'
    }
}

function ConvertTo-EnvironmentHtml {
    param([Parameter(Mandatory)]$Report)
    $encode = { param($value) [Net.WebUtility]::HtmlEncode([string]$value) }
    $hostLabel = & $encode $Report.ComputerName
    $blocks = if ($Report.ReportType -eq 'WinDefState.EnvironmentDiff') {
        '<p>' + (& $encode $Report.Note) + '</p>'
        foreach ($change in $Report.Changes) {
            '<details open><summary>' + (& $encode "$($change.Change) / $($change.Section) / $($change.Key)") + '</summary><pre>' + (& $encode (ConvertTo-Json -InputObject $change -Depth 14)) + '</pre></details>'
        }
        if (@($Report.Changes).Count -eq 0) { '<p>No inventory differences detected in the captured sections.</p>' }
    } else {
        '<p>Captured ' + (& $encode $Report.CapturedAtUtc) + ' / Elevated: ' + (& $encode $Report.Elevated) + '</p>'
        '<p>' + (& $encode ($Report.Limitations -join ' ')) + '</p>'
        foreach ($section in $Report.Sections) {
            '<details><summary>' + (& $encode "$($section.Name) / $($section.Status) / $(@($section.Items).Count) records") + '</summary><p>' + (& $encode $section.Error) + '</p><pre>' + (& $encode (ConvertTo-Json -InputObject @($section.Items.Data) -Depth 14)) + '</pre></details>'
        }
        '<details><summary>Windows protection health</summary><pre>' + (& $encode (ConvertTo-Json -InputObject $Report.Health -Depth 14)) + '</pre></details>'
    }
    @"
<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>WinDefState environment - $hostLabel</title>
<style>body{font:15px/1.6 system-ui,Segoe UI,sans-serif;background:#f3f6fa;color:#172c3e;margin:0}main{max-width:1300px;margin:auto;padding:36px}h1{font-size:34px}small{color:#08705c;letter-spacing:2px;font-weight:700}p{color:#425870}details{background:white;border:1px solid #d4dfe8;border-radius:8px;padding:14px 18px;margin:12px 0}summary{cursor:pointer;font-weight:600}pre{white-space:pre-wrap;overflow-wrap:anywhere;font:13px/1.6 Consolas,monospace}input{font:inherit;padding:12px;width:90%;max-width:520px;border:1px solid #8ca0b3;border-radius:6px}button{font:inherit;padding:10px;margin:8px;cursor:pointer}@media print{input,button,label{display:none}main{padding:0}}</style></head>
<body><main><small>WINDEFSTATE / ENVIRONMENT</small><h1>$hostLabel baseline inventory</h1><label for="search">Search inventory and evidence</label><p><input id="search" type="search" placeholder="Port, rule, service or application"><button id="expand">Expand all</button><button id="collapse">Collapse all</button></p><div id="records">$($blocks -join "`n")</div><footer>Read-only configuration evidence. Keep the JSON baseline for comparison. This is not a disk image or restore archive.</footer></main>
<script>const sections=Array.from(document.querySelectorAll('details'));document.querySelector('#search').addEventListener('input',e=>{const tokens=e.target.value.toLowerCase().trim().split(/\s+/).filter(Boolean);sections.forEach(s=>{s.hidden=!tokens.every(t=>s.textContent.toLowerCase().includes(t));if(tokens.length&&!s.hidden)s.open=true})});document.querySelector('#expand').onclick=()=>sections.forEach(s=>s.open=true);document.querySelector('#collapse').onclick=()=>sections.forEach(s=>s.open=false);</script></body></html>
"@
}

function Export-EnvironmentReport {
    param([Parameter(Mandatory)]$Report, [Parameter(Mandatory)][string]$Path,
        [ValidateSet('Json', 'Html')][string]$Format = 'Json')
    if ($Format -eq 'Json') {
        $content = ConvertTo-Json -InputObject $Report -Depth 18
    } else { $content = ConvertTo-EnvironmentHtml $Report }
    $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    $stream = [IO.File]::Open($fullPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $bytes = [Text.UTF8Encoding]::new($false).GetBytes($content)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally { $stream.Dispose() }
    $fullPath
}

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    if ($EnvironmentCurrentPath -and -not $EnvironmentBaselinePath) { throw '-CurrentPath requires -BaselinePath.' }
    if ($EnvironmentBaselinePath) {
        $before = Read-EnvironmentBaseline $EnvironmentBaselinePath
        $after = if ($EnvironmentCurrentPath) { Read-EnvironmentBaseline $EnvironmentCurrentPath } else { Get-WindowsEnvironmentBaseline }
        $result = Compare-EnvironmentBaseline $before $after
    } else { $result = Get-WindowsEnvironmentBaseline }
    if ($EnvironmentOutputPath) { $null = Export-EnvironmentReport $result $EnvironmentOutputPath $EnvironmentFormat }
    $result
}

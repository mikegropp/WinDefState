[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Snapshot', 'Permissive', 'Restore')]
    [string]$Command,

    [string]$SnapshotPath,

    [string]$StateRoot = (Join-Path $PSScriptRoot 'state')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Run this script from an elevated PowerShell session.'
    }
}

function Ensure-Directory {
    param([Parameter(Mandatory)] [string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
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

    $tempPath = Join-Path $parent ((Split-Path -Leaf $resolvedPath) + '.tmp')
    $json = $InputObject | ConvertTo-Json -Depth 12
    [System.IO.File]::WriteAllText($tempPath, $json, [System.Text.UTF8Encoding]::new($false))

    if (Test-Path -LiteralPath $resolvedPath) {
        try {
            [System.IO.File]::Replace($tempPath, $resolvedPath, $null)
        } catch {
            Remove-Item -LiteralPath $resolvedPath -Force -ErrorAction SilentlyContinue
            [System.IO.File]::Move($tempPath, $resolvedPath)
        }
    } else {
        [System.IO.File]::Move($tempPath, $resolvedPath)
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

    $tempPath = Join-Path $parent ((Split-Path -Leaf $resolvedPath) + '.tmp')
    [System.IO.File]::WriteAllText($tempPath, $Content, [System.Text.UTF8Encoding]::new($false))

    if (Test-Path -LiteralPath $resolvedPath) {
        try {
            [System.IO.File]::Replace($tempPath, $resolvedPath, $null)
        } catch {
            Remove-Item -LiteralPath $resolvedPath -Force -ErrorAction SilentlyContinue
            [System.IO.File]::Move($tempPath, $resolvedPath)
        }
    } else {
        [System.IO.File]::Move($tempPath, $resolvedPath)
    }
}

function Read-JsonFile {
    param([Parameter(Mandatory)] [string]$Path)

    Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
}

function Get-DefaultSnapshotPath {
    param([Parameter(Mandatory)] [string]$Root)

    $snapshotDir = Join-Path $Root 'snapshots'
    Ensure-Directory -Path $snapshotDir
    Join-Path $snapshotDir "$($env:COMPUTERNAME)-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
}

function Get-SnapshotReportPath {
    param([Parameter(Mandatory)] [string]$SnapshotPath)

    [IO.Path]::ChangeExtension([IO.Path]::GetFullPath($SnapshotPath), 'txt')
}

function Get-VerificationReportPath {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $verificationDir = Join-Path $Root 'verification'
    Ensure-Directory -Path $verificationDir

    $baseName = [IO.Path]::GetFileNameWithoutExtension($SnapshotPath)
    Join-Path $verificationDir "$baseName-restore-check-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
}

function Get-OperationPath {
    param([Parameter(Mandatory)] [string]$Root)

    Ensure-Directory -Path $Root
    Join-Path $Root 'current-operation.json'
}

function Write-OperationState {
    param(
        [Parameter(Mandatory)] [string]$Root,
        [Parameter(Mandatory)] [string]$SnapshotPath,
        [Parameter(Mandatory)] [string]$Mode
    )

    $state = [PSCustomObject]@{
        Mode         = $Mode
        SnapshotPath = [IO.Path]::GetFullPath($SnapshotPath)
        StartedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        ComputerName = $env:COMPUTERNAME
    }

    Write-JsonAtomic -Path (Get-OperationPath -Root $Root) -InputObject $state
}

function Get-OperationState {
    param([Parameter(Mandatory)] [string]$Root)

    $path = Get-OperationPath -Root $Root
    if (-not (Test-Path -LiteralPath $path)) {
        return $null
    }

    Read-JsonFile -Path $path
}

function Clear-OperationState {
    param([Parameter(Mandatory)] [string]$Root)

    $path = Get-OperationPath -Root $Root
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

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

function Get-RegistryValueEntries {
    param([Parameter(Mandatory)] [string]$Path)

    if (-not (Test-Path -Path $Path)) {
        return @()
    }

    $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue
    $key = Get-Item -Path $Path -ErrorAction SilentlyContinue

    if ($null -eq $item -or $null -eq $key) {
        return @()
    }

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

    [PSCustomObject]@{
        Id             = $Id
        Type           = 'RegistryKeyFlat'
        Path           = $Path
        Exists         = Test-Path -Path $Path
        CurrentValue   = @(Get-RegistryValueEntries -Path $Path)
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

    [PSCustomObject]@{
        Id             = 'powershell.module_logging'
        Type           = 'PowerShellModuleLogging'
        BasePath       = $basePath
        Exists         = Test-Path -Path $basePath
        CurrentValue   = [PSCustomObject]@{
            BaseValues        = @(Get-RegistryValueEntries -Path $basePath)
            ModuleNamesExists = Test-Path -Path $moduleNamesPath
            ModuleNamesValues = @(Get-RegistryValueEntries -Path $moduleNamesPath)
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

function Get-LoadedUserRegistryValueStates {
    param([Parameter(Mandatory)] [object[]]$Items)

    $entries = @()

    foreach ($sid in @(Get-LoadedUserSids)) {
        foreach ($item in @($Items)) {
            $path = "Registry::HKEY_USERS\$sid\$($item.RelativePath)"
            $property = Get-ItemProperty -Path $path -Name $item.Name -ErrorAction SilentlyContinue
            $exists = $null -ne $property -and $null -ne $property.$($item.Name)
            $kind = if ($exists) {
                try { (Get-Item -Path $path).GetValueKind($item.Name).ToString() } catch { $item.ValueKind }
            } else {
                $item.ValueKind
            }

            $entries += [PSCustomObject]@{
                Sid          = $sid
                RelativePath = $item.RelativePath
                Name         = $item.Name
                Exists       = $exists
                CurrentValue = if ($exists) { $property.$($item.Name) } else { $null }
                ValueKind    = $kind
            }
        }
    }

    @($entries)
}

function Set-Permissive-LoadedUserRegistryValues {
    param([Parameter(Mandatory)] [object[]]$Items)

    foreach ($sid in @(Get-LoadedUserSids)) {
        foreach ($item in @($Items)) {
            $path = "Registry::HKEY_USERS\$sid\$($item.RelativePath)"

            if ($item.PermissiveExists) {
                Ensure-RegistryPath -Path $path
                New-ItemProperty -Path $path -Name $item.Name -PropertyType $item.ValueKind -Value $item.PermissiveValue -Force | Out-Null
            } else {
                Remove-ItemProperty -Path $path -Name $item.Name -ErrorAction SilentlyContinue
            }
        }
    }
}

function Restore-LoadedUserRegistryValues {
    param([Parameter(Mandatory)] [object[]]$Entries)

    foreach ($entry in @($Entries)) {
        $path = "Registry::HKEY_USERS\$($entry.Sid)\$($entry.RelativePath)"

        if ($entry.Exists) {
            Ensure-RegistryPath -Path $path
            New-ItemProperty -Path $path -Name $entry.Name -PropertyType $entry.ValueKind -Value $entry.CurrentValue -Force | Out-Null
        } else {
            Remove-ItemProperty -Path $path -Name $entry.Name -ErrorAction SilentlyContinue
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
        return
    }

    if ($Enabled) {
        Enable-LocalUser -SID $user.SID -ErrorAction SilentlyContinue
    } else {
        Disable-LocalUser -SID $user.SID -ErrorAction SilentlyContinue
    }
}

function Get-NetBiosAdapterStates {
    $adapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = TRUE' -ErrorAction SilentlyContinue

    $states = foreach ($adapter in @($adapters)) {
        [PSCustomObject]@{
            Index               = [int]$adapter.Index
            Description         = $adapter.Description
            TcpipNetbiosOptions = [int]$adapter.TcpipNetbiosOptions
        }
    }

    @($states)
}

function Set-NetBiosAdapterOption {
    param(
        [Parameter(Mandatory)] [int]$Index,
        [Parameter(Mandatory)] [uint32]$Option
    )

    $adapter = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "Index = $Index" -ErrorAction SilentlyContinue
    if ($null -eq $adapter) {
        return
    }

    Invoke-CimMethod -InputObject $adapter -MethodName SetTcpipNetbios -Arguments @{ TcpipNetbiosOptions = $Option } | Out-Null
}

function Set-Permissive-NetBiosAdapters {
    foreach ($adapter in @(Get-NetBiosAdapterStates)) {
        Set-NetBiosAdapterOption -Index $adapter.Index -Option 1
    }
}

function Restore-NetBiosAdapters {
    param([Parameter(Mandatory)] [object[]]$Adapters)

    foreach ($adapter in @($Adapters)) {
        Set-NetBiosAdapterOption -Index ([int]$adapter.Index) -Option ([uint32]$adapter.TcpipNetbiosOptions)
    }
}

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
    param([Parameter(Mandatory)] [object]$Action)

    switch ([string]$Action) {
        '0' { 'Disabled' }
        '1' { 'Block' }
        '2' { 'Audit' }
        '6' { 'Warn' }
        'Disabled' { 'Disabled' }
        'Enabled' { 'Block' }
        'AuditMode' { 'Audit' }
        'Warn' { 'Warn' }
        default { "Unknown ($Action)" }
    }
}

function Get-AsrRestoreAction {
    param([Parameter(Mandatory)] [object]$Action)

    switch ([string]$Action) {
        '0' { 'Disabled' }
        '1' { 'Enabled' }
        '2' { 'AuditMode' }
        '6' { 'Warn' }
        'Disabled' { 'Disabled' }
        'Enabled' { 'Enabled' }
        'AuditMode' { 'AuditMode' }
        'Warn' { 'Warn' }
        default { $null }
    }
}

function Get-ConfiguredAsrRules {
    $catalog = Get-AsrRuleCatalog
    $mp = Get-MpPreference -ErrorAction SilentlyContinue
    if ($null -eq $mp) {
        return @()
    }

    $ids = @($mp.AttackSurfaceReductionRules_Ids)
    $actions = @($mp.AttackSurfaceReductionRules_Actions)
    $rules = @()

    for ($i = 0; $i -lt $ids.Count; $i++) {
        $id = [string]$ids[$i]
        $action = if ($actions.Count -gt $i) { [string]$actions[$i] } else { $null }

        $rules += [PSCustomObject]@{
            Id          = $id
            Name        = if ($catalog.ContainsKey($id)) { $catalog[$id] } else { 'Unknown / custom rule' }
            Action      = $action
            ActionLabel = Get-AsrActionLabel -Action $action
        }
    }

    $rules
}

function Disable-ConfiguredAsrRules {
    $mp = Get-MpPreference -ErrorAction SilentlyContinue
    if ($null -eq $mp) {
        return
    }

    foreach ($id in @($mp.AttackSurfaceReductionRules_Ids)) {
        Remove-MpPreference -AttackSurfaceReductionRules_Ids $id -ErrorAction SilentlyContinue
    }
}

function Restore-AsrRules {
    param([Parameter(Mandatory)] [object[]]$Rules)

    Disable-ConfiguredAsrRules

    $ids = @()
    $actions = @()

    foreach ($rule in @($Rules)) {
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

function Set-MpPreferencePropertyValue {
    param(
        [Parameter(Mandatory)] [string]$Property,
        [Parameter(Mandatory)] [object]$Value
    )

    $params = @{}
    $params[$Property] = $Value
    Set-MpPreference @params
}

function Resolve-MpPreferenceValue {
    param(
        [Parameter(Mandatory)] [object]$Definition,
        [Parameter(Mandatory)] [object]$Value
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

function Get-AuditPolicyState {
    param([Parameter(Mandatory)] [string]$Subcategory)

    $csv = auditpol /get /subcategory:"$Subcategory" /r
    $rows = $csv | Where-Object { $_ -match ',' } | ConvertFrom-Csv
    $row = $rows | Where-Object { $_.Subcategory -eq $Subcategory } | Select-Object -First 1

    if ($null -eq $row) {
        return [PSCustomObject]@{
            Success = $false
            Failure = $false
        }
    }

    $inclusion = [string]$row.'Inclusion Setting'
    [PSCustomObject]@{
        Success = $inclusion -match 'Success'
        Failure = $inclusion -match 'Failure'
    }
}

function Set-AuditPolicyState {
    param(
        [Parameter(Mandatory)] [string]$Subcategory,
        [Parameter(Mandatory)] [bool]$Success,
        [Parameter(Mandatory)] [bool]$Failure
    )

    $successValue = if ($Success) { 'enable' } else { 'disable' }
    $failureValue = if ($Failure) { 'enable' } else { 'disable' }
    auditpol /set /subcategory:"$Subcategory" /success:$successValue /failure:$failureValue | Out-Null
}

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
        [Parameter(Mandatory)] [object]$Entry
    )

    $Lines.Add("[$($Entry.Id)] $($Entry.Type)")
    Add-ReportKeyValueLine -Lines $Lines -Label 'Requires reboot' -Value $Entry.RequiresReboot

    switch ($Entry.Type) {
        'RegistryValue' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Path' -Value $Entry.Path
            Add-ReportKeyValueLine -Lines $Lines -Label 'Name' -Value $Entry.Name
            Add-ReportKeyValueLine -Lines $Lines -Label 'Exists' -Value $Entry.Exists
            Add-ReportKeyValueLine -Lines $Lines -Label 'Value kind' -Value $Entry.ValueKind
            Add-ReportKeyValueLine -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
        }
        'RegistryKeyFlat' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Path' -Value $Entry.Path
            Add-ReportKeyValueLine -Lines $Lines -Label 'Exists' -Value $Entry.Exists
            Add-ReportJsonBlock -Lines $Lines -Label 'Values' -Value @($Entry.CurrentValue)
        }
        'MpPreferenceValue' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Property' -Value $Entry.Property
            Add-ReportKeyValueLine -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
            Add-ReportKeyValueLine -Lines $Lines -Label 'Restore value' -Value $Entry.RestoreValue
        }
        'AsrRules' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Configured rule count' -Value @($Entry.CurrentValue).Count
            foreach ($rule in @($Entry.CurrentValue)) {
                $Lines.Add(('  - {0} | {1} | {2}' -f $rule.Id, $rule.Name, $rule.ActionLabel))
            }
        }
        'PowerShellModuleLogging' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Base path' -Value $Entry.BasePath
            Add-ReportKeyValueLine -Lines $Lines -Label 'Exists' -Value $Entry.Exists
            Add-ReportJsonBlock -Lines $Lines -Label 'Base values' -Value @($Entry.CurrentValue.BaseValues)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Module names key exists' -Value $Entry.CurrentValue.ModuleNamesExists
            Add-ReportJsonBlock -Lines $Lines -Label 'Module names values' -Value @($Entry.CurrentValue.ModuleNamesValues)
        }
        'FirewallProfiles' {
            foreach ($profile in @($Entry.CurrentValue)) {
                $Lines.Add(('  - {0}: {1}' -f $profile.Profile, $profile.Enabled))
            }
        }
        'NetBiosAdapters' {
            foreach ($adapter in @($Entry.CurrentValue)) {
                $Lines.Add(('  - Index {0} | {1} | TcpipNetbiosOptions={2}' -f $adapter.Index, $adapter.Description, $adapter.TcpipNetbiosOptions))
            }
        }
        'LoadedUserRegistryValues' {
            foreach ($value in @($Entry.CurrentValue)) {
                $displayValue = if ($value.Exists) { ConvertTo-DisplayString -Value $value.CurrentValue } else { '<absent>' }
                $Lines.Add(('  - {0} | {1} | {2} | {3}' -f $value.Sid, $value.RelativePath, $value.Name, $displayValue))
            }
        }
        'MachineEnvironmentValue' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Name' -Value $Entry.Name
            Add-ReportKeyValueLine -Lines $Lines -Label 'Exists' -Value $Entry.Exists
            Add-ReportKeyValueLine -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
        }
        'ServiceConfig' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Name' -Value $Entry.Name
            Add-ReportKeyValueLine -Lines $Lines -Label 'Start mode' -Value $Entry.CurrentValue.StartMode
            Add-ReportKeyValueLine -Lines $Lines -Label 'State' -Value $Entry.CurrentValue.State
        }
        'LocalUser' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Name' -Value $Entry.Name
            Add-ReportKeyValueLine -Lines $Lines -Label 'SID' -Value $Entry.Sid
            Add-ReportKeyValueLine -Lines $Lines -Label 'Enabled' -Value $Entry.CurrentValue
        }
        'WsManValue' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Path' -Value $Entry.Path
            Add-ReportKeyValueLine -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
        }
        'AuditPolicy' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Subcategory' -Value $Entry.Subcategory
            Add-ReportKeyValueLine -Lines $Lines -Label 'Success' -Value $Entry.CurrentValue.Success
            Add-ReportKeyValueLine -Lines $Lines -Label 'Failure' -Value $Entry.CurrentValue.Failure
        }
        'SmbClientConfig' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Require security signature' -Value $Entry.CurrentValue
        }
        'SmbServerConfig' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Require security signature' -Value $Entry.CurrentValue
        }
        default {
            Add-ReportJsonBlock -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
        }
    }
}

function Get-SnapshotReportLines {
    param(
        [Parameter(Mandatory)] [object]$Snapshot,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $settings = @($Snapshot.Settings)

    $lines.Add('WinDefState Snapshot Report')
    $lines.Add(('Snapshot JSON: {0}' -f $SnapshotPath))
    $lines.Add(('ComputerName: {0}' -f $Snapshot.ComputerName))
    $lines.Add(('CapturedAtUtc: {0}' -f $Snapshot.CapturedAtUtc))
    $lines.Add(('Settings captured: {0}' -f $settings.Count))
    $lines.Add(('Reboot-required settings: {0}' -f (@($settings | Where-Object { $_.RequiresReboot }).Count)))

    foreach ($entry in $settings) {
        $lines.Add(' ')
        Add-SnapshotEntryReportLines -Lines $lines -Entry $entry
    }

    [string[]]$lines
}

function Show-ReportLines {
    param([AllowEmptyString()] [string[]]$Lines)

    foreach ($line in @($Lines)) {
        Write-Host $line
    }
}

function ConvertTo-ComparableSnapshotEntry {
    param([Parameter(Mandatory)] [object]$Entry)

    switch ($Entry.Type) {
        'RegistryKeyFlat' {
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Path           = $Entry.Path
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
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Property       = $Entry.Property
                RestoreValue   = $Entry.RestoreValue
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'AsrRules' {
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = @(
                    foreach ($rule in @($Entry.CurrentValue)) {
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
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                BasePath       = $Entry.BasePath
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
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = @(
                    foreach ($profile in @($Entry.CurrentValue)) {
                        [PSCustomObject]@{
                            Profile = $profile.Profile
                            Enabled = [bool]$profile.Enabled
                        }
                    }
                )
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'NetBiosAdapters' {
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
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
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = @(
                    foreach ($value in @($Entry.CurrentValue)) {
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
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        default {
            return ConvertTo-CanonicalValue -Value $Entry
        }
    }
}

function Test-DefenseSnapshot {
    param([Parameter(Mandatory)] [object]$Snapshot)

    $definitions = @{}
    foreach ($definition in Get-DefenseDefinitions) {
        $definitions[[string]$definition.Id] = $definition
    }

    $results = @()
    foreach ($entry in @($Snapshot.Settings)) {
        $entryId = [string]$entry.Id
        $expectedComparable = ConvertTo-ComparableSnapshotEntry -Entry $entry

        if (-not $definitions.ContainsKey($entryId)) {
            $results += [PSCustomObject]@{
                Id             = $entryId
                Type           = $entry.Type
                RequiresReboot = $entry.RequiresReboot
                Matches        = $false
                Reason         = 'This snapshot entry no longer has a matching definition in the current script.'
                Expected       = $expectedComparable
                Actual         = $null
            }
            continue
        }

        try {
            $liveEntry = Capture-Definition -Definition $definitions[$entryId]
            $actualComparable = ConvertTo-ComparableSnapshotEntry -Entry $liveEntry
            $matches = (Get-CanonicalJson -Value $expectedComparable) -eq (Get-CanonicalJson -Value $actualComparable)

            $results += [PSCustomObject]@{
                Id             = $entryId
                Type           = $entry.Type
                RequiresReboot = $entry.RequiresReboot
                Matches        = $matches
                Reason         = if ($matches) { $null } else { 'Live state does not match the snapshot entry.' }
                Expected       = $expectedComparable
                Actual         = $actualComparable
            }
        } catch {
            $results += [PSCustomObject]@{
                Id             = $entryId
                Type           = $entry.Type
                RequiresReboot = $entry.RequiresReboot
                Matches        = $false
                Reason         = $_.Exception.Message
                Expected       = $expectedComparable
                Actual         = $null
            }
        }
    }

    [PSCustomObject]@{
        Tool          = 'WinDefState'
        ComputerName  = $env:COMPUTERNAME
        VerifiedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        MatchedCount  = @($results | Where-Object { $_.Matches }).Count
        MismatchCount = @($results | Where-Object { -not $_.Matches }).Count
        Results       = @($results)
    }
}

function Get-VerificationReportLines {
    param(
        [Parameter(Mandatory)] [object]$Verification,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $mismatches = @($Verification.Results | Where-Object { -not $_.Matches })

    $lines.Add('WinDefState Restore Verification')
    $lines.Add(('Snapshot JSON: {0}' -f $SnapshotPath))
    $lines.Add(('ComputerName: {0}' -f $Verification.ComputerName))
    $lines.Add(('VerifiedAtUtc: {0}' -f $Verification.VerifiedAtUtc))
    $lines.Add(('Matched settings: {0}' -f $Verification.MatchedCount))
    $lines.Add(('Mismatched settings: {0}' -f $Verification.MismatchCount))

    if ($mismatches.Count -eq 0) {
        $lines.Add(' ')
        $lines.Add('All captured settings match the requested snapshot.')
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

function Get-DefenseDefinitions {
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

    @(
        [PSCustomObject]@{ Id = 'defender.disable_realtime_monitoring'; Type = 'MpPreferenceValue'; Property = 'DisableRealtimeMonitoring'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.enable_network_protection'; Type = 'MpPreferenceValue'; Property = 'EnableNetworkProtection'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.enable_controlled_folder_access'; Type = 'MpPreferenceValue'; Property = 'EnableControlledFolderAccess'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.asr_rules'; Type = 'AsrRules'; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'firewall.profiles'; Type = 'FirewallProfiles'; Profiles = @('Domain', 'Private', 'Public'); PermissiveValue = $false; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'powershell.lockdown'; Type = 'MachineEnvironmentValue'; Name = '__PSLockdownPolicy'; PermissiveExists = $true; PermissiveValue = '0'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'powershell.script_block_logging'; Type = 'RegistryKeyFlat'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'powershell.module_logging'; Type = 'PowerShellModuleLogging'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'powershell.transcription'; Type = 'RegistryKeyFlat'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\Transcription'; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'applocker.service'; Type = 'ServiceConfig'; Name = 'AppIDSvc'; PermissiveStartup = 'demand'; PermissiveState = 'Stopped'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'print.spooler_service'; Type = 'ServiceConfig'; Name = 'Spooler'; PermissiveStartup = 'auto'; PermissiveState = 'Running'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'localuser.administrator'; Type = 'LocalUser'; Rid = 500; Name = 'Built-in local administrator'; PermissiveValue = $true; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'uac.prompt_on_secure_desktop'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'PromptOnSecureDesktop'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'uac.enable_lua'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'EnableLUA'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'uac.consent_prompt_behavior_admin'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'ConsentPromptBehaviorAdmin'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }

        [PSCustomObject]@{ Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'; Name = 'UserAuthentication'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'wsh.enabled'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows Script Host\Settings'; Name = 'Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'smartscreen.enable'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Name = 'EnableSmartScreen'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'sehop.disable_exception_chain_validation'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\kernel'; Name = 'DisableExceptionChainValidation'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }

        [PSCustomObject]@{ Id = 'lsa.run_as_ppl'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'RunAsPPL'; ValueKind = 'DWord'; PermissiveExists = $false; PermissiveValue = $null; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'lsa.no_lm_hash'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'NoLMHash'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'lsa.lsa_cfg_flags'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'LsaCfgFlags'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'rdp.disable_restricted_admin'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name = 'DisableRestrictedAdmin'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }

        [PSCustomObject]@{ Id = 'deviceguard.enable_vbs'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard'; Name = 'EnableVirtualizationBasedSecurity'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'deviceguard.require_platform_security_features'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard'; Name = 'RequirePlatformSecurityFeatures'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'hvci.enabled'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity'; Name = 'Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'wdigest.use_logon_credential'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest'; Name = 'UseLogonCredential'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }

        [PSCustomObject]@{ Id = 'network.netbios_adapters'; Type = 'NetBiosAdapters'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'wpad.disable_wpad'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp'; Name = 'DisableWpad'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'wpad.user_auto_detect'; Type = 'LoadedUserRegistryValues'; Items = @([PSCustomObject]@{ RelativePath = 'Software\Microsoft\Windows\CurrentVersion\Internet Settings'; Name = 'AutoDetect'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1 }); RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'llmnr.enable_multicast'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient'; Name = 'EnableMulticast'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'mdns.enable'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters'; Name = 'EnableMDNS'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'telemetry.allow'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'; Name = 'AllowTelemetry'; ValueKind = 'DWord'; PermissiveExists = $false; PermissiveValue = $null; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'audit.process_cmdline'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit'; Name = 'ProcessCreationIncludeCmdLine_Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'audit.process_creation'; Type = 'AuditPolicy'; Subcategory = 'Process Creation'; PermissiveSuccess = $false; PermissiveFailure = $false; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'print.register_spooler_remote_rpc_endpoint'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers'; Name = 'RegisterSpoolerRemoteRpcEndPoint'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'winrm.service.allow_unencrypted'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\AllowUnencrypted'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.allow_unencrypted'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\AllowUnencrypted'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.basic'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Basic'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.basic'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Basic'; PermissiveValue = $true; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'smb.client.require_security_signature'; Type = 'SmbClientConfig'; PermissiveValue = $false; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'smb.server.require_security_signature'; Type = 'SmbServerConfig'; PermissiveValue = $false; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'office.block_macros_from_internet'; Type = 'LoadedUserRegistryValues'; Items = @($officeMacroItems); RequiresReboot = $false }
    )
}

function Capture-Definition {
    param([Parameter(Mandatory)] [object]$Definition)

    switch ($Definition.Type) {
        'RegistryValue' {
            $item = Get-ItemProperty -Path $Definition.Path -Name $Definition.Name -ErrorAction SilentlyContinue
            $exists = $null -ne $item -and $null -ne $item.$($Definition.Name)
            $valueKind = if ($exists) {
                try { (Get-Item -Path $Definition.Path).GetValueKind($Definition.Name).ToString() } catch { $Definition.ValueKind }
            } else {
                $Definition.ValueKind
            }

            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'RegistryValue'
                Path           = $Definition.Path
                Name           = $Definition.Name
                ValueKind      = $valueKind
                Exists         = $exists
                CurrentValue   = if ($exists) { $item.$($Definition.Name) } else { $null }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'RegistryKeyFlat' {
            return Capture-RegistryKeyFlatState -Id $Definition.Id -Path $Definition.Path -RequiresReboot $Definition.RequiresReboot
        }
        'MpPreferenceValue' {
            $mp = Get-MpPreference -ErrorAction SilentlyContinue
            $rawValue = if ($null -ne $mp) { $mp.($Definition.Property) } else { $null }

            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'MpPreferenceValue'
                Property       = $Definition.Property
                CurrentValue   = $rawValue
                RestoreValue   = Resolve-MpPreferenceValue -Definition $Definition -Value $rawValue
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'AsrRules' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'AsrRules'
                CurrentValue   = @(Get-ConfiguredAsrRules)
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'PowerShellModuleLogging' {
            return Capture-PowerShellModuleLoggingState
        }
        'FirewallProfiles' {
            $profiles = foreach ($profile in $Definition.Profiles) {
                $fw = Get-NetFirewallProfile -Profile $profile -ErrorAction SilentlyContinue
                [PSCustomObject]@{
                    Profile = $profile
                    Enabled = $fw.Enabled
                }
            }

            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'FirewallProfiles'
                CurrentValue   = @($profiles)
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'NetBiosAdapters' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'NetBiosAdapters'
                CurrentValue   = @(Get-NetBiosAdapterStates)
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'LoadedUserRegistryValues' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'LoadedUserRegistryValues'
                CurrentValue   = @(Get-LoadedUserRegistryValueStates -Items @($Definition.Items))
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
            $service = Get-CimInstance Win32_Service -Filter "Name='$($Definition.Name)'" -ErrorAction SilentlyContinue
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'ServiceConfig'
                Name           = $Definition.Name
                CurrentValue   = [PSCustomObject]@{
                    StartMode = if ($null -ne $service) { $service.StartMode } else { $null }
                    State     = if ($null -ne $service) { $service.State } else { $null }
                }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'LocalUser' {
            $user = Resolve-LocalUserTarget -Reference $Definition
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'LocalUser'
                Name           = if ($null -ne $user) { $user.Name } else { $Definition.Name }
                Sid            = if ($null -ne $user -and $null -ne $user.SID) { $user.SID.Value } else { $null }
                Rid            = if ($Definition.PSObject.Properties['Rid']) { $Definition.Rid } else { $null }
                CurrentValue   = if ($null -ne $user) { $user.Enabled } else { $null }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'WsManValue' {
            $item = Get-Item -Path $Definition.Path -ErrorAction SilentlyContinue
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'WsManValue'
                Path           = $Definition.Path
                CurrentValue   = if ($null -ne $item) { $item.Value } else { $null }
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'AuditPolicy' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'AuditPolicy'
                Subcategory    = $Definition.Subcategory
                CurrentValue   = Get-AuditPolicyState -Subcategory $Definition.Subcategory
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'SmbClientConfig' {
            $config = Get-SmbClientConfiguration
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'SmbClientConfig'
                CurrentValue   = $config.RequireSecuritySignature
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'SmbServerConfig' {
            $config = Get-SmbServerConfiguration
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'SmbServerConfig'
                CurrentValue   = $config.RequireSecuritySignature
                RequiresReboot = $Definition.RequiresReboot
            }
        }
    }
}

function Apply-PermissiveDefinition {
    param([Parameter(Mandatory)] [object]$Definition)

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
        'AsrRules' {
            Disable-ConfiguredAsrRules
        }
        'PowerShellModuleLogging' {
            Set-Permissive-PowerShellModuleLogging
        }
        'FirewallProfiles' {
            Set-NetFirewallProfile -Profile $Definition.Profiles -Enabled (ConvertTo-FirewallProfileEnabledValue -Value $Definition.PermissiveValue)
        }
        'NetBiosAdapters' {
            Set-Permissive-NetBiosAdapters
        }
        'LoadedUserRegistryValues' {
            Set-Permissive-LoadedUserRegistryValues -Items @($Definition.Items)
        }
        'MachineEnvironmentValue' {
            if ($Definition.PermissiveExists) {
                [Environment]::SetEnvironmentVariable($Definition.Name, $Definition.PermissiveValue, 'Machine')
            } else {
                [Environment]::SetEnvironmentVariable($Definition.Name, $null, 'Machine')
            }
        }
        'ServiceConfig' {
            & sc.exe config $Definition.Name "start= $($Definition.PermissiveStartup)" | Out-Null
            if ([string]$Definition.PermissiveState -eq 'Running') {
                Start-Service -Name $Definition.Name -ErrorAction SilentlyContinue
            } else {
                Stop-Service -Name $Definition.Name -Force -ErrorAction SilentlyContinue
            }
        }
        'LocalUser' {
            Set-LocalUserEnabledState -Reference $Definition -Enabled ([bool]$Definition.PermissiveValue)
        }
        'WsManValue' {
            Set-Item -Path $Definition.Path -Value $Definition.PermissiveValue
        }
        'AuditPolicy' {
            Set-AuditPolicyState -Subcategory $Definition.Subcategory -Success $Definition.PermissiveSuccess -Failure $Definition.PermissiveFailure
        }
        'SmbClientConfig' {
            Set-SmbClientConfiguration -RequireSecuritySignature $Definition.PermissiveValue -Confirm:$false | Out-Null
        }
        'SmbServerConfig' {
            Set-SmbServerConfiguration -RequireSecuritySignature $Definition.PermissiveValue -Confirm:$false | Out-Null
        }
    }
}

function Restore-SnapshotEntry {
    param([Parameter(Mandatory)] [object]$Entry)

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
        'AsrRules' {
            Restore-AsrRules -Rules @($Entry.CurrentValue)
        }
        'PowerShellModuleLogging' {
            Restore-PowerShellModuleLogging -Entry $Entry
        }
        'FirewallProfiles' {
            foreach ($profile in @($Entry.CurrentValue)) {
                Set-NetFirewallProfile -Profile $profile.Profile -Enabled (ConvertTo-FirewallProfileEnabledValue -Value $profile.Enabled)
            }
        }
        'NetBiosAdapters' {
            Restore-NetBiosAdapters -Adapters @($Entry.CurrentValue)
        }
        'LoadedUserRegistryValues' {
            Restore-LoadedUserRegistryValues -Entries @($Entry.CurrentValue)
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
            & sc.exe config $Entry.Name "start= $startMode" | Out-Null

            if ([string]$Entry.CurrentValue.State -eq 'Running') {
                Start-Service -Name $Entry.Name -ErrorAction SilentlyContinue
            } else {
                Stop-Service -Name $Entry.Name -Force -ErrorAction SilentlyContinue
            }
        }
        'LocalUser' {
            if ($null -eq $Entry.CurrentValue) {
                return
            }

            Set-LocalUserEnabledState -Reference $Entry -Enabled ([bool]$Entry.CurrentValue)
        }
        'WsManValue' {
            Set-Item -Path $Entry.Path -Value $Entry.CurrentValue
        }
        'AuditPolicy' {
            Set-AuditPolicyState -Subcategory $Entry.Subcategory -Success ([bool]$Entry.CurrentValue.Success) -Failure ([bool]$Entry.CurrentValue.Failure)
        }
        'SmbClientConfig' {
            Set-SmbClientConfiguration -RequireSecuritySignature ([bool]$Entry.CurrentValue) -Confirm:$false | Out-Null
        }
        'SmbServerConfig' {
            Set-SmbServerConfiguration -RequireSecuritySignature ([bool]$Entry.CurrentValue) -Confirm:$false | Out-Null
        }
    }
}

function Export-DefenseSnapshot {
    param([Parameter(Mandatory)] [string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    $snapshot = [PSCustomObject]@{
        SchemaVersion = 1
        Tool          = 'WinDefState'
        ComputerName  = $env:COMPUTERNAME
        CapturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Settings      = @(foreach ($definition in Get-DefenseDefinitions) { Capture-Definition -Definition $definition })
    }

    $reportPath = Get-SnapshotReportPath -SnapshotPath $fullPath
    $reportLines = Get-SnapshotReportLines -Snapshot $snapshot -SnapshotPath $fullPath

    Write-JsonAtomic -Path $fullPath -InputObject $snapshot
    Write-TextAtomic -Path $reportPath -Content ($reportLines -join [Environment]::NewLine)

    [PSCustomObject]@{
        JsonPath    = $fullPath
        ReportPath  = $reportPath
        Snapshot    = $snapshot
        ReportLines = @($reportLines)
    }
}

function Set-DefensePermissive {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Get-DefaultSnapshotPath -Root $StateRoot
    }

    $export = Export-DefenseSnapshot -Path $Path
    Write-OperationState -Root $StateRoot -SnapshotPath $export.JsonPath -Mode 'Permissive'

    foreach ($definition in Get-DefenseDefinitions) {
        Apply-PermissiveDefinition -Definition $definition
    }

    Write-Host "Permissive mode applied. Snapshot JSON saved to: $($export.JsonPath)"
    Write-Host "Snapshot report saved to: $($export.ReportPath)"
}

function Restore-DefenseSnapshot {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $operation = Get-OperationState -Root $StateRoot
        if ($null -eq $operation) {
            throw 'No snapshot path was provided and no current operation file exists.'
        }

        $Path = [string]$operation.SnapshotPath
    }

    $fullPath = [IO.Path]::GetFullPath($Path)
    $snapshot = Read-JsonFile -Path $fullPath
    foreach ($entry in @($snapshot.Settings)) {
        Restore-SnapshotEntry -Entry $entry
    }

    $verification = Test-DefenseSnapshot -Snapshot $snapshot
    $verificationPath = Get-VerificationReportPath -Root $StateRoot -SnapshotPath $fullPath
    $verificationLines = Get-VerificationReportLines -Verification $verification -SnapshotPath $fullPath
    Write-TextAtomic -Path $verificationPath -Content ($verificationLines -join [Environment]::NewLine)

    if ($verification.MismatchCount -gt 0) {
        $mismatchIds = @($verification.Results | Where-Object { -not $_.Matches } | ForEach-Object { $_.Id })
        Write-Warning "Restore verification found $($verification.MismatchCount) mismatched setting(s)."
        Write-Warning "Verification report saved to: $verificationPath"
        Write-Warning ("Mismatched IDs: {0}" -f ($mismatchIds -join ', '))
        throw 'Restore verification failed. current-operation.json was left in place so you can retry the same snapshot.'
    }

    Clear-OperationState -Root $StateRoot
    Write-Host "Restore completed and verified from snapshot: $fullPath"
    Write-Host "Verification report saved to: $verificationPath"
}

Assert-Administrator
Ensure-Directory -Path $StateRoot

switch ($Command) {
    'Snapshot' {
        if ([string]::IsNullOrWhiteSpace($SnapshotPath)) {
            $SnapshotPath = Get-DefaultSnapshotPath -Root $StateRoot
        }

        $export = Export-DefenseSnapshot -Path $SnapshotPath
        Write-Host "Snapshot JSON saved to: $($export.JsonPath)"
        Write-Host "Snapshot report saved to: $($export.ReportPath)"
        Write-Host ''
        Show-ReportLines -Lines $export.ReportLines
    }
    'Permissive' {
        Set-DefensePermissive -Path $SnapshotPath
    }
    'Restore' {
        Restore-DefenseSnapshot -Path $SnapshotPath
    }
}

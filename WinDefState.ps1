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

function Write-JsonAtomic {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [object]$InputObject
    )

    $parent = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Ensure-Directory -Path $parent
    }

    $tempPath = "$Path.tmp"
    $json = $InputObject | ConvertTo-Json -Depth 12
    [System.IO.File]::WriteAllText($tempPath, $json, [System.Text.UTF8Encoding]::new($false))

    if (Test-Path -LiteralPath $Path) {
        [System.IO.File]::Replace($tempPath, $Path, $null)
    } else {
        [System.IO.File]::Move($tempPath, $Path)
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

function Get-DefenseDefinitions {
    @(
        [PSCustomObject]@{ Id = 'defender.disable_realtime_monitoring'; Type = 'MpPreferenceValue'; Property = 'DisableRealtimeMonitoring'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.enable_network_protection'; Type = 'MpPreferenceValue'; Property = 'EnableNetworkProtection'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.enable_controlled_folder_access'; Type = 'MpPreferenceValue'; Property = 'EnableControlledFolderAccess'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.asr_rules'; Type = 'AsrRules'; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'firewall.profiles'; Type = 'FirewallProfiles'; Profiles = @('Domain', 'Private', 'Public'); PermissiveValue = $false; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'powershell.lockdown'; Type = 'MachineEnvironmentValue'; Name = '__PSLockdownPolicy'; PermissiveExists = $true; PermissiveValue = '0'; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'applocker.service'; Type = 'ServiceConfig'; Name = 'AppIDSvc'; PermissiveStartup = 'demand'; PermissiveState = 'Stopped'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'localuser.administrator'; Type = 'LocalUser'; Name = 'Administrator'; PermissiveValue = $true; RequiresReboot = $false }

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

        [PSCustomObject]@{ Id = 'deviceguard.enable_vbs'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard'; Name = 'EnableVirtualizationBasedSecurity'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'deviceguard.require_platform_security_features'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard'; Name = 'RequirePlatformSecurityFeatures'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'hvci.enabled'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity'; Name = 'Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'wdigest.use_logon_credential'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest'; Name = 'UseLogonCredential'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }

        [PSCustomObject]@{ Id = 'llmnr.enable_multicast'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient'; Name = 'EnableMulticast'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'mdns.enable'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Services\Dnscache\Parameters'; Name = 'EnableMDNS'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'telemetry.allow'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection'; Name = 'AllowTelemetry'; ValueKind = 'DWord'; PermissiveExists = $false; PermissiveValue = $null; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'audit.process_cmdline'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\Audit'; Name = 'ProcessCreationIncludeCmdLine_Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'audit.process_creation'; Type = 'AuditPolicy'; Subcategory = 'Process Creation'; PermissiveSuccess = $false; PermissiveFailure = $false; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'winrm.service.allow_unencrypted'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\AllowUnencrypted'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.allow_unencrypted'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\AllowUnencrypted'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.basic'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Basic'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.basic'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Basic'; PermissiveValue = $true; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'smb.client.require_security_signature'; Type = 'SmbClientConfig'; PermissiveValue = $false; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'smb.server.require_security_signature'; Type = 'SmbServerConfig'; PermissiveValue = $false; RequiresReboot = $false }
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
            $user = Get-LocalUser -Name $Definition.Name -ErrorAction SilentlyContinue
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'LocalUser'
                Name           = $Definition.Name
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
        'MpPreferenceValue' {
            Set-MpPreferencePropertyValue -Property $Definition.Property -Value $Definition.PermissiveValue
        }
        'AsrRules' {
            Disable-ConfiguredAsrRules
        }
        'FirewallProfiles' {
            Set-NetFirewallProfile -Profile $Definition.Profiles -Enabled $Definition.PermissiveValue
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
            Stop-Service -Name $Definition.Name -Force -ErrorAction SilentlyContinue
        }
        'LocalUser' {
            if ($Definition.PermissiveValue) {
                Enable-LocalUser -Name $Definition.Name -ErrorAction SilentlyContinue
            } else {
                Disable-LocalUser -Name $Definition.Name -ErrorAction SilentlyContinue
            }
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
        'MpPreferenceValue' {
            Set-MpPreferencePropertyValue -Property $Entry.Property -Value $Entry.RestoreValue
        }
        'AsrRules' {
            Restore-AsrRules -Rules @($Entry.CurrentValue)
        }
        'FirewallProfiles' {
            foreach ($profile in @($Entry.CurrentValue)) {
                Set-NetFirewallProfile -Profile $profile.Profile -Enabled ([bool]$profile.Enabled)
            }
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

            if ([bool]$Entry.CurrentValue) {
                Enable-LocalUser -Name $Entry.Name -ErrorAction SilentlyContinue
            } else {
                Disable-LocalUser -Name $Entry.Name -ErrorAction SilentlyContinue
            }
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

    $snapshot = [PSCustomObject]@{
        SchemaVersion = 1
        Tool          = 'WinDefState'
        ComputerName  = $env:COMPUTERNAME
        CapturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Settings      = @(foreach ($definition in Get-DefenseDefinitions) { Capture-Definition -Definition $definition })
    }

    Write-JsonAtomic -Path $Path -InputObject $snapshot
    [IO.Path]::GetFullPath($Path)
}

function Set-DefensePermissive {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Get-DefaultSnapshotPath -Root $StateRoot
    }

    $fullPath = Export-DefenseSnapshot -Path $Path
    Write-OperationState -Root $StateRoot -SnapshotPath $fullPath -Mode 'Permissive'

    foreach ($definition in Get-DefenseDefinitions) {
        Apply-PermissiveDefinition -Definition $definition
    }

    Write-Host "Permissive mode applied. Snapshot saved to: $fullPath"
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

    $snapshot = Read-JsonFile -Path $Path
    foreach ($entry in @($snapshot.Settings)) {
        Restore-SnapshotEntry -Entry $entry
    }

    Clear-OperationState -Root $StateRoot
    Write-Host "Restore completed from snapshot: $([IO.Path]::GetFullPath($Path))"
}

Assert-Administrator
Ensure-Directory -Path $StateRoot

switch ($Command) {
    'Snapshot' {
        if ([string]::IsNullOrWhiteSpace($SnapshotPath)) {
            $SnapshotPath = Get-DefaultSnapshotPath -Root $StateRoot
        }

        $fullPath = Export-DefenseSnapshot -Path $SnapshotPath
        Write-Host "Snapshot saved to: $fullPath"
    }
    'Permissive' {
        Set-DefensePermissive -Path $SnapshotPath
    }
    'Restore' {
        Restore-DefenseSnapshot -Path $SnapshotPath
    }
}

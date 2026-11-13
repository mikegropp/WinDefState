BeforeAll {
    . (Join-Path (Split-Path -Parent $PSScriptRoot) 'WinDefState.ps1')

    $script:TestCreatedGetWinEventStub = $false
    if ($null -eq (Get-Command Get-WinEvent -ErrorAction SilentlyContinue)) {
        $script:TestCreatedGetWinEventStub = $true
        Set-Item -Path Function:\global:Get-WinEvent -Value {
            [CmdletBinding()]
            param([string]$ListLog, [hashtable]$FilterHashtable, [int]$MaxEvents)
            throw 'The test-only Get-WinEvent stub must be mocked before use.'
        }
    }

    function New-TestLapsRoot {
        param(
            [string]$Source,
            [object[]]$Settings = @(),
            [bool]$Captured = $true,
            [bool]$Exists = $false
        )
        [PSCustomObject]@{
            Source                 = $Source
            Path                   = "HKLM:\Test\$Source"
            Exists                 = $Exists
            Captured               = $Captured
            KnownSettingCount      = @($Settings).Count
            UnrecognizedValueCount = 0
            Settings               = @($Settings)
            ErrorType              = if ($Captured) { $null } else { 'UnauthorizedAccessException' }
        }
    }

    function New-TestLapsSetting {
        param([string]$Name, [AllowNull()] [object]$Value)
        [PSCustomObject]@{ Name = $Name; ValueKind = 'DWord'; Value = $Value }
    }

    function New-TestLapsEvents {
        [PSCustomObject]@{ CommandAvailable = $true; Captured = $true; Events = @(); ErrorType = $null }
    }
}

AfterAll {
    if ($script:TestCreatedGetWinEventStub) {
        Remove-Item -Path Function:\global:Get-WinEvent -ErrorAction SilentlyContinue
    }
}

Describe 'Windows LAPS policy inventory' {
    It 'skips an empty LocalConfig key and reports no active policy' {
        $roots = @(
            New-TestLapsRoot -Source CSP
            New-TestLapsRoot -Source GroupPolicy
            New-TestLapsRoot -Source LocalConfig -Exists $true
            New-TestLapsRoot -Source Legacy
        )

        $state = Resolve-LapsPolicyStatus -Roots $roots -EventStatus (New-TestLapsEvents)

        $state.Captured | Should -BeTrue
        $state.ActivePolicySource | Should -Be 'None'
        $state.BackupDirectory | Should -Be 0
        $state.BackupDirectoryLabel | Should -Be 'Disabled (no policy)'
        (Test-LapsPolicyStatusCapturedExactly -State $state) | Should -BeTrue
    }

    It 'uses the first configured root and never merges lower-priority values' {
        $roots = @(
            New-TestLapsRoot -Source CSP -Exists $true -Settings @(New-TestLapsSetting -Name PasswordLength -Value 20)
            New-TestLapsRoot -Source GroupPolicy -Exists $true -Settings @(New-TestLapsSetting -Name BackupDirectory -Value 2)
            New-TestLapsRoot -Source LocalConfig
            New-TestLapsRoot -Source Legacy
        )

        $state = Resolve-LapsPolicyStatus -Roots $roots -EventStatus (New-TestLapsEvents)

        $state.ActivePolicySource | Should -Be 'CSP'
        $state.BackupDirectory | Should -Be 0
        $state.BackupDirectoryLabel | Should -Be 'Disabled'
    }

    It 'selects Group Policy after an empty CSP root' {
        $roots = @(
            New-TestLapsRoot -Source CSP -Exists $true
            New-TestLapsRoot -Source GroupPolicy -Exists $true -Settings @(New-TestLapsSetting -Name BackupDirectory -Value 2)
            New-TestLapsRoot -Source LocalConfig
            New-TestLapsRoot -Source Legacy
        )

        $state = Resolve-LapsPolicyStatus -Roots $roots -EventStatus (New-TestLapsEvents)

        $state.ActivePolicySource | Should -Be 'GroupPolicy'
        $state.BackupDirectoryLabel | Should -Be 'Active Directory'
    }

    It 'fails closed when a higher-priority root cannot be read' {
        $roots = @(
            New-TestLapsRoot -Source CSP -Captured $false
            New-TestLapsRoot -Source GroupPolicy -Exists $true -Settings @(New-TestLapsSetting -Name BackupDirectory -Value 2)
            New-TestLapsRoot -Source LocalConfig
            New-TestLapsRoot -Source Legacy
        )

        $state = Resolve-LapsPolicyStatus -Roots $roots -EventStatus (New-TestLapsEvents)

        $state.Captured | Should -BeFalse
        $state.ActivePolicySource | Should -Be 'Unknown'
        $state.BackupDirectory | Should -BeNullOrEmpty
        (Test-LapsPolicyStatusCapturedExactly -State $state) | Should -BeFalse
    }

    It 'keeps legacy policy separate from the modern backup-directory enum' {
        $roots = @(
            New-TestLapsRoot -Source CSP
            New-TestLapsRoot -Source GroupPolicy
            New-TestLapsRoot -Source LocalConfig
            New-TestLapsRoot -Source Legacy -Exists $true -Settings @(New-TestLapsSetting -Name AdmPwdEnabled -Value 1)
        )

        $state = Resolve-LapsPolicyStatus -Roots $roots -EventStatus (New-TestLapsEvents)

        $state.ActivePolicySource | Should -Be 'Legacy'
        $state.BackupDirectory | Should -BeNullOrEmpty
        $state.BackupDirectoryLabel | Should -Be 'Legacy Active Directory policy'
    }

    It 'does not label an invalid backup-directory value as enabled' {
        $roots = @(
            New-TestLapsRoot -Source CSP -Exists $true -Settings @(New-TestLapsSetting -Name BackupDirectory -Value 9)
            New-TestLapsRoot -Source GroupPolicy
            New-TestLapsRoot -Source LocalConfig
            New-TestLapsRoot -Source Legacy
        )

        $state = Resolve-LapsPolicyStatus -Roots $roots -EventStatus (New-TestLapsEvents)

        $state.ActivePolicySource | Should -Be 'CSP'
        $state.BackupDirectoryLabel | Should -Be 'Unknown'
    }

    It 'does not read or serialize unknown registry value data' {
        $script:TestLapsReadNames = [System.Collections.Generic.List[string]]::new()
        $script:TestLapsKey = [PSCustomObject]@{}
        $script:TestLapsKey | Add-Member -MemberType ScriptMethod -Name GetValueNames -Value { @('BackupDirectory', 'PasswordBlob') }
        $script:TestLapsKey | Add-Member -MemberType ScriptMethod -Name GetValueKind -Value {
            param($Name)
            $script:TestLapsReadNames.Add("kind:$Name") | Out-Null
            if ($Name -ne 'BackupDirectory') { throw 'Unknown value kind was read.' }
            [Microsoft.Win32.RegistryValueKind]::DWord
        }
        $script:TestLapsKey | Add-Member -MemberType ScriptMethod -Name GetValue -Value {
            param($Name)
            $script:TestLapsReadNames.Add("value:$Name") | Out-Null
            if ($Name -ne 'BackupDirectory') { throw 'Unknown value data was read.' }
            1
        }
        Mock Get-Item { $script:TestLapsKey } -ParameterFilter { $LiteralPath -like 'HKLM:\*' }

        $root = Get-LapsPolicyRootState -Definition @(Get-LapsPolicyRootDefinitions)[0]

        $root.Captured | Should -BeTrue
        $root.KnownSettingCount | Should -Be 1
        $root.UnrecognizedValueCount | Should -Be 1
        [string]::Join(',', $script:TestLapsReadNames.ToArray()) | Should -Be 'kind:BackupDirectory,value:BackupDirectory'
        ($root | ConvertTo-Json -Depth 6) | Should -Not -Match 'PasswordBlob'
    }

    It 'persists only event ID, record ID, and timestamp metadata' {
        Mock Test-CommandAvailable { $true } -ParameterFilter { $Name -eq 'Get-WinEvent' }
        Mock Get-WinEvent { [PSCustomObject]@{ LogName = 'Microsoft-Windows-LAPS/Operational' } } -ParameterFilter { [bool]$ListLog }
        Mock Get-WinEvent {
            [PSCustomObject]@{
                Id = 10020
                RecordId = 42
                TimeCreated = [datetime]'2026-11-12T13:00:00Z'
                Message = 'Synthetic secret that must not be saved'
            }
        } -ParameterFilter { $null -ne $FilterHashtable }

        $status = Get-LapsPolicyEventStatus
        $json = $status | ConvertTo-Json -Depth 6

        $status.Captured | Should -BeTrue
        @($status.Events).Count | Should -Be 1
        $status.Events[0].Id | Should -Be 10020
        $json | Should -Not -Match 'Synthetic secret'
        $json | Should -Not -Match 'Message'
    }

    It 'is inventory only across permissive and restore capabilities' {
        $definition = (Get-DefenseDefinitionMap)['laps.policy_status']
        $capabilities = Get-DefinitionCapabilities -Definition $definition

        $definition.RequiresReboot | Should -BeFalse
        $capabilities.Permissive | Should -BeFalse
        $capabilities.Restore | Should -BeFalse
        $capabilities.InventoryOnly | Should -BeTrue
    }
}

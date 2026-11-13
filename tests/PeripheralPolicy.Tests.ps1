BeforeAll {
    . (Join-Path (Split-Path -Parent $PSScriptRoot) 'WinDefState.Environment.ps1')

    function New-TestPolicyKey {
        param([string]$Name, [string]$PSPath, [string]$PSChildName,
            [string]$ValueName = '', [object]$Value = $null)
        $key = [pscustomobject]@{
            Name = $Name; PSPath = $PSPath; PSChildName = $PSChildName
            TestValueName = $ValueName; TestValue = $Value
        }
        $key | Add-Member -MemberType ScriptMethod -Name GetValueNames -Value {
            if ($this.TestValueName) { @($this.TestValueName) } else { @() }
        }
        $key | Add-Member -MemberType ScriptMethod -Name GetValueKind -Value {
            param($Name)
            [Microsoft.Win32.RegistryValueKind]::DWord
        }
        $key | Add-Member -MemberType ScriptMethod -Name GetValue -Value {
            param($Name, $Default, $Options)
            $this.TestValue
        }
        $key
    }
}

Describe 'Peripheral policy environment inventory' {
    It 'distinguishes absent machine and current-user roots from captured empty policy' {
        Mock Test-Path { $false }

        $machine = Get-EnvironmentPeripheralPolicy -Scope Machine -PolicyFamily RemovableStorage
        $user = Get-EnvironmentPeripheralPolicy -Scope CurrentUser -PolicyFamily RemovableStorage -UserSid 'S-1-5-21-test'
        $device = Get-EnvironmentPeripheralPolicy -Scope Machine -PolicyFamily DeviceInstallation

        foreach ($result in @($machine, $user, $device)) {
            $result.ConfigurationState | Should -Be Absent
            $result.RootPresent | Should -BeFalse
            @($result.Values).Count | Should -Be 0
        }
        $user.UserSid | Should -Be 'S-1-5-21-test'
        $device.RootPath | Should -Match 'DeviceInstall\\Restrictions$'
    }

    It 'records an empty key without calling it an enforced policy' {
        $script:PolicyTestKey = New-TestPolicyKey -Name 'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DeviceInstall\Restrictions' -PSPath 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceInstall\Restrictions' -PSChildName Restrictions
        Mock Test-Path { $true }
        Mock Get-Item { $script:PolicyTestKey }
        Mock Get-ChildItem { @() }

        $result = Get-EnvironmentPeripheralPolicy -Scope Machine -PolicyFamily DeviceInstallation

        $result.RootPresent | Should -BeTrue
        $result.ConfigurationState | Should -Be KeysPresentWithoutValues
        @($result.Values).Count | Should -Be 0
    }

    It 'reports explicit enabled and disabled DWORDs without inferring USB behavior' {
        $script:PolicyTestKey = New-TestPolicyKey -Name 'HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\DeviceInstall\Restrictions' -PSPath 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceInstall\Restrictions' -PSChildName Restrictions -ValueName DenyRemovableDevices -Value 1
        Mock Test-Path { $true }
        Mock Get-Item { $script:PolicyTestKey }
        Mock Get-ChildItem { @() }

        $enabled = Get-EnvironmentPeripheralPolicy -Scope Machine -PolicyFamily DeviceInstallation
        $enabled.ConfigurationState | Should -Be ValuesPresent
        $enabled.Values[0].FlagState | Should -Be EnabledInRegistry
        $enabled.Values[0].Value | Should -Be 1

        $script:PolicyTestKey.TestValue = 0
        $disabled = Get-EnvironmentPeripheralPolicy -Scope Machine -PolicyFamily DeviceInstallation
        $disabled.Values[0].FlagState | Should -Be DisabledInRegistry
        $disabled.Values[0].Value | Should -Be 0
    }

    It 'finds a class-specific removable-storage key even when the family root is absent' {
        $base = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows'
        $classKey = 'RemovableStorageDevices{53f5630d-b6bf-11d0-94f2-00a0c91efb8b}'
        $script:PolicyTestKey = New-TestPolicyKey -Name "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\$classKey" -PSPath "$base\$classKey" -PSChildName $classKey -ValueName Deny_Write -Value 1
        Mock Test-Path { $LiteralPath -eq $base }
        Mock Get-ChildItem {
            if ($LiteralPath -eq $base) { $script:PolicyTestKey } else { @() }
        }
        Mock Get-Item { $script:PolicyTestKey }

        $result = Get-EnvironmentPeripheralPolicy -Scope Machine -PolicyFamily RemovableStorage

        $result.RootPresent | Should -BeFalse
        $result.ConfigurationState | Should -Be ValuesPresent
        $result.Values[0].ValueName | Should -Be Deny_Write
        $result.Values[0].FlagState | Should -Be EnabledInRegistry
    }

    It 'makes the section unknown on registry read failure instead of calling it absent' {
        Mock Test-Path { throw 'Registry access denied' }

        $section = New-EnvironmentSection 'policy.device_installation' 'Device installation restriction policy registry' @('Scope') {
            Get-EnvironmentPeripheralPolicy -Scope Machine -PolicyFamily DeviceInstallation
        }

        $section.Status | Should -Be Unknown
        $section.Error | Should -Match 'Registry access denied'
    }
}

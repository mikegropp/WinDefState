BeforeAll {
    . (Join-Path (Split-Path -Parent $PSScriptRoot) 'WinDefState.ps1')
}

Describe 'Windows Time inventory' {
    It 'parses the unsynchronized status reported by the test image' {
        $text = @'
Leap Indicator: 3(not synchronized)
Stratum: 0 (unspecified)
Last Successful Sync Time: unspecified
Source: Local CMOS Clock
'@

        $status = ConvertFrom-W32tmStatusText -Text $text

        $status.Parsed | Should -BeTrue
        $status.LeapIndicator | Should -Be 3
        $status.Stratum | Should -Be 0
        $status.Source | Should -Be 'Local CMOS Clock'
        $status.LastSuccessfulSyncTime | Should -Be 'unspecified'
        $status.UnsynchronizedFlag | Should -BeTrue
        $status.PSObject.Properties['ClockOffset'] | Should -BeNullOrEmpty
    }

    It 'does not treat a prior sync timestamp as current synchronization' {
        $text = @'
Leap Indicator: 3(not synchronized)
Stratum: 0 (unspecified)
Last Successful Sync Time: 10/9/2025 11:28:21 PM
Source: Free-running System Clock
'@

        $status = ConvertFrom-W32tmStatusText -Text $text

        $status.Parsed | Should -BeTrue
        $status.UnsynchronizedFlag | Should -BeTrue
        $status.LastSuccessfulSyncTime | Should -Be '10/9/2025 11:28:21 PM'
    }

    It 'marks unrecognized localized labels partial rather than inventing values' {
        $status = ConvertFrom-W32tmStatusText -Text "Indicateur de saut : 3`nStrate : 0"

        $status.Parsed | Should -BeFalse
        $status.LeapIndicator | Should -BeNullOrEmpty
        $status.UnsynchronizedFlag | Should -BeNullOrEmpty
    }

    It 'accepts a framed bounded w32tm status result without invoking synchronization' {
        Mock Test-CommandAvailable { $true } -ParameterFilter { $Name -eq 'w32tm.exe' }
        Mock Invoke-ChildPowerShell {
            [PSCustomObject]@{
                CommandAvailable = $true
                TimedOut = $false
                ExitCode = 0
                StdOut = 'WDS_TIME_JSON:{"ExitCode":0,"Text":"Leap Indicator: 3(not synchronized)\nStratum: 0 (unspecified)\nLast Successful Sync Time: unspecified\nSource: Local CMOS Clock"}'
                StdErr = ''
            }
        }

        $query = Get-W32tmStatusQuery

        $query.Captured | Should -BeTrue
        $query.ExitCode | Should -Be 0
        $query.RawText | Should -Match 'Local CMOS Clock'
        Assert-MockCalled Invoke-ChildPowerShell -Times 1 -ParameterFilter { $TimeoutSeconds -eq 12 -and $ScriptText -match '/query /status /verbose' -and $ScriptText -notmatch '/resync' }
    }

    It 'marks a timed-out status query incomplete' {
        Mock Test-CommandAvailable { $true } -ParameterFilter { $Name -eq 'w32tm.exe' }
        Mock Invoke-ChildPowerShell {
            [PSCustomObject]@{ CommandAvailable = $true; TimedOut = $true; ExitCode = $null; StdOut = ''; StdErr = '' }
        }

        $query = Get-W32tmStatusQuery

        $query.Captured | Should -BeFalse
        $query.TimedOut | Should -BeTrue
        $query.ErrorType | Should -Be 'TimedOut'
    }

    It 'omits unrecognized registry value data from peer configuration' {
        $script:TimeTestReads = [System.Collections.Generic.List[string]]::new()
        $script:TimeTestKey = [PSCustomObject]@{}
        $script:TimeTestKey | Add-Member -MemberType ScriptMethod -Name GetValueNames -Value { @('Type', 'NtpServer', 'SecretValue') }
        $script:TimeTestKey | Add-Member -MemberType ScriptMethod -Name GetValueKind -Value {
            param($Name)
            $script:TimeTestReads.Add("kind:$Name") | Out-Null
            if ($Name -eq 'SecretValue') { throw 'Unknown value kind was read.' }
            [Microsoft.Win32.RegistryValueKind]::String
        }
        $script:TimeTestKey | Add-Member -MemberType ScriptMethod -Name GetValue -Value {
            param($Name, $Default, $Options)
            $script:TimeTestReads.Add("value:$Name") | Out-Null
            if ($Name -eq 'SecretValue') { throw 'Unknown value data was read.' }
            if ($Name -eq 'Type') { return 'NTP' }
            'time.windows.com,0x9'
        }
        Mock Get-Item { $script:TimeTestKey } -ParameterFilter { $LiteralPath -like 'HKLM:\*' }

        $root = Get-TimeSyncPeerRootState -Source Local -Path 'HKLM:\Test\W32Time'

        $root.Captured | Should -BeTrue
        $root.Type | Should -Be 'NTP'
        $root.NtpServer | Should -Be 'time.windows.com,0x9'
        @($root.PeerTokens).Count | Should -Be 1
        ($root | ConvertTo-Json -Depth 5) | Should -Not -Match 'SecretValue'
        [string]::Join(',', $script:TimeTestReads.ToArray()) | Should -Be 'kind:Type,value:Type,kind:NtpServer,value:NtpServer'
    }

    It 'requires service, parsed live status, and both configuration roots for a complete inventory' {
        Mock Get-TimeSyncServiceState {
            [PSCustomObject]@{ Captured = $true; Exists = $true; Status = 'Running'; StartMode = 'Auto' }
        }
        Mock Get-W32tmStatusQuery {
            [PSCustomObject]@{
                Captured = $true
                RawText = "Leap Indicator: 0(no warning)`nStratum: 3`nLast Successful Sync Time: 11/12/2026 2:00:00 PM`nSource: time.example.test"
            }
        }
        Mock Get-TimeSyncPeerRootState {
            [PSCustomObject]@{ Source = $Source; Captured = $true; Exists = $false; Type = $null; NtpServer = $null; PeerTokens = @() }
        }

        $state = Get-TimeSyncStatus

        $state.Captured | Should -BeTrue
        (Test-TimeSyncStatusCapturedExactly -State $state) | Should -BeTrue
        $state.Status.Source | Should -Be 'time.example.test'
        @($state.PeerRoots).Count | Should -Be 2
    }

    It 'marks a localized live status incomplete while preserving raw evidence' {
        Mock Get-TimeSyncServiceState {
            [PSCustomObject]@{ Captured = $true; Exists = $true; Status = 'Running'; StartMode = 'Auto' }
        }
        Mock Get-W32tmStatusQuery {
            [PSCustomObject]@{ Captured = $true; RawText = 'Source de temps : horloge locale' }
        }
        Mock Get-TimeSyncPeerRootState {
            [PSCustomObject]@{ Source = $Source; Captured = $true; Exists = $false; Type = $null; NtpServer = $null; PeerTokens = @() }
        }

        $state = Get-TimeSyncStatus

        $state.Captured | Should -BeFalse
        (Test-TimeSyncStatusCapturedExactly -State $state) | Should -BeFalse
        $state.Query.RawText | Should -Be 'Source de temps : horloge locale'
    }

    It 'keeps time status inventory only' {
        $definition = (Get-DefenseDefinitionMap)['time.sync_status']
        $capabilities = Get-DefinitionCapabilities -Definition $definition

        $definition.RequiresReboot | Should -BeFalse
        $capabilities.Permissive | Should -BeFalse
        $capabilities.Restore | Should -BeFalse
        $capabilities.InventoryOnly | Should -BeTrue
    }
}

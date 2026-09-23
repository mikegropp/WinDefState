BeforeAll {
    $repositoryRoot = Split-Path -Parent $PSScriptRoot
    . (Join-Path $repositoryRoot 'WinDefState.Environment.ps1')
    function New-TestProbe {
        param($Value, [bool]$Success = $true)
        [pscustomobject]@{ Success = $Success; Value = $Value; Error = $(if ($Success) { '' } else { 'Test provider unavailable' }); DurationMs = 1 }
    }
    function New-TestHealth {
        $probes = [ordered]@{
            OS = New-TestProbe ([pscustomobject]@{ BuildNumber = '26200'; ProductType = 1; OSArchitecture = '64-bit' })
            Version = New-TestProbe ([pscustomobject]@{ EditionID = 'Professional'; UBR = 100; ProductName = 'Windows 10 Pro' })
            Defender = New-TestProbe ([pscustomobject]@{ AMRunningMode = 'Normal'; RealTimeProtectionEnabled = $true; BehaviorMonitorEnabled = $true; IsTamperProtected = $true; AntivirusSignatureLastUpdated = [datetime]'2026-09-22T10:00:00Z' })
            SecureBoot = New-TestProbe ([pscustomobject]@{ Enabled = $true })
            Tpm = New-TestProbe ([pscustomobject]@{ TpmReady = $true })
            DeviceGuard = New-TestProbe ([pscustomobject]@{ SecurityServicesRunning = @(1, 2); SecurityServicesConfigured = @(1, 2); VirtualizationBasedSecurityStatus = 2 })
            Firewall = New-TestProbe @([pscustomobject]@{ Name = 'Domain'; Enabled = 'True' }, [pscustomobject]@{ Name = 'Private'; Enabled = 'True' }, [pscustomobject]@{ Name = 'Public'; Enabled = 'True' })
            Reboot = New-TestProbe ([pscustomobject]@{ Reasons = @() })
        }
        ConvertTo-WindowsHealthReport $probes ([datetime]'2026-09-23T10:00:00Z') 'DEMO-PC'
    }
    function New-TestBaseline {
        param([object[]]$Sections)
        [pscustomobject]@{
            SchemaVersion = 1; ReportType = 'WinDefState.Environment'; ComputerName = 'DEMO-PC'
            StartedAtUtc = '2026-09-23T10:00:00Z'; CapturedAtUtc = '2026-09-23T10:00:10Z'
            Elevated = $false; Health = New-TestHealth; ProcessInventoryError = ''; Sections = $Sections
            Limitations = @('Synthetic fixture. No live device data.')
        }
    }
    function New-TestSection {
        param([string]$Value = 'Running')
        New-EnvironmentSection 'system.services' Services @('Name') { [pscustomobject]@{ Name = 'ExampleService'; State = $Value } }
    }
}

Describe 'Windows client identity and lifecycle' {
    It 'uses the build instead of the stale Windows 10 product name' {
        $report = New-TestHealth
        $report.Windows.Family | Should -Be 'Windows 11'
        $report.Windows.Release | Should -Be '25H2'
    }
    It 'does not mistake Server 2025 for Windows 11' {
        $identity = Get-WindowsHealthIdentity ([pscustomobject]@{ BuildNumber = '26100'; ProductType = 3 }) ([pscustomobject]@{ EditionID = 'ServerStandard' })
        $identity.Family | Should -Be 'Windows Server'
        $identity.LifecycleStatus | Should -Be Unknown
    }
    It 'keeps unmapped future builds and unknown editions unknown' {
        foreach ($case in @(@('29000', 'Professional'), @('26200', 'UnmappedEdition'))) {
            $identity = Get-WindowsHealthIdentity ([pscustomobject]@{ BuildNumber = $case[0]; ProductType = 1 }) ([pscustomobject]@{ EditionID = $case[1] })
            $identity.LifecycleStatus | Should -Be Unknown
        }
    }
    It 'distinguishes enterprise servicing from consumer servicing' {
        $os = [pscustomobject]@{ BuildNumber = '22631'; ProductType = 1 }
        (Get-WindowsHealthIdentity $os ([pscustomobject]@{ EditionID = 'Professional' }) ([datetime]'2026-09-23')).LifecycleStatus | Should -Be Attention
        (Get-WindowsHealthIdentity $os ([pscustomobject]@{ EditionID = 'Enterprise' }) ([datetime]'2026-09-23')).LifecycleStatus | Should -Be Healthy
    }
    It 'does not infer ESU enrollment for Windows 10 22H2' {
        $identity = Get-WindowsHealthIdentity ([pscustomobject]@{ BuildNumber = '19045'; ProductType = 1 }) ([pscustomobject]@{ EditionID = 'Professional' }) ([datetime]'2026-09-23')
        $identity.LifecycleStatus | Should -Be Attention
        $identity.LifecycleDetail | Should -Match 'ESU enrollment is not checked'
    }
    It 'uses separate Enterprise and IoT LTSC lifecycle dates' {
        $os = [pscustomobject]@{ BuildNumber = '19044'; ProductType = 1 }
        (Get-WindowsHealthIdentity $os ([pscustomobject]@{ EditionID = 'EnterpriseS' })).ServicingEnd | Should -Be '2027-01-12'
        (Get-WindowsHealthIdentity $os ([pscustomobject]@{ EditionID = 'IoTEnterpriseS' })).ServicingEnd | Should -Be '2032-01-13'
    }
    It 'treats the published end date as the final supported day' {
        $os = [pscustomobject]@{ BuildNumber = '26100'; ProductType = 1 }
        $version = [pscustomobject]@{ EditionID = 'Professional' }
        (Get-WindowsHealthIdentity $os $version ([datetime]'2026-10-13T23:00:00')).LifecycleStatus | Should -Be Healthy
        (Get-WindowsHealthIdentity $os $version ([datetime]'2026-10-14')).LifecycleStatus | Should -Be Attention
    }
}

Describe 'Health evidence quality' {
    It 'never treats string false, absent properties or failed probes as enabled' {
        foreach ($probe in @((New-TestProbe ([pscustomobject]@{ Enabled = 'False' })), (New-TestProbe ([pscustomobject]@{})), (New-TestProbe $null $false))) {
            (New-HealthBooleanCheck 'test' Test Test $probe Enabled Context Source).Status | Should -Be Unknown
        }
        (New-HealthBooleanCheck 'test' Test Test (New-TestProbe ([pscustomobject]@{ Enabled = $false })) Enabled Context Source).Status | Should -Be Attention
    }
    It 'does not report a pending restart for an empty marker list' {
        $report = New-TestHealth
        ($report.Checks | Where-Object Id -eq windows.reboot).Value | Should -Be 'No common markers'
        $report.Summary.Unknown | Should -Be 0
    }
    It 'retains provider errors without discarding the report' {
        $probe = Invoke-HealthProbe { throw 'Access denied fixture' }
        $probe.Success | Should -BeFalse
        $probe.Error | Should -Match 'Access denied fixture'
    }
    It 'HTML-encodes host names and evidence text' {
        $report = New-TestHealth
        $report.ComputerName = '<script>bad()</script>'
        $report.Checks[0].Detail = '<img src=x onerror=bad()>'
        $html = ConvertTo-HealthHtml $report
        $html | Should -Not -Match '<script>bad'
        $html | Should -Match '&lt;img'
    }
}

Describe 'Environment baseline and drift' {
    It 'distinguishes an empty section from a failed provider' {
        (New-EnvironmentSection empty Empty @('Name') {}).Status | Should -Be Captured
        $failed = New-EnvironmentSection failed Failed @('Name') { throw 'access denied' }
        $failed.Status | Should -Be Unknown
        $failed.Error | Should -Be 'access denied'
    }
    It 'rejects duplicate and missing identities' {
        (New-EnvironmentSection duplicate Duplicate @('Name') { [pscustomobject]@{ Name = 'same' }; [pscustomobject]@{ Name = 'same' } }).Status | Should -Be Unknown
        (New-EnvironmentSection missing Missing @('Name') { [pscustomobject]@{ Other = 'value' } }).Status | Should -Be Unknown
    }
    It 'identifies changed configuration with both values' {
        $diff = Compare-EnvironmentBaseline (New-TestBaseline @((New-TestSection Running))) (New-TestBaseline @((New-TestSection Stopped)))
        $diff.Changes.Count | Should -Be 1
        $diff.Changes[0].Change | Should -Be Changed
        $diff.Changes[0].Before[0].State | Should -Be Running
        $diff.Changes[0].After[0].State | Should -Be Stopped
    }
    It 'detects additions and removals' {
        $empty = New-EnvironmentSection system.services Services @('Name') {}
        (Compare-EnvironmentBaseline (New-TestBaseline @($empty)) (New-TestBaseline @((New-TestSection)))).Changes[0].Change | Should -Be Added
        (Compare-EnvironmentBaseline (New-TestBaseline @((New-TestSection))) (New-TestBaseline @($empty))).Changes[0].Change | Should -Be Removed
    }
    It 'does not turn unreadable or missing sections into mass removals' {
        $unknown = New-EnvironmentSection system.services Services @('Name') { throw 'Access denied' }
        foreach ($after in @((New-TestBaseline @($unknown)), (New-TestBaseline @()))) {
            $diff = Compare-EnvironmentBaseline (New-TestBaseline @((New-TestSection))) $after
            $diff.Changes.Count | Should -Be 1
            $diff.Changes[0].Change | Should -Be Unknown
        }
    }
    It 'ignores property order but preserves ordered values such as DNS priority' {
        $a = New-EnvironmentSection dns DNS @('Name') { [pscustomobject]@{ Name = 'LAN'; Servers = @('192.0.2.1', '192.0.2.2') } }
        $b = New-EnvironmentSection dns DNS @('Name') { [pscustomobject]@{ Servers = @('192.0.2.1', '192.0.2.2'); Name = 'LAN' } }
        (Compare-EnvironmentBaseline (New-TestBaseline @($a)) (New-TestBaseline @($b))).Changes.Count | Should -Be 0
        $b.Items[0].Data.Servers = @('192.0.2.2', '192.0.2.1')
        (Compare-EnvironmentBaseline (New-TestBaseline @($a)) (New-TestBaseline @($b))).Changes[0].Change | Should -Be Changed
    }
    It 'ignores listener PID churn but retains executable ownership changes' {
        $a = New-EnvironmentSection network.tcp TCP @('LocalPort', 'ProcessId') { [pscustomobject]@{ Protocol = 'TCP'; LocalAddress = '127.0.0.1'; LocalPort = 443; ProcessId = 100; ProcessName = 'demo'; ExecutablePath = 'C:\demo.exe' } }
        $b = New-EnvironmentSection network.tcp TCP @('LocalPort', 'ProcessId') { [pscustomobject]@{ Protocol = 'TCP'; LocalAddress = '127.0.0.1'; LocalPort = 443; ProcessId = 200; ProcessName = 'demo'; ExecutablePath = 'C:\demo.exe' } }
        (Compare-EnvironmentBaseline (New-TestBaseline @($a)) (New-TestBaseline @($b))).Changes.Count | Should -Be 0
        $b.Items[0].Data.ExecutablePath = 'C:\other.exe'
        (Compare-EnvironmentBaseline (New-TestBaseline @($a)) (New-TestBaseline @($b))).Changes[0].Change | Should -Be Changed
        $b.Items[0].Data.ExecutablePath = ''
        (Compare-EnvironmentBaseline (New-TestBaseline @($a)) (New-TestBaseline @($b))).Changes[0].Change | Should -Be Unknown
    }
    It 'refuses cross-computer and non-environment inputs' {
        $a = New-TestBaseline @((New-TestSection))
        $b = New-TestBaseline @((New-TestSection))
        $b.ComputerName = 'ANOTHER-PC'
        { Compare-EnvironmentBaseline $a $b } | Should -Throw '*different computer*'
        { Assert-EnvironmentBaseline ([pscustomobject]@{ SchemaVersion = 2; ReportType = 'WinDefState' }) } | Should -Throw '*supported*'
    }
    It 'round-trips JSON without overwriting an existing file' {
        $report = New-TestBaseline @((New-TestSection))
        $path = Join-Path $TestDrive 'before.json'
        $null = Export-EnvironmentReport $report $path
        $read = Read-EnvironmentBaseline $path
        (Compare-EnvironmentBaseline $report $read).Changes.Count | Should -Be 0
        { Export-EnvironmentReport $report $path } | Should -Throw
    }
    It 'escapes untrusted inventory in standalone HTML' {
        $report = New-TestBaseline @((New-TestSection '<script>bad()</script>'))
        $html = ConvertTo-EnvironmentHtml $report
        $html | Should -Not -Match '<script>bad'
        $html | Should -Match '(&lt;|\\u003c)script(&gt;|\\u003e)bad'
    }
    It 'passes native CLI parameters through dependency imports' {
        $beforePath = Join-Path $TestDrive 'cli-before.json'
        $afterPath = Join-Path $TestDrive 'cli-after.json'
        $diffPath = Join-Path $TestDrive 'cli-diff.json'
        $null = Export-EnvironmentReport (New-TestBaseline @((New-TestSection Running))) $beforePath
        $null = Export-EnvironmentReport (New-TestBaseline @((New-TestSection Stopped))) $afterPath
        $null = & (Join-Path $repositoryRoot 'WinDefState.Environment.ps1') -BaselinePath $beforePath -CurrentPath $afterPath -OutputPath $diffPath
        (Get-Content -LiteralPath $diffPath -Raw | ConvertFrom-Json).Changes[0].Change | Should -Be Changed
    }
}

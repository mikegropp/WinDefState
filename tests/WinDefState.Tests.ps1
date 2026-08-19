BeforeAll {
    $enginePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'WinDefState.ps1'
    . $enginePath

    $guiPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'WinDefState.Gui.ps1'
    $tokens = $null
    $parseErrors = $null
    $guiAst = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -gt 0) {
        throw ($parseErrors.Message -join [Environment]::NewLine)
    }
    $argumentHelper = $guiAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'ConvertTo-WindowsCommandLineArgument'
    }, $true)
    Invoke-Expression $argumentHelper.Extent.Text
    $guiArgumentsFunction = $guiAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-WinDefStateArguments'
    }, $true)
    Invoke-Expression $guiArgumentsFunction.Extent.Text
    $script:EnginePath = $enginePath
    $script:StateRoot = 'C:\ProgramData\WinDefState'
    foreach ($helperName in @(
        'ConvertTo-ShortText',
        'Get-ItemCount',
        'Get-EntryCategory',
        'Get-EntryCurrentSummary',
        'Get-EntryCapabilities',
        'Get-EntryBadges',
        'ConvertTo-GuiCanonicalValue',
        'Get-GuiEntryFingerprint',
        'Get-EntryPermissiveTargetSummary',
        'New-SnapshotRow',
        'Test-SnapshotRowVisible',
        'Get-SnapshotRowSummary',
        'Rebuild-SnapshotRows',
        'Get-MutationPreviewRows',
        'Get-GuiOperationResultValue'
    )) {
        $helper = $guiAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $helperName
        }, $true)
        if ($null -eq $helper) {
            throw "GUI helper was not found: $helperName"
        }
        Invoke-Expression $helper.Extent.Text
    }

    if ($null -eq (Get-Command Get-CimInstance -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Get-CimInstance -Value {
            [CmdletBinding()]
            param([string]$ClassName, [string]$Filter)
            throw 'The test-only Get-CimInstance stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Invoke-CimMethod -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Invoke-CimMethod -Value {
            [CmdletBinding()]
            param([object]$InputObject, [string]$MethodName, [hashtable]$Arguments)
            throw 'The test-only Invoke-CimMethod stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Get-BitLockerVolume -Value {
            [CmdletBinding()]
            param([string[]]$MountPoint)
            throw 'The test-only Get-BitLockerVolume stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Resume-BitLocker -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Resume-BitLocker -Value {
            [CmdletBinding()]
            param([string]$MountPoint)
            throw 'The test-only Resume-BitLocker stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Suspend-BitLocker -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Suspend-BitLocker -Value {
            [CmdletBinding()]
            param([string]$MountPoint, [int]$RebootCount)
            throw 'The test-only Suspend-BitLocker stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Get-MpPreference -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Get-MpPreference -Value {
            [CmdletBinding()]
            param()
            throw 'The test-only Get-MpPreference stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Set-MpPreference -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Set-MpPreference -Value {
            [CmdletBinding()]
            param(
                [object]$DisableRealtimeMonitoring,
                [object]$DisableBehaviorMonitoring,
                [object]$MAPSReporting,
                [object]$SubmitSamplesConsent,
                [object]$PUAProtection,
                [object]$DisableScriptScanning,
                [object]$DisableIOAVProtection,
                [object]$DisableIntrusionPreventionSystem,
                [object]$EnableNetworkProtection,
                [object]$EnableControlledFolderAccess
            )
            throw 'The test-only Set-MpPreference stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Remove-MpPreference -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Remove-MpPreference -Value {
            [CmdletBinding()]
            param(
                [string[]]$AttackSurfaceReductionRules_Ids,
                [object[]]$AttackSurfaceReductionRules_Actions,
                [string[]]$ExclusionPath,
                [string[]]$ExclusionProcess,
                [string[]]$ExclusionExtension,
                [string[]]$ControlledFolderAccessAllowedApplications,
                [string[]]$ControlledFolderAccessProtectedFolders
            )
            throw 'The test-only Remove-MpPreference stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Add-MpPreference -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Add-MpPreference -Value {
            [CmdletBinding()]
            param(
                [string[]]$AttackSurfaceReductionRules_Ids,
                [object[]]$AttackSurfaceReductionRules_Actions,
                [string[]]$ExclusionPath,
                [string[]]$ExclusionProcess,
                [string[]]$ExclusionExtension,
                [string[]]$ControlledFolderAccessAllowedApplications,
                [string[]]$ControlledFolderAccessProtectedFolders
            )
            throw 'The test-only Add-MpPreference stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Get-WSManInstance -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Get-WSManInstance -Value {
            [CmdletBinding()]
            param([string]$ResourceURI, [switch]$Enumerate)
            throw 'The test-only Get-WSManInstance stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Set-WSManInstance -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Set-WSManInstance -Value {
            [CmdletBinding()]
            param([string]$ResourceURI, [hashtable]$ValueSet)
            throw 'The test-only Set-WSManInstance stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Get-NetFirewallProfile -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Get-NetFirewallProfile -Value {
            [CmdletBinding()]
            param([string[]]$Profile)
            throw 'The test-only Get-NetFirewallProfile stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Set-NetFirewallProfile -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Set-NetFirewallProfile -Value {
            [CmdletBinding()]
            param(
                [string[]]$Profile,
                [object]$Enabled,
                [object]$DefaultInboundAction,
                [object]$DefaultOutboundAction,
                [object]$AllowUnicastResponseToMulticast,
                [object]$NotifyOnListen,
                [object]$LogAllowed,
                [object]$LogBlocked,
                [object]$LogIgnored
            )
            throw 'The test-only Set-NetFirewallProfile stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Set-NetFirewallRule -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Set-NetFirewallRule -Value {
            [CmdletBinding()]
            param([string[]]$Name, [string]$Group, [object]$Enabled)
            throw 'The test-only Set-NetFirewallRule stub must be mocked before use.'
        }
    }
    if ($null -eq (Get-Command Start-Service -ErrorAction SilentlyContinue)) {
        Set-Item -Path Function:\global:Start-Service -Value {
            [CmdletBinding()]
            param([string]$Name)
            throw 'The test-only Start-Service stub must be mocked before use.'
        }
    }

    function New-TestNetworkAdapterInstance {
        param(
            [Parameter(Mandatory)] [uint32]$Index,
            [Parameter(Mandatory)] [string]$Description
        )

        if ($null -ne (Get-Command New-CimInstance -ErrorAction SilentlyContinue)) {
            return New-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Property @{
                Index       = $Index
                Description = $Description
            } -ClientOnly
        }

        [PSCustomObject]@{ Index = $Index; Description = $Description }
    }
}

Describe 'Definition catalog' {
    It 'reuses the immutable catalog and ID map throughout one process' {
        $script:WinDefStateDefinitionCache = $null
        $script:WinDefStateDefinitionMapCache = $null

        $firstDefinitions = @(Get-DefenseDefinitions)
        $secondDefinitions = @(Get-DefenseDefinitions)
        $firstMap = Get-DefenseDefinitionMap
        $secondMap = Get-DefenseDefinitionMap

        $firstDefinitions.Count | Should -Be 94
        [object]::ReferenceEquals($firstDefinitions[0], $secondDefinitions[0]) | Should -BeTrue
        [object]::ReferenceEquals($firstMap, $secondMap) | Should -BeTrue
        [object]::ReferenceEquals($firstMap['defender.runtime_status'], $firstDefinitions[0]) | Should -BeTrue
    }

    It 'contains 94 unique stable setting IDs' {
        $definitions = @(Get-DefenseDefinitions)

        $definitions.Count | Should -Be 94
        @($definitions.Id | Sort-Object -Unique).Count | Should -Be $definitions.Count
    }

    It 'includes the extended credential, remote-access, and network controls' {
        $definitions = Get-DefenseDefinitionMap

        $definitions['ntlm.lm_compatibility_level'].PermissiveValue | Should -Be 0
        $definitions['ntlm.lm_compatibility_level'].RequiresReboot | Should -BeFalse
        $definitions['ntlm.minimum_client_session_security'].PermissiveValue | Should -Be 0
        $definitions['ntlm.minimum_server_session_security'].PermissiveValue | Should -Be 0
        $definitions['ldap.client_signing'].PermissiveValue | Should -Be 0
        $definitions['smb.client.allow_insecure_guest_auth'].PermissiveValue | Should -Be 1
        $definitions['smb.client.require_encryption'].PermissiveValue | Should -Be 0
        $definitions['rdp.min_encryption_level'].PermissiveValue | Should -Be 1
        $definitions['rdp.min_encryption_level'].RequiresReboot | Should -BeTrue
        $definitions['rdp.allow_clipboard_redirection'].PermissiveValue | Should -Be 0
        $definitions['rdp.allow_clipboard_redirection'].RequiresReboot | Should -BeTrue
        $definitions['rdp.allow_drive_redirection'].PermissiveValue | Should -Be 0
        $definitions['uac.local_account_token_filter_policy'].PermissiveValue | Should -Be 1
        $definitions['accounts.limit_blank_password_use'].PermissiveValue | Should -Be 0
        $definitions['network.restrict_anonymous'].PermissiveValue | Should -Be 0
        $definitions['network.restrict_anonymous_sam'].PermissiveValue | Should -Be 0
        $definitions['network.everyone_includes_anonymous'].PermissiveValue | Should -Be 1
        $definitions['logon.cached_domain_logons'].ValueKind | Should -Be 'String'
        $definitions['logon.cached_domain_logons'].PermissiveValue | Should -Be '50'
        $definitions['logon.cached_domain_logons'].RequiresReboot | Should -BeTrue
    }

    It 'provides a reviewable target descriptor for every permissive definition' {
        $definitions = @(Get-DefenseDefinitions | Where-Object { Test-DefinitionHasPermissiveAction -Definition $_ })

        foreach ($definition in $definitions) {
            $descriptor = Get-DefinitionPermissiveTargetDescriptor -Definition $definition -Entry $null
            $descriptor | Should -Not -BeNullOrEmpty -Because $definition.Id
            [string]$descriptor.Summary | Should -Not -BeNullOrEmpty -Because $definition.Id
            [string]$descriptor.Mode | Should -BeIn @('Exact', 'BaselineDependent') -Because $definition.Id
        }
    }

    It 'selects only requested definitions before capture begins' {
        $selected = @(Get-SelectedDefenseDefinitions -IncludeId @('rdp.user_authentication', 'winrm.service'))
        $excluded = @(Get-SelectedDefenseDefinitions -ExcludeId @('wdac.policies'))

        $selected.Count | Should -Be 2
        @($selected.Id) | Should -Contain 'rdp.user_authentication'
        @($selected.Id) | Should -Contain 'winrm.service'
        $excluded.Count | Should -Be 93
        @($excluded.Id) | Should -Not -Contain 'wdac.policies'
    }

    It 'supports validated wildcard and category scopes' {
        $rdpDefinitions = @(Get-SelectedDefenseDefinitions -IncludeId 'rdp.*' -ExcludeId 'rdp.firewall_rules')
        $categoryFilters = @(Merge-SettingIdFilter -Id 'rdp.allow_connections' -Category @('Defender, FIREWALL'))
        $categoryDefinitions = @(Get-SelectedDefenseDefinitions -IncludeId $categoryFilters)

        $rdpDefinitions.Count | Should -Be 8
        @($rdpDefinitions.Id | Where-Object { $_ -notlike 'rdp.*' }).Count | Should -Be 0
        @($rdpDefinitions.Id) | Should -Not -Contain 'rdp.firewall_rules'
        $categoryFilters | Should -Contain 'defender.*'
        $categoryFilters | Should -Contain 'firewall.*'
        $categoryFilters | Should -Contain 'rdp.allow_connections'
        @($categoryDefinitions.Id) | Should -Contain 'defender.runtime_status'
        @($categoryDefinitions.Id) | Should -Contain 'firewall.profiles'
        @($categoryDefinitions.Id) | Should -Contain 'rdp.allow_connections'
    }

    It 'rejects unknown categories and unmatched wildcard filters' {
        { ConvertTo-CategoryIdFilter -Category 'not-a-category' } | Should -Throw '*Unknown setting category*'
        { Get-SelectedDefenseDefinitions -IncludeId 'missing.*' } | Should -Throw '*Unknown setting ID or unmatched pattern*'
    }

    It 'passes selected scope into permissive snapshot export' {
        (Get-Command Set-DefensePermissive -CommandType Function).Definition |
            Should -Match 'Export-DefenseSnapshot\s+-Path\s+\$Path\s+-IncludeId\s+\$IncludeId\s+-ExcludeId\s+\$ExcludeId'
    }

    It 'dispatches every definition type through each lifecycle phase' {
        $definitionTypes = @(Get-DefenseDefinitions | ForEach-Object { [string]$_.Type } | Sort-Object -Unique)
        $dispatchFunctions = @(
            'Capture-Definition'
            'Apply-PermissiveDefinition'
            'Restore-SnapshotEntry'
            'Add-SnapshotEntryReportLines'
            'ConvertTo-ComparableSnapshotEntry'
            'Test-SnapshotEntryCapturedExactly'
            'Test-DefinitionHasRestoreAction'
            'Test-PermissiveDefinitionState'
        )

        $dispatchGaps = @()
        foreach ($functionName in $dispatchFunctions) {
            $functionText = (Get-Command -Name $functionName -CommandType Function).Definition
            foreach ($definitionType in $definitionTypes) {
                if ($functionText -notmatch ("(?m)^\s*'{0}'\s*\{{" -f [regex]::Escape($definitionType))) {
                    $dispatchGaps += "$functionName -> $definitionType"
                }
            }
        }
        $dispatchGaps | Should -BeNullOrEmpty
    }
}

Describe 'Stable state root' {
    It 'uses the same ProgramData default in the engine and GUI' {
        $engineText = Get-Content -LiteralPath $enginePath -Raw
        $guiText = Get-Content -LiteralPath $guiPath -Raw

        $engineText | Should -Match "Join-Path\s+\`$env:ProgramData\s+'WinDefState'"
        $guiText | Should -Match "Join-Path\s+\`$env:ProgramData\s+'WinDefState'"
    }

    It 'protects the trusted state root before acquiring the operation lock' {
        $engineText = Get-Content -LiteralPath $enginePath -Raw
        $protectIndex = $engineText.LastIndexOf('Protect-StateRoot -Path $StateRoot')
        $lockIndex = $engineText.LastIndexOf('$operationLock = Enter-WinDefStateOperationLock')

        $protectIndex | Should -BeGreaterThan -1
        $lockIndex | Should -BeGreaterThan $protectIndex
        (Get-Command Protect-StateRoot -CommandType Function).Definition |
            Should -Match 'SetAccessRuleProtection\(\$true,\s*\$false\)'
    }
}

Describe 'PowerShell safety semantics' {
    It 'gates all three public command paths with ShouldProcess' {
        $engineText = Get-Content -LiteralPath $enginePath -Raw

        $engineText | Should -Match 'CmdletBinding\(SupportsShouldProcess\s*=\s*\$true'
        @([regex]::Matches($engineText, '\$PSCmdlet\.ShouldProcess\(')).Count | Should -Be 3
    }
}

Describe 'Safe operation cancellation' {
    It 'recognizes a marker file and reports a no-mutation cancellation' {
        $cancellationPath = Join-Path $TestDrive 'cancel.requested'

        Test-OperationCancellationRequested -Path $cancellationPath | Should -BeFalse
        Set-Content -LiteralPath $cancellationPath -Value 'cancel'
        Test-OperationCancellationRequested -Path $cancellationPath | Should -BeTrue
        { Assert-OperationNotCancelled -Path $cancellationPath -Stage 'test boundary' } |
            Should -Throw '*cancelled at a safe boundary*No defense setting was changed*'
    }

    It 'checks permissive cancellation after persistence but before journaling or mutation' {
        $functionText = (Get-Command Set-DefensePermissive -CommandType Function).Definition
        $exportIndex = $functionText.IndexOf('Export-DefenseSnapshot')
        $cancelIndex = $functionText.IndexOf("Assert-OperationNotCancelled -Path `$CancellationPath -Stage 'before permissive mutation'")
        $reviewIndex = $functionText.IndexOf('Wait-MutationApproval')
        $journalIndex = $functionText.IndexOf('Write-OperationState')
        $mutationIndex = $functionText.IndexOf('Invoke-PermissiveMutationWorkItem')

        $exportIndex | Should -BeGreaterOrEqual 0
        $cancelIndex | Should -BeGreaterThan $exportIndex
        $reviewIndex | Should -BeGreaterThan $cancelIndex
        $journalIndex | Should -BeGreaterThan $reviewIndex
        $mutationIndex | Should -BeGreaterThan $journalIndex
    }

    It 'checks restore cancellation before changing journal status or live state' {
        $functionText = (Get-Command Restore-DefenseSnapshot -CommandType Function).Definition
        $cancelIndex = $functionText.IndexOf("Assert-OperationNotCancelled -Path `$CancellationPath -Stage 'before restore mutation'")
        $reviewIndex = $functionText.IndexOf('Wait-MutationApproval')
        $statusIndex = $functionText.IndexOf("-Status 'Restoring'")
        $mutationIndex = $functionText.IndexOf('Invoke-RestoreMutationWorkItem')

        $cancelIndex | Should -BeGreaterOrEqual 0
        $reviewIndex | Should -BeGreaterThan $cancelIndex
        $statusIndex | Should -BeGreaterThan $reviewIndex
        $mutationIndex | Should -BeGreaterThan $statusIndex
    }

    It 'does not journal or mutate when permissive cancellation is accepted after snapshot persistence' {
        $script:StateRoot = Join-Path $TestDrive 'cancelled-permissive-state'
        $snapshotPath = Join-Path $TestDrive 'cancelled-permissive.json'
        $cancellationPath = Join-Path $TestDrive 'cancelled-permissive.request'
        Set-Content -LiteralPath $cancellationPath -Value 'cancel'
        $definition = [PSCustomObject]@{
            Id = 'test.registry'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Test'; Name = 'Enabled'
            ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false
        }
        $entry = [PSCustomObject]@{
            Id = 'test.registry'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Test'; Name = 'Enabled'
            ValueKind = 'DWord'; Captured = $true; Exists = $true; CurrentValue = 1; RequiresReboot = $false
        }
        Mock Assert-NoActiveOperation {}
        Mock Get-SelectedDefenseDefinitions { @($definition) }
        Mock Export-DefenseSnapshot {
            [PSCustomObject]@{
                JsonPath = $snapshotPath; ReportPath = ($snapshotPath + '.txt')
                Snapshot = [PSCustomObject]@{ Settings = @($entry) }
            }
        }
        Mock Write-OperationState {}
        Mock Apply-PermissiveDefinition {}

        { Set-DefensePermissive -Path $snapshotPath -CancellationPath $cancellationPath } |
            Should -Throw '*cancelled at a safe boundary*'
        Should -Invoke Export-DefenseSnapshot -Times 1 -Exactly
        Should -Invoke Write-OperationState -Times 0 -Exactly
        Should -Invoke Apply-PermissiveDefinition -Times 0 -Exactly
    }

    It 'accepts only an explicit approval marker after announcing the persisted baseline' {
        $script:StateRoot = $TestDrive
        $approvalPath = Join-Path $TestDrive ('.review-{0}.approve' -f ([guid]::NewGuid().ToString('N')))
        $script:approvalProbeCount = 0
        $script:EmitProgress = $true
        Mock Test-Path {
            if ([string]$LiteralPath -ne $approvalPath) {
                return $false
            }
            $script:approvalProbeCount++
            if ($script:approvalProbeCount -eq 1) {
                return $false
            }
            if (-not [IO.File]::Exists($approvalPath)) {
                [IO.File]::WriteAllText($approvalPath, 'APPROVE')
            }
            return $true
        }

        try {
            $output = @(Wait-MutationApproval -Path $approvalPath -Action Permissive -SnapshotPath (Join-Path $TestDrive 'baseline.json') -TrustedRoot $TestDrive 6>&1 | ForEach-Object { [string]$_ })

            $output | Should -Contain ("WDS_REVIEW|Permissive|{0}" -f (Join-Path $TestDrive 'baseline.json'))
            $output | Should -Contain 'WDS_APPROVED|Permissive'
            [IO.File]::Exists($approvalPath) | Should -BeFalse
        } finally {
            $script:EmitProgress = $false
        }
    }

    It 'rejects mutation approval markers outside the protected state root' {
        $script:StateRoot = Join-Path $TestDrive 'trusted-state'
        $outsidePath = Join-Path $TestDrive ('.review-{0}.approve' -f ([guid]::NewGuid().ToString('N')))

        { Wait-MutationApproval -Path $outsidePath -Action Restore -SnapshotPath (Join-Path $TestDrive 'baseline.json') -TrustedRoot $script:StateRoot } |
            Should -Throw '*directly under the protected state root*'
    }

    It 'treats an explicit protected CANCEL decision as a no-mutation rejection' {
        $script:StateRoot = $TestDrive
        $decisionPath = Join-Path $TestDrive ('.review-{0}.approve' -f ([guid]::NewGuid().ToString('N')))
        $script:decisionProbeCount = 0
        Mock Test-Path {
            if ([string]$LiteralPath -ne $decisionPath) {
                return $false
            }
            $script:decisionProbeCount++
            if ($script:decisionProbeCount -eq 1) {
                return $false
            }
            if (-not [IO.File]::Exists($decisionPath)) {
                [IO.File]::WriteAllText($decisionPath, 'CANCEL')
            }
            return $true
        }

        { Wait-MutationApproval -Path $decisionPath -Action Restore -SnapshotPath (Join-Path $TestDrive 'baseline.json') -TrustedRoot $TestDrive } |
            Should -Throw '*pre-change review was rejected*no defense setting was changed*'
        [IO.File]::Exists($decisionPath) | Should -BeFalse
    }
}

Describe 'Command discovery cache' {
    It 'resolves each command once and suppresses module-autoload verbosity' {
        Clear-WinDefStateCommandCache
        $script:observedCommandVerbosePreference = $null
        Mock Get-Command {
            $script:observedCommandVerbosePreference = [string]$VerbosePreference
            [PSCustomObject]@{ Name = $Name; Source = 'test' }
        }

        $first = Get-WinDefStateCommand -Name 'Test-ProviderCommand'
        $second = Get-WinDefStateCommand -Name 'TEST-PROVIDERCOMMAND'

        $first.Name | Should -Be 'Test-ProviderCommand'
        $second | Should -Be $first
        $script:observedCommandVerbosePreference | Should -Be 'SilentlyContinue'
        Should -Invoke Get-Command -Times 1 -Exactly
    }

    It 'also caches unavailable commands' {
        Clear-WinDefStateCommandCache
        Mock Get-Command { $null }

        Test-CommandAvailable -Name 'Missing-ProviderCommand' | Should -BeFalse
        Test-CommandAvailable -Name 'missing-providercommand' | Should -BeFalse

        Should -Invoke Get-Command -Times 1 -Exactly
    }
}

Describe 'Generated artifact paths' {
    It 'uses a numeric suffix instead of overwriting an existing artifact' {
        $artifactRoot = Join-Path $TestDrive 'artifacts'
        $firstPath = Get-AvailableArtifactPath -Directory $artifactRoot -BaseName 'snapshot' -Extension 'json'
        Set-Content -LiteralPath $firstPath -Value '{}'

        $secondPath = Get-AvailableArtifactPath -Directory $artifactRoot -BaseName 'snapshot' -Extension '.json'

        [IO.Path]::GetFileName($firstPath) | Should -Be 'snapshot.json'
        [IO.Path]::GetFileName($secondPath) | Should -Be 'snapshot-1.json'
    }

    It 'releases verification capture resources even when verification aborts' {
        (Get-Command Test-DefenseSnapshot -CommandType Function).Definition |
            Should -Match 'finally\s*\{\s*\$captureMetrics\s*=\s*Complete-CaptureSession'
    }

    It 'renders empty and singleton capture-scope collections under strict mode' {
        $definition = @(Get-DefenseDefinitions | Where-Object { $_.Id -eq 'rdp.user_authentication' })[0]
        $entry = [PSCustomObject]@{
            Id = $definition.Id; Type = $definition.Type; Path = $definition.Path; Name = $definition.Name
            ValueKind = $definition.ValueKind; Captured = $true; Exists = $true; CurrentValue = 1; RequiresReboot = $false
        }
        $scopes = @(
            [PSCustomObject]@{ IsFiltered = $false; IncludeId = @(); ExcludeId = @() }
            [PSCustomObject]@{ IsFiltered = $true; IncludeId = @($definition.Id); ExcludeId = @() }
        )

        foreach ($scope in $scopes) {
            $snapshot = [PSCustomObject]@{
                Tool = 'WinDefState'; ComputerName = 'HOST'; CapturedAtUtc = '2026-01-01T00:00:00Z'
                CaptureScope = $scope; Settings = @($entry)
            }
            $reportLines = @(Get-SnapshotReportLines -Snapshot $snapshot -SnapshotPath (Join-Path $TestDrive 'scope.json'))
            $reportLines.Count | Should -BeGreaterThan 0
        }
    }

    It 'persists source-of-truth JSON before report presentation can fail' {
        $snapshotPath = Join-Path $TestDrive 'durable-before-report.json'
        $script:persistenceDefinition = @(Get-DefenseDefinitions | Where-Object { $_.Id -eq 'rdp.user_authentication' })[0]
        Mock Get-SelectedDefenseDefinitions { @($script:persistenceDefinition) }
        Mock Capture-Definition {
            param($Definition, $CaptureSession)

            [PSCustomObject]@{
                Id = $Definition.Id; Type = $Definition.Type; Path = $Definition.Path; Name = $Definition.Name
                ValueKind = $Definition.ValueKind; Captured = $true; Exists = $true; CurrentValue = 1; RequiresReboot = $false
            }
        }
        Mock Get-SnapshotReportLines { throw 'expected report failure' }

        { Export-DefenseSnapshot -Path $snapshotPath } | Should -Throw '*Snapshot JSON was saved successfully*No defense setting was changed*expected report failure*'

        Test-Path -LiteralPath $snapshotPath -PathType Leaf | Should -BeTrue
        $persisted = Read-JsonFile -Path $snapshotPath
        @($persisted.Settings).Count | Should -Be 1
        $persisted.Settings[0].Id | Should -Be 'rdp.user_authentication'
    }
}

Describe 'Snapshot sidecar cache' {
    It 'reuses one immutable text and byte read throughout an operation' {
        Clear-SnapshotAssetCache
        $textPath = Join-Path $TestDrive 'policy.xml'
        $bytesPath = Join-Path $TestDrive 'policy.cip'
        [IO.File]::WriteAllText($textPath, '<Policy />')
        [IO.File]::WriteAllBytes($bytesPath, [byte[]]@(1, 2, 3, 4))

        Read-SnapshotAssetText -Path $textPath | Should -Be '<Policy />'
        $firstBytes = [byte[]](Read-SnapshotAssetBytes -Path $bytesPath)
        @($firstBytes) | Should -Be @(1, 2, 3, 4)
        $script:WinDefStateSnapshotAssetCache.Count | Should -Be 2

        Remove-Item -LiteralPath $textPath, $bytesPath -Force
        Read-SnapshotAssetText -Path $textPath | Should -Be '<Policy />'
        $cachedBytes = [byte[]](Read-SnapshotAssetBytes -Path $bytesPath)
        @($cachedBytes) | Should -Be @(1, 2, 3, 4)
        $script:WinDefStateSnapshotAssetCache.Count | Should -Be 2
    }

    It 'records and enforces hashes for text sidecars in new snapshots' {
        Clear-SnapshotAssetCache
        $snapshotPath = Join-Path $TestDrive 'hashed.json'
        $snapshot = [PSCustomObject]@{
            Settings = @(
                [PSCustomObject]@{
                    Type = 'AppLockerPolicy'
                    CurrentValue = [PSCustomObject]@{
                        CommandAvailable = $true
                        LocalCaptured = $true
                        EffectiveCaptured = $true
                        LocalMatchesEffective = $true
                        CaptureIssues = @()
                        CollectionSummaries = @()
                        LocalXml = '<AppLockerPolicy Version="1" />'
                        EffectiveXml = '<AppLockerPolicy Version="1" />'
                    }
                }
                [PSCustomObject]@{
                    Type = 'ExploitProtectionPolicy'
                    CurrentValue = [PSCustomObject]@{
                        CommandAvailable = $true
                        Xml = '<MitigationPolicy />'
                    }
                }
            )
        }

        Persist-SnapshotExternalAssets -Snapshot $snapshot -SnapshotPath $snapshotPath
        $appLockerState = $snapshot.Settings[0].CurrentValue
        $exploitState = $snapshot.Settings[1].CurrentValue
        $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $snapshotPath
        $localPath = Join-Path $assetRoot $appLockerState.LocalSnapshotAssetRelativePath
        $exploitPath = Join-Path $assetRoot $exploitState.SnapshotAssetRelativePath

        $appLockerState.LocalSnapshotAssetSha256 | Should -Be (Get-FileHash -LiteralPath $localPath -Algorithm SHA256).Hash
        $appLockerState.EffectiveSnapshotAssetSha256 | Should -Match '^[A-F0-9]{64}$'
        $exploitState.SnapshotAssetSha256 | Should -Be (Get-FileHash -LiteralPath $exploitPath -Algorithm SHA256).Hash

        Clear-SnapshotAssetCache
        Write-TextAtomic -Path $localPath -Content '<AppLockerPolicy Version="2" />'
        { Get-AppLockerPolicyXml -State $appLockerState -PolicyScope Local -SnapshotPath $snapshotPath } |
            Should -Throw '*AppLocker snapshot asset no longer matches its recorded SHA-256*'
    }

    It 'rejects WDAC bytes that do not match their recorded hash' {
        $content = [byte[]]@(1, 2, 3, 4)
        $file = [PSCustomObject]@{
            RelativePath = 'CiPolicies/Active/test.cip'
            Base64 = [Convert]::ToBase64String($content)
            Sha256 = ('0' * 64)
        }

        { Get-WdacSnapshotFileBytes -File $file } |
            Should -Throw '*WDAC snapshot content no longer matches its recorded SHA-256*'
    }

    It 'preloads only selected sidecar-backed entries before restore mutation' {
        Clear-SnapshotAssetCache
        $snapshotPath = Join-Path $TestDrive 'selected.json'
        $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $snapshotPath
        $appLockerRoot = Join-Path $assetRoot 'applocker'
        New-Item -ItemType Directory -Path $appLockerRoot -Force | Out-Null
        [IO.File]::WriteAllText((Join-Path $appLockerRoot 'local.xml'), '<AppLockerPolicy Version="1" />')
        [IO.File]::WriteAllText((Join-Path $appLockerRoot 'effective.xml'), '<AppLockerPolicy Version="1" />')
        $entries = @(
            [PSCustomObject]@{
                Id = 'applocker.policy'; Type = 'AppLockerPolicy'
                CurrentValue = [PSCustomObject]@{
                    CommandAvailable = $true
                    LocalCaptured = $true
                    EffectiveCaptured = $true
                    LocalMatchesEffective = $true
                    CaptureIssues = @()
                    LocalSnapshotAssetRelativePath = 'applocker/local.xml'
                    EffectiveSnapshotAssetRelativePath = 'applocker/effective.xml'
                }
            }
            [PSCustomObject]@{
                Id = 'exploit_protection.policy'; Type = 'ExploitProtectionPolicy'
                CurrentValue = [PSCustomObject]@{ SnapshotAssetRelativePath = 'exploit/missing.xml' }
            }
        )

        { Initialize-SnapshotAssetCache -Entries @($entries[0]) -SnapshotPath $snapshotPath } | Should -Not -Throw
        $script:WinDefStateSnapshotAssetCache.Count | Should -Be 2
        { Initialize-SnapshotAssetCache -Entries @($entries[1]) -SnapshotPath $snapshotPath } | Should -Throw '*Exploit protection snapshot asset is missing*'
    }

    It 'does not require sidecars for baselines already marked incomplete' {
        Clear-SnapshotAssetCache
        $snapshotPath = Join-Path $TestDrive 'incomplete.json'
        $entries = @(
            [PSCustomObject]@{
                Id = 'applocker.policy'; Type = 'AppLockerPolicy'
                CurrentValue = [PSCustomObject]@{
                    CommandAvailable = $false
                    LocalCaptured = $false
                    EffectiveCaptured = $false
                    LocalMatchesEffective = $false
                    CaptureIssues = @('unavailable')
                    LocalSnapshotAssetRelativePath = 'applocker/missing-local.xml'
                    EffectiveSnapshotAssetRelativePath = 'applocker/missing-effective.xml'
                }
            }
            [PSCustomObject]@{
                Id = 'exploit_protection.policy'; Type = 'ExploitProtectionPolicy'
                CurrentValue = [PSCustomObject]@{
                    CommandAvailable = $false
                    SnapshotAssetRelativePath = 'exploit/missing.xml'
                }
            }
            [PSCustomObject]@{
                Id = 'wdac.policies'; Type = 'WdacPolicies'
                CurrentValue = [PSCustomObject]@{
                    Captured = $false
                    CaptureIssues = @('unavailable')
                    Files = @(
                        [PSCustomObject]@{
                            RelativePath = 'CiPolicies/Active/missing.cip'
                            SnapshotAssetRelativePath = 'wdac/missing.cip'
                        }
                    )
                }
            }
        )

        { Initialize-SnapshotAssetCache -Entries $entries -SnapshotPath $snapshotPath } | Should -Not -Throw
        $script:WinDefStateSnapshotAssetCache.Count | Should -Be 0
    }

    It 'rejects AppLocker and exploit-protection sidecar traversal' {
        $snapshotPath = Join-Path $TestDrive 'traversal.json'
        $outsidePath = Join-Path $TestDrive 'outside.xml'
        [IO.File]::WriteAllText($outsidePath, '<Policy />')
        $appLockerState = [PSCustomObject]@{ LocalSnapshotAssetRelativePath = '../outside.xml' }
        $exploitState = [PSCustomObject]@{ SnapshotAssetRelativePath = '../outside.xml' }

        { Get-AppLockerPolicyXml -State $appLockerState -PolicyScope Local -SnapshotPath $snapshotPath } |
            Should -Throw '*AppLocker snapshot asset path escapes its trusted root*'
        { Get-ExploitProtectionPolicyXml -State $exploitState -SnapshotPath $snapshotPath } |
            Should -Throw '*Exploit protection snapshot asset path escapes its trusted root*'
    }
}

Describe 'Baseline completeness' {
    It 'rejects unavailable or failed mutable-provider baselines' {
        $entries = @(
            [PSCustomObject]@{ Type = 'RegistryValue'; Captured = $false }
            [PSCustomObject]@{ Type = 'RegistryKeyFlat'; Captured = $false }
            [PSCustomObject]@{ Type = 'PowerShellModuleLogging'; Captured = $false }
            [PSCustomObject]@{ Type = 'MpPreferenceValue'; CommandAvailable = $false; Captured = $false; RestoreValue = $null }
            [PSCustomObject]@{ Type = 'AsrRules'; CommandAvailable = $false; Captured = $false; CurrentValue = @(); InvalidEntries = @() }
            [PSCustomObject]@{ Type = 'ServiceConfig'; Captured = $false; CurrentValue = [PSCustomObject]@{ StartMode = $null; State = $null } }
            [PSCustomObject]@{ Type = 'LocalUser'; Captured = $false; Sid = $null; CurrentValue = $null }
            [PSCustomObject]@{ Type = 'NetBiosAdapters'; CommandAvailable = $true; Captured = $false; CurrentValue = @() }
            [PSCustomObject]@{ Type = 'AuditPolicy'; CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; Captured = $false } }
            [PSCustomObject]@{ Type = 'WdacPolicies'; CurrentValue = [PSCustomObject]@{ Captured = $false; CaptureIssues = @('failed') } }
        )

        foreach ($entry in $entries) {
            Test-SnapshotEntryCapturedExactly -Entry $entry | Should -BeFalse -Because "$($entry.Type) was not captured exactly"
        }
    }

    It 'keeps complete legacy scalar, service, and local-user entries restorable' {
        Test-SnapshotEntryCapturedExactly -Entry ([PSCustomObject]@{ Type = 'MpPreferenceValue'; RestoreValue = $false }) | Should -BeTrue
        Test-SnapshotEntryCapturedExactly -Entry ([PSCustomObject]@{ Type = 'ServiceConfig'; CurrentValue = [PSCustomObject]@{ StartMode = 'Auto'; State = 'Running' } }) | Should -BeTrue
        Test-SnapshotEntryCapturedExactly -Entry ([PSCustomObject]@{ Type = 'LocalUser'; Sid = 'S-1-5-21-1-500'; CurrentValue = $false }) | Should -BeTrue
    }

    It 'renders a single WDAC capture issue under strict mode' {
        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add('Report')
        $entry = [PSCustomObject]@{
            Id = 'wdac.policies'; Type = 'WdacPolicies'; RequiresReboot = $true
            CurrentValue = [PSCustomObject]@{ CiToolAvailable = $true; Captured = $false; CaptureIssues = @('failed'); Policies = @(); Files = @() }
        }

        { Add-SnapshotEntryReportLines -Lines $lines -Entry $entry } | Should -Not -Throw
        $lines | Should -Contain '  Capture issue count: 1'
    }
}

Describe 'Audit policy interop' {
    It 'uses the native API and resolves legacy Process Creation snapshots by GUID' {
        $engineText = Get-Content -LiteralPath $enginePath -Raw

        $engineText | Should -Not -Match '(?i)\bauditpol(?:\.exe)?\b'
        (Resolve-AuditSubcategoryGuid -Subcategory 'Process Creation').ToString() |
            Should -Be '0cce922b-69ae-11d9-bed3-505054503030'
    }
}

Describe 'Diagnostic output' {
    It 'suppresses verbose module auto-import chatter during command discovery' {
        $noisyCommandProbes = @(Get-Content -LiteralPath $enginePath | Where-Object {
            $_ -match '\bGet-Command\b' -and $_ -notmatch '-Verbose:\$false'
        })

        $noisyCommandProbes | Should -BeNullOrEmpty
    }
}

Describe 'Pre-mutation validation' {
    It 'rejects unknown, overlapping, and empty setting filters' {
        $availableIds = @('first.setting', 'second.setting')

        { Assert-ValidSettingIdFilter -AvailableId $availableIds -IncludeId 'missing.setting' } | Should -Throw '*Unknown setting ID*'
        { Assert-ValidSettingIdFilter -AvailableId $availableIds -IncludeId 'first.setting' -ExcludeId 'first.setting' } | Should -Throw '*both included and excluded*'
        { Assert-ValidSettingIdFilter -AvailableId $availableIds -ExcludeId $availableIds } | Should -Throw '*selects no settings*'
        { Assert-ValidSettingIdFilter -AvailableId $availableIds -IncludeId 'first.setting' } | Should -Not -Throw
    }

    It 'distinguishes mutable definitions from capture-only inventory' {
        $definitions = @{}
        foreach ($definition in Get-DefenseDefinitions) {
            $definitions[[string]$definition.Id] = $definition
        }

        (Test-DefinitionHasPermissiveAction -Definition $definitions['defender.runtime_status']) | Should -BeFalse
        (Test-DefinitionHasPermissiveAction -Definition $definitions['defender.exclusion_paths']) | Should -BeFalse
        (Test-DefinitionHasPermissiveAction -Definition $definitions['defender.disable_realtime_monitoring']) | Should -BeTrue
        (Test-DefinitionHasRestoreAction -Definition $definitions['defender.runtime_status']) | Should -BeFalse
        (Test-DefinitionHasRestoreAction -Definition $definitions['defender.exclusion_paths']) | Should -BeTrue
        (Test-DefinitionHasRestoreAction -Definition $definitions['defender.disable_realtime_monitoring']) | Should -BeTrue

        $runtimeCapabilities = Get-DefinitionCapabilities -Definition $definitions['defender.runtime_status']
        $runtimeCapabilities.InventoryOnly | Should -BeTrue
        $runtimeCapabilities.Permissive | Should -BeFalse
        $runtimeCapabilities.Restore | Should -BeFalse
    }

    It 'accepts a catalog-aligned snapshot and rejects an immutable target change' {
        $computerName = if (-not [string]::IsNullOrWhiteSpace([string]$env:COMPUTERNAME)) { [string]$env:COMPUTERNAME } else { [Environment]::MachineName }
        $snapshot = [PSCustomObject]@{
            SchemaVersion = 2
            Tool          = 'WinDefState'
            ComputerName  = $computerName
            Settings      = @(
                [PSCustomObject]@{
                    Id             = 'uac.enable_lua'
                    Type           = 'RegistryValue'
                    Path           = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
                    Name           = 'EnableLUA'
                    ValueKind      = 'DWord'
                    Exists         = $true
                    CurrentValue   = 1
                    RequiresReboot = $true
                }
            )
        }

        { Assert-ValidDefenseSnapshot -Snapshot $snapshot } | Should -Not -Throw
        $snapshot.Settings[0].Path = 'HKLM:\SOFTWARE\Unexpected'
        { Assert-ValidDefenseSnapshot -Snapshot $snapshot } | Should -Throw '*unexpected Path*'
    }

    It 'requires an explicit override for a different computer' {
        $snapshot = [PSCustomObject]@{
            SchemaVersion = 2
            Tool          = 'WinDefState'
            ComputerName  = 'A-DIFFERENT-COMPUTER'
            Settings      = @([PSCustomObject]@{ Id = 'defender.runtime_status'; Type = 'DefenderRuntimeStatus'; CurrentValue = $null; RequiresReboot = $false })
        }

        { Assert-ValidDefenseSnapshot -Snapshot $snapshot } | Should -Throw '*does not match this computer*'
        { Assert-ValidDefenseSnapshot -Snapshot $snapshot -AllowDifferentComputer } | Should -Not -Throw
    }
}

Describe 'Permissive target verification' {
    BeforeAll {
        $definitionsById = @{}
        foreach ($definition in Get-DefenseDefinitions) {
            $definitionsById[[string]$definition.Id] = $definition
        }
    }

    It 'distinguishes immediate verification, reboot-pending configuration, and mismatch' {
        $uacDefinition = $definitionsById['uac.enable_lua']
        $baseline = [PSCustomObject]@{
            Id = $uacDefinition.Id; Type = $uacDefinition.Type; Path = $uacDefinition.Path; Name = $uacDefinition.Name
            ValueKind = 'DWord'; Captured = $true; Exists = $true; CurrentValue = 1; RequiresReboot = $true
        }
        $live = [PSCustomObject]@{
            Id = $uacDefinition.Id; Type = $uacDefinition.Type; Path = $uacDefinition.Path; Name = $uacDefinition.Name
            ValueKind = 'DWord'; Captured = $true; Exists = $true; CurrentValue = 0; RequiresReboot = $true
        }

        $pending = Test-PermissiveDefinitionState -Definition $uacDefinition -LiveEntry $live -BaselineEntry $baseline
        $pending.Status | Should -Be 'ConfiguredPendingReboot'
        $pending.Matches | Should -BeTrue

        $immediateDefinition = $definitionsById['rdp.user_authentication']
        $immediateLive = [PSCustomObject]@{
            Id = $immediateDefinition.Id; Type = $immediateDefinition.Type; Path = $immediateDefinition.Path; Name = $immediateDefinition.Name
            ValueKind = 'DWord'; Captured = $true; Exists = $true; CurrentValue = 0; RequiresReboot = $false
        }
        (Test-PermissiveDefinitionState -Definition $immediateDefinition -LiveEntry $immediateLive -BaselineEntry $immediateLive).Status |
            Should -Be 'Verified'

        $mpDefinition = $definitionsById['defender.disable_realtime_monitoring']
        $mpLive = [PSCustomObject]@{
            Id = $mpDefinition.Id; Type = $mpDefinition.Type; Property = $mpDefinition.Property
            CommandAvailable = $true; Captured = $true; CurrentValue = $false; RestoreValue = $false; RequiresReboot = $false
        }
        (Test-PermissiveDefinitionState -Definition $mpDefinition -LiveEntry $mpLive -BaselineEntry $mpLive).Status |
            Should -Be 'Mismatch'
    }

    It 'passes the persisted snapshot path to AppLocker sidecar-backed apply' {
        (Get-Command Apply-PermissiveDefinition -CommandType Function).Definition |
            Should -Match 'Set-Permissive-AppLockerPolicy\s+-State\s+\$appLockerState\s+-SnapshotPath\s+\$SnapshotPath'
        (Get-Command Set-DefensePermissive -CommandType Function).Definition |
            Should -Match 'Invoke-PermissiveMutationWorkItem\s+-WorkItem\s+\$workItem\s+-SnapshotPath\s+\$export\.JsonPath'
        (Get-Command Invoke-PermissiveMutationWorkItem -CommandType Function).Definition |
            Should -Match 'Apply-PermissiveDefinition[^\r\n]+-SnapshotPath\s+\$SnapshotPath'
    }

    It 'never mutates a definition missing from the persisted baseline' {
        (Get-Command Set-DefensePermissive -CommandType Function).Definition |
            Should -Match 'if\s*\(\$null\s+-eq\s+\$entry\)\s*\{\s*throw\s+"The persisted baseline is missing setting'
    }

    It 'verifies the explicit exploit-protection system target' {
        $xml = '<MitigationPolicy><SystemConfig><DEP Enable="false" EmulateAtlThunks="false"/><ControlFlowGuard Enable="false"/><ASLR ForceRelocateImages="false" BottomUp="false" HighEntropy="false"/><SEHOP Enable="false"/></SystemConfig></MitigationPolicy>'
        $live = [PSCustomObject]@{
            CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; Xml = $xml }
        }

        (Get-PermissiveExploitProtectionEvaluation -LiveEntry $live).Matches | Should -BeTrue
    }

    It 'preserves platform-managed WDAC policies and fails closed without CiTool for removable policies' {
        $platformId = '0283ac0f-fff1-49ae-ada1-8a933130cad6'
        $platformState = [PSCustomObject]@{
            CiToolAvailable = $false
            Policies = @([PSCustomObject]@{
                PolicyID = $platformId; FriendlyName = 'VerifiedAndReputableDesktop'
                HasFileOnDisk = $true; IsCurrentlyEnforced = $true
            })
            Files = @()
        }
        { Remove-WdacPolicies -State $platformState } | Should -Not -Throw

        $removableState = [PSCustomObject]@{
            CiToolAvailable = $false
            Policies = @([PSCustomObject]@{
                PolicyID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'; FriendlyName = 'Test policy'
                HasFileOnDisk = $true; IsCurrentlyEnforced = $true
            })
            Files = @()
        }
        { Remove-WdacPolicies -State $removableState } | Should -Throw '*CiTool is unavailable*'

        $unclassifiedFileState = [PSCustomObject]@{
            CiToolAvailable = $false
            Policies = @()
            Files = @([PSCustomObject]@{ FileName = 'SiPolicy.p7b'; RelativePath = 'SiPolicy.p7b' })
        }
        { Remove-WdacPolicies -State $unclassifiedFileState } | Should -Throw '*No raw policy files were deleted*'
        (Get-Command Remove-WdacPolicies -CommandType Function).Definition | Should -Not -Match 'Remove-WdacPolicyFiles'
    }

    It 'reconciles WDAC restore by policy identity without removing baseline or platform policy' {
        $baselineId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        $extraId = '11111111-2222-3333-4444-555555555555'
        $platformId = '0283ac0f-fff1-49ae-ada1-8a933130cad6'
        $baseline = [PSCustomObject]@{
            Policies = @([PSCustomObject]@{ PolicyID = $baselineId })
            Files = @([PSCustomObject]@{ FileName = "{$baselineId}.cip"; RelativePath = "CiPolicies\Active\{$baselineId}.cip" })
        }
        $live = [PSCustomObject]@{
            CiToolAvailable = $true
            Policies = @(
                [PSCustomObject]@{ PolicyID = $baselineId; FriendlyName = 'Baseline'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true }
                [PSCustomObject]@{ PolicyID = $extraId; FriendlyName = 'Extra'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true }
                [PSCustomObject]@{ PolicyID = $platformId; FriendlyName = 'VerifiedAndReputableDesktop'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true }
            )
            Files = @(
                [PSCustomObject]@{ FileName = "{$baselineId}.cip"; RelativePath = "CiPolicies\Active\{$baselineId}.cip" }
                [PSCustomObject]@{ FileName = "{$extraId}.cip"; RelativePath = "CiPolicies\Active\{$extraId}.cip" }
                [PSCustomObject]@{ FileName = "{$platformId}.cip"; RelativePath = "CiPolicies\Active\{$platformId}.cip" }
            )
        }

        $removal = Get-WdacRestoreRemovalState -BaselineState $baseline -LiveState $live

        @($removal.Policies).Count | Should -Be 1
        (Get-WdacNormalizedPolicyId -Value $removal.Policies[0].PolicyID) | Should -Be $extraId
        @($removal.Files).Count | Should -Be 1
        (Get-WdacPolicyIdFromFileName -FileName $removal.Files[0].FileName) | Should -Be $extraId
        (Get-Command Restore-WdacPolicies -CommandType Function).Definition | Should -Not -Match 'Remove-WdacPolicyFiles|Restore-WdacPolicyFiles'
    }

    It 'restricts WDAC restore files to supported Code Integrity destinations' {
        Mock Get-WdacCodeIntegrityRoot { $TestDrive }
        $policyId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        $valid = [PSCustomObject]@{
            FileName = "{$policyId}.cip"
            RelativePath = "CiPolicies\Active\{$policyId}.cip"
        }

        (Split-Path -Leaf (Get-WdacPolicyDestinationPath -File $valid)) | Should -Be "{$policyId}.cip"
        { Get-WdacPolicyDestinationPath -File ([PSCustomObject]@{ FileName = 'escape.cip'; RelativePath = '..\escape.cip' }) } |
            Should -Throw '*unsupported policy path*'
        { Get-WdacPolicyDestinationPath -File ([PSCustomObject]@{ FileName = 'other.cip'; RelativePath = "CiPolicies\Active\{$policyId}.cip" }) } |
            Should -Throw '*does not match*'
    }

    It 'compares deterministic custom WDAC state without treating provider or platform state as restorable' {
        $customId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        $platformId = '0283ac0f-fff1-49ae-ada1-8a933130cad6'
        $customPolicy = [PSCustomObject]@{ PolicyID = $customId; FriendlyName = 'Custom'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true }
        $baselineState = [PSCustomObject]@{
            CiToolAvailable = $false; Captured = $true; CaptureIssues = @()
            Policies = @(
                [PSCustomObject]@{ PolicyID = $platformId; FriendlyName = 'VerifiedAndReputableDesktop'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true }
                $customPolicy
            )
            Files = @(
                [PSCustomObject]@{ FileName = "{$platformId}.cip"; RelativePath = "CiPolicies\Active\{$platformId}.cip"; Sha256 = 'PLATFORM-OLD' }
                [PSCustomObject]@{ FileName = "{$customId}.cip"; RelativePath = "CiPolicies\Active\{$customId}.cip"; Sha256 = 'CUSTOM' }
            )
        }
        $liveState = [PSCustomObject]@{
            CiToolAvailable = $true; Captured = $true; CaptureIssues = @()
            Policies = @(
                $customPolicy
                [PSCustomObject]@{ PolicyID = $platformId; FriendlyName = 'VerifiedAndReputableDesktop updated'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true }
            )
            Files = @(
                [PSCustomObject]@{ FileName = "{$customId}.cip"; RelativePath = "CiPolicies\Active\{$customId}.cip"; Sha256 = 'CUSTOM' }
                [PSCustomObject]@{ FileName = "{$platformId}.cip"; RelativePath = "CiPolicies\Active\{$platformId}.cip"; Sha256 = 'PLATFORM-NEW' }
            )
        }
        $baselineEntry = [PSCustomObject]@{ Id = 'wdac.policies'; Type = 'WdacPolicies'; CurrentValue = $baselineState; RequiresReboot = $true }
        $liveEntry = [PSCustomObject]@{ Id = 'wdac.policies'; Type = 'WdacPolicies'; CurrentValue = $liveState; RequiresReboot = $true }

        (Get-CanonicalJson -Value (ConvertTo-ComparableSnapshotEntry -Entry $baselineEntry)) |
            Should -Be (Get-CanonicalJson -Value (ConvertTo-ComparableSnapshotEntry -Entry $liveEntry -ReferenceEntry $baselineEntry))
    }

    It 'renders a durable permissive verification report' {
        $verification = [PSCustomObject]@{
            ComputerName = 'HOST'; VerifiedAtUtc = '2026-01-01T00:00:00Z'
            BaselineProducer = [PSCustomObject]@{ ScriptFileName = 'WinDefState.ps1'; ScriptSha256 = 'BASELINE'; PowerShellVersion = '5.1'; PowerShellEdition = 'Desktop'; ProcessArchitecture = 'AMD64' }
            Verifier = [PSCustomObject]@{ ScriptFileName = 'WinDefState.ps1'; ScriptSha256 = 'VERIFIER'; PowerShellVersion = '5.1'; PowerShellEdition = 'Desktop'; ProcessArchitecture = 'AMD64' }
            VerifiedCount = 1; PendingRebootCount = 1; MismatchCount = 0
            MutationMetrics = [PSCustomObject]@{
                DurationMs = 1250; ProviderQueryCount = 2; CacheHitCount = 3; ProviderQueries = @()
                Settings = @([PSCustomObject]@{ Id = 'rdp.user_authentication'; DurationMs = 42.5 })
            }
            Results = @(
                [PSCustomObject]@{ Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Status = 'Verified'; Changed = $true; RequiresReboot = $false; Reason = $null; Expected = 0; Actual = 0 }
                [PSCustomObject]@{ Id = 'uac.enable_lua'; Type = 'RegistryValue'; Status = 'ConfiguredPendingReboot'; Changed = $true; RequiresReboot = $true; Reason = 'Reboot required.'; Expected = 0; Actual = 0 }
            )
        }

        $lines = @(Get-PermissiveVerificationReportLines -Verification $verification -SnapshotPath 'C:\state\snapshot.json')
        $lines | Should -Contain 'Configured, pending reboot: 1'
        $lines | Should -Contain 'Baseline producer script SHA-256: BASELINE'
        $lines | Should -Contain 'Verifier script SHA-256: VERIFIER'
        $lines | Should -Contain 'Mutation duration: 1.25 seconds'
        $lines | Should -Contain 'Slowest mutations:'
        $lines | Should -Contain '  - rdp.user_authentication: 42.5 ms'
        $lines | Should -Contain '[uac.enable_lua] RegistryValue'
    }

    It 'accepts a representative permissive target for every mutable provider type' {
        $emptyAppLockerXml = Get-EmptyAppLockerPolicyXml
        $exploitXml = '<MitigationPolicy><SystemConfig><DEP Enable="false" EmulateAtlThunks="false"/><ControlFlowGuard Enable="false"/><ASLR ForceRelocateImages="false" BottomUp="false" HighEntropy="false"/><SEHOP Enable="false"/></SystemConfig></MitigationPolicy>'
        $firewallProfiles = @(
            foreach ($profile in @('Domain', 'Private', 'Public')) {
                [PSCustomObject]@{
                    Profile = $profile; Enabled = 'False'; DefaultInboundAction = 'Allow'; DefaultOutboundAction = 'Allow'
                    AllowUnicastResponseToMulticast = 'True'; NotifyOnListen = 'False'; LogAllowed = 'False'; LogBlocked = 'False'; LogIgnored = 'False'
                }
            }
        )
        $listener = [PSCustomObject]@{ Address = '*'; Transport = 'HTTP'; Port = 5985; Hostname = ''; Enabled = $true; URLPrefix = 'wsman'; CertificateThumbprint = '' }
        $userItem = $definitionsById['wpad.user_auto_detect'].Items[0]

        $cases = @(
            [PSCustomObject]@{ Id = 'rdp.user_authentication'; Live = [PSCustomObject]@{ Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Path = $definitionsById['rdp.user_authentication'].Path; Name = 'UserAuthentication'; ValueKind = 'DWord'; Captured = $true; Exists = $true; CurrentValue = 0; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'powershell.script_block_logging'; Live = [PSCustomObject]@{ Id = 'powershell.script_block_logging'; Type = 'RegistryKeyFlat'; Path = $definitionsById['powershell.script_block_logging'].Path; Captured = $true; Exists = $false; CurrentValue = @(); RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'defender.disable_realtime_monitoring'; Live = [PSCustomObject]@{ Id = 'defender.disable_realtime_monitoring'; Type = 'MpPreferenceValue'; Property = 'DisableRealtimeMonitoring'; CommandAvailable = $true; Captured = $true; CurrentValue = $true; RestoreValue = $true; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'defender.asr_rules'; Live = [PSCustomObject]@{ Id = 'defender.asr_rules'; Type = 'AsrRules'; CommandAvailable = $true; Captured = $true; CurrentValue = @(); InvalidEntries = @(); RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'powershell.module_logging'; Live = [PSCustomObject]@{ Id = 'powershell.module_logging'; Type = 'PowerShellModuleLogging'; BasePath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ModuleLogging'; Captured = $true; Exists = $false; CurrentValue = [PSCustomObject]@{ BaseValues = @(); ModuleNamesExists = $false; ModuleNamesValues = @() }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'powershell.lockdown'; Live = [PSCustomObject]@{ Id = 'powershell.lockdown'; Type = 'MachineEnvironmentValue'; Name = '__PSLockdownPolicy'; Exists = $true; CurrentValue = '0'; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'applocker.service'; Live = [PSCustomObject]@{ Id = 'applocker.service'; Type = 'ServiceConfig'; Name = 'AppIDSvc'; Captured = $true; CurrentValue = [PSCustomObject]@{ StartMode = 'Manual'; State = 'Stopped' }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'localuser.administrator'; Live = [PSCustomObject]@{ Id = 'localuser.administrator'; Type = 'LocalUser'; Name = 'Administrator'; Sid = 'S-1-5-21-1-500'; Rid = 500; Captured = $true; CurrentValue = $true; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'network.netbios_adapters'; Live = [PSCustomObject]@{ Id = 'network.netbios_adapters'; Type = 'NetBiosAdapters'; CommandAvailable = $true; Captured = $true; CurrentValue = @([PSCustomObject]@{ Index = 1; TcpipNetbiosOptions = 1 }); RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'wpad.user_auto_detect'; Live = [PSCustomObject]@{ Id = 'wpad.user_auto_detect'; Type = 'LoadedUserRegistryValues'; CurrentValue = [PSCustomObject]@{ Entries = @([PSCustomObject]@{ Sid = 'S-1-5-21-1-1000'; RelativePath = $userItem.RelativePath; Name = $userItem.Name; Exists = $true; CurrentValue = 1; ValueKind = 'DWord' }); CaptureIssues = @() }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'audit.process_creation'; Live = [PSCustomObject]@{ Id = 'audit.process_creation'; Type = 'AuditPolicy'; Subcategory = 'Process Creation'; CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; Captured = $true; Success = $false; Failure = $false }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'winrm.service.allow_unencrypted'; Live = [PSCustomObject]@{ Id = 'winrm.service.allow_unencrypted'; Type = 'WsManValue'; Path = $definitionsById['winrm.service.allow_unencrypted'].Path; CommandAvailable = $true; Captured = $true; CurrentValue = $true; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'winrm.listeners'; Live = [PSCustomObject]@{ Id = 'winrm.listeners'; Type = 'WinRmListeners'; CommandAvailable = $true; Captured = $true; CurrentValue = @($listener); RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'smb.client.require_security_signature'; Live = [PSCustomObject]@{ Id = 'smb.client.require_security_signature'; Type = 'SmbClientConfig'; CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; TimedOut = $false; RequireSecuritySignature = $false }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'smb.server.require_security_signature'; Live = [PSCustomObject]@{ Id = 'smb.server.require_security_signature'; Type = 'SmbServerConfig'; CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; TimedOut = $false; RequireSecuritySignature = $false }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'firewall.profiles'; Live = [PSCustomObject]@{ Id = 'firewall.profiles'; Type = 'FirewallProfiles'; CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; CaptureIssues = @(); Profiles = @($firewallProfiles) }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'rdp.firewall_rules'; Live = [PSCustomObject]@{ Id = 'rdp.firewall_rules'; Type = 'FirewallRules'; CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; Group = '@FirewallAPI.dll,-28752'; CaptureIssues = @(); Rules = @([PSCustomObject]@{ Name = 'RemoteDesktop-Test'; Enabled = 'True'; Group = '@FirewallAPI.dll,-28752'; Direction = 'Inbound'; Action = 'Allow'; Profile = 'Any' }) }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'bitlocker.volumes'; Live = [PSCustomObject]@{ Id = 'bitlocker.volumes'; Type = 'BitLockerVolumes'; CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; TimedOutMountPoints = @(); CaptureIssues = @(); Volumes = @([PSCustomObject]@{ MountPoint = 'C:'; VolumeType = 'OperatingSystem'; ProtectionStatus = 'Off'; ProtectionMode = 'Suspended'; VolumeStatus = 'FullyEncrypted'; LockStatus = 'Unlocked'; EncryptionMethod = 'XtsAes256'; EncryptionPercentage = 100; KeyProtectorCount = 1; KeyProtectors = @([PSCustomObject]@{ KeyProtectorId = 'test'; KeyProtectorType = 'Tpm' }); AutoUnlockEnabled = $null }) }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'applocker.policy'; Live = [PSCustomObject]@{ Id = 'applocker.policy'; Type = 'AppLockerPolicy'; CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; LocalCaptured = $true; EffectiveCaptured = $true; LocalMatchesEffective = $true; CaptureIssues = @(); CollectionSummaries = @(); LocalXml = $emptyAppLockerXml; EffectiveXml = $emptyAppLockerXml }; RequiresReboot = $false } }
            [PSCustomObject]@{ Id = 'exploit_protection.policy'; Live = [PSCustomObject]@{ Id = 'exploit_protection.policy'; Type = 'ExploitProtectionPolicy'; CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; Xml = $exploitXml }; RequiresReboot = $true } }
            [PSCustomObject]@{ Id = 'wdac.policies'; Live = [PSCustomObject]@{ Id = 'wdac.policies'; Type = 'WdacPolicies'; CurrentValue = [PSCustomObject]@{ CiToolAvailable = $true; Captured = $true; CaptureIssues = @(); Policies = @(); Files = @() }; RequiresReboot = $true } }
        )

        @($cases.Live.Type | Sort-Object -Unique).Count | Should -Be 21
        foreach ($case in $cases) {
            $result = Test-PermissiveDefinitionState -Definition $definitionsById[$case.Id] -LiveEntry $case.Live -BaselineEntry $case.Live
            $result.Matches | Should -BeTrue -Because "$($case.Live.Type) should accept its representative target"
        }
    }
}

Describe 'Restore capability verification' {
    It 'classifies an incomplete sidecar-backed baseline before canonicalization or provider capture' {
        $snapshot = [PSCustomObject]@{
            SchemaVersion = 2
            Tool = 'WinDefState'
            ComputerName = 'TEST'
            Settings = @(
                [PSCustomObject]@{
                    Id = 'applocker.policy'
                    Type = 'AppLockerPolicy'
                    RequiresReboot = $false
                    CurrentValue = $null
                }
            )
        }

        $verification = Test-DefenseSnapshot -Snapshot $snapshot -SnapshotPath (Join-Path $TestDrive 'incomplete-applocker.json')

        $verification.MismatchCount | Should -Be 0
        $verification.IncompleteCount | Should -Be 1
        $verification.InventoryCount | Should -Be 0
        $verification.SkippedCount | Should -Be 1
        $verification.Results[0].SkipCategory | Should -Be 'IncompleteBaseline'
        $verification.Results[0].Expected | Should -BeNullOrEmpty
        @($verification.CaptureMetrics.Settings).Count | Should -Be 0
    }

    It 'does not mutate or compare inventory-only runtime state as a restore target' {
        $snapshot = [PSCustomObject]@{
            SchemaVersion = 2
            Tool = 'WinDefState'
            ComputerName = 'TEST'
            Settings = @(
                [PSCustomObject]@{
                    Id = 'defender.runtime_status'
                    Type = 'DefenderRuntimeStatus'
                    RequiresReboot = $false
                    CurrentValue = [PSCustomObject]@{
                        CommandAvailable = $true
                        Captured = $true
                        AMRunningMode = 'Normal'
                        RealTimeProtectionEnabled = $true
                        AntivirusEnabled = $true
                        IsTamperProtected = $false
                    }
                }
            )
        }

        $verification = Test-DefenseSnapshot -Snapshot $snapshot -SnapshotPath (Join-Path $TestDrive 'runtime.json')

        $verification.MismatchCount | Should -Be 0
        $verification.IncompleteCount | Should -Be 0
        $verification.InventoryCount | Should -Be 1
        $verification.SkippedCount | Should -Be 1
        $verification.Results[0].SkipCategory | Should -Be 'InventoryOnly'
        @($verification.CaptureMetrics.Settings).Count | Should -Be 0
        (Get-RestoreCompletionMessage -SnapshotPath 'C:\state\runtime.json' -Verification $verification) |
            Should -Be 'Restore completed and verified all restorable settings from snapshot: C:\state\runtime.json'
    }
}

Describe 'Permissive orchestration' {
    It 'persists a baseline, times the mutation, verifies live state, and journals success' {
        $script:StateRoot = Join-Path $TestDrive 'permissive-state'
        $script:orchestrationSnapshotPath = Join-Path $TestDrive 'baseline.json'
        $script:orchestrationDefinition = [PSCustomObject]@{
            Id = 'test.registry'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Test'; Name = 'Enabled'
            ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false
        }
        $script:orchestrationEntry = [PSCustomObject]@{
            Id = 'test.registry'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Test'; Name = 'Enabled'
            ValueKind = 'DWord'; Captured = $true; Exists = $true; CurrentValue = 1; RequiresReboot = $false
        }
        Mock Assert-NoActiveOperation {}
        Mock Get-SelectedDefenseDefinitions { @($script:orchestrationDefinition) }
        Mock Export-DefenseSnapshot {
            [PSCustomObject]@{
                JsonPath = $script:orchestrationSnapshotPath
                ReportPath = ($script:orchestrationSnapshotPath + '.txt')
                Snapshot = [PSCustomObject]@{ Settings = @($script:orchestrationEntry) }
            }
        }
        Mock Write-OperationState { [PSCustomObject]@{ Status = 'Applying'; SnapshotPath = $script:orchestrationSnapshotPath } }
        Mock Test-SnapshotEntryCapturedExactly { $true }
        Mock Apply-PermissiveDefinition {}
        Mock Test-DefensePermissiveState {
            [PSCustomObject]@{
                ComputerName = 'HOST'; VerifiedAtUtc = '2026-01-01T00:00:00Z'
                VerifiedCount = 1; PendingRebootCount = 0; MismatchCount = 0
                CaptureMetrics = $null
                Results = @([PSCustomObject]@{ Id = 'test.registry'; Status = 'Verified' })
            }
        }
        Mock Get-PermissiveVerificationReportPath { Join-Path $TestDrive 'permissive-check.txt' }
        Mock Get-PermissiveVerificationReportLines { @('report') }
        Mock Write-TextAtomic {}
        Mock Update-OperationStateStatus {
            $Operation.Status = $Status
            $Operation
        }
        Mock Write-Host {}
        Mock Write-Warning {}

        Set-DefensePermissive -Path $script:orchestrationSnapshotPath

        Should -Invoke Export-DefenseSnapshot -Times 1 -Exactly
        Should -Invoke Apply-PermissiveDefinition -Times 1 -Exactly
        Should -Invoke Test-DefensePermissiveState -Times 1 -Exactly
        Should -Invoke Get-PermissiveVerificationReportLines -Times 1 -Exactly -ParameterFilter {
            $Verification.PSObject.Properties['MutationMetrics'] -and @($Verification.MutationMetrics.Settings).Count -eq 1
        }
        Should -Invoke Update-OperationStateStatus -Times 1 -Exactly -ParameterFilter { $Status -eq 'AppliedVerified' }
    }

    It 'dispatches multiple Defender scalar settings as one provider mutation' {
        $script:StateRoot = Join-Path $TestDrive 'permissive-defender-batch-state'
        $script:orchestrationSnapshotPath = Join-Path $TestDrive 'defender-batch.json'
        $script:orchestrationDefinitions = @(
            Get-DefenseDefinitions | Where-Object { [string]$_.Id -in @('defender.disable_realtime_monitoring', 'defender.disable_behavior_monitoring') }
        )
        $script:orchestrationEntries = @(
            foreach ($definition in $script:orchestrationDefinitions) {
                [PSCustomObject]@{
                    Id = [string]$definition.Id; Type = 'MpPreferenceValue'; Property = [string]$definition.Property
                    CommandAvailable = $true; Captured = $true; CurrentValue = $false; RestoreValue = $false; RequiresReboot = $false
                }
            }
        )
        Mock Assert-NoActiveOperation {}
        Mock Get-SelectedDefenseDefinitions { @($script:orchestrationDefinitions) }
        Mock Export-DefenseSnapshot {
            [PSCustomObject]@{
                JsonPath = $script:orchestrationSnapshotPath
                ReportPath = ($script:orchestrationSnapshotPath + '.txt')
                Snapshot = [PSCustomObject]@{ Settings = @($script:orchestrationEntries) }
            }
        }
        Mock Write-OperationState { [PSCustomObject]@{ Status = 'Applying'; SnapshotPath = $script:orchestrationSnapshotPath } }
        Mock Test-SnapshotEntryCapturedExactly { $true }
        Mock Set-MpPreferencePropertyValues {}
        Mock Apply-PermissiveDefinition {}
        Mock Test-DefensePermissiveState {
            [PSCustomObject]@{
                ComputerName = 'HOST'; VerifiedAtUtc = '2026-01-01T00:00:00Z'
                VerifiedCount = 2; PendingRebootCount = 0; MismatchCount = 0; CaptureMetrics = $null
                Results = @(
                    foreach ($definition in $script:orchestrationDefinitions) {
                        [PSCustomObject]@{ Id = [string]$definition.Id; Status = 'Verified' }
                    }
                )
            }
        }
        Mock Get-PermissiveVerificationReportPath { Join-Path $TestDrive 'defender-batch-check.txt' }
        Mock Get-PermissiveVerificationReportLines { @('report') }
        Mock Write-TextAtomic {}
        Mock Update-OperationStateStatus { $Operation.Status = $Status; $Operation }
        Mock Write-Host {}
        Mock Write-Warning {}

        Set-DefensePermissive -Path $script:orchestrationSnapshotPath

        Should -Invoke Set-MpPreferencePropertyValues -Times 1 -Exactly -ParameterFilter { @($Items).Count -eq 2 }
        Should -Invoke Apply-PermissiveDefinition -Times 0 -Exactly
        Should -Invoke Test-DefensePermissiveState -Times 1 -Exactly -ParameterFilter { @($Definitions).Count -eq 2 }
    }
}

Describe 'Simulated lifecycle integration' {
    It 'explains how to restore an explicit snapshot when no active journal exists' {
        $StateRoot = Join-Path $TestDrive 'no-active-operation-state'

        { Restore-DefenseSnapshot } |
            Should -Throw '*No active permissive operation was found*successful restore clears current-operation.json*-SnapshotPath*'
    }

    It 'rejects a missing complete-baseline sidecar before any restore mutation' {
        $StateRoot = Join-Path $TestDrive 'missing-sidecar-state'
        $snapshotPath = Join-Path $StateRoot 'snapshots/missing-sidecar.json'
        $snapshot = [PSCustomObject]@{
            SchemaVersion = 2
            Tool = 'WinDefState'
            ComputerName = $env:COMPUTERNAME
            CapturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
            Settings = @(
                [PSCustomObject]@{
                    Id = 'applocker.policy'
                    Type = 'AppLockerPolicy'
                    RequiresReboot = $false
                    CurrentValue = [PSCustomObject]@{
                        CommandAvailable = $true
                        LocalCaptured = $true
                        EffectiveCaptured = $true
                        LocalMatchesEffective = $true
                        CaptureIssues = @()
                        LocalSnapshotAssetRelativePath = 'applocker/missing-local.xml'
                        EffectiveSnapshotAssetRelativePath = 'applocker/missing-effective.xml'
                    }
                }
            )
        }
        Write-SnapshotJsonAtomic -Path $snapshotPath -Snapshot $snapshot
        Mock Restore-SnapshotEntry {}

        { Restore-DefenseSnapshot -Path $snapshotPath } | Should -Throw '*AppLocker snapshot asset is missing*'
        Should -Invoke Restore-SnapshotEntry -Times 0 -Exactly
    }

    It 'round trips a real snapshot, permissive verification, journal, and restore verification' {
        $StateRoot = Join-Path $TestDrive 'lifecycle-state'
        $snapshotPath = Join-Path $StateRoot 'snapshots/lifecycle.json'
        $script:lifecycleValue = 1
        $script:lifecycleDefinition = @(Get-DefenseDefinitions | Where-Object { $_.Id -eq 'rdp.user_authentication' })[0]

        Mock Get-SelectedDefenseDefinitions { @($script:lifecycleDefinition) }
        Mock Capture-Definition {
            param($Definition, $CaptureSession)

            [PSCustomObject]@{
                Id             = [string]$Definition.Id
                Type           = 'RegistryValue'
                Path           = [string]$Definition.Path
                Name           = [string]$Definition.Name
                ValueKind      = [string]$Definition.ValueKind
                Captured       = $true
                CaptureError   = $null
                Exists         = $true
                CurrentValue   = $script:lifecycleValue
                RequiresReboot = [bool]$Definition.RequiresReboot
            }
        }
        Mock Apply-PermissiveDefinition {
            param($Definition, $Entry, $SnapshotPath, $CaptureSession)

            $script:lifecycleValue = [int]$Definition.PermissiveValue
        }
        Mock Restore-SnapshotEntry {
            param($Entry, $SnapshotPath, $CaptureSession)

            $script:lifecycleValue = [int]$Entry.CurrentValue
        }
        Mock Write-Host {}
        Mock Write-Warning {}

        Set-DefensePermissive -Path $snapshotPath -IncludeId 'rdp.*'

        $script:lifecycleValue | Should -Be 0
        Test-Path -LiteralPath $snapshotPath -PathType Leaf | Should -BeTrue
        Test-Path -LiteralPath ([IO.Path]::ChangeExtension($snapshotPath, 'txt')) -PathType Leaf | Should -BeTrue
        $snapshot = Read-JsonFile -Path $snapshotPath
        $snapshot.CaptureMetrics.Phase | Should -Be 'Snapshot'
        $snapshot.CaptureScope.IsFiltered | Should -BeTrue
        @($snapshot.CaptureScope.IncludeId) | Should -Contain 'rdp.*'
        @($snapshot.CaptureMetrics.Settings).Count | Should -Be 1

        $operation = Get-OperationState -Root $StateRoot
        $operation.Status | Should -Be 'AppliedVerified'
        @($operation.IncludeId) | Should -Contain 'rdp.*'
        $operation.PermissiveVerification.VerifiedCount | Should -Be 1
        { Assert-OperationSnapshotIntegrity -Operation $operation } | Should -Not -Throw
        @(Get-ChildItem -LiteralPath (Join-Path $StateRoot 'verification') -Filter '*-permissive-check-*.txt').Count | Should -Be 1

        Restore-DefenseSnapshot

        $script:lifecycleValue | Should -Be 1
        Test-Path -LiteralPath (Get-OperationPath -Root $StateRoot) | Should -BeFalse
        $restoreReports = @(
            Get-ChildItem -LiteralPath (Join-Path $StateRoot 'verification') -Filter '*-restore-check-*.txt' |
                Where-Object { $_.Name -notlike '*-wdac.txt' }
        )
        $restoreReports.Count | Should -Be 1
        (Get-Content -LiteralPath $restoreReports[0].FullName -Raw) | Should -Match 'Matched settings: 1'

        Should -Invoke Capture-Definition -Times 3 -Exactly
        Should -Invoke Apply-PermissiveDefinition -Times 1 -Exactly
        Should -Invoke Restore-SnapshotEntry -Times 1 -Exactly
    }

    It 'records a restore verification exception instead of leaving the journal at Restoring' {
        $StateRoot = Join-Path $TestDrive 'restore-verification-failure-state'
        $snapshotPath = Join-Path $StateRoot 'snapshots/baseline.json'
        $definition = @(Get-DefenseDefinitions | Where-Object { $_.Id -eq 'rdp.user_authentication' })[0]
        $snapshot = [PSCustomObject]@{
            SchemaVersion = 2
            Tool           = 'WinDefState'
            ComputerName   = $env:COMPUTERNAME
            CapturedAtUtc  = (Get-Date).ToUniversalTime().ToString('o')
            Settings       = @([PSCustomObject]@{
                Id = $definition.Id; Type = $definition.Type; Path = $definition.Path; Name = $definition.Name
                ValueKind = $definition.ValueKind; Captured = $true; Exists = $true; CurrentValue = 1; RequiresReboot = $false
            })
        }
        Write-SnapshotJsonAtomic -Path $snapshotPath -Snapshot $snapshot
        $null = Write-OperationState -Root $StateRoot -SnapshotPath $snapshotPath -Mode Permissive
        Mock Restore-SnapshotEntry {}
        Mock Test-DefenseSnapshot { throw 'expected verification exception' }

        { Restore-DefenseSnapshot } | Should -Throw '*post-restore verification or report persistence could not finish*'
        (Get-OperationState -Root $StateRoot).Status | Should -Be 'RestoreVerificationFailed'
    }

    It 'preserves the persisted baseline and records an apply mutation failure' {
        $StateRoot = Join-Path $TestDrive 'apply-failure-state'
        $snapshotPath = Join-Path $StateRoot 'snapshots/baseline.json'
        $script:lifecycleValue = 1
        $script:lifecycleDefinition = @(Get-DefenseDefinitions | Where-Object { $_.Id -eq 'rdp.user_authentication' })[0]
        Mock Get-SelectedDefenseDefinitions { @($script:lifecycleDefinition) }
        Mock Capture-Definition {
            param($Definition, $CaptureSession)

            [PSCustomObject]@{
                Id = $Definition.Id; Type = $Definition.Type; Path = $Definition.Path; Name = $Definition.Name
                ValueKind = $Definition.ValueKind; Captured = $true; Exists = $true; CurrentValue = $script:lifecycleValue; RequiresReboot = $false
            }
        }
        Mock Apply-PermissiveDefinition { throw 'expected apply failure' }

        { Set-DefensePermissive -Path $snapshotPath } | Should -Throw '*Permissive apply failed*expected apply failure*'

        Test-Path -LiteralPath $snapshotPath -PathType Leaf | Should -BeTrue
        $operation = Get-OperationState -Root $StateRoot
        $operation.Status | Should -Be 'ApplyFailed'
        { Assert-OperationSnapshotIntegrity -Operation $operation } | Should -Not -Throw
    }

    It 'records a rejected permissive target as a verification failure' {
        $StateRoot = Join-Path $TestDrive 'apply-mismatch-state'
        $snapshotPath = Join-Path $StateRoot 'snapshots/baseline.json'
        $script:lifecycleValue = 1
        $script:lifecycleDefinition = @(Get-DefenseDefinitions | Where-Object { $_.Id -eq 'rdp.user_authentication' })[0]
        Mock Get-SelectedDefenseDefinitions { @($script:lifecycleDefinition) }
        Mock Capture-Definition {
            param($Definition, $CaptureSession)

            [PSCustomObject]@{
                Id = $Definition.Id; Type = $Definition.Type; Path = $Definition.Path; Name = $Definition.Name
                ValueKind = $Definition.ValueKind; Captured = $true; Exists = $true; CurrentValue = $script:lifecycleValue; RequiresReboot = $false
            }
        }
        Mock Apply-PermissiveDefinition {}
        Mock Write-Warning {}

        { Set-DefensePermissive -Path $snapshotPath } | Should -Throw '*Permissive verification failed*'

        $operation = Get-OperationState -Root $StateRoot
        $operation.Status | Should -Be 'ApplyVerificationFailed'
        $operation.PermissiveVerification.MismatchCount | Should -Be 1
        Test-Path -LiteralPath $operation.PermissiveVerification.ReportPath -PathType Leaf | Should -BeTrue
        { Assert-OperationSnapshotIntegrity -Operation $operation } | Should -Not -Throw
    }

    It 'records a restore mutation failure and keeps the active operation journal' {
        $StateRoot = Join-Path $TestDrive 'restore-failure-state'
        $snapshotPath = Join-Path $StateRoot 'snapshots/baseline.json'
        $definition = @(Get-DefenseDefinitions | Where-Object { $_.Id -eq 'rdp.user_authentication' })[0]
        $snapshot = [PSCustomObject]@{
            SchemaVersion = 2; Tool = 'WinDefState'; ComputerName = $env:COMPUTERNAME
            CapturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
            Settings = @([PSCustomObject]@{
                Id = $definition.Id; Type = $definition.Type; Path = $definition.Path; Name = $definition.Name
                ValueKind = $definition.ValueKind; Captured = $true; Exists = $true; CurrentValue = 1; RequiresReboot = $false
            })
        }
        Write-SnapshotJsonAtomic -Path $snapshotPath -Snapshot $snapshot
        $null = Write-OperationState -Root $StateRoot -SnapshotPath $snapshotPath -Mode Permissive
        Mock Restore-SnapshotEntry { throw 'expected restore failure' }

        { Restore-DefenseSnapshot } | Should -Throw '*Restore failed while applying snapshot*expected restore failure*'
        (Get-OperationState -Root $StateRoot).Status | Should -Be 'RestoreFailed'
    }
}

Describe 'Capture session cache' {
    It 'accepts every orchestration phase including post-permissive verification' {
        foreach ($phase in @('Snapshot', 'Permissive', 'Restore', 'Verification', 'PermissiveVerification')) {
            $session = New-CaptureSession -Phase $phase
            $session.Phase | Should -Be $phase
            $null = Complete-CaptureSession -Session $session
        }
    }

    It 'evaluates a provider factory once per cache key' {
        $session = New-CaptureSession -Phase Snapshot
        $script:factoryCallCount = 0

        $first = Get-CaptureSessionValue -Session $session -Key 'sample.provider' -Factory {
            $script:factoryCallCount++
            [PSCustomObject]@{ Value = 42 }
        }
        $second = Get-CaptureSessionValue -Session $session -Key 'sample.provider' -Factory {
            throw 'The cached provider factory should not run twice.'
        }
        $metrics = Complete-CaptureSession -Session $session

        $first.Value | Should -Be 42
        $second.Value | Should -Be 42
        $script:factoryCallCount | Should -Be 1
        $metrics.ProviderQueryCount | Should -Be 1
        $metrics.CacheHitCount | Should -Be 1
    }

    It 'records per-setting mutation timing and failure state' {
        $session = New-CaptureSession -Phase Permissive

        Invoke-TimedSettingOperation -Phase Permissive -Id 'test.setting' -Type 'RegistryValue' -Session $session -Action {}
        { Invoke-TimedSettingOperation -Phase Permissive -Id 'test.failure' -Type 'RegistryValue' -Session $session -Action { throw 'expected failure' } } |
            Should -Throw '*expected failure*'
        $metrics = Complete-CaptureSession -Session $session

        @($metrics.Settings).Count | Should -Be 2
        $metrics.Settings[0].Id | Should -Be 'test.setting'
        $metrics.Settings[0].Succeeded | Should -BeTrue
        $metrics.Settings[1].Succeeded | Should -BeFalse
        $metrics.Settings[1].Error | Should -Be 'expected failure'
    }
}

Describe 'Mutation work planning' {
    It 'collapses ten Defender values and seventeen WSMan values into five provider work items' {
        $items = @(
            foreach ($definition in @(Get-DefenseDefinitions | Where-Object { [string]$_.Type -in @('MpPreferenceValue', 'WsManValue') })) {
                [PSCustomObject]@{
                    Id = [string]$definition.Id; Type = [string]$definition.Type
                    Definition = $definition; Entry = [PSCustomObject]@{}; CanBatch = $true
                }
            }
        )

        $workItems = @(Get-MutationWorkItems -Items $items)
        $defenderWork = @($workItems | Where-Object { [string]$_.BatchKind -eq 'DefenderPreferenceValues' })
        $wsManWork = @($workItems | Where-Object { [string]$_.BatchKind -eq 'WsManValues' })

        $defenderWork.Count | Should -Be 1
        @($defenderWork[0].Items).Count | Should -Be 10
        $wsManWork.Count | Should -Be 4
        @($wsManWork.Items | ForEach-Object { @($_).Count } | Measure-Object -Sum).Sum | Should -Be 17
        $workItems.Count | Should -Be 5
    }

    It 'submits multiple Defender scalar properties in one setter call' {
        Clear-WinDefStateCommandCache
        Mock Set-MpPreference {}

        Set-MpPreferencePropertyValues -Items @(
            [PSCustomObject]@{ Property = 'DisableRealtimeMonitoring'; Value = $true }
            [PSCustomObject]@{ Property = 'DisableBehaviorMonitoring'; Value = $true }
        )

        Should -Invoke Set-MpPreference -Times 1 -Exactly -ParameterFilter {
            $DisableRealtimeMonitoring -eq $true -and $DisableBehaviorMonitoring -eq $true
        }
    }

    It 'submits properties sharing one WSMan resource in one setter call' {
        Clear-WinDefStateCommandCache
        Mock Get-CimInstance { [PSCustomObject]@{ Name = 'WinRM'; StartMode = 'Auto'; State = 'Running' } }
        Mock Set-WSManInstance {}

        Set-WsManConfigValues -Items @(
            [PSCustomObject]@{ Path = 'WSMan:\localhost\Service\AllowUnencrypted'; Value = $true }
            [PSCustomObject]@{ Path = 'WSMan:\localhost\Service\IPv4Filter'; Value = '*' }
        )

        Should -Invoke Set-WSManInstance -Times 1 -Exactly -ParameterFilter {
            $ResourceURI -eq 'winrm/config/service' -and
            $ValueSet.AllowUnencrypted -eq 'true' -and
            $ValueSet.IPv4Filter -eq '*'
        }
    }

    It 'rejects mixed WSMan resource batches before opening the write scope' {
        Clear-WinDefStateCommandCache
        Mock Get-CimInstance { throw 'The WinRM scope must not open for an invalid batch.' }
        Mock Set-WSManInstance {}

        {
            Set-WsManConfigValues -Items @(
                [PSCustomObject]@{ Path = 'WSMan:\localhost\Service\AllowUnencrypted'; Value = $true }
                [PSCustomObject]@{ Path = 'WSMan:\localhost\Client\AllowUnencrypted'; Value = $true }
            )
        } | Should -Throw '*mixes resource URIs*'
        Should -Invoke Get-CimInstance -Times 0 -Exactly
        Should -Invoke Set-WSManInstance -Times 0 -Exactly
    }

    It 'wires both mutation orchestrators through the grouped work plan' {
        (Get-Command Set-DefensePermissive -CommandType Function).Definition |
            Should -Match 'Get-MutationWorkItems[^\r\n]+\$mutationItems'
        (Get-Command Set-DefensePermissive -CommandType Function).Definition |
            Should -Match 'Invoke-PermissiveMutationWorkItem[^\r\n]+\$workItem'
        (Get-Command Restore-DefenseSnapshot -CommandType Function).Definition |
            Should -Match 'Invoke-RestoreMutationWorkItem[^\r\n]+\$workItem'
    }

    It 'submits all Defender restore values from a grouped work item' {
        Mock Set-MpPreferencePropertyValues {}
        $items = @(
            [PSCustomObject]@{
                Id = 'defender.one'; Type = 'MpPreferenceValue'; CanBatch = $true
                Entry = [PSCustomObject]@{ Property = 'DisableRealtimeMonitoring'; RestoreValue = $false }
            }
            [PSCustomObject]@{
                Id = 'defender.two'; Type = 'MpPreferenceValue'; CanBatch = $true
                Entry = [PSCustomObject]@{ Property = 'DisableBehaviorMonitoring'; RestoreValue = $false }
            }
        )
        $workItem = @(Get-MutationWorkItems -Items $items)[0]

        Invoke-RestoreMutationWorkItem -WorkItem $workItem

        Should -Invoke Set-MpPreferencePropertyValues -Times 1 -Exactly -ParameterFilter {
            @($Items).Count -eq 2 -and
            @($Items | Where-Object { $_.Property -eq 'DisableRealtimeMonitoring' -and $_.Value -eq $false }).Count -eq 1 -and
            @($Items | Where-Object { $_.Property -eq 'DisableBehaviorMonitoring' -and $_.Value -eq $false }).Count -eq 1
        }
    }
}

Describe 'User registry mutation session' {
    It 'reuses one profile enumeration and hive mount across permissive user settings' {
        $target = [PSCustomObject]@{ Sid = 'S-1-5-21-1-1000'; ProfilePath = 'C:\Users\Test'; HivePath = 'C:\Users\Test\NTUSER.DAT'; Loaded = $false }
        Mock Get-UserProfileRegistryTargets { @($target) }
        Mock Open-UserRegistryTarget { [PSCustomObject]@{ Sid = 'S-1-5-21-1-1000'; RootPath = 'Test:\User'; MountName = 'WinDefState_Test'; MountedByTool = $true } }
        Mock Close-UserRegistryTarget {}
        Mock Ensure-RegistryPath {}
        Mock New-ItemProperty {}
        $item = [PSCustomObject]@{ RelativePath = 'Software\Test'; Name = 'Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1 }
        $session = New-CaptureSession -Phase Permissive
        $session.UserRegistryDefinitionsRemaining = 2

        Set-Permissive-LoadedUserRegistryValues -Items @($item) -CaptureSession $session
        Complete-UserRegistrySessionDefinition -Session $session
        Should -Invoke Close-UserRegistryTarget -Times 0 -Exactly
        Set-Permissive-LoadedUserRegistryValues -Items @($item) -CaptureSession $session
        Complete-UserRegistrySessionDefinition -Session $session
        Should -Invoke Close-UserRegistryTarget -Times 1 -Exactly
        $metrics = Complete-CaptureSession -Session $session

        Should -Invoke Get-UserProfileRegistryTargets -Times 1 -Exactly
        Should -Invoke Open-UserRegistryTarget -Times 1 -Exactly
        Should -Invoke New-ItemProperty -Times 2 -Exactly
        Should -Invoke Close-UserRegistryTarget -Times 1 -Exactly
        $metrics.CacheHitCount | Should -BeGreaterOrEqual 2
    }

    It 'reuses one profile enumeration and hive mount across restore user settings' {
        $target = [PSCustomObject]@{ Sid = 'S-1-5-21-1-1000'; ProfilePath = 'C:\Users\Test'; HivePath = 'C:\Users\Test\NTUSER.DAT'; Loaded = $false }
        Mock Get-UserProfileRegistryTargets { @($target) }
        Mock Open-UserRegistryTarget { [PSCustomObject]@{ Sid = 'S-1-5-21-1-1000'; RootPath = 'Test:\User'; MountName = 'WinDefState_Test'; MountedByTool = $true } }
        Mock Close-UserRegistryTarget {}
        Mock Ensure-RegistryPath {}
        Mock New-ItemProperty {}
        $entry = [PSCustomObject]@{
            Sid = $target.Sid; ProfilePath = $target.ProfilePath; HivePath = $target.HivePath
            RelativePath = 'Software\Test'; Name = 'Enabled'; ValueKind = 'DWord'; Exists = $true; CurrentValue = 1
        }
        $session = New-CaptureSession -Phase Restore
        $session.UserRegistryDefinitionsRemaining = 2

        Restore-LoadedUserRegistryValues -Entries @($entry) -CaptureSession $session
        Complete-UserRegistrySessionDefinition -Session $session
        Should -Invoke Close-UserRegistryTarget -Times 0 -Exactly
        Restore-LoadedUserRegistryValues -Entries @($entry) -CaptureSession $session
        Complete-UserRegistrySessionDefinition -Session $session
        Should -Invoke Close-UserRegistryTarget -Times 1 -Exactly
        $metrics = Complete-CaptureSession -Session $session

        Should -Invoke Get-UserProfileRegistryTargets -Times 1 -Exactly
        Should -Invoke Open-UserRegistryTarget -Times 1 -Exactly
        Should -Invoke New-ItemProperty -Times 2 -Exactly
        Should -Invoke Close-UserRegistryTarget -Times 1 -Exactly
        $metrics.CacheHitCount | Should -BeGreaterOrEqual 2
    }
}

Describe 'WinRM mutation session' {
    BeforeEach {
        Mock Get-CimInstance {
            [PSCustomObject]@{ Name = 'WinRM'; StartMode = 'Disabled'; State = 'Stopped' }
        }
        Mock Set-ServiceStartModeValue {}
        Mock Start-Service {}
        Mock Set-ServiceRunningState {}
    }

    It 'reuses one temporary service scope across all selected WSMan writes' {
        $session = New-CaptureSession -Phase Restore
        $session.WinRmMutationDefinitionsRemaining = 2

        Invoke-WithTemporaryWinRmServiceForWrite -CaptureSession $session -ScriptBlock {}
        Complete-WinRmMutationSessionDefinition -Session $session
        $session.Resources.Count | Should -Be 1

        Invoke-WithTemporaryWinRmServiceForWrite -CaptureSession $session -ScriptBlock {}
        Complete-WinRmMutationSessionDefinition -Session $session
        $session.Resources.Count | Should -Be 0
        $metrics = Complete-CaptureSession -Session $session

        Should -Invoke Get-CimInstance -Times 1 -Exactly
        Should -Invoke Start-Service -Times 1 -Exactly -ParameterFilter { $Name -eq 'WinRM' }
        Should -Invoke Set-ServiceRunningState -Times 1 -Exactly -ParameterFilter { $Name -eq 'WinRM' -and $Running -eq $false }
        Should -Invoke Set-ServiceStartModeValue -Times 1 -Exactly -ParameterFilter { $Name -eq 'WinRM' -and $StartModeValue -eq 'demand' }
        Should -Invoke Set-ServiceStartModeValue -Times 1 -Exactly -ParameterFilter { $Name -eq 'WinRM' -and $StartModeValue -eq 'disabled' }
        @($metrics.ProviderQueries | Where-Object { $_.Key -eq 'winrm.write-scope' }).Count | Should -Be 1
        $metrics.CacheHitCount | Should -Be 1
    }

    It 'defers only the WinRM service baseline until after WSMan writes' {
        $entries = @(
            [PSCustomObject]@{ Id = 'winrm.service'; Type = 'ServiceConfig'; Name = 'WinRM' }
            [PSCustomObject]@{ Id = 'winrm.service.basic'; Type = 'WsManValue' }
            [PSCustomObject]@{ Id = 'rdp.user_authentication'; Type = 'RegistryValue' }
            [PSCustomObject]@{ Id = 'spooler.service'; Type = 'ServiceConfig'; Name = 'Spooler' }
        )

        $ordered = @(Get-OrderedRestoreEntries -Entries $entries)

        $ordered[0].Id | Should -Be 'winrm.service.basic'
        $ordered[1].Id | Should -Be 'rdp.user_authentication'
        $ordered[2].Id | Should -Be 'spooler.service'
        $ordered[3].Id | Should -Be 'winrm.service'
    }

    It 'still restores the startup mode when stopping the temporary service fails' {
        Mock Set-ServiceRunningState { throw 'expected stop failure' }
        Mock Write-Warning {}
        $session = New-CaptureSession -Phase Restore
        $session.WinRmMutationDefinitionsRemaining = 1
        Invoke-WithTemporaryWinRmServiceForWrite -CaptureSession $session -ScriptBlock {}

        { Complete-WinRmMutationSessionDefinition -Session $session } | Should -Throw '*Could not stop the temporary WinRM service*'

        Should -Invoke Set-ServiceStartModeValue -Times 1 -Exactly -ParameterFilter { $Name -eq 'WinRM' -and $StartModeValue -eq 'disabled' }
        $session.Resources.Count | Should -Be 0
    }
}

Describe 'Operation journal integrity' {
    It 'replaces existing files without leaving temporary artifacts' {
        $path = Join-Path $TestDrive 'atomic.txt'
        Write-TextAtomic -Path $path -Content 'before'
        Write-TextAtomic -Path $path -Content 'after'

        (Get-Content -LiteralPath $path -Raw) | Should -Be 'after'
        @(Get-ChildItem -LiteralPath $TestDrive -Filter '*.tmp' -Force).Count | Should -Be 0
    }

    It 'removes a partial snapshot file when serialization fails before publish' {
        $path = Join-Path $TestDrive 'serialization-failure.json'
        $snapshot = [PSCustomObject]@{
            SchemaVersion = 2
            Tool = 'WinDefState'
            ComputerName = 'TEST'
            CapturedAtUtc = '2026-01-01T00:00:00Z'
            Settings = @()
        }
        Mock ConvertTo-Json { throw 'expected serialization failure' }

        { Write-SnapshotJsonAtomic -Path $path -Snapshot $snapshot } | Should -Throw '*expected serialization failure*'

        Test-Path -LiteralPath $path | Should -BeFalse
        @(Get-ChildItem -LiteralPath $TestDrive -Filter '*.tmp' -Force).Count | Should -Be 0
    }

    It 'cleans each temporary writer artifact when publishing fails' {
        Mock Publish-FileAtomic { throw 'expected publish failure' }

        { Write-JsonAtomic -Path (Join-Path $TestDrive 'failed.json') -InputObject ([PSCustomObject]@{ Value = 1 }) } |
            Should -Throw '*expected publish failure*'
        { Write-TextAtomic -Path (Join-Path $TestDrive 'failed.txt') -Content 'value' } |
            Should -Throw '*expected publish failure*'
        { Write-BytesAtomic -Path (Join-Path $TestDrive 'failed.bin') -Content ([byte[]]@(1, 2, 3)) } |
            Should -Throw '*expected publish failure*'

        @(Get-ChildItem -LiteralPath $TestDrive -Filter '*.tmp' -Force).Count | Should -Be 0
    }

    It 'fails closed when an existing operation journal contains JSON null' {
        $stateRoot = Join-Path $TestDrive 'null-journal'
        Ensure-Directory -Path $stateRoot
        Write-TextAtomic -Path (Get-OperationPath -Root $stateRoot) -Content 'null'

        { Get-OperationState -Root $stateRoot } | Should -Throw '*did not contain a JSON object*'
        { Assert-NoActiveOperation -Root $stateRoot } | Should -Throw '*did not contain a JSON object*'
    }

    It 'fails closed when an existing operation journal has no snapshot path' {
        $stateRoot = Join-Path $TestDrive 'incomplete-journal'
        Ensure-Directory -Path $stateRoot
        Write-JsonAtomic -Path (Get-OperationPath -Root $stateRoot) -InputObject ([PSCustomObject]@{ Status = 'Applying' })

        { Get-OperationState -Root $stateRoot } | Should -Throw '*missing its snapshot path*'
        { Assert-NoActiveOperation -Root $stateRoot } | Should -Throw '*missing its snapshot path*'
    }

    It 'records status, scope, and snapshot sidecar hashes' {
        $stateRoot = Join-Path $TestDrive 'state'
        $snapshotPath = Join-Path $TestDrive 'snapshot.json'
        Write-TextAtomic -Path $snapshotPath -Content '{"snapshot":true}'
        $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $snapshotPath
        Ensure-Directory -Path $assetRoot
        Write-TextAtomic -Path (Join-Path $assetRoot 'policy.xml') -Content '<Policy />'

        $operation = Write-OperationState -Root $stateRoot -SnapshotPath $snapshotPath -Mode Permissive -IncludeId @('rdp.user_authentication')
        $loaded = Get-OperationState -Root $stateRoot

        $loaded.SchemaVersion | Should -Be 1
        $loaded.Status | Should -Be 'Applying'
        $loaded.Producer.ScriptSha256 | Should -Be (Get-FileHash -LiteralPath $enginePath -Algorithm SHA256).Hash
        @($loaded.IncludeId) | Should -Contain 'rdp.user_authentication'
        @($loaded.SnapshotIntegrity.Assets).Count | Should -Be 1
        Clear-SnapshotAssetCache
        { Assert-OperationSnapshotIntegrity -Operation $loaded } | Should -Not -Throw
        $script:WinDefStateSnapshotAssetCache.Count | Should -Be 1

        $operation = Update-OperationStateStatus -Root $stateRoot -Operation $operation -Status Applied
        (Get-OperationState -Root $stateRoot).Status | Should -Be 'Applied'

        foreach ($failureStatus in @('ApplyFailed', 'RestoreFailed')) {
            $operation = Update-OperationStateStatus -Root $stateRoot -Operation $operation -Status $failureStatus
            (Get-OperationState -Root $stateRoot).Status | Should -Be $failureStatus
        }

        $verification = [PSCustomObject]@{
            VerifiedAtUtc = '2026-01-01T00:00:00Z'; VerifiedCount = 1; PendingRebootCount = 1; MismatchCount = 0
            MutationMetrics = [PSCustomObject]@{ DurationMs = 1250 }
            Results = @([PSCustomObject]@{ Id = 'uac.enable_lua'; Status = 'ConfiguredPendingReboot' })
        }
        $verificationPath = Join-Path $stateRoot 'permissive-check.txt'
        $operation = Update-OperationStateStatus -Root $stateRoot -Operation $operation -Status AppliedPendingReboot -PermissiveVerification $verification -PermissiveVerificationReportPath $verificationPath
        $loadedVerification = (Get-OperationState -Root $stateRoot).PermissiveVerification
        $loadedVerification.PendingRebootCount | Should -Be 1
        $loadedVerification.MutationDurationMs | Should -Be 1250
        @($loadedVerification.PendingRebootIds) | Should -Contain 'uac.enable_lua'
    }

    It 'rejects a snapshot changed after the journal was written' {
        $stateRoot = Join-Path $TestDrive 'changed-state'
        $snapshotPath = Join-Path $TestDrive 'changed-snapshot.json'
        Write-TextAtomic -Path $snapshotPath -Content '{"snapshot":true}'
        $operation = Write-OperationState -Root $stateRoot -SnapshotPath $snapshotPath -Mode Permissive
        Write-TextAtomic -Path $snapshotPath -Content '{"snapshot":false}'

        { Assert-OperationSnapshotIntegrity -Operation $operation } | Should -Throw '*no longer matches*'
    }

    It 'distinguishes unrelated explicit snapshots' {
        $operation = [PSCustomObject]@{ SnapshotPath = Join-Path $TestDrive 'first.json' }

        (Test-OperationTargetsSnapshot -Operation $operation -SnapshotPath (Join-Path $TestDrive 'first.json')) | Should -BeTrue
        (Test-OperationTargetsSnapshot -Operation $operation -SnapshotPath (Join-Path $TestDrive 'second.json')) | Should -BeFalse
    }

    It 'prevents a second permissive run from replacing an active baseline' {
        $stateRoot = Join-Path $TestDrive 'active-operation'
        Ensure-Directory -Path $stateRoot
        Write-JsonAtomic -Path (Get-OperationPath -Root $stateRoot) -InputObject ([PSCustomObject]@{
            Status       = 'Applied'
            SnapshotPath = 'C:\state\original.json'
        })

        { Assert-NoActiveOperation -Root $stateRoot } | Should -Throw '*Restore it before starting another permissive operation*'
        Clear-OperationState -Root $stateRoot
        { Assert-NoActiveOperation -Root $stateRoot } | Should -Not -Throw
    }

    It 'persists per-setting restore checkpoints and the in-flight failure' {
        $stateRoot = Join-Path $TestDrive 'restore-checkpoint-state'
        $snapshotPath = Join-Path $TestDrive 'restore-checkpoint-snapshot.json'
        Write-TextAtomic -Path $snapshotPath -Content '{"snapshot":true}'
        $operation = Write-OperationState -Root $stateRoot -SnapshotPath $snapshotPath -Mode Permissive

        $operation = Initialize-RestoreCheckpoint -Root $stateRoot -Operation $operation -RequestedIds @('rdp.user_authentication', 'firewall.profiles')
        $operation = Start-RestoreCheckpointWorkItem -Root $stateRoot -Operation $operation -WorkItemId 'rdp.user_authentication' -SettingIds @('rdp.user_authentication')
        $operation = Complete-RestoreCheckpointWorkItem -Root $stateRoot -Operation $operation -SettingIds @('rdp.user_authentication')
        $operation = Start-RestoreCheckpointWorkItem -Root $stateRoot -Operation $operation -WorkItemId 'firewall.profiles' -SettingIds @('firewall.profiles')
        $operation = Set-RestoreCheckpointFailure -Root $stateRoot -Operation $operation -Message 'expected provider failure'

        $loaded = Get-OperationState -Root $stateRoot
        $loaded.RestoreCheckpoint.AttemptNumber | Should -Be 1
        @($loaded.RestoreCheckpoint.RequestedIds) | Should -Contain 'rdp.user_authentication'
        @($loaded.RestoreCheckpoint.CompletedIds) | Should -Contain 'rdp.user_authentication'
        @($loaded.RestoreCheckpoint.CompletedIds) | Should -Not -Contain 'firewall.profiles'
        $loaded.RestoreCheckpoint.CurrentWorkItemId | Should -Be 'firewall.profiles'
        @($loaded.RestoreCheckpoint.CurrentIds) | Should -Contain 'firewall.profiles'
        $loaded.RestoreCheckpoint.LastFailure.Message | Should -Be 'expected provider failure'

        $operation = Initialize-RestoreCheckpoint -Root $stateRoot -Operation $loaded -RequestedIds @('rdp.user_authentication', 'firewall.profiles')
        $operation.RestoreCheckpoint.AttemptNumber | Should -Be 2
        @($operation.RestoreCheckpoint.CompletedIds) | Should -Contain 'rdp.user_authentication'
        $operation.RestoreCheckpoint.CurrentWorkItemId | Should -BeNullOrEmpty
    }

    It 'resumes only checkpointed settings that still match the snapshot' {
        $snapshotPath = Join-Path $TestDrive 'resume-snapshot.json'
        $entries = @(
            [PSCustomObject]@{ Id = 'rdp.user_authentication'; Type = 'RegistryValue' },
            [PSCustomObject]@{ Id = 'firewall.profiles'; Type = 'FirewallProfiles' }
        )
        $snapshot = [PSCustomObject]@{ Settings = @($entries) }
        $operation = [PSCustomObject]@{
            RestoreCheckpoint = [PSCustomObject]@{ CompletedIds = @('rdp.user_authentication', 'firewall.profiles') }
        }
        Mock Test-DefenseSnapshot {
            [PSCustomObject]@{
                Results = @(
                    [PSCustomObject]@{ Id = 'rdp.user_authentication'; Matches = $true; Skipped = $false },
                    [PSCustomObject]@{ Id = 'firewall.profiles'; Matches = $false; Skipped = $false }
                )
            }
        }

        $resume = Get-RestoreCheckpointResumeState -Operation $operation -Snapshot $snapshot -SnapshotPath $snapshotPath -MutationEntries $entries

        @($resume.CandidateIds).Count | Should -Be 2
        @($resume.VerifiedIds) | Should -Contain 'rdp.user_authentication'
        @($resume.VerifiedIds) | Should -Not -Contain 'firewall.profiles'
        @($resume.RetryIds) | Should -Contain 'firewall.profiles'
        Should -Invoke Test-DefenseSnapshot -Times 1 -Exactly -ParameterFilter {
            @($IncludeId).Count -eq 2
        }
    }

    It 'keeps restore resume evidence in the verification report' {
        $verification = [PSCustomObject]@{
            ComputerName = 'HOST'; VerifiedAtUtc = '2026-01-01T00:00:00Z'
            MatchedCount = 2; SkippedCount = 0; IncompleteCount = 0; InventoryCount = 0; MismatchCount = 0
            Results = @()
            RestoreCheckpointSummary = [PSCustomObject]@{
                AttemptNumber = 2; RequestedMutationCount = 3; PreviouslyCompletedCount = 2
                RevalidatedAndSkippedCount = 1; ReappliedCheckpointCount = 1; ScheduledMutationCount = 2
            }
        }

        $report = @(Get-VerificationReportLines -Verification $verification -SnapshotPath 'C:\state\baseline.json') -join "`n"
        $report | Should -Match 'Restore attempt: 2'
        $report | Should -Match 'Checkpointed settings skipped as matching: 1'
        $report | Should -Match 'Checkpointed settings scheduled again: 1'
    }
}

Describe 'Automatic preflight diagnostics' {
    It 'maps selected providers to capture and mutation command requirements' {
        $items = @(
            [PSCustomObject]@{ Type = 'MpPreferenceValue' },
            [PSCustomObject]@{ Type = 'FirewallProfiles' },
            [PSCustomObject]@{ Type = 'RegistryValue' }
        )

        $snapshotRequirements = @(Get-WinDefStatePreflightProviderRequirements -Items $items -Action Snapshot)
        @($snapshotRequirements.Command) | Should -Contain 'Get-MpPreference'
        @($snapshotRequirements.Command) | Should -Contain 'Get-NetFirewallProfile'
        @($snapshotRequirements.Command) | Should -Not -Contain 'Set-MpPreference'

        $restoreRequirements = @(Get-WinDefStatePreflightProviderRequirements -Items $items -Action Restore)
        @($restoreRequirements.Command) | Should -Contain 'Set-MpPreference'
        @($restoreRequirements.Command) | Should -Contain 'Set-NetFirewallProfile'

        $restoreOnlyList = @([PSCustomObject]@{ Type = 'MpPreferenceList' })
        @((Get-WinDefStatePreflightProviderRequirements -Items $restoreOnlyList -Action Permissive).Command) | Should -Not -Contain 'Add-MpPreference'
        @((Get-WinDefStatePreflightProviderRequirements -Items $restoreOnlyList -Action Restore).Command) | Should -Contain 'Add-MpPreference'
    }

    It 'renders warning details in a stable text report' {
        $preflight = [PSCustomObject]@{
            Action = 'Restore'; ComputerName = 'HOST'; CheckedAtUtc = '2026-01-01T00:00:00Z'
            OverallStatus = 'Warning'; WarningCount = 1
            Checks = @([PSCustomObject]@{ Status = 'Warning'; Name = 'Provider command: Example'; Value = 'Missing'; Detail = 'ExampleProvider' })
        }

        $lines = @(Get-WinDefStatePreflightReportLines -Preflight $preflight)
        ($lines -join "`n") | Should -Match 'Overall status: Warning'
        ($lines -join "`n") | Should -Match '\[WARNING\] Provider command: Example: Missing'
        ($lines -join "`n") | Should -Match 'ExampleProvider'
    }
}

Describe 'Registry capture batching' {
    BeforeEach {
        $script:testRegistryKey = [PSCustomObject]@{}
        $script:testRegistryKey | Add-Member -MemberType ScriptMethod -Name GetValueKind -Value { param($name) 'DWord' }
        Mock Get-Item { $script:testRegistryKey }
        Mock Get-ItemProperty {
            [PSCustomObject]@{
                FirstValue  = 1
                SecondValue = 0
            }
        }
    }

    It 'reads a shared registry key once per capture session' {
        $session = New-CaptureSession -Phase Snapshot
        $firstDefinition = [PSCustomObject]@{ Id = 'registry.first'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\WinDefStateTest'; Name = 'FirstValue'; ValueKind = 'DWord'; RequiresReboot = $false }
        $secondDefinition = [PSCustomObject]@{ Id = 'registry.second'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\WinDefStateTest'; Name = 'SecondValue'; ValueKind = 'DWord'; RequiresReboot = $false }

        $first = Capture-Definition -Definition $firstDefinition -CaptureSession $session
        $second = Capture-Definition -Definition $secondDefinition -CaptureSession $session

        $first.CurrentValue | Should -Be 1
        $second.CurrentValue | Should -Be 0
        $first.ValueKind | Should -Be 'DWord'
        Assert-MockCalled Get-Item -Times 1 -Exactly
        Assert-MockCalled Get-ItemProperty -Times 1 -Exactly
        $session.CacheHitCount | Should -Be 1
    }
}

Describe 'Service capture batching' {
    BeforeEach {
        Mock Get-CimInstance {
            param($ClassName, $Filter)
            $services = @(
                [PSCustomObject]@{ Name = 'AppIDSvc'; StartMode = 'Manual'; State = 'Stopped' }
                [PSCustomObject]@{ Name = 'Spooler'; StartMode = 'Auto'; State = 'Running' }
                [PSCustomObject]@{ Name = 'WinRM'; StartMode = 'Auto'; State = 'Running' }
            )
            @($services | Where-Object { $Filter -match ("Name='{0}'" -f [regex]::Escape([string]$_.Name)) })
        }
    }

    It 'queries all tracked services once per capture session' {
        $session = New-CaptureSession -Phase Verification
        $session.ServiceNames = @('AppIDSvc', 'Spooler', 'WinRM')
        $definitions = @(
            [PSCustomObject]@{ Id = 'applocker.service'; Type = 'ServiceConfig'; Name = 'AppIDSvc'; RequiresReboot = $false }
            [PSCustomObject]@{ Id = 'print.spooler_service'; Type = 'ServiceConfig'; Name = 'Spooler'; RequiresReboot = $false }
            [PSCustomObject]@{ Id = 'winrm.service'; Type = 'ServiceConfig'; Name = 'WinRM'; RequiresReboot = $false }
        )

        $entries = @($definitions | ForEach-Object { Capture-Definition -Definition $_ -CaptureSession $session })

        $entries[0].CurrentValue.State | Should -Be 'Stopped'
        $entries[1].CurrentValue.StartMode | Should -Be 'Auto'
        $entries[2].CurrentValue.State | Should -Be 'Running'
        Assert-MockCalled Get-CimInstance -Times 1 -Exactly
        $session.CacheHitCount | Should -Be 2
    }

    It 'extends an unseeded session without losing later service state' {
        $session = New-CaptureSession -Phase Snapshot

        $appIdService = Get-ServiceCaptureState -Name 'AppIDSvc' -CaptureSession $session
        $spoolerService = Get-ServiceCaptureState -Name 'Spooler' -CaptureSession $session

        $appIdService.State | Should -Be 'Stopped'
        $spoolerService.State | Should -Be 'Running'
        Assert-MockCalled Get-CimInstance -Times 2 -Exactly
    }
}

Describe 'Operational output' {
    BeforeEach {
        Mock Write-Progress {}
    }

    It 'emits structured progress only when requested' {
        $definition = (Get-Command Write-OperationProgress -CommandType Function).Definition

        $definition | Should -Match 'if\s*\(\$EmitProgress\)\s*\{\s*Write-Host'
        $definition | Should -Match 'WDS_PROGRESS\|\{0\}\|\{1\}\|\{2\}\|\{3\}'
    }

    It 'uses native progress without requiring verbose output' {
        $script:EmitProgress = $false

        Write-OperationProgress -Phase 'Capture' -Current 20 -Total 80 -Id 'registry.test'

        Assert-MockCalled Write-Progress -Times 1 -ParameterFilter {
            $Activity -eq 'WinDefState Capture' -and $Status -eq 'registry.test' -and $PercentComplete -eq 25
        }
    }

    It 'emits structured results only when requested' {
        $definition = (Get-Command Write-OperationResult -CommandType Function).Definition

        $definition | Should -Match 'if\s*\(-not\s+\$EmitProgress\)\s*\{\s*return'
        $definition | Should -Match 'WDS_RESULT\|\{0\}\|\{1\}'
        $script:EmitProgress = $false
        @(Write-OperationResult -Name 'SnapshotPath' -Value 'C:\State\snapshot.json').Count | Should -Be 0
    }

    It 'returns only the report header block for summary output' {
        $summary = @(Get-ReportSummaryLines -Lines @('Heading', 'Value: 1', ' ', 'Detail'))

        $summary.Count | Should -Be 2
        $summary[1] | Should -Be 'Value: 1'
    }

    It 'does not describe an incomplete-baseline restore as fully verified' {
        $snapshotPath = 'C:\state\baseline.json'
        (Get-RestoreCompletionMessage -SnapshotPath $snapshotPath -Verification ([PSCustomObject]@{ SkippedCount = 0 })) |
            Should -Be "Restore completed and verified from snapshot: $snapshotPath"
        (Get-RestoreCompletionMessage -SnapshotPath $snapshotPath -Verification ([PSCustomObject]@{ SkippedCount = 2 })) |
            Should -Be "Restore completed and verified all settings with complete baselines from snapshot: $snapshotPath"
    }
}

Describe 'Operation concurrency guard' {
    It 'wraps the public command dispatcher in a machine-wide lock' {
        $engineText = Get-Content -LiteralPath $enginePath -Raw

        $engineText | Should -Match '\$operationLock\s*=\s*Enter-WinDefStateOperationLock'
        $engineText | Should -Match 'finally\s*\{\s*Exit-WinDefStateOperationLock\s+-Lock\s+\$operationLock'
        (Get-Command Enter-WinDefStateOperationLock -CommandType Function).Definition | Should -Match 'Global\\WinDefState\.Operation'
    }
}

Describe 'Isolated provider process' {
    It 'drains redirected streams asynchronously before waiting and always disposes the process' {
        $functionText = (Get-Command Invoke-ChildPowerShell -CommandType Function).Definition
        $readIndex = $functionText.IndexOf('ReadToEndAsync')
        $waitIndex = $functionText.IndexOf('WaitForExit($TimeoutSeconds * 1000)')

        $readIndex | Should -BeGreaterThan -1
        $waitIndex | Should -BeGreaterThan $readIndex
        $functionText | Should -Match 'finally\s*\{\s*\$process\.Dispose\(\)'
    }

    It 'runs keyed child PowerShell requests through the concurrent batch runner' {
        $originalPath = $env:PATH
        if ($env:OS -ne 'Windows_NT') {
            $pwshPath = (Get-Process -Id $PID).Path
            $shimPath = Join-Path $TestDrive 'powershell.exe'
            New-Item -ItemType SymbolicLink -Path $shimPath -Target $pwshPath | Out-Null
            $env:PATH = "$TestDrive$([IO.Path]::PathSeparator)$originalPath"
        }

        try {
            $results = @(Invoke-ChildPowerShellBatch -Requests @(
                [PSCustomObject]@{ Key = 'first'; TimeoutSeconds = 5; ScriptText = "[Console]::Out.Write('one')" }
                [PSCustomObject]@{ Key = 'second'; TimeoutSeconds = 5; ScriptText = "[Console]::Out.Write('two')" }
                [PSCustomObject]@{ Key = 'slow'; TimeoutSeconds = 1; ScriptText = 'Start-Sleep -Seconds 5' }
            ))

            $results.Count | Should -Be 3
            $results[0].Key | Should -Be 'first'
            $results[0].StdOut | Should -Be 'one'
            $results[0].ExitCode | Should -Be 0
            $results[1].Key | Should -Be 'second'
            $results[1].StdOut | Should -Be 'two'
            $results[1].ExitCode | Should -Be 0
            $results[2].Key | Should -Be 'slow'
            $results[2].TimedOut | Should -BeTrue
            $results[2].ExitCode | Should -BeNullOrEmpty
        } finally {
            $env:PATH = $originalPath
        }
    }

    It 'starts and drains all batch processes before waiting and always disposes them' {
        $functionText = (Get-Command Invoke-ChildPowerShellBatch -CommandType Function).Definition
        $startIndex = $functionText.IndexOf('$invocations.Add')
        $waitIndex = $functionText.IndexOf('foreach ($invocation in $invocations)')

        $startIndex | Should -BeGreaterThan -1
        $waitIndex | Should -BeGreaterThan $startIndex
        $functionText | Should -Match 'ReadToEndAsync'
        $functionText | Should -Match '\$invocation\.Process\.Dispose\(\)'
    }
}

Describe 'BitLocker capture batching' {
    It 'submits mounted volumes as one bounded child-process batch' {
        Mock Test-CommandAvailable { $Name -eq 'Get-BitLockerVolume' }
        Mock Get-CimInstance {
            @(
                [PSCustomObject]@{ DeviceID = 'C:' }
                [PSCustomObject]@{ DeviceID = 'D:' }
            )
        }
        Mock Invoke-ChildPowerShellBatch {
            $script:bitLockerProbeRequests = @($Requests)
            foreach ($request in @($Requests)) {
                $volumeType = if ([string]$request.Key -eq 'C:') { 'OperatingSystem' } else { 'Data' }
                $json = [PSCustomObject]@{
                    MountPoint = [string]$request.Key; VolumeType = $volumeType; ProtectionStatus = 'On'
                    VolumeStatus = 'FullyEncrypted'; LockStatus = 'Unlocked'; EncryptionMethod = 'XtsAes256'
                    EncryptionPercentage = 100; KeyProtectorCount = 0; KeyProtectors = @(); KeyProtectorsCaptured = $true
                    AutoUnlockEnabled = $false; AutoUnlockCaptured = $true
                } | ConvertTo-Json -Compress -Depth 6
                [PSCustomObject]@{
                    Key = [string]$request.Key; CommandAvailable = $true; TimedOut = $false
                    ExitCode = 0; StdOut = $json; StdErr = ''; DurationMs = 10
                }
            }
        }

        $state = Get-BitLockerVolumeStates

        Should -Invoke Invoke-ChildPowerShellBatch -Times 1 -Exactly
        @($script:bitLockerProbeRequests).Count | Should -Be 2
        @($state.Volumes).Count | Should -Be 2
        @($state.TimedOutMountPoints).Count | Should -Be 0
        @($state.CaptureIssues).Count | Should -Be 0
    }

    It 'returns an empty complete inventory when no eligible volumes are mounted' {
        Mock Test-CommandAvailable { $Name -eq 'Get-BitLockerVolume' }
        Mock Get-CimInstance { @() }

        $state = Get-BitLockerVolumeStates

        $state.CommandAvailable | Should -BeTrue
        @($state.Volumes).Count | Should -Be 0
        @($state.TimedOutMountPoints).Count | Should -Be 0
        @($state.CaptureIssues).Count | Should -Be 0
    }
}

Describe 'BitLocker restore batching' {
    BeforeEach {
        Clear-WinDefStateCommandCache
        Mock Resume-BitLocker {}
        Mock Suspend-BitLocker {}
    }

    It 'queries all target mount points once before restoring protection state' {
        $state = [PSCustomObject]@{
            CommandAvailable = $true
            TimedOutMountPoints = @()
            CaptureIssues = @()
            Volumes = @(
                [PSCustomObject]@{ MountPoint = 'C:'; VolumeType = 'OperatingSystem'; ProtectionStatus = 'On'; VolumeStatus = 'FullyEncrypted' }
                [PSCustomObject]@{ MountPoint = 'D:'; VolumeType = 'OperatingSystem'; ProtectionStatus = 'Off'; VolumeStatus = 'FullyEncrypted' }
            )
        }
        Mock Get-BitLockerVolume {
            @(
                [PSCustomObject]@{ MountPoint = 'C:'; ProtectionStatus = 'Off' }
                [PSCustomObject]@{ MountPoint = 'D:'; ProtectionStatus = 'On' }
            )
        }

        Restore-BitLockerVolumes -State $state

        Should -Invoke Get-BitLockerVolume -Times 1 -Exactly -ParameterFilter {
            @($MountPoint).Count -eq 2 -and $MountPoint -contains 'C:' -and $MountPoint -contains 'D:'
        }
        Should -Invoke Resume-BitLocker -Times 1 -Exactly -ParameterFilter { $MountPoint -eq 'C:' }
        Should -Invoke Suspend-BitLocker -Times 1 -Exactly -ParameterFilter { $MountPoint -eq 'D:' -and $RebootCount -eq 0 }
    }

    It 'resolves every target mount point before performing the first mutation' {
        $state = [PSCustomObject]@{
            CommandAvailable = $true
            TimedOutMountPoints = @()
            CaptureIssues = @()
            Volumes = @(
                [PSCustomObject]@{ MountPoint = 'C:'; VolumeType = 'OperatingSystem'; ProtectionStatus = 'On'; VolumeStatus = 'FullyEncrypted' }
                [PSCustomObject]@{ MountPoint = 'D:'; VolumeType = 'OperatingSystem'; ProtectionStatus = 'Off'; VolumeStatus = 'FullyEncrypted' }
            )
        }
        Mock Get-BitLockerVolume { @([PSCustomObject]@{ MountPoint = 'C:'; ProtectionStatus = 'Off' }) }

        { Restore-BitLockerVolumes -State $state } | Should -Throw "*volume 'D:' was not returned*"
        Should -Invoke Resume-BitLocker -Times 0 -Exactly
        Should -Invoke Suspend-BitLocker -Times 0 -Exactly
    }

    It 'rejects duplicate baseline mount points as incomplete' {
        $state = [PSCustomObject]@{
            CommandAvailable = $true
            TimedOutMountPoints = @()
            CaptureIssues = @()
            Volumes = @(
                [PSCustomObject]@{ MountPoint = 'C:'; ProtectionStatus = 'On' }
                [PSCustomObject]@{ MountPoint = 'C:'; ProtectionStatus = 'Off' }
            )
        }

        Test-BitLockerStateCapturedExactly -State $state | Should -BeFalse
    }
}

Describe 'Capture resource lifetime' {
    It 'releases registered temporary resources when the session completes' {
        Mock Close-UserRegistryTarget {}
        $session = New-CaptureSession -Phase Snapshot
        Register-CaptureSessionResource -Session $session -Kind 'UserRegistryTarget' -Value ([PSCustomObject]@{ Sid = 'S-1-5-21-test' })

        $null = Complete-CaptureSession -Session $session

        $session.Resources.Count | Should -Be 0
        Assert-MockCalled Close-UserRegistryTarget -Times 1 -Exactly
    }
}

Describe 'Defender capture batching' {
    BeforeEach {
        Mock Test-CommandAvailable {
            param($Name)
            $Name -eq 'Get-MpPreference'
        }
        Mock Get-MpPreference {
            [PSCustomObject]@{
                DisableRealtimeMonitoring           = $false
                ExclusionPath                        = @('C:\Tools')
                AttackSurfaceReductionRules_Ids     = @('56a863a9-875e-4185-98a7-b882c64b5ce5')
                AttackSurfaceReductionRules_Actions = @(1)
            }
        }
    }

    It 'uses one Get-MpPreference call for scalar, list, and ASR state' {
        $session = New-CaptureSession -Phase Snapshot

        $scalar = Get-MpPreferencePropertyRawValue -Property 'DisableRealtimeMonitoring' -CaptureSession $session
        $list = Get-MpPreferenceListState -Property 'ExclusionPath' -CaptureSession $session
        $asr = Get-AsrRuleCaptureState -CaptureSession $session

        $scalar | Should -BeFalse
        $list.Items | Should -Contain 'C:\Tools'
        $asr.Rules.Count | Should -Be 1
        Assert-MockCalled Get-MpPreference -Times 1 -Exactly
    }
}

Describe 'NetBIOS mutation batching' {
    BeforeEach {
        Mock Get-CimInstance {
            @(
                New-TestNetworkAdapterInstance -Index 4 -Description 'Ethernet'
                New-TestNetworkAdapterInstance -Index 9 -Description 'Wi-Fi'
            )
        }
        Mock Invoke-CimMethod {}
    }

    It 'uses one indexed CIM query for every permissive adapter mutation' {
        $adapters = @(
            [PSCustomObject]@{ Index = 4; TcpipNetbiosOptions = 0 }
            [PSCustomObject]@{ Index = 9; TcpipNetbiosOptions = 2 }
        )

        Set-Permissive-NetBiosAdapters -Adapters $adapters

        Should -Invoke Get-CimInstance -Times 1 -Exactly -ParameterFilter {
            $ClassName -eq 'Win32_NetworkAdapterConfiguration' -and $Filter -match 'Index = 4' -and $Filter -match 'Index = 9'
        }
        Should -Invoke Invoke-CimMethod -Times 2 -Exactly -ParameterFilter {
            $MethodName -eq 'SetTcpipNetbios' -and $Arguments.TcpipNetbiosOptions -eq 1
        }
    }

    It 'uses one indexed CIM query while restoring distinct adapter options' {
        $adapters = @(
            [PSCustomObject]@{ Index = 4; TcpipNetbiosOptions = 0 }
            [PSCustomObject]@{ Index = 9; TcpipNetbiosOptions = 2 }
        )

        Restore-NetBiosAdapters -Adapters $adapters

        Should -Invoke Get-CimInstance -Times 1 -Exactly
        Should -Invoke Invoke-CimMethod -Times 1 -Exactly -ParameterFilter {
            $InputObject.Index -eq 4 -and $Arguments.TcpipNetbiosOptions -eq 0
        }
        Should -Invoke Invoke-CimMethod -Times 1 -Exactly -ParameterFilter {
            $InputObject.Index -eq 9 -and $Arguments.TcpipNetbiosOptions -eq 2
        }
    }

    It 'resolves every requested adapter before performing the first mutation' {
        Mock Get-CimInstance { @([PSCustomObject]@{ Index = 4 }) }
        $targets = @(
            [PSCustomObject]@{ Index = 4; Option = 1 }
            [PSCustomObject]@{ Index = 9; Option = 1 }
        )

        { Set-NetBiosAdapterOptions -Targets $targets } | Should -Throw '*adapter index 9 was not found*'
        Should -Invoke Invoke-CimMethod -Times 0 -Exactly
    }
}

Describe 'ASR mutation batching' {
    It 'removes all valid configured ASR pairs in one Defender call' {
        $firstId = '56a863a9-875e-4185-98a7-b882c64b5ce5'
        $secondId = 'd4f940ab-401b-4efc-aadc-ad5f3c50688a'
        Mock Get-MpPreference {
            [PSCustomObject]@{
                AttackSurfaceReductionRules_Ids = @($firstId, $secondId)
                AttackSurfaceReductionRules_Actions = @(1, 2)
            }
        }
        Mock Remove-MpPreference {}

        Disable-ConfiguredAsrRules

        Should -Invoke Remove-MpPreference -Times 1 -Exactly -ParameterFilter {
            @($AttackSurfaceReductionRules_Ids).Count -eq 2 -and
            @($AttackSurfaceReductionRules_Actions).Count -eq 2 -and
            $AttackSurfaceReductionRules_Actions[0] -eq 'Enabled' -and
            $AttackSurfaceReductionRules_Actions[1] -eq 'AuditMode'
        }
    }

    It 'marks unsupported actions incomplete and refuses malformed live pairs' {
        $ruleId = '56a863a9-875e-4185-98a7-b882c64b5ce5'
        Mock Get-MpPreference {
            [PSCustomObject]@{
                AttackSurfaceReductionRules_Ids = @($ruleId)
                AttackSurfaceReductionRules_Actions = @('UnsupportedAction')
            }
        }
        Mock Remove-MpPreference {}

        $capture = Get-AsrRuleCaptureState
        @($capture.Rules).Count | Should -Be 0
        @($capture.InvalidEntries).Count | Should -Be 1
        $legacyEntry = [PSCustomObject]@{
            Type = 'AsrRules'; CommandAvailable = $true; Captured = $true
            CurrentValue = @([PSCustomObject]@{ Id = $ruleId; Action = 'UnsupportedAction' })
            InvalidEntries = @()
        }
        (Test-SnapshotEntryCapturedExactly -Entry $legacyEntry) | Should -BeFalse
        { Disable-ConfiguredAsrRules } | Should -Throw '*Cannot safely clear malformed ASR rule entry*'
        Should -Invoke Remove-MpPreference -Times 0 -Exactly
    }

    It 'captures missing actions as invalid and refuses unpaired live arrays' {
        $ruleId = '56a863a9-875e-4185-98a7-b882c64b5ce5'
        Mock Get-MpPreference {
            [PSCustomObject]@{
                AttackSurfaceReductionRules_Ids = @($ruleId)
                AttackSurfaceReductionRules_Actions = @()
            }
        }
        Mock Remove-MpPreference {}

        $capture = Get-AsrRuleCaptureState
        @($capture.Rules).Count | Should -Be 0
        @($capture.InvalidEntries).Count | Should -Be 1
        $capture.InvalidEntries[0].Id | Should -Be $ruleId
        $capture.InvalidEntries[0].Action | Should -BeNullOrEmpty
        { Disable-ConfiguredAsrRules } | Should -Throw '*Defender returned 1 ID(s) and 0 action(s)*'
        Should -Invoke Remove-MpPreference -Times 0 -Exactly
    }
}

Describe 'Defender mutation read batching' {
    It 'shares one preference baseline across independent list and ASR mutation planning' {
        $ruleId = '56a863a9-875e-4185-98a7-b882c64b5ce5'
        Clear-WinDefStateCommandCache
        Mock Get-MpPreference {
            [PSCustomObject]@{
                ExclusionPath = @('C:\Tools')
                AttackSurfaceReductionRules_Ids = @($ruleId)
                AttackSurfaceReductionRules_Actions = @(1)
            }
        }
        Mock Add-MpPreference {}
        Mock Remove-MpPreference {}

        $session = New-CaptureSession -Phase Permissive
        Set-MpPreferenceListValue -Property 'ExclusionPath' -DesiredItems @('C:\Tools') -CaptureSession $session
        Disable-ConfiguredAsrRules -CaptureSession $session

        Should -Invoke Get-MpPreference -Times 1 -Exactly
        Should -Invoke Add-MpPreference -Times 0 -Exactly
        Should -Invoke Remove-MpPreference -Times 1 -Exactly -ParameterFilter {
            @($AttackSurfaceReductionRules_Ids).Count -eq 1 -and
            $AttackSurfaceReductionRules_Ids[0] -eq $ruleId
        }
        $metrics = Complete-CaptureSession -Session $session
        $metrics.ProviderQueryCount | Should -Be 1
        $metrics.CacheHitCount | Should -Be 1
    }

    It 'wires the mutation session through permissive and restore dispatch' {
        (Get-Command Apply-PermissiveDefinition -CommandType Function).Definition |
            Should -Match 'Set-MpPreferenceListValue[^\r\n]+-CaptureSession\s+\$CaptureSession'
        (Get-Command Apply-PermissiveDefinition -CommandType Function).Definition |
            Should -Match 'Disable-ConfiguredAsrRules\s+-CaptureSession\s+\$CaptureSession'
        (Get-Command Restore-SnapshotEntry -CommandType Function).Definition |
            Should -Match 'Restore-AsrRules[^\r\n]+-CaptureSession\s+\$CaptureSession'
    }
}

Describe 'WSMan capture batching' {
    BeforeEach {
        Mock Test-CommandAvailable {
            param($Name)
            $Name -eq 'Get-WSManInstance'
        }
        Mock Get-WSManInstance {
            param($ResourceURI)
            if ($ResourceURI -eq 'winrm/config/service/auth') {
                return [PSCustomObject]@{
                    Basic       = 'true'
                    Kerberos    = 'false'
                    Negotiate   = 'true'
                    Certificate = 'false'
                    CredSSP     = 'false'
                }
            }

            [PSCustomObject]@{
                AllowUnencrypted = 'false'
                IPv4Filter       = '*'
                IPv6Filter       = '*'
                CbtHardeningLevel = 'Strict'
            }
        }
    }

    It 'queries each WSMan resource URI only once' {
        $session = New-CaptureSession -Phase Verification

        $basic = Get-WsManConfigValueState -Path 'WSMan:\localhost\Service\Auth\Basic' -CaptureSession $session
        $kerberos = Get-WsManConfigValueState -Path 'WSMan:\localhost\Service\Auth\Kerberos' -CaptureSession $session
        $allowUnencrypted = Get-WsManConfigValueState -Path 'WSMan:\localhost\Service\AllowUnencrypted' -CaptureSession $session

        $basic.Value | Should -BeTrue
        $kerberos.Value | Should -BeFalse
        $allowUnencrypted.Value | Should -BeFalse
        Assert-MockCalled Get-WSManInstance -Times 2 -Exactly
    }
}

Describe 'SMB capture batching' {
    BeforeEach {
        Mock Invoke-ChildPowerShell {
            [PSCustomObject]@{
                CommandAvailable = $true
                TimedOut         = $false
                ExitCode         = 0
                StdErr           = ''
                StdOut           = "module chatter`r`nWDS_SMB_JSON:{`"Client`":{`"CommandAvailable`":true,`"Captured`":true,`"Error`":null,`"RequireSecuritySignature`":true},`"Server`":{`"CommandAvailable`":true,`"Captured`":true,`"Error`":null,`"RequireSecuritySignature`":false}}"
            }
        }
    }

    It 'captures client and server state through one framed child probe' {
        $session = New-CaptureSession -Phase Snapshot
        $clientDefinition = [PSCustomObject]@{ Id = 'smb.client.require_security_signature'; Type = 'SmbClientConfig'; RequiresReboot = $false }
        $serverDefinition = [PSCustomObject]@{ Id = 'smb.server.require_security_signature'; Type = 'SmbServerConfig'; RequiresReboot = $false }

        $client = Capture-Definition -Definition $clientDefinition -CaptureSession $session
        $server = Capture-Definition -Definition $serverDefinition -CaptureSession $session

        $client.CurrentValue.RequireSecuritySignature | Should -BeTrue
        $server.CurrentValue.RequireSecuritySignature | Should -BeFalse
        Assert-MockCalled Invoke-ChildPowerShell -Times 1 -Exactly
        $session.CacheHitCount | Should -Be 1
    }
}

Describe 'Firewall profile capture batching' {
    BeforeEach {
        Mock Test-CommandAvailable {
            param($Name)
            $Name -eq 'Get-NetFirewallProfile'
        }
        Mock Get-NetFirewallProfile {
            @(
                [PSCustomObject]@{ Name = 'Domain'; Enabled = $true; DefaultInboundAction = 'Block'; DefaultOutboundAction = 'Allow' }
                [PSCustomObject]@{ Name = 'Private'; Enabled = $true; DefaultInboundAction = 'Block'; DefaultOutboundAction = 'Allow' }
                [PSCustomObject]@{ Name = 'Public'; Enabled = $false; DefaultInboundAction = 'Block'; DefaultOutboundAction = 'Allow' }
            )
        }
    }

    It 'queries all requested firewall profiles once' {
        $state = Get-FirewallProfileStates -Profiles @('Domain', 'Private', 'Public')

        @($state.Profiles).Count | Should -Be 3
        @($state.CaptureIssues).Count | Should -Be 0
        (ConvertTo-NullableBoolean -Value (@($state.Profiles | Where-Object { $_.Profile -eq 'Public' })[0].Enabled)) | Should -BeFalse
        Assert-MockCalled Get-NetFirewallProfile -Times 1 -Exactly
    }
}

Describe 'Firewall profile mutation batching' {
    It 'applies the common permissive posture to every profile in one call' {
        Clear-WinDefStateCommandCache
        Mock Set-NetFirewallProfile {}
        $definition = [PSCustomObject]@{
            Profiles = @('Domain', 'Private', 'Public')
            PermissiveValue = $false
            PermissiveDefaultInboundAction = 'Allow'
            PermissiveDefaultOutboundAction = 'Allow'
            PermissiveAllowUnicastResponseToMulticast = $true
            PermissiveNotifyOnListen = $false
            PermissiveLogAllowed = $false
            PermissiveLogBlocked = $false
            PermissiveLogIgnored = $false
        }

        Set-Permissive-FirewallProfiles -Definition $definition

        Should -Invoke Set-NetFirewallProfile -Times 1 -Exactly -ParameterFilter {
            @($Profile).Count -eq 3 -and
            $Profile -contains 'Domain' -and
            $Profile -contains 'Private' -and
            $Profile -contains 'Public'
        }
    }
}

Describe 'Firewall rule restore batching' {
    BeforeEach {
        Clear-WinDefStateCommandCache
        Mock Set-NetFirewallRule {}
    }

    It 'restores any number of captured rules in at most two enabled-state calls' {
        $state = [PSCustomObject]@{
            CommandAvailable = $true
            Group = '@FirewallAPI.dll,-28752'
            CaptureIssues = @()
            Rules = @(
                [PSCustomObject]@{ Name = 'RDP-In-TCP'; Enabled = 'True' }
                [PSCustomObject]@{ Name = 'RDP-In-UDP'; Enabled = 1 }
                [PSCustomObject]@{ Name = 'RDP-Shadow-In-TCP'; Enabled = 'False' }
            )
        }

        Restore-FirewallRules -State $state

        Should -Invoke Set-NetFirewallRule -Times 2 -Exactly
        Should -Invoke Set-NetFirewallRule -Times 1 -Exactly -ParameterFilter {
            $Enabled -eq 'True' -and @($Name).Count -eq 2 -and $Name -contains 'RDP-In-TCP' -and $Name -contains 'RDP-In-UDP'
        }
        Should -Invoke Set-NetFirewallRule -Times 1 -Exactly -ParameterFilter {
            $Enabled -eq 'False' -and @($Name).Count -eq 1 -and $Name[0] -eq 'RDP-Shadow-In-TCP'
        }
    }

    It 'rejects malformed or duplicate rule identities before mutation' {
        $state = [PSCustomObject]@{
            CommandAvailable = $true
            CaptureIssues = @()
            Rules = @(
                [PSCustomObject]@{ Name = 'Duplicate'; Enabled = 'True' }
                [PSCustomObject]@{ Name = 'Duplicate'; Enabled = 'False' }
            )
        }

        Test-FirewallRuleStateCapturedExactly -State $state | Should -BeFalse
        Restore-FirewallRules -State $state
        Should -Invoke Set-NetFirewallRule -Times 0 -Exactly
    }
}

Describe 'Snapshot performance metadata' {
    It 'identifies the exact engine and PowerShell runtime used for capture' {
        $runtimeInfo = Get-WinDefStateRuntimeInfo

        $runtimeInfo.ScriptFileName | Should -Be 'WinDefState.ps1'
        $runtimeInfo.ScriptSha256 | Should -Be (Get-FileHash -LiteralPath $enginePath -Algorithm SHA256).Hash
        $runtimeInfo.PowerShellVersion | Should -Be ([string]$PSVersionTable.PSVersion)
        $runtimeInfo.PowerShellEdition | Should -Not -BeNullOrEmpty
    }

    It 'annotates captured entries with authoritative lifecycle capabilities' {
        $definition = [PSCustomObject]@{
            Id = 'defender.exclusion_paths'; Type = 'MpPreferenceList'; Property = 'ExclusionPath'; RequiresReboot = $false
        }
        $session = New-CaptureSession -Phase Snapshot
        Mock Capture-Definition {
            [PSCustomObject]@{
                Id = $Definition.Id; Type = $Definition.Type; Property = $Definition.Property
                CommandAvailable = $true; Captured = $true; CurrentValue = @(); RequiresReboot = $false
            }
        }

        $entry = Invoke-TimedDefinitionCapture -Definition $definition -CaptureSession $session
        $null = Complete-CaptureSession -Session $session

        $entry.Capabilities.Permissive | Should -BeFalse
        $entry.Capabilities.Restore | Should -BeTrue
        $entry.Capabilities.InventoryOnly | Should -BeFalse
    }

    It 'round-trips schema version 2 capture metrics through the snapshot writer' {
        $snapshotPath = Join-Path $TestDrive 'snapshot.json'
        $snapshot = [PSCustomObject]@{
            SchemaVersion  = 2
            Tool           = 'WinDefState'
            Producer       = Get-WinDefStateRuntimeInfo
            ComputerName   = 'TESTHOST'
            CapturedAtUtc  = '2026-08-09T00:00:00.0000000Z'
            CaptureMetrics = [PSCustomObject]@{
                Phase              = 'Snapshot'
                DurationMs         = 125.5
                ProviderQueryCount = 2
                CacheHitCount      = 4
                ProviderQueries    = @()
                Settings           = @()
            }
            CaptureScope   = [PSCustomObject]@{
                IsFiltered = $true
                IncludeId  = @('rdp.user_authentication')
                ExcludeId  = @()
            }
            Settings       = @()
        }

        Write-SnapshotJsonAtomic -Path $snapshotPath -Snapshot $snapshot
        $saved = Get-Content -LiteralPath $snapshotPath -Raw | ConvertFrom-Json

        $saved.SchemaVersion | Should -Be 2
        $saved.Producer.ScriptSha256 | Should -Be (Get-FileHash -LiteralPath $enginePath -Algorithm SHA256).Hash
        $saved.CaptureMetrics.DurationMs | Should -Be 125.5
        $saved.CaptureMetrics.CacheHitCount | Should -Be 4
        $saved.CaptureScope.IsFiltered | Should -BeTrue
        @($saved.CaptureScope.IncludeId) | Should -Contain 'rdp.user_authentication'
        @($saved.Settings).Count | Should -Be 0
    }
}

Describe 'GUI snapshot row capabilities' {
    It 'disables incomplete settings and explains the capture failure' {
        $entry = [PSCustomObject]@{
            Id = 'winrm.service.basic'; Type = 'WsManValue'; Captured = $false
            CaptureError = 'provider failed'; CommandAvailable = $true; CurrentValue = $null; RequiresReboot = $false
        }

        $row = New-SnapshotRow -Entry $entry

        $row.Current | Should -Match '^Incomplete: provider failed'
        $row.CanRun | Should -BeFalse
        $row.Action | Should -Be 'Unavailable'
    }

    It 'presents Defender runtime state as inventory only' {
        $entry = [PSCustomObject]@{
            Id = 'defender.runtime_status'; Type = 'DefenderRuntimeStatus'; RequiresReboot = $false
            CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; Captured = $true; AMRunningMode = 'Normal' }
        }

        $row = New-SnapshotRow -Entry $entry

        $row.CanRun | Should -BeFalse
        $row.Action | Should -Be 'Inventory only'
        $row.Badges | Should -Match 'Inventory only'
    }

    It 'offers exact Defender list baselines for restore only' {
        $entry = [PSCustomObject]@{
            Id = 'defender.exclusion_paths'; Type = 'MpPreferenceList'; Captured = $true
            CommandAvailable = $true; CurrentValue = @('C:\Tools'); RequiresReboot = $false
        }

        $row = New-SnapshotRow -Entry $entry

        $row.CanRun | Should -BeTrue
        $row.Action | Should -Be 'Restore captured'
        @($row.ActionOptions).Count | Should -Be 1
        @($row.ActionOptions) | Should -Contain 'Restore captured'
    }

    It 'uses persisted engine capabilities instead of guessing from type' {
        $entry = [PSCustomObject]@{
            Id = 'custom.restore_only'; Type = 'RegistryValue'; Captured = $true
            Exists = $true; CurrentValue = 1; RequiresReboot = $false
            Capabilities = [PSCustomObject]@{ Permissive = $false; Restore = $true; InventoryOnly = $false }
        }

        $row = New-SnapshotRow -Entry $entry

        $row.SupportsPermissive | Should -BeFalse
        $row.SupportsRestore | Should -BeTrue
        $row.Action | Should -Be 'Restore captured'
        $row.Badges | Should -Match 'Restore only'
    }

    It 'offers mutable settings both supported actions' {
        $entry = [PSCustomObject]@{
            Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Captured = $true
            Exists = $true; CurrentValue = 1; RequiresReboot = $false
        }

        $row = New-SnapshotRow -Entry $entry

        $row.CanRun | Should -BeTrue
        $row.Action | Should -Be 'Permissive target'
        @($row.ActionOptions).Count | Should -Be 2
        @($row.ActionOptions) | Should -Contain 'Restore captured'
    }

    It 'summarizes a legacy null structured value without throwing' {
        $entry = [PSCustomObject]@{
            Id = 'applocker.service'; Type = 'ServiceConfig'; CurrentValue = $null; RequiresReboot = $false
        }

        { New-SnapshotRow -Entry $entry } | Should -Not -Throw
        $row = New-SnapshotRow -Entry $entry
        $row.Current | Should -Be '<not captured>'
        $row.CanRun | Should -BeFalse
        $row.Action | Should -Be 'Unavailable'
    }

    It 'filters rows by category and every search token' {
        $row = [PSCustomObject]@{
            Category = 'rdp'; Id = 'rdp.allow_clipboard_redirection'; Type = 'RegistryValue'
            Reboot = 'Yes'; Badges = 'Requires reboot'; Current = '0'; Compare = '1'
            Difference = 'Changed'; PermissiveTarget = 'Set to 0'; Action = 'Permissive target'
        }

        (Test-SnapshotRowVisible -Row $row -SearchText 'clipboard reboot' -Category 'RDP') | Should -BeTrue
        (Test-SnapshotRowVisible -Row $row -SearchText 'clipboard missing' -Category 'rdp') | Should -BeFalse
        (Test-SnapshotRowVisible -Row $row -SearchText $null -Category 'defender') | Should -BeFalse
        (Test-SnapshotRowVisible -Row $row -SearchText 'registryvalue' -Category 'All categories') | Should -BeTrue
        (Test-SnapshotRowVisible -Row $row -SearchText 'changed set' -Category 'All categories' -ChangedOnly $true) | Should -BeTrue
        $row.Difference = 'Unchanged'
        (Test-SnapshotRowVisible -Row $row -SearchText $null -Category 'All categories' -ChangedOnly $true) | Should -BeFalse
    }

    It 'summarizes visible, runnable, reboot, and selected rows independently' {
        $rows = @(
            [PSCustomObject]@{ CanRun = $true; Reboot = 'Yes'; Selected = $true; Difference = 'Changed' },
            [PSCustomObject]@{ CanRun = $false; Reboot = 'No'; Selected = $true; Difference = 'Removed' },
            [PSCustomObject]@{ CanRun = $true; Reboot = 'No'; Selected = $false; Difference = 'Added' }
        )

        $summary = Get-SnapshotRowSummary -Rows $rows -VisibleRows @($rows[0], $rows[2])

        $summary.Total | Should -Be 3
        $summary.Visible | Should -Be 2
        $summary.Runnable | Should -Be 2
        $summary.Reboot | Should -Be 1
        $summary.Selected | Should -Be 2
        $summary.Changed | Should -Be 1
        $summary.Added | Should -Be 1
        $summary.Removed | Should -Be 1
    }

    It 'compares snapshot state canonically without treating preview metadata as drift' {
        $first = [PSCustomObject]@{
            Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Exists = $true; CurrentValue = 1
            Capabilities = [PSCustomObject]@{ Permissive = $true; Restore = $true }
            PermissiveTarget = [PSCustomObject]@{ Summary = 'Set to 0.' }
        }
        $second = [PSCustomObject]@{
            PermissiveTarget = [PSCustomObject]@{ Summary = 'Legacy wording.' }
            CurrentValue = 1; Exists = $true; Type = 'RegistryValue'; Id = 'rdp.user_authentication'
            Capabilities = [PSCustomObject]@{ InventoryOnly = $false }
        }

        (Get-GuiEntryFingerprint -Entry $first) | Should -Be (Get-GuiEntryFingerprint -Entry $second)
        $second.CurrentValue = 0
        (Get-GuiEntryFingerprint -Entry $first) | Should -Not -Be (Get-GuiEntryFingerprint -Entry $second)
    }

    It 'builds changed, added, and removed rows across two snapshots' {
        $script:Rows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
        $script:Snapshot = [PSCustomObject]@{ Settings = @(
            [PSCustomObject]@{ Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Captured = $true; Exists = $true; CurrentValue = 0; RequiresReboot = $false },
            [PSCustomObject]@{ Id = 'rdp.allow_connections'; Type = 'RegistryValue'; Captured = $true; Exists = $true; CurrentValue = 0; RequiresReboot = $false }
        ) }
        $script:ComparisonSnapshot = [PSCustomObject]@{ Settings = @(
            [PSCustomObject]@{ Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Captured = $true; Exists = $true; CurrentValue = 1; RequiresReboot = $false },
            [PSCustomObject]@{ Id = 'defender.pua_protection'; Type = 'MpPreferenceValue'; Captured = $true; CurrentValue = 'Enabled'; RequiresReboot = $false }
        ) }

        Rebuild-SnapshotRows

        @($script:Rows | Where-Object Difference -eq 'Changed').Id | Should -Contain 'rdp.user_authentication'
        @($script:Rows | Where-Object Difference -eq 'Added').Id | Should -Contain 'rdp.allow_connections'
        $removed = @($script:Rows | Where-Object Difference -eq 'Removed')
        $removed.Id | Should -Contain 'defender.pua_protection'
        $removed[0].CanRun | Should -BeFalse
    }

    It 'builds a scoped pre-change preview from persisted target descriptors' {
        $snapshot = [PSCustomObject]@{ Settings = @(
            [PSCustomObject]@{
                Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Captured = $true; Exists = $true
                CurrentValue = 1; RequiresReboot = $false
                Capabilities = [PSCustomObject]@{ Permissive = $true; Restore = $true; InventoryOnly = $false }
                PermissiveTarget = [PSCustomObject]@{ Mode = 'Exact'; Summary = 'Set UserAuthentication to 0 (DWord).' }
            },
            [PSCustomObject]@{
                Id = 'defender.runtime_status'; Type = 'DefenderRuntimeStatus'; RequiresReboot = $false
                CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; Captured = $true; AMRunningMode = 'Normal' }
                Capabilities = [PSCustomObject]@{ Permissive = $false; Restore = $false; InventoryOnly = $true }
            }
        ) }

        $rows = @(Get-MutationPreviewRows -Snapshot $snapshot -Command Permissive -IncludeId 'rdp.user_authentication')

        $rows.Count | Should -Be 1
        $rows[0].Id | Should -Be 'rdp.user_authentication'
        $rows[0].Target | Should -Be 'Set UserAuthentication to 0 (DWord).'
        $rows[0].Recovery | Should -Be '1'
    }
}

Describe 'GUI process argument quoting' {
    It 'quotes spaces, embedded quotes, and trailing backslashes for powershell.exe' {
        (ConvertTo-WindowsCommandLineArgument -Argument 'plain') | Should -Be 'plain'
        (ConvertTo-WindowsCommandLineArgument -Argument 'C:\Tools With Spaces\WinDefState.ps1') | Should -Be '"C:\Tools With Spaces\WinDefState.ps1"'
        (ConvertTo-WindowsCommandLineArgument -Argument 'C:\Tools With Spaces\') | Should -Be '"C:\Tools With Spaces\\"'
        (ConvertTo-WindowsCommandLineArgument -Argument 'a"b') | Should -Be '"a\"b"'
    }

    It 'uses structured progress instead of global verbose output' {
        $guiArgumentsFunction.Extent.Text | Should -Match "'-EmitProgress'"
        $guiArgumentsFunction.Extent.Text | Should -Match "'-CancellationPath'"
        $guiArgumentsFunction.Extent.Text | Should -Match "'-MutationApprovalPath'"
        $guiArgumentsFunction.Extent.Text | Should -Not -Match "'-Verbose'"
    }

    It 'passes a protected one-time mutation approval path to the engine' {
        $approvalPath = 'C:\ProgramData\WinDefState\.review-0123456789abcdef0123456789abcdef.approve'
        $arguments = @(Get-WinDefStateArguments -Command Restore -SnapshotPath 'C:\State\baseline.json' -MutationApprovalPath $approvalPath)
        $approvalIndex = [Array]::IndexOf($arguments, '-MutationApprovalPath')

        $approvalIndex | Should -BeGreaterThan -1
        $arguments[$approvalIndex + 1] | Should -Be $approvalPath
        $guiText = Get-Content -LiteralPath $guiPath -Raw
        $guiText | Should -Match 'Join-Path\s+\$script:StateRoot\s+\("\.review-'
        $guiText | Should -Match 'WDS_REVIEW\\\|'
        $guiText | Should -Match 'Show-MutationPreview'
    }

    It 'passes selected IDs as one native-process argument' {
        $arguments = @(Get-WinDefStateArguments -Command Permissive -IncludeId @(
            'defender.enable_network_protection',
            'rdp.user_authentication'
        ))
        $includeIndex = [Array]::IndexOf($arguments, '-IncludeId')

        $includeIndex | Should -BeGreaterThan -1
        $arguments[$includeIndex + 1] | Should -Be 'defender.enable_network_protection,rdp.user_authentication'
        $arguments | Should -Not -Contain 'rdp.user_authentication'
        @(Get-NormalizedIdFilter -Ids $arguments[$includeIndex + 1]) | Should -Be @(
            'defender.enable_network_protection',
            'rdp.user_authentication'
        )
    }

    It 'uses cooperative safe-boundary cancellation without killing the engine' {
        $guiText = Get-Content -LiteralPath $guiPath -Raw

        $guiText | Should -Match 'x:Name="CancelButton"'
        $guiText | Should -Match 'WDS_CANCELLED'
        $guiText | Should -Match 'CancellationObserved'
        $guiText | Should -Not -Match '\.Kill\('
    }

    It 'loads the exact snapshot path reported by the child operation' {
        $operation = [PSCustomObject]@{ Results = @{ SnapshotPath = 'C:\State\exact.json' } }
        Get-GuiOperationResultValue -Operation $operation -Name SnapshotPath | Should -Be 'C:\State\exact.json'

        $guiText = Get-Content -LiteralPath $guiPath -Raw
        $guiText | Should -Match 'WDS_RESULT\\\|'
        $guiText | Should -Match '\$script:ActiveOperation\.Results\[\$name\]\s*=\s*\$value'
        $guiText | Should -Match 'Load-CompletedOperationSnapshot\s+-Operation\s+\$completedOperation'
    }

    It 'renders reboot as read-only text rather than an editable checkbox' {
        $guiText = Get-Content -LiteralPath $guiPath -Raw

        $guiText | Should -Match '<DataGridTextColumn Header="REBOOT"[^>]+IsReadOnly="True"'
        $guiText | Should -Not -Match '<DataGridCheckBoxColumn Header="REBOOT"'
    }

    It 'uses the polished runbook, explorer filters, and virtualized grid shell' {
        $guiText = Get-Content -LiteralPath $guiPath -Raw

        $guiText | Should -Match 'Text="Three deliberate operations"'
        $guiText | Should -Match 'FontFamily="Bahnschrift SemiCondensed"'
        $guiText | Should -Match 'x:Name="SearchBox"'
        $guiText | Should -Match 'x:Name="CategoryFilter"'
        $guiText | Should -Match 'x:Name="VisibleCountText"'
        $guiText | Should -Match 'x:Name="HistoryCombo"'
        $guiText | Should -Match 'x:Name="CompareButton"'
        $guiText | Should -Match 'x:Name="ChangedOnlyCheckBox"'
        $guiText | Should -Match 'Header="CHANGE"'
        $guiText | Should -Match 'Header="COMPARE"'
        $guiText | Should -Match 'Get-Content\s+-LiteralPath\s+\$file\.FullName\s+-TotalCount\s+40'
        $guiText | Should -Match '\$script:SnapshotCache\.ContainsKey\(\$cacheKey\)'
        $guiText | Should -Match 'VirtualizingPanel\.VirtualizationMode="Recycling"'
        $guiText | Should -Match '<GridSplitter '
        $guiText | Should -Match '\$LogBox\.Text\.Length\s+-gt\s+600000'
    }

    It 'preserves visible keyboard focus and accessible operation labels' {
        $guiText = Get-Content -LiteralPath $guiPath -Raw

        $guiText | Should -Not -Match 'FocusVisualStyle" Value="\{x:Null\}"'
        $guiText | Should -Match 'x:Name="SnapshotButton" AutomationProperties.Name="Snapshot only"'
        $guiText | Should -Match 'x:Name="PermissiveButton" AutomationProperties.Name="Snapshot and apply permissive settings"'
        $guiText | Should -Match 'x:Name="RestoreButton" AutomationProperties.Name="Restore baseline"'
        $guiText | Should -Match 'x:Name="SearchBox" AutomationProperties.Name="Search settings"'
        $guiText | Should -Match 'x:Name="HistoryCombo" AutomationProperties.Name="Snapshot comparison history"'
        $guiText | Should -Match 'x:Name="RunSelectedButton"[^>]+AutomationProperties.Name="Run selected setting actions"'
    }

    It 'instantiates the WPF tree noninteractively in Windows test and release gates' {
        $repositoryRoot = Split-Path -Parent $PSScriptRoot
        $guiText = Get-Content -LiteralPath $guiPath -Raw
        $testWorkflow = Get-Content -LiteralPath (Join-Path $repositoryRoot '.github/workflows/test.yml') -Raw
        $releaseWorkflow = Get-Content -LiteralPath (Join-Path $repositoryRoot '.github/workflows/release.yml') -Raw
        $snapshotWorkflow = Get-Content -LiteralPath (Join-Path $repositoryRoot '.github/workflows/windows-snapshot.yml') -Raw

        $guiText | Should -Match '\[switch\]\$ValidateOnly'
        $guiText | Should -Match '\$requiredControls\s*=\s*\[ordered\]@\{'
        $guiText | Should -Match 'The WPF visual tree is missing required control'
        $guiText | Should -Match 'New-MutationPreviewWindow'
        $guiText | Should -Match 'mutation-preview visual tree is missing required control'
        $testWorkflow | Should -Match 'WinDefState\.Gui\.ps1 -ValidateOnly'
        $releaseWorkflow | Should -Match 'WinDefState\.Gui\.ps1 -ValidateOnly'
        $testWorkflow | Should -Match 'windows-2022'
        $testWorkflow | Should -Match 'windows-2025'
        $testWorkflow | Should -Not -Match 'actions/checkout@v4'
        $releaseWorkflow | Should -Not -Match 'actions/checkout@v4'
        $snapshotWorkflow | Should -Not -Match 'actions/checkout@v4'
    }

    It 'disables selection and actions for rows that cannot run' {
        $guiText = Get-Content -LiteralPath $guiPath -Raw

        @([regex]::Matches($guiText, 'IsEnabled="\{Binding CanRun\}"')).Count | Should -Be 2
        $guiText | Should -Match 'Content="Select permissive"'
        $guiText | Should -Match '\$supportsPermissive\s*=\s*\[bool\]\$row\.CanRun'
        $guiText | Should -Match '\$row\.Selected\s*=\s*\$supportsPermissive'
        $guiText | Should -Match '\$unavailableRows\s*=\s*@\('
    }

    It 'disables a second permissive action while an operation journal exists' {
        $guiText = Get-Content -LiteralPath $guiPath -Raw

        $guiText | Should -Match '\$PermissiveButton\.IsEnabled\s*=\s*-not\s+\$Running\s+-and\s+-not\s+\$journalExists'
        $guiText | Should -Match '\$RestoreButton\.IsEnabled\s*=\s*-not\s+\$Running'
    }

    It 'badges nested incomplete provider state' {
        $guiText = Get-Content -LiteralPath $guiPath -Raw

        $guiText | Should -Match '\$state\.PSObject\.Properties\[''Captured''\].*-not\s+\[bool\]\$state\.Captured'
        $guiText | Should -Match '\$state\.PSObject\.Properties\[''CommandAvailable''\].*-not\s+\[bool\]\$state\.CommandAvailable'
    }
}

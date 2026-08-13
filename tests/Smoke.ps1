[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$enginePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'WinDefState.ps1'
Write-Host '[smoke 1/6] Loading engine and parsing GUI'
. $enginePath

$repositoryRoot = Split-Path -Parent $PSScriptRoot
$guiPath = Join-Path $repositoryRoot 'WinDefState.Gui.ps1'
$guiTokens = $null
$guiParseErrors = $null
$guiAst = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$guiTokens, [ref]$guiParseErrors)
if ($guiParseErrors.Count -gt 0) {
    throw ($guiParseErrors.Message -join [Environment]::NewLine)
}
$xamlAssignments = @($guiAst.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $node.Left.Extent.Text -in @('$xaml', '$script:MutationPreviewXaml')
}, $true))
if ($xamlAssignments.Count -ne 2) {
    throw "Expected two embedded GUI XAML assignments, found $($xamlAssignments.Count)."
}
foreach ($xamlAssignment in $xamlAssignments) {
    $xamlExtentLines = @($xamlAssignment.Right.Extent.Text -split '\r?\n')
    $xamlText = @($xamlExtentLines[1..($xamlExtentLines.Count - 2)]) -join [Environment]::NewLine
    try {
        $xamlDocument = [xml]$xamlText
    } catch {
        throw "Embedded GUI XAML '$($xamlAssignment.Left.Extent.Text)' is not valid XML: $($_.Exception.Message)"
    }
    if ([string]$xamlDocument.DocumentElement.LocalName -ne 'Window') {
        throw "Expected the GUI XAML root to be Window, received '$($xamlDocument.DocumentElement.LocalName)'."
    }
    $xamlNames = @([regex]::Matches($xamlText, 'x:Name="([^"]+)"') | ForEach-Object { [string]$_.Groups[1].Value })
    $duplicateXamlNames = @($xamlNames | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { [string]$_.Name })
    if ($duplicateXamlNames.Count -gt 0) {
        throw "GUI XAML '$($xamlAssignment.Left.Extent.Text)' contains duplicate x:Name values: $($duplicateXamlNames -join ', ')"
    }
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

$engineText = Get-Content -LiteralPath $enginePath -Raw
$guiText = Get-Content -LiteralPath $guiPath -Raw
$releaseBuilderPath = Join-Path $repositoryRoot 'build/Build-Release.ps1'
$releaseBuilderTokens = $null
$releaseBuilderParseErrors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($releaseBuilderPath, [ref]$releaseBuilderTokens, [ref]$releaseBuilderParseErrors)
if ($releaseBuilderParseErrors.Count -gt 0) {
    throw ($releaseBuilderParseErrors.Message -join [Environment]::NewLine)
}
$releaseBuilderText = Get-Content -LiteralPath $releaseBuilderPath -Raw
$windowsSnapshotText = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'WindowsSnapshot.ps1') -Raw
$engineRegionNames = @(
    'Core runtime, persistence, and operation state'
    'Registry, service, and user-profile providers'
    'Remote management and network providers'
    'Defender provider'
    'BitLocker provider'
    'Application control, exploit protection, and audit providers'
    'Firewall and service mutation providers'
    'Canonicalization, reporting, and verification'
    'Definition catalog and lifecycle dispatch'
    'Public orchestration'
    'Script entry point'
)

$collectorTypeCommand = $guiAst.Find({
    param($node)
    $node -is [System.Management.Automation.Language.CommandAst] -and $node.Extent.Text -match '^Add-Type\s+-TypeDefinition'
}, $true)
if ($null -eq $collectorTypeCommand) {
    throw 'The GUI process collector type definition was not found.'
}
Write-Host '[smoke 2/6] Validating GUI process collector'
$collectorSource = $collectorTypeCommand.CommandElements[-1].Value
if ($env:OS -eq 'Windows_NT') {
    Add-Type -TypeDefinition $collectorSource
    if ($null -eq ('WinDefState.Gui.ProcessOutputCollector' -as [type])) {
        throw 'The GUI process collector type did not compile.'
    }
} elseif ($collectorSource -notmatch 'class\s+ProcessOutputCollector') {
    throw 'The GUI process collector source declaration was not found.'
}

function Assert-SmokeEqual {
    param(
        [Parameter(Mandatory)] [AllowNull()] [object]$Actual,
        [Parameter(Mandatory)] [AllowNull()] [object]$Expected,
        [Parameter(Mandatory)] [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected', received '$Actual'."
    }
}

Write-Host '[smoke 3/6] Checking static contracts and validation rules'
Assert-SmokeEqual -Actual (ConvertTo-WindowsCommandLineArgument -Argument 'plain') -Expected 'plain' -Message 'Plain GUI argument quoting failed.'
Assert-SmokeEqual -Actual (ConvertTo-WindowsCommandLineArgument -Argument 'C:\Tools With Spaces\WinDefState.ps1') -Expected '"C:\Tools With Spaces\WinDefState.ps1"' -Message 'Spaced GUI argument quoting failed.'
Assert-SmokeEqual -Actual (ConvertTo-WindowsCommandLineArgument -Argument 'C:\Tools With Spaces\') -Expected '"C:\Tools With Spaces\\"' -Message 'Trailing slash GUI argument quoting failed.'
Assert-SmokeEqual -Actual (ConvertTo-WindowsCommandLineArgument -Argument 'a"b') -Expected '"a\"b"' -Message 'Embedded quote GUI argument quoting failed.'
Assert-SmokeEqual -Actual ($guiArgumentsFunction.Extent.Text -match "'-EmitProgress'") -Expected $true -Message 'The GUI does not request structured progress.'
Assert-SmokeEqual -Actual ($guiArgumentsFunction.Extent.Text -match "'-CancellationPath'") -Expected $true -Message 'The GUI does not provide a cooperative cancellation marker.'
Assert-SmokeEqual -Actual ($guiArgumentsFunction.Extent.Text -match "'-MutationApprovalPath'") -Expected $true -Message 'The GUI does not provide a protected pre-change approval marker.'
Assert-SmokeEqual -Actual ($guiArgumentsFunction.Extent.Text -match "'-Verbose'") -Expected $false -Message 'The GUI still enables noisy global verbose output.'
$selectedArguments = @(Get-WinDefStateArguments -Command Permissive -IncludeId @('defender.enable_network_protection', 'rdp.user_authentication'))
$selectedIncludeIndex = [Array]::IndexOf($selectedArguments, '-IncludeId')
Assert-SmokeEqual -Actual ($selectedIncludeIndex -gt -1) -Expected $true -Message 'The GUI omitted its selected-ID filter.'
Assert-SmokeEqual -Actual $selectedArguments[$selectedIncludeIndex + 1] -Expected 'defender.enable_network_protection,rdp.user_authentication' -Message 'The GUI does not transport multiple IDs as one native-process argument.'
Assert-SmokeEqual -Actual (@(Get-NormalizedIdFilter -Ids $selectedArguments[$selectedIncludeIndex + 1]).Count) -Expected 2 -Message 'The engine cannot recover the GUI selected-ID transport token.'
$approvalArguments = @(Get-WinDefStateArguments -Command Restore -SnapshotPath 'C:\State\baseline.json' -MutationApprovalPath 'C:\ProgramData\WinDefState\.review-0123456789abcdef0123456789abcdef.approve')
$approvalIndex = [Array]::IndexOf($approvalArguments, '-MutationApprovalPath')
Assert-SmokeEqual -Actual ($approvalIndex -gt -1) -Expected $true -Message 'The GUI omitted its mutation approval path.'
Assert-SmokeEqual -Actual $approvalArguments[$approvalIndex + 1] -Expected 'C:\ProgramData\WinDefState\.review-0123456789abcdef0123456789abcdef.approve' -Message 'The GUI changed its mutation approval path in transport.'
Assert-SmokeEqual -Actual ($guiText -notmatch '\.Kill\(') -Expected $true -Message 'The GUI can terminate the engine during an unsafe mutation phase.'
Assert-SmokeEqual -Actual ($guiText -match '\$PermissiveButton\.IsEnabled\s*=\s*-not\s+\$Running\s+-and\s+-not\s+\$journalExists') -Expected $true -Message 'GUI permits a second permissive run while a journal exists.'
Assert-SmokeEqual -Actual ($guiText -match '\$state\.PSObject\.Properties\[''Captured''\].*-not\s+\[bool\]\$state\.Captured') -Expected $true -Message 'GUI does not flag nested incomplete provider state.'
Assert-SmokeEqual -Actual ((@([regex]::Matches($guiText, 'IsEnabled="\{Binding CanRun\}"'))).Count) -Expected 2 -Message 'GUI does not disable both controls for unavailable rows.'
Assert-SmokeEqual -Actual ($guiText -match 'Content="Select permissive"') -Expected $true -Message 'GUI bulk selection does not communicate its permissive-only scope.'
Assert-SmokeEqual -Actual ($guiText -match '\$supportsPermissive\s*=\s*\[bool\]\$row\.CanRun') -Expected $true -Message 'GUI permissive selection includes unavailable or restore-only rows.'
Assert-SmokeEqual -Actual ($guiText -match '\$row\.Selected\s*=\s*\$supportsPermissive') -Expected $true -Message 'GUI bulk selection does not enforce permissive capability.'
Assert-SmokeEqual -Actual ($guiText -match '\$unavailableRows\s*=\s*@\(') -Expected $true -Message 'GUI selected-action handler does not reject unavailable rows.'
Assert-SmokeEqual -Actual ($guiText -match 'WDS_REVIEW\\\|') -Expected $true -Message 'GUI does not consume the engine pre-change review handshake.'
Assert-SmokeEqual -Actual ($guiText -match 'x:Name="HistoryCombo"') -Expected $true -Message 'GUI snapshot history is missing.'
Assert-SmokeEqual -Actual ($guiText -match 'x:Name="ChangedOnlyCheckBox"') -Expected $true -Message 'GUI changed-only comparison filter is missing.'
Write-Host '[smoke 3/6] GUI command contracts passed'

$childProcessFunction = (Get-Command Invoke-ChildPowerShell -CommandType Function).Definition
$asyncReadIndex = $childProcessFunction.IndexOf('ReadToEndAsync')
$timedWaitIndex = $childProcessFunction.IndexOf('WaitForExit($TimeoutSeconds * 1000)')
Assert-SmokeEqual -Actual ($asyncReadIndex -ge 0 -and $timedWaitIndex -gt $asyncReadIndex) -Expected $true -Message 'Child provider output is not drained before waiting.'
Assert-SmokeEqual -Actual ($childProcessFunction -match 'finally\s*\{\s*\$process\.Dispose\(\)') -Expected $true -Message 'Child provider process is not disposed deterministically.'
$verificationFunction = (Get-Command Test-DefenseSnapshot -CommandType Function).Definition
Assert-SmokeEqual -Actual ($verificationFunction -match 'finally\s*\{\s*\$captureMetrics\s*=\s*Complete-CaptureSession') -Expected $true -Message 'Verification capture resources are not released deterministically.'
$baselineClassificationIndex = $verificationFunction.IndexOf('$baselineComplete = Test-SnapshotEntryCapturedExactly')
$baselineCanonicalizationIndex = $verificationFunction.IndexOf('$expectedComparable = ConvertTo-ComparableSnapshotEntry')
Assert-SmokeEqual -Actual ($baselineClassificationIndex -ge 0 -and $baselineCanonicalizationIndex -gt $baselineClassificationIndex) -Expected $true -Message 'Verification canonicalizes snapshot entries before classifying incomplete baselines.'
foreach ($atomicWriterName in @('Write-JsonAtomic', 'Write-SnapshotJsonAtomic', 'Write-TextAtomic', 'Write-BytesAtomic')) {
    $atomicWriter = (Get-Command $atomicWriterName -CommandType Function).Definition
    Assert-SmokeEqual -Actual ($atomicWriter -match 'finally[\s\S]*Remove-Item\s+-LiteralPath\s+\$tempPath') -Expected $true -Message "$atomicWriterName does not clean a pre-publish temporary file."
}
Write-Host '[smoke 3/6] Child-process contracts passed'

Assert-SmokeEqual -Actual ($engineText -match '\$operationLock\s*=\s*Enter-WinDefStateOperationLock') -Expected $true -Message 'Public commands do not acquire the machine-wide operation lock.'
Assert-SmokeEqual -Actual ($engineText -match 'finally\s*\{\s*Exit-WinDefStateOperationLock\s+-Lock\s+\$operationLock') -Expected $true -Message 'Public commands do not release the operation lock in a finally block.'
Assert-SmokeEqual -Actual ($engineText -match 'CmdletBinding\(SupportsShouldProcess\s*=\s*\$true') -Expected $true -Message 'Public commands do not support -WhatIf and -Confirm.'
$shouldProcessCount = (@([regex]::Matches($engineText, '\$PSCmdlet\.ShouldProcess\('))).Count
Assert-SmokeEqual -Actual $shouldProcessCount -Expected 3 -Message 'One or more public command paths bypass ShouldProcess.'
Assert-SmokeEqual -Actual ($engineText -match "Join-Path\s+\`$env:ProgramData\s+'WinDefState'") -Expected $true -Message 'Engine state root is not stable across download folders.'
Assert-SmokeEqual -Actual ($guiText -match "Join-Path\s+\`$env:ProgramData\s+'WinDefState'") -Expected $true -Message 'GUI and engine state roots do not share the ProgramData default.'
Assert-SmokeEqual -Actual ($engineText -match 'Protect-StateRoot\s+-Path\s+\$StateRoot') -Expected $true -Message 'Public commands do not secure the trusted state root.'
Assert-SmokeEqual -Actual ($engineText -match 'Clear-SnapshotAssetCache\s*\r?\n\s*\$operationLock') -Expected $true -Message 'Public operations can reuse stale snapshot sidecar content.'
Assert-SmokeEqual -Actual ((Get-Command Protect-StateRoot -CommandType Function).Definition -match 'SetAccessRuleProtection\(\$true,\s*\$false\)') -Expected $true -Message 'State-root ACL still inherits broad parent permissions.'
Clear-WinDefStateCommandCache
$firstCommandResolution = Get-WinDefStateCommand -Name 'Get-Command'
$secondCommandResolution = Get-WinDefStateCommand -Name 'GET-COMMAND'
Assert-SmokeEqual -Actual ($null -ne $firstCommandResolution) -Expected $true -Message 'Command discovery could not resolve a core command.'
Assert-SmokeEqual -Actual ([object]::ReferenceEquals($firstCommandResolution, $secondCommandResolution)) -Expected $true -Message 'Command discovery did not reuse its cached command object.'
Assert-SmokeEqual -Actual $script:WinDefStateCommandCache.Count -Expected 1 -Message 'Command discovery cache key normalization failed.'
Assert-SmokeEqual -Actual ($engineText -match '(?i)\bauditpol(?:\.exe)?\b') -Expected $false -Message 'Audit policy capture still depends on locale-sensitive auditpol text.'
$processCreationGuid = Resolve-AuditSubcategoryGuid -Subcategory 'Process Creation'
Assert-SmokeEqual -Actual $processCreationGuid.ToString() -Expected '0cce922b-69ae-11d9-bed3-505054503030' -Message 'Legacy audit subcategory GUID resolution failed.'
Write-Host '[smoke 3/6] Operation and state-root contracts passed'
$noisyCommandProbes = @(Get-Content -LiteralPath $enginePath | Where-Object {
    $_ -match '\bGet-Command\b' -and $_ -notmatch '-Verbose:\$false'
})
Assert-SmokeEqual -Actual $noisyCommandProbes.Count -Expected 0 -Message 'A capability probe can leak module auto-import chatter into verbose output.'
Write-Host '[smoke 3/6] Diagnostic-output contracts passed'

if ($env:OS -eq 'Windows_NT') {
    $script:EmitProgress = $true
    try {
        $progressOutput = @(Write-OperationProgress -Phase 'Capture' -Current 3 -Total 94 -Id 'defender.test' 6>&1 | ForEach-Object { [string]$_ })
        Assert-SmokeEqual -Actual ($progressOutput -contains 'WDS_PROGRESS|Capture|3|94|defender.test') -Expected $true -Message 'Structured progress output failed.'
    } finally {
        $script:EmitProgress = $false
    }
} else {
    $progressFunction = (Get-Command Write-OperationProgress -CommandType Function).Definition
    Assert-SmokeEqual -Actual ($progressFunction -match 'WDS_PROGRESS\|\{0\}\|\{1\}\|\{2\}\|\{3\}') -Expected $true -Message 'Structured progress format changed.'
}
$resultFunction = (Get-Command Write-OperationResult -CommandType Function).Definition
Assert-SmokeEqual -Actual ($resultFunction -match 'WDS_RESULT\|\{0\}\|\{1\}') -Expected $true -Message 'Structured result format changed.'
Assert-SmokeEqual -Actual ((Get-Command Export-DefenseSnapshot -CommandType Function).Definition -match "Write-OperationResult\s+-Name\s+'SnapshotPath'\s+-Value\s+\`$fullPath") -Expected $true -Message 'Snapshot persistence does not publish its exact result path.'
Assert-SmokeEqual -Actual ($windowsSnapshotText -match 'InvocationDurationMs') -Expected $true -Message 'Native Windows integration does not record end-to-end command duration.'
Assert-SmokeEqual -Actual ($windowsSnapshotText -match 'NonCaptureOverheadMs') -Expected $true -Message 'Native Windows integration cannot distinguish provider capture from orchestration overhead.'
Assert-SmokeEqual -Actual ($releaseBuilderText -match "1980-01-01T00:00:00\+00:00") -Expected $true -Message 'Release archives do not use a reproducible fixed timestamp.'
Assert-SmokeEqual -Actual ($releaseBuilderText -match 'CompressionLevel\]::NoCompression') -Expected $true -Message 'Release archive compression can vary across build runtimes.'
Assert-SmokeEqual -Actual ($releaseBuilderText -match 'Release target already exists') -Expected $true -Message 'Release builds can overwrite an existing artifact set.'
$previousRegionIndex = -1
foreach ($regionName in $engineRegionNames) {
    $regionIndex = $engineText.IndexOf("#region $regionName", [System.StringComparison]::Ordinal)
    Assert-SmokeEqual -Actual ($regionIndex -gt $previousRegionIndex) -Expected $true -Message "Engine source region is missing or out of order: $regionName"
    $previousRegionIndex = $regionIndex
}
Assert-SmokeEqual -Actual ((@([regex]::Matches($engineText, '(?m)^#region '))).Count) -Expected $engineRegionNames.Count -Message 'Engine source has an unregistered region boundary.'
Assert-SmokeEqual -Actual ((@([regex]::Matches($engineText, '(?m)^#endregion\s*$'))).Count) -Expected $engineRegionNames.Count -Message 'Engine source region boundaries are unbalanced.'
Write-Host '[smoke 3/6] Structured progress contract passed'

$summaryLines = @(Get-ReportSummaryLines -Lines @('Heading', 'Value: 1', ' ', 'Detail'))
Assert-SmokeEqual -Actual $summaryLines.Count -Expected 2 -Message 'Console report summary truncation failed.'
Assert-SmokeEqual -Actual $summaryLines[1] -Expected 'Value: 1' -Message 'Console report summary content changed.'
Write-Host '[smoke 3/6] Report summary contract passed'

$definitions = @(Get-DefenseDefinitions)
$definitionsAgain = @(Get-DefenseDefinitions)
$definitionMap = Get-DefenseDefinitionMap
$definitionMapAgain = Get-DefenseDefinitionMap
$uniqueDefinitionIds = @($definitions.Id | Sort-Object -Unique)
Assert-SmokeEqual -Actual $definitions.Count -Expected 94 -Message 'The defense definition count changed unexpectedly.'
Assert-SmokeEqual -Actual $uniqueDefinitionIds.Count -Expected $definitions.Count -Message 'Defense definition IDs are not unique.'
Assert-SmokeEqual -Actual ([object]::ReferenceEquals($definitions[0], $definitionsAgain[0])) -Expected $true -Message 'Definition catalog objects are rebuilt on repeated reads.'
Assert-SmokeEqual -Actual ([object]::ReferenceEquals($definitionMap, $definitionMapAgain)) -Expected $true -Message 'Definition ID map is rebuilt on repeated reads.'
$selectedDefinitions = @(Get-SelectedDefenseDefinitions -IncludeId @('rdp.user_authentication', 'winrm.service'))
Assert-SmokeEqual -Actual $selectedDefinitions.Count -Expected 2 -Message 'Selected definition filtering did not happen before capture.'
Assert-SmokeEqual -Actual (@($selectedDefinitions.Id) -contains 'rdp.user_authentication') -Expected $true -Message 'Selected RDP definition was lost.'
$wildcardDefinitions = @(Get-SelectedDefenseDefinitions -IncludeId 'rdp.*' -ExcludeId 'rdp.firewall_rules')
Assert-SmokeEqual -Actual $wildcardDefinitions.Count -Expected 8 -Message 'Wildcard definition filtering did not resolve the RDP category safely.'
$categoryFilters = @(Merge-SettingIdFilter -Id $null -Category @('defender', 'firewall'))
Assert-SmokeEqual -Actual (@($categoryFilters) -contains 'defender.*') -Expected $true -Message 'Category filtering did not resolve Defender to a stable ID pattern.'
Assert-SmokeEqual -Actual ($engineText -match '\$IncludeId\s*=\s*@\(Merge-SettingIdFilter\s+-Id\s+\$IncludeId\s+-Category\s+\$IncludeCategory\)') -Expected $true -Message 'The public CLI does not merge category scopes before orchestration.'
Assert-SmokeEqual -Actual ((Get-Command Set-DefensePermissive -CommandType Function).Definition -match 'Export-DefenseSnapshot\s+-Path\s+\$Path\s+-IncludeId\s+\$IncludeId\s+-ExcludeId\s+\$ExcludeId') -Expected $true -Message 'Permissive mode does not pass its selected scope into snapshot export.'
Write-Host '[smoke 3/6] Definition catalog contracts passed'
$definitionsById = @{}
foreach ($definition in $definitions) {
    $definitionsById[[string]$definition.Id] = $definition
}
$definitionTypes = @($definitions | ForEach-Object { [string]$_.Type } | Sort-Object -Unique)
$dispatchGaps = @()
foreach ($dispatchFunction in @('Capture-Definition', 'Apply-PermissiveDefinition', 'Restore-SnapshotEntry', 'Add-SnapshotEntryReportLines', 'ConvertTo-ComparableSnapshotEntry', 'Test-SnapshotEntryCapturedExactly', 'Test-DefinitionHasRestoreAction', 'Test-PermissiveDefinitionState')) {
    $functionText = (Get-Command -Name $dispatchFunction -CommandType Function).Definition
    foreach ($definitionType in $definitionTypes) {
        if ($functionText -notmatch ("(?m)^\s*'{0}'\s*\{{" -f [regex]::Escape($definitionType))) {
            $dispatchGaps += "$dispatchFunction -> $definitionType"
        }
    }
}
if ($dispatchGaps.Count -gt 0) {
    throw "Definition lifecycle dispatch gaps: $($dispatchGaps -join '; ')"
}
Write-Host '[smoke 3/6] Definition lifecycle contracts passed'
$permissiveVerificationSession = New-CaptureSession -Phase PermissiveVerification
Assert-SmokeEqual -Actual $permissiveVerificationSession.Phase -Expected 'PermissiveVerification' -Message 'Post-permissive verification is not an allowed capture phase.'
$null = Complete-CaptureSession -Session $permissiveVerificationSession
$orderedRestoreEntries = @(Get-OrderedRestoreEntries -Entries @(
    [PSCustomObject]@{ Id = 'winrm.service'; Type = 'ServiceConfig'; Name = 'WinRM' }
    [PSCustomObject]@{ Id = 'winrm.service.basic'; Type = 'WsManValue' }
    [PSCustomObject]@{ Id = 'spooler.service'; Type = 'ServiceConfig'; Name = 'Spooler' }
))
Assert-SmokeEqual -Actual $orderedRestoreEntries[0].Id -Expected 'winrm.service.basic' -Message 'WSMan restore is not ordered before WinRM service restoration.'
Assert-SmokeEqual -Actual $orderedRestoreEntries[1].Id -Expected 'spooler.service' -Message 'Restore ordering changed an unrelated service entry.'
Assert-SmokeEqual -Actual $orderedRestoreEntries[2].Id -Expected 'winrm.service' -Message 'WinRM service restoration is not deferred until after WSMan writes.'
Assert-SmokeEqual -Actual (Test-DefinitionHasPermissiveAction -Definition $definitionsById['defender.runtime_status']) -Expected $false -Message 'Defender status is not marked capture-only.'
Assert-SmokeEqual -Actual (Test-DefinitionHasPermissiveAction -Definition $definitionsById['defender.exclusion_paths']) -Expected $false -Message 'Defender exclusions are not marked capture-only.'
Assert-SmokeEqual -Actual (Test-DefinitionHasPermissiveAction -Definition $definitionsById['defender.disable_realtime_monitoring']) -Expected $true -Message 'Mutable Defender preference lost its permissive action.'
Assert-SmokeEqual -Actual (Test-DefinitionHasRestoreAction -Definition $definitionsById['defender.runtime_status']) -Expected $false -Message 'Defender runtime inventory is incorrectly marked as restorable.'
Assert-SmokeEqual -Actual (Test-DefinitionHasRestoreAction -Definition $definitionsById['defender.exclusion_paths']) -Expected $true -Message 'Defender exclusion baselines lost their restore action.'
Assert-SmokeEqual -Actual ((Get-Command Invoke-TimedDefinitionCapture -CommandType Function).Definition -match 'Get-DefinitionCapabilities') -Expected $true -Message 'Captured settings do not persist authoritative engine capabilities.'
Assert-SmokeEqual -Actual ((Get-Command Invoke-TimedDefinitionCapture -CommandType Function).Definition -match 'Get-DefinitionPermissiveTargetDescriptor') -Expected $true -Message 'Captured settings do not persist reviewable permissive targets.'
Assert-SmokeEqual -Actual ((Get-Command Restore-DefenseSnapshot -CommandType Function).Definition -match '\$mutationEntries\s*=\s*@\(\$entries\s*\|\s*Where-Object\s*\{\s*Test-DefinitionHasRestoreAction') -Expected $true -Message 'Restore still schedules inventory-only entries for mutation.'
Assert-SmokeEqual -Actual ((Get-Command Restore-DefenseSnapshot -CommandType Function).Definition -match 'Initialize-SnapshotAssetCache\s+-Entries\s+\$entries\s+-SnapshotPath\s+\$fullPath[\s\S]*\$mutationEntries') -Expected $true -Message 'Restore does not preload selected snapshot assets before mutation planning.'
Assert-SmokeEqual -Actual ((Get-Command Get-AppLockerPolicyXml -CommandType Function).Definition -match "Resolve-ContainedFileSystemPath[\s\S]*AppLocker snapshot asset path") -Expected $true -Message 'AppLocker sidecar references are not constrained to the snapshot asset root.'
Assert-SmokeEqual -Actual ((Get-Command Get-ExploitProtectionPolicyXml -CommandType Function).Definition -match "Resolve-ContainedFileSystemPath[\s\S]*Exploit protection snapshot asset path") -Expected $true -Message 'Exploit-protection sidecar references are not constrained to the snapshot asset root.'
Assert-SmokeEqual -Actual ((Get-Command Test-DefenseSnapshot -CommandType Function).Definition -match "SkipCategory\s*=\s*'InventoryOnly'") -Expected $true -Message 'Restore verification still treats inventory-only drift as a mismatch.'
$permissiveFunction = (Get-Command Set-DefensePermissive -CommandType Function).Definition
$permissiveReviewIndex = $permissiveFunction.IndexOf('Wait-MutationApproval')
$permissiveJournalIndex = $permissiveFunction.IndexOf('Write-OperationState')
Assert-SmokeEqual -Actual ($permissiveReviewIndex -ge 0 -and $permissiveJournalIndex -gt $permissiveReviewIndex) -Expected $true -Message 'Permissive mode journals or mutates before GUI review approval.'
$restoreFunction = (Get-Command Restore-DefenseSnapshot -CommandType Function).Definition
$restoreReviewIndex = $restoreFunction.IndexOf('Wait-MutationApproval')
$restoreStatusIndex = $restoreFunction.IndexOf("-Status 'Restoring'")
Assert-SmokeEqual -Actual ($restoreReviewIndex -ge 0 -and $restoreStatusIndex -gt $restoreReviewIndex) -Expected $true -Message 'Restore changes journal status or live state before GUI review approval.'
Assert-SmokeEqual -Actual ($permissiveFunction -match 'Test-DefensePermissiveState') -Expected $true -Message 'Permissive mode does not verify live state after mutation.'
Assert-SmokeEqual -Actual ($permissiveFunction -match 'Get-PermissiveVerificationReportLines') -Expected $true -Message 'Permissive mode does not persist a verification report.'
Assert-SmokeEqual -Actual ($permissiveFunction -match 'Invoke-PermissiveMutationWorkItem\s+-WorkItem\s+\$workItem\s+-SnapshotPath\s+\$export\.JsonPath') -Expected $true -Message 'Permissive work planning does not pass the persisted snapshot path to sidecar-backed providers.'
Assert-SmokeEqual -Actual ($permissiveFunction -match 'if\s*\(\$null\s+-eq\s+\$entry\)\s*\{\s*throw\s+"The persisted baseline is missing setting') -Expected $true -Message 'Permissive mode can mutate a setting with no persisted baseline entry.'
Assert-SmokeEqual -Actual ((Get-Command Apply-PermissiveDefinition -CommandType Function).Definition -match 'Set-Permissive-AppLockerPolicy\s+-State\s+\$appLockerState\s+-SnapshotPath\s+\$SnapshotPath') -Expected $true -Message 'AppLocker permissive apply cannot resolve its persisted sidecar XML.'
Assert-SmokeEqual -Actual ((Get-Command Set-WsManConfigValues -CommandType Function).Definition -match 'Invoke-WithTemporaryWinRmServiceForWrite\s+-CaptureSession\s+\$CaptureSession') -Expected $true -Message 'WSMan writes do not reuse their mutation-session service scope.'

$uacDefinition = $definitionsById['uac.enable_lua']
$uacBaseline = [PSCustomObject]@{
    Id = $uacDefinition.Id; Type = $uacDefinition.Type; Path = $uacDefinition.Path; Name = $uacDefinition.Name
    ValueKind = 'DWord'; Captured = $true; Exists = $true; CurrentValue = 1; RequiresReboot = $true
}
$uacLive = [PSCustomObject]@{
    Id = $uacDefinition.Id; Type = $uacDefinition.Type; Path = $uacDefinition.Path; Name = $uacDefinition.Name
    ValueKind = 'DWord'; Captured = $true; Exists = $true; CurrentValue = 0; RequiresReboot = $true
}
$uacVerification = Test-PermissiveDefinitionState -Definition $uacDefinition -LiveEntry $uacLive -BaselineEntry $uacBaseline
Assert-SmokeEqual -Actual $uacVerification.Status -Expected 'ConfiguredPendingReboot' -Message 'Reboot-pending permissive state was not classified correctly.'

$mpDefinition = $definitionsById['defender.disable_realtime_monitoring']
$mpLive = [PSCustomObject]@{
    Id = $mpDefinition.Id; Type = $mpDefinition.Type; Property = $mpDefinition.Property
    CommandAvailable = $true; Captured = $true; CurrentValue = $false; RestoreValue = $false; RequiresReboot = $false
}
$mpVerification = Test-PermissiveDefinitionState -Definition $mpDefinition -LiveEntry $mpLive -BaselineEntry $mpLive
Assert-SmokeEqual -Actual $mpVerification.Status -Expected 'Mismatch' -Message 'A rejected Defender permissive target was reported as verified.'

$exploitXml = '<MitigationPolicy><SystemConfig><DEP Enable="false" EmulateAtlThunks="false"/><ControlFlowGuard Enable="false"/><ASLR ForceRelocateImages="false" BottomUp="false" HighEntropy="false"/><SEHOP Enable="false"/></SystemConfig></MitigationPolicy>'
$exploitEvaluation = Get-PermissiveExploitProtectionEvaluation -LiveEntry ([PSCustomObject]@{ CurrentValue = [PSCustomObject]@{ CommandAvailable = $true; Xml = $exploitXml } })
Assert-SmokeEqual -Actual $exploitEvaluation.Matches -Expected $true -Message 'Exploit-protection permissive target projection failed.'

$platformWdacState = [PSCustomObject]@{
    CiToolAvailable = $false
    Policies = @([PSCustomObject]@{ PolicyID = '0283ac0f-fff1-49ae-ada1-8a933130cad6'; FriendlyName = 'VerifiedAndReputableDesktop'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true })
    Files = @()
}
Remove-WdacPolicies -State $platformWdacState
$removableWdacState = [PSCustomObject]@{
    CiToolAvailable = $false
    Policies = @([PSCustomObject]@{ PolicyID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'; FriendlyName = 'Test policy'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true })
    Files = @()
}
$wdacFailedClosed = $false
try {
    Remove-WdacPolicies -State $removableWdacState
} catch {
    $wdacFailedClosed = $_.Exception.Message -like '*CiTool is unavailable*'
}
Assert-SmokeEqual -Actual $wdacFailedClosed -Expected $true -Message 'WDAC permissive removal did not fail closed without CiTool.'
$unclassifiedWdacFailedClosed = $false
try {
    Remove-WdacPolicies -State ([PSCustomObject]@{
        CiToolAvailable = $false; Policies = @()
        Files = @([PSCustomObject]@{ FileName = 'SiPolicy.p7b'; RelativePath = 'SiPolicy.p7b' })
    })
} catch {
    $unclassifiedWdacFailedClosed = $_.Exception.Message -like '*No raw policy files were deleted*'
}
Assert-SmokeEqual -Actual $unclassifiedWdacFailedClosed -Expected $true -Message 'An unclassified WDAC policy file did not fail closed.'
Assert-SmokeEqual -Actual ((Get-Command Remove-WdacPolicies -CommandType Function).Definition -match 'Remove-WdacPolicyFiles') -Expected $false -Message 'WDAC permissive removal can still fall back to indiscriminate file deletion.'
$baselineWdacId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
$extraWdacId = '11111111-2222-3333-4444-555555555555'
$wdacRestoreRemoval = Get-WdacRestoreRemovalState -BaselineState ([PSCustomObject]@{
    Policies = @([PSCustomObject]@{ PolicyID = $baselineWdacId })
    Files = @([PSCustomObject]@{ FileName = "{$baselineWdacId}.cip"; RelativePath = "CiPolicies\Active\{$baselineWdacId}.cip" })
}) -LiveState ([PSCustomObject]@{
    CiToolAvailable = $true
    Policies = @(
        [PSCustomObject]@{ PolicyID = $baselineWdacId; FriendlyName = 'Baseline'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true }
        [PSCustomObject]@{ PolicyID = $extraWdacId; FriendlyName = 'Extra'; HasFileOnDisk = $true; IsCurrentlyEnforced = $true }
        $platformWdacState.Policies[0]
    )
    Files = @(
        [PSCustomObject]@{ FileName = "{$baselineWdacId}.cip"; RelativePath = "CiPolicies\Active\{$baselineWdacId}.cip" }
        [PSCustomObject]@{ FileName = "{$extraWdacId}.cip"; RelativePath = "CiPolicies\Active\{$extraWdacId}.cip" }
    )
})
Assert-SmokeEqual -Actual @($wdacRestoreRemoval.Policies).Count -Expected 1 -Message 'WDAC restore removal did not isolate the extra custom policy.'
Assert-SmokeEqual -Actual (Get-WdacNormalizedPolicyId -Value $wdacRestoreRemoval.Policies[0].PolicyID) -Expected $extraWdacId -Message 'WDAC restore selected the wrong policy for removal.'
Assert-SmokeEqual -Actual ((Get-Command Restore-WdacPolicies -CommandType Function).Definition -match 'Remove-WdacPolicyFiles|Restore-WdacPolicyFiles') -Expected $false -Message 'WDAC restore can still clear policy files wholesale.'
Assert-SmokeEqual -Actual ((Get-Command Apply-PermissiveDefinition -CommandType Function).Definition -match 'Set-Permissive-LoadedUserRegistryValues\s+-Items\s+@\(\$Definition\.Items\)\s+-CaptureSession\s+\$CaptureSession') -Expected $true -Message 'Permissive user-registry mutation does not reuse its phase session.'
Assert-SmokeEqual -Actual ((Get-Command Restore-SnapshotEntry -CommandType Function).Definition -match 'Restore-LoadedUserRegistryValues\s+-Entries\s+@\(\$state\.Entries\)\s+-CaptureSession\s+\$CaptureSession') -Expected $true -Message 'Restore user-registry mutation does not reuse its phase session.'
Assert-SmokeEqual -Actual ((Get-Command Get-BitLockerVolumeStates -CommandType Function).Definition -match 'Invoke-ChildPowerShellBatch\s+-Requests\s+\$probeRequests') -Expected $true -Message 'BitLocker volume probes are not submitted as one concurrent batch.'
Write-Host '[smoke 3/6] Permissive verification contracts passed'

$incompleteEntries = @(
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
foreach ($entry in $incompleteEntries) {
    Assert-SmokeEqual -Actual (Test-SnapshotEntryCapturedExactly -Entry $entry) -Expected $false -Message "Incomplete $($entry.Type) baseline was treated as exact."
}
$wdacReportLines = [System.Collections.Generic.List[string]]::new()
$wdacReportLines.Add('Report')
Add-SnapshotEntryReportLines -Lines $wdacReportLines -Entry ([PSCustomObject]@{
    Id = 'wdac.policies'; Type = 'WdacPolicies'; RequiresReboot = $true
    CurrentValue = [PSCustomObject]@{ CiToolAvailable = $true; Captured = $false; CaptureIssues = @('failed'); Policies = @(); Files = @() }
})
Assert-SmokeEqual -Actual ($wdacReportLines -contains '  Capture issue count: 1') -Expected $true -Message 'Single-issue WDAC report rendering failed under strict mode.'

$filterValidationRejected = $false
try {
    Assert-ValidSettingIdFilter -AvailableId @($definitions.Id) -IncludeId 'missing.setting'
} catch {
    $filterValidationRejected = $true
}
Assert-SmokeEqual -Actual $filterValidationRejected -Expected $true -Message 'Unknown setting filter was not rejected.'
Write-Host '[smoke 3/6] Mutation validation contracts passed'

$smokeComputerName = if (-not [string]::IsNullOrWhiteSpace([string]$env:COMPUTERNAME)) { [string]$env:COMPUTERNAME } else { [Environment]::MachineName }
$validSnapshot = [PSCustomObject]@{
    SchemaVersion = 2
    Tool          = 'WinDefState'
    ComputerName  = $smokeComputerName
    Settings      = @([PSCustomObject]@{ Id = 'defender.runtime_status'; Type = 'DefenderRuntimeStatus'; CurrentValue = $null; RequiresReboot = $false })
}
Assert-ValidDefenseSnapshot -Snapshot $validSnapshot
$validSnapshot.Settings[0].Type = 'RegistryValue'
$snapshotValidationRejected = $false
try {
    Assert-ValidDefenseSnapshot -Snapshot $validSnapshot
} catch {
    $snapshotValidationRejected = $true
}
Assert-SmokeEqual -Actual $snapshotValidationRejected -Expected $true -Message 'Snapshot type mismatch was not rejected.'

function Test-CommandAvailable {
    param([Parameter(Mandatory)] [string]$Name)

    $Name -in @('Get-MpPreference', 'Get-WSManInstance', 'Get-NetFirewallProfile')
}

function Get-MpPreference {
    [CmdletBinding()]
    param()

    $script:mpPreferenceCalls++
    [PSCustomObject]@{
        DisableRealtimeMonitoring           = $false
        ExclusionPath                        = @('C:\Tools')
        AttackSurfaceReductionRules_Ids     = @('56a863a9-875e-4185-98a7-b882c64b5ce5')
        AttackSurfaceReductionRules_Actions = @(1)
    }
}

function Get-WSManInstance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$ResourceURI,
        [switch]$Enumerate
    )

    $script:wsManCalls++
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

function Get-NetFirewallProfile {
    [CmdletBinding()]
    param([string[]]$Profile)

    $script:firewallProfileCalls++
    @(
        [PSCustomObject]@{ Name = 'Domain'; Enabled = $true; DefaultInboundAction = 'Block'; DefaultOutboundAction = 'Allow' }
        [PSCustomObject]@{ Name = 'Private'; Enabled = $true; DefaultInboundAction = 'Block'; DefaultOutboundAction = 'Allow' }
        [PSCustomObject]@{ Name = 'Public'; Enabled = $false; DefaultInboundAction = 'Block'; DefaultOutboundAction = 'Allow' }
    )
}

Write-Host '[smoke 4/6] Exercising shared provider caches'
$script:mpPreferenceCalls = 0
$defenderSession = New-CaptureSession -Phase Snapshot
$scalar = Get-MpPreferencePropertyRawValue -Property 'DisableRealtimeMonitoring' -CaptureSession $defenderSession
$list = Get-MpPreferenceListState -Property 'ExclusionPath' -CaptureSession $defenderSession
$asr = Get-AsrRuleCaptureState -CaptureSession $defenderSession
$defenderMetrics = Complete-CaptureSession -Session $defenderSession

Assert-SmokeEqual -Actual $scalar -Expected $false -Message 'Defender scalar capture failed.'
Assert-SmokeEqual -Actual $list.Items[0] -Expected 'C:\Tools' -Message 'Defender list capture failed.'
Assert-SmokeEqual -Actual $asr.Rules.Count -Expected 1 -Message 'ASR capture failed.'
Assert-SmokeEqual -Actual $script:mpPreferenceCalls -Expected 1 -Message 'Defender preferences were not batched.'
Assert-SmokeEqual -Actual $defenderMetrics.CacheHitCount -Expected 2 -Message 'Defender cache hit telemetry is incorrect.'

$script:wsManCalls = 0
$wsManSession = New-CaptureSession -Phase Verification
$basic = Get-WsManConfigValueState -Path 'WSMan:\localhost\Service\Auth\Basic' -CaptureSession $wsManSession
$kerberos = Get-WsManConfigValueState -Path 'WSMan:\localhost\Service\Auth\Kerberos' -CaptureSession $wsManSession
$allowUnencrypted = Get-WsManConfigValueState -Path 'WSMan:\localhost\Service\AllowUnencrypted' -CaptureSession $wsManSession
$wsManMetrics = Complete-CaptureSession -Session $wsManSession

Assert-SmokeEqual -Actual $basic.Value -Expected $true -Message 'WSMan Basic capture failed.'
Assert-SmokeEqual -Actual $kerberos.Value -Expected $false -Message 'WSMan Kerberos capture failed.'
Assert-SmokeEqual -Actual $allowUnencrypted.Value -Expected $false -Message 'WSMan AllowUnencrypted capture failed.'
Assert-SmokeEqual -Actual $script:wsManCalls -Expected 2 -Message 'WSMan resource reads were not batched.'
Assert-SmokeEqual -Actual $wsManMetrics.CacheHitCount -Expected 1 -Message 'WSMan cache hit telemetry is incorrect.'

function Invoke-ChildPowerShell {
    [CmdletBinding()]
    param(
        [string]$ScriptText,
        [int]$TimeoutSeconds
    )

    $script:smbChildCalls++
    [PSCustomObject]@{
        CommandAvailable = $true
        TimedOut         = $false
        ExitCode         = 0
        StdErr           = ''
        StdOut           = "module chatter`r`nWDS_SMB_JSON:{`"Client`":{`"CommandAvailable`":true,`"Captured`":true,`"Error`":null,`"RequireSecuritySignature`":true},`"Server`":{`"CommandAvailable`":true,`"Captured`":true,`"Error`":null,`"RequireSecuritySignature`":false}}"
    }
}

$script:smbChildCalls = 0
$smbSession = New-CaptureSession -Phase Snapshot
$smbClientDefinition = [PSCustomObject]@{ Id = 'smb.client.require_security_signature'; Type = 'SmbClientConfig'; RequiresReboot = $false }
$smbServerDefinition = [PSCustomObject]@{ Id = 'smb.server.require_security_signature'; Type = 'SmbServerConfig'; RequiresReboot = $false }
$smbClient = Capture-Definition -Definition $smbClientDefinition -CaptureSession $smbSession
$smbServer = Capture-Definition -Definition $smbServerDefinition -CaptureSession $smbSession
Assert-SmokeEqual -Actual $smbClient.CurrentValue.RequireSecuritySignature -Expected $true -Message 'SMB client signing capture failed.'
Assert-SmokeEqual -Actual $smbServer.CurrentValue.RequireSecuritySignature -Expected $false -Message 'SMB server signing capture failed.'
Assert-SmokeEqual -Actual $script:smbChildCalls -Expected 1 -Message 'SMB client/server capture did not share one child process.'
Assert-SmokeEqual -Actual $smbSession.CacheHitCount -Expected 1 -Message 'SMB cache hit telemetry is incorrect.'

$script:firewallProfileCalls = 0
$firewallState = Get-FirewallProfileStates -Profiles @('Domain', 'Private', 'Public')
Assert-SmokeEqual -Actual @($firewallState.Profiles).Count -Expected 3 -Message 'Firewall profile capture lost a requested profile.'
Assert-SmokeEqual -Actual @($firewallState.CaptureIssues).Count -Expected 0 -Message 'Firewall profile capture reported an unexpected issue.'
Assert-SmokeEqual -Actual $script:firewallProfileCalls -Expected 1 -Message 'Firewall profile reads were not batched.'

function Get-CimInstance {
    [CmdletBinding()]
    param(
        [string]$ClassName,
        [string]$Filter
    )

    $script:serviceCimCalls++
    @(
        [PSCustomObject]@{ Name = 'AppIDSvc'; StartMode = 'Manual'; State = 'Stopped' }
        [PSCustomObject]@{ Name = 'Spooler'; StartMode = 'Auto'; State = 'Running' }
        [PSCustomObject]@{ Name = 'WinRM'; StartMode = 'Auto'; State = 'Running' }
    )
}

$script:serviceCimCalls = 0
$serviceSession = New-CaptureSession -Phase Verification
$serviceSession.ServiceNames = @('AppIDSvc', 'Spooler', 'WinRM')
$appIdService = Get-ServiceCaptureState -Name 'AppIDSvc' -CaptureSession $serviceSession
$spoolerService = Get-ServiceCaptureState -Name 'Spooler' -CaptureSession $serviceSession
$winRmService = Get-ServiceCaptureState -Name 'WinRM' -CaptureSession $serviceSession

Assert-SmokeEqual -Actual $appIdService.State -Expected 'Stopped' -Message 'AppLocker service capture failed.'
Assert-SmokeEqual -Actual $spoolerService.StartMode -Expected 'Auto' -Message 'Spooler service capture failed.'
Assert-SmokeEqual -Actual $winRmService.State -Expected 'Running' -Message 'WinRM service capture failed.'
Assert-SmokeEqual -Actual $script:serviceCimCalls -Expected 1 -Message 'Service CIM reads were not batched.'
Assert-SmokeEqual -Actual $serviceSession.CacheHitCount -Expected 2 -Message 'Service cache hit telemetry is incorrect.'

function Get-Item {
    [CmdletBinding()]
    param([string]$Path)

    $script:registryItemCalls++
    $key = [PSCustomObject]@{}
    $key | Add-Member -MemberType ScriptMethod -Name GetValueKind -Value { param($name) 'DWord' }
    $key
}

function Get-ItemProperty {
    [CmdletBinding()]
    param([string]$Path)

    $script:registryPropertyCalls++
    [PSCustomObject]@{ FirstValue = 1; SecondValue = 0 }
}

$script:registryItemCalls = 0
$script:registryPropertyCalls = 0
$registrySession = New-CaptureSession -Phase Snapshot
$firstRegistryValue = Get-RegistryValueCaptureState -Path 'HKLM:\SOFTWARE\WinDefStateTest' -Name 'FirstValue' -DefaultValueKind 'DWord' -CaptureSession $registrySession
$secondRegistryValue = Get-RegistryValueCaptureState -Path 'HKLM:\SOFTWARE\WinDefStateTest' -Name 'SecondValue' -DefaultValueKind 'DWord' -CaptureSession $registrySession

Assert-SmokeEqual -Actual $firstRegistryValue.CurrentValue -Expected 1 -Message 'First registry value capture failed.'
Assert-SmokeEqual -Actual $secondRegistryValue.CurrentValue -Expected 0 -Message 'Second registry value capture failed.'
Assert-SmokeEqual -Actual $script:registryItemCalls -Expected 1 -Message 'Registry key metadata reads were not batched.'
Assert-SmokeEqual -Actual $script:registryPropertyCalls -Expected 1 -Message 'Registry property reads were not batched.'
Assert-SmokeEqual -Actual $registrySession.CacheHitCount -Expected 1 -Message 'Registry cache hit telemetry is incorrect.'

$journalRoot = New-TemporaryFilePath -Extension '.state'
$journalSnapshot = Join-Path $journalRoot 'snapshot.json'
Write-Host '[smoke 5/6] Exercising atomic journal and integrity handling'
try {
    Ensure-Directory -Path $journalRoot
    $firstArtifactPath = Get-AvailableArtifactPath -Directory $journalRoot -BaseName 'artifact' -Extension 'json'
    Write-TextAtomic -Path $firstArtifactPath -Content '{}'
    $secondArtifactPath = Get-AvailableArtifactPath -Directory $journalRoot -BaseName 'artifact' -Extension '.json'
    Assert-SmokeEqual -Actual ([IO.Path]::GetFileName($secondArtifactPath)) -Expected 'artifact-1.json' -Message 'Generated artifact paths can overwrite an existing file.'

    Write-TextAtomic -Path $journalSnapshot -Content '{"snapshot":true}'
    $journalAssetRoot = Get-SnapshotAssetRoot -SnapshotPath $journalSnapshot
    Ensure-Directory -Path $journalAssetRoot
    Write-TextAtomic -Path (Join-Path $journalAssetRoot 'policy.xml') -Content '<Policy />'
    $operation = Write-OperationState -Root $journalRoot -SnapshotPath $journalSnapshot -Mode Permissive -IncludeId @('rdp.user_authentication')
    $loadedOperation = Get-OperationState -Root $journalRoot

    Assert-SmokeEqual -Actual $loadedOperation.Status -Expected 'Applying' -Message 'Operation journal status was not persisted.'
    Assert-SmokeEqual -Actual $loadedOperation.Producer.ScriptSha256 -Expected (Get-FileHash -LiteralPath $enginePath -Algorithm SHA256).Hash -Message 'Operation journal producer hash is incorrect.'
    Assert-SmokeEqual -Actual @($loadedOperation.SnapshotIntegrity.Assets).Count -Expected 1 -Message 'Operation journal asset integrity metadata is incomplete.'
    Assert-OperationSnapshotIntegrity -Operation $loadedOperation
    $operation = Update-OperationStateStatus -Root $journalRoot -Operation $operation -Status Applied
    Assert-SmokeEqual -Actual (Get-OperationState -Root $journalRoot).Status -Expected 'Applied' -Message 'Operation journal status update failed.'
    $operation = Update-OperationStateStatus -Root $journalRoot -Operation $operation -Status ApplyFailed
    Assert-SmokeEqual -Actual (Get-OperationState -Root $journalRoot).Status -Expected 'ApplyFailed' -Message 'Apply failure status was not persisted.'
    $operation = Update-OperationStateStatus -Root $journalRoot -Operation $operation -Status RestoreFailed
    Assert-SmokeEqual -Actual (Get-OperationState -Root $journalRoot).Status -Expected 'RestoreFailed' -Message 'Restore failure status was not persisted.'
    $permissiveVerification = [PSCustomObject]@{
        VerifiedAtUtc = '2026-01-01T00:00:00Z'; VerifiedCount = 1; PendingRebootCount = 1; MismatchCount = 0
        MutationMetrics = [PSCustomObject]@{ DurationMs = 1250 }
        Results = @([PSCustomObject]@{ Id = 'uac.enable_lua'; Status = 'ConfiguredPendingReboot' })
    }
    $permissiveReportPath = Join-Path $journalRoot 'permissive-check.txt'
    $operation = Update-OperationStateStatus -Root $journalRoot -Operation $operation -Status AppliedPendingReboot -PermissiveVerification $permissiveVerification -PermissiveVerificationReportPath $permissiveReportPath
    $loadedPermissiveVerification = (Get-OperationState -Root $journalRoot).PermissiveVerification
    Assert-SmokeEqual -Actual $loadedPermissiveVerification.PendingRebootCount -Expected 1 -Message 'Permissive verification summary was not journaled.'
    Assert-SmokeEqual -Actual $loadedPermissiveVerification.MutationDurationMs -Expected 1250 -Message 'Permissive mutation duration was not journaled.'
    Assert-SmokeEqual -Actual (@($loadedPermissiveVerification.PendingRebootIds) -contains 'uac.enable_lua') -Expected $true -Message 'Pending-reboot IDs were not journaled.'

    $invalidJournalRoot = Join-Path $journalRoot 'invalid-journal'
    Ensure-Directory -Path $invalidJournalRoot
    Write-TextAtomic -Path (Get-OperationPath -Root $invalidJournalRoot) -Content 'null'
    $invalidJournalRejected = $false
    try {
        Assert-NoActiveOperation -Root $invalidJournalRoot
    } catch {
        $invalidJournalRejected = $_.Exception.Message -match 'did not contain a JSON object'
    }
    Assert-SmokeEqual -Actual $invalidJournalRejected -Expected $true -Message 'A null operation journal can be mistaken for no active operation.'

    $activeOperationRejected = $false
    try {
        Assert-NoActiveOperation -Root $journalRoot
    } catch {
        $activeOperationRejected = $true
    }
    Assert-SmokeEqual -Actual $activeOperationRejected -Expected $true -Message 'A second permissive operation could replace the active journal.'

    Write-TextAtomic -Path $journalSnapshot -Content '{"snapshot":false}'
    $integrityRejected = $false
    try {
        Assert-OperationSnapshotIntegrity -Operation $operation
    } catch {
        $integrityRejected = $true
    }
    Assert-SmokeEqual -Actual $integrityRejected -Expected $true -Message 'Changed operation snapshot was not rejected.'
} finally {
    if (Test-Path -LiteralPath $journalRoot) {
        Remove-Item -LiteralPath $journalRoot -Recurse -Force
    }
}

$script:closedCaptureResources = 0
function Close-UserRegistryTarget {
    param([AllowNull()] [object]$Target)

    $script:closedCaptureResources++
}

$resourceSession = New-CaptureSession -Phase Snapshot
Register-CaptureSessionResource -Session $resourceSession -Kind 'UserRegistryTarget' -Value ([PSCustomObject]@{ Sid = 'S-1-5-21-smoke' })
$null = Complete-CaptureSession -Session $resourceSession
Assert-SmokeEqual -Actual $script:closedCaptureResources -Expected 1 -Message 'Capture resources were not released.'
Assert-SmokeEqual -Actual $resourceSession.Resources.Count -Expected 0 -Message 'Released capture resources were retained.'

$snapshotPath = New-TemporaryFilePath -Extension '.json'
Write-Host '[smoke 6/6] Exercising snapshot serialization'
try {
    $snapshot = [PSCustomObject]@{
        SchemaVersion  = 2
        Tool           = 'WinDefState'
        Producer       = Get-WinDefStateRuntimeInfo
        ComputerName   = 'SMOKE'
        CapturedAtUtc  = '2026-08-09T00:00:00.0000000Z'
        CaptureMetrics = $defenderMetrics
        CaptureScope   = [PSCustomObject]@{
            IsFiltered = $true
            IncludeId  = @('defender.disable_realtime_monitoring')
            ExcludeId  = @()
        }
        Settings       = @()
    }

    Write-SnapshotJsonAtomic -Path $snapshotPath -Snapshot $snapshot
    $saved = Get-Content -LiteralPath $snapshotPath -Raw | ConvertFrom-Json
    Assert-SmokeEqual -Actual $saved.SchemaVersion -Expected 2 -Message 'Snapshot schema version did not round-trip.'
    Assert-SmokeEqual -Actual $saved.Producer.ScriptSha256 -Expected (Get-FileHash -LiteralPath $enginePath -Algorithm SHA256).Hash -Message 'Snapshot producer hash did not round-trip.'
    Assert-SmokeEqual -Actual $saved.CaptureMetrics.CacheHitCount -Expected 2 -Message 'Snapshot metrics did not round-trip.'
    Assert-SmokeEqual -Actual $saved.CaptureScope.IsFiltered -Expected $true -Message 'Snapshot capture scope did not round-trip.'
} finally {
    if (Test-Path -LiteralPath $snapshotPath) {
        Remove-Item -LiteralPath $snapshotPath -Force
    }
}

Write-Host 'WinDefState smoke tests passed.'

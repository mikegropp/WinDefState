[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Snapshot', 'Permissive', 'Restore')]
    [string]$Command,

    [string]$SnapshotPath,

    [string]$StateRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptRoot)) {
    if (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
        $scriptRoot = Split-Path -Parent $PSCommandPath
    } elseif ($null -ne $MyInvocation.MyCommand -and -not [string]::IsNullOrWhiteSpace([string]$MyInvocation.MyCommand.Path)) {
        $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        $scriptRoot = (Get-Location).Path
    }
}

if ([string]::IsNullOrWhiteSpace($StateRoot)) {
    $StateRoot = Join-Path $scriptRoot 'state'
}

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
        Remove-Item -LiteralPath $resolvedPath -Force -ErrorAction SilentlyContinue
    }

    Move-Item -LiteralPath $tempPath -Destination $resolvedPath -Force
}

function Write-SnapshotJsonAtomic {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [object]$Snapshot
    )

    $resolvedPath = Resolve-FileSystemPath -Path $Path
    $parent = Split-Path -Parent $resolvedPath
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Ensure-Directory -Path $parent
    }

    $tempPath = Join-Path $parent ((Split-Path -Leaf $resolvedPath) + '.tmp')
    $encoding = [System.Text.UTF8Encoding]::new($false)
    $stream = [System.IO.File]::Open($tempPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    $writer = New-Object System.IO.StreamWriter($stream, $encoding)

    try {
        $writer.WriteLine('{')
        $writer.WriteLine(('  "SchemaVersion": {0},' -f (ConvertTo-Json -InputObject $Snapshot.SchemaVersion -Compress)))
        $writer.WriteLine(('  "Tool": {0},' -f (ConvertTo-Json -InputObject $Snapshot.Tool -Compress)))
        $writer.WriteLine(('  "ComputerName": {0},' -f (ConvertTo-Json -InputObject $Snapshot.ComputerName -Compress)))
        $writer.WriteLine(('  "CapturedAtUtc": {0},' -f (ConvertTo-Json -InputObject $Snapshot.CapturedAtUtc -Compress)))
        $writer.WriteLine('  "Settings": [')

        $settings = @($Snapshot.Settings)
        for ($i = 0; $i -lt $settings.Count; $i++) {
            $entry = $settings[$i]
            $entryId = if ($null -ne $entry -and $entry.PSObject.Properties['Id']) { [string]$entry.Id } else { "<entry-$i>" }
            Write-Verbose ("[json {0}/{1}] Serializing {2}" -f ($i + 1), $settings.Count, $entryId)
            $entryJson = ConvertTo-Json -InputObject $entry -Depth 12 -Compress
            $suffix = if ($i -lt ($settings.Count - 1)) { ',' } else { '' }
            $writer.WriteLine(('    {0}{1}' -f $entryJson, $suffix))
        }

        $writer.WriteLine('  ]')
        $writer.WriteLine('}')
    } finally {
        if ($null -ne $writer) {
            $writer.Dispose()
        } elseif ($null -ne $stream) {
            $stream.Dispose()
        }
    }

    if (Test-Path -LiteralPath $resolvedPath) {
        Remove-Item -LiteralPath $resolvedPath -Force -ErrorAction SilentlyContinue
    }

    Move-Item -LiteralPath $tempPath -Destination $resolvedPath -Force
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
        Remove-Item -LiteralPath $resolvedPath -Force -ErrorAction SilentlyContinue
    }

    Move-Item -LiteralPath $tempPath -Destination $resolvedPath -Force
}

function Write-BytesAtomic {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [byte[]]$Content
    )

    $resolvedPath = Resolve-FileSystemPath -Path $Path
    $parent = Split-Path -Parent $resolvedPath
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        Ensure-Directory -Path $parent
    }

    $tempPath = Join-Path $parent ((Split-Path -Leaf $resolvedPath) + '.tmp')
    [System.IO.File]::WriteAllBytes($tempPath, $Content)

    if (Test-Path -LiteralPath $resolvedPath) {
        Remove-Item -LiteralPath $resolvedPath -Force -ErrorAction SilentlyContinue
    }

    Move-Item -LiteralPath $tempPath -Destination $resolvedPath -Force
}

function New-TemporaryFilePath {
    param([string]$Extension = '.tmp')

    if (-not $Extension.StartsWith('.')) {
        $Extension = ".$Extension"
    }

    Join-Path ([System.IO.Path]::GetTempPath()) ("WinDefState-{0}{1}" -f ([guid]::NewGuid().ToString('N')), $Extension)
}

function Test-CommandAvailable {
    param([Parameter(Mandatory)] [string]$Name)

    $null -ne (Get-Command -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1)
}

function Get-CiToolCommand {
    foreach ($name in @('CiTool.exe', 'CiTool')) {
        $command = Get-Command -Name $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command) {
            return $command
        }
    }

    $null
}

function Invoke-ChildPowerShell {
    param(
        [Parameter(Mandatory)] [string]$ScriptText,
        [int]$TimeoutSeconds = 15
    )

    $powershellCommand = Get-Command -Name 'powershell.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $powershellCommand) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            TimedOut         = $false
            ExitCode         = $null
            StdOut           = ''
            StdErr           = 'powershell.exe was not found.'
        }
    }

    $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($ScriptText))
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $powershellCommand.Source
    $startInfo.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encodedCommand"
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    [void]$process.Start()

    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        try {
            $process.Kill()
        } catch {
        }

        return [PSCustomObject]@{
            CommandAvailable = $true
            TimedOut         = $true
            ExitCode         = $null
            StdOut           = ''
            StdErr           = ''
        }
    }

    [PSCustomObject]@{
        CommandAvailable = $true
        TimedOut         = $false
        ExitCode         = $process.ExitCode
        StdOut           = $process.StandardOutput.ReadToEnd()
        StdErr           = $process.StandardError.ReadToEnd()
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

function Get-SnapshotAssetRoot {
    param([Parameter(Mandatory)] [string]$SnapshotPath)

    $fullPath = [IO.Path]::GetFullPath($SnapshotPath)
    $snapshotDir = Split-Path -Parent $fullPath
    $baseName = [IO.Path]::GetFileNameWithoutExtension($fullPath)
    Join-Path $snapshotDir ($baseName + '.assets')
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

function Get-WdacVerificationReportPath {
    param([Parameter(Mandatory)] [string]$VerificationPath)

    $fullPath = [IO.Path]::GetFullPath($VerificationPath)
    $parent = Split-Path -Parent $fullPath
    $baseName = [IO.Path]::GetFileNameWithoutExtension($fullPath)
    Join-Path $parent ($baseName + '-wdac.txt')
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

function Get-UserProfileRegistryTargets {
    $targetsBySid = @{}
    $loadedSidSet = @{}

    foreach ($sid in @(Get-LoadedUserSids)) {
        $loadedSidSet[[string]$sid] = $true
    }

    if (Test-CommandAvailable -Name 'Get-CimInstance') {
        try {
            foreach ($profile in @(Get-CimInstance -ClassName Win32_UserProfile -ErrorAction Stop)) {
                $sid = if ($null -ne $profile.PSObject.Properties['SID']) { [string]$profile.SID } else { '' }
                if ([string]::IsNullOrWhiteSpace($sid) -or $sid -notmatch '^S-\d-\d+-.+') {
                    continue
                }

                $isSpecial = $false
                if ($null -ne $profile.PSObject.Properties['Special'] -and $null -ne $profile.Special) {
                    $isSpecial = [bool]$profile.Special
                }

                if ($isSpecial) {
                    continue
                }

                $profilePath = $null
                if ($null -ne $profile.PSObject.Properties['LocalPath'] -and -not [string]::IsNullOrWhiteSpace([string]$profile.LocalPath)) {
                    $profilePath = Resolve-FileSystemPath -Path ([string]$profile.LocalPath)
                }

                $targetsBySid[$sid] = [PSCustomObject]@{
                    Sid         = $sid
                    ProfilePath = $profilePath
                    HivePath    = if (-not [string]::IsNullOrWhiteSpace($profilePath)) { Join-Path $profilePath 'NTUSER.DAT' } else { $null }
                    Loaded      = ($loadedSidSet.ContainsKey($sid) -or ($null -ne $profile.PSObject.Properties['Loaded'] -and [bool]$profile.Loaded))
                }
            }
        } catch {
            Write-Verbose ("Failed to enumerate Win32_UserProfile instances: {0}" -f $_.Exception.Message)
        }
    }

    foreach ($profileKey in @(Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' -ErrorAction SilentlyContinue)) {
        $sid = [string]$profileKey.PSChildName
        if ([string]::IsNullOrWhiteSpace($sid) -or $sid -notmatch '^S-\d-\d+-.+') {
            continue
        }

        if ($targetsBySid.ContainsKey($sid)) {
            continue
        }

        $profilePath = $null
        $properties = Get-ItemProperty -Path $profileKey.PSPath -Name 'ProfileImagePath' -ErrorAction SilentlyContinue
        if ($null -ne $properties -and -not [string]::IsNullOrWhiteSpace([string]$properties.ProfileImagePath)) {
            $profilePath = Resolve-FileSystemPath -Path ([string]$properties.ProfileImagePath)
        }

        $targetsBySid[$sid] = [PSCustomObject]@{
            Sid         = $sid
            ProfilePath = $profilePath
            HivePath    = if (-not [string]::IsNullOrWhiteSpace($profilePath)) { Join-Path $profilePath 'NTUSER.DAT' } else { $null }
            Loaded      = $loadedSidSet.ContainsKey($sid)
        }
    }

    foreach ($sid in @($loadedSidSet.Keys)) {
        if ($targetsBySid.ContainsKey($sid)) {
            continue
        }

        $targetsBySid[$sid] = [PSCustomObject]@{
            Sid         = $sid
            ProfilePath = $null
            HivePath    = $null
            Loaded      = $true
        }
    }

    @(
        foreach ($entry in @($targetsBySid.GetEnumerator() | Sort-Object -Property Name)) {
            $entry.Value
        }
    )
}

function Open-UserRegistryTarget {
    param([Parameter(Mandatory)] [object]$Target)

    $sid = [string]$Target.Sid
    if ([string]::IsNullOrWhiteSpace($sid)) {
        throw 'User registry target SID is missing.'
    }

    $loadedRootPath = "Registry::HKEY_USERS\$sid"
    if (Test-Path -LiteralPath $loadedRootPath) {
        return [PSCustomObject]@{
            Sid           = $sid
            ProfilePath   = if ($null -ne $Target.PSObject.Properties['ProfilePath']) { [string]$Target.ProfilePath } else { $null }
            HivePath      = if ($null -ne $Target.PSObject.Properties['HivePath']) { [string]$Target.HivePath } else { $null }
            RootPath      = $loadedRootPath
            MountName     = $sid
            MountedByTool = $false
        }
    }

    $hivePath = if ($null -ne $Target.PSObject.Properties['HivePath']) { [string]$Target.HivePath } else { $null }
    if ([string]::IsNullOrWhiteSpace($hivePath)) {
        throw ("User profile hive path is unavailable for SID {0}." -f $sid)
    }

    if (-not (Test-Path -LiteralPath $hivePath)) {
        throw ("User profile hive file was not found for SID {0} at {1}." -f $sid, $hivePath)
    }

    $mountName = "WinDefState_{0}" -f ([guid]::NewGuid().ToString('N'))
    $loadOutput = & reg.exe load "HKU\$mountName" $hivePath 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = (@($loadOutput | ForEach-Object { $_.ToString().Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = "reg.exe load failed with exit code $LASTEXITCODE."
        }

        throw ("Failed to load user hive for SID {0} from {1}. {2}" -f $sid, $hivePath, $message)
    }

    [PSCustomObject]@{
        Sid           = $sid
        ProfilePath   = if ($null -ne $Target.PSObject.Properties['ProfilePath']) { [string]$Target.ProfilePath } else { $null }
        HivePath      = $hivePath
        RootPath      = "Registry::HKEY_USERS\$mountName"
        MountName     = $mountName
        MountedByTool = $true
    }
}

function Close-UserRegistryTarget {
    param([AllowNull()] [object]$Target)

    if ($null -eq $Target) {
        return
    }

    $mountedByTool = $false
    if ($null -ne $Target.PSObject.Properties['MountedByTool'] -and $Target.MountedByTool) {
        $mountedByTool = $true
    }

    if (-not $mountedByTool) {
        return
    }

    $mountName = if ($null -ne $Target.PSObject.Properties['MountName']) { [string]$Target.MountName } else { $null }
    if ([string]::IsNullOrWhiteSpace($mountName)) {
        return
    }

    $unloadOutput = & reg.exe unload "HKU\$mountName" 2>&1
    if ($LASTEXITCODE -ne 0) {
        $message = (@($unloadOutput | ForEach-Object { $_.ToString().Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($message)) {
            $message = "reg.exe unload failed with exit code $LASTEXITCODE."
        }

        Write-Warning ("Failed to unload temporary user hive mount HKU\{0} for SID {1}. {2}" -f $mountName, ([string]$Target.Sid), $message)
    }
}

function New-UserRegistryCaptureIssue {
    param(
        [Parameter(Mandatory)] [string]$Sid,
        [string]$ProfilePath,
        [string]$HivePath,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        Sid         = $Sid
        ProfilePath = $ProfilePath
        HivePath    = $HivePath
        Message     = $Message
    }
}

function Normalize-UserRegistryValueState {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return [PSCustomObject]@{
            Entries       = @()
            CaptureIssues = @()
        }
    }

    $entries = @()
    $captureIssues = @()

    if ($null -ne $State.PSObject.Properties['Entries']) {
        $entries = @($State.Entries)
    } else {
        $entries = @($State)
    }

    if ($null -ne $State.PSObject.Properties['CaptureIssues']) {
        $captureIssues = @($State.CaptureIssues)
    }

    [PSCustomObject]@{
        Entries       = @($entries)
        CaptureIssues = @($captureIssues)
    }
}

function Get-LoadedUserRegistryValueStates {
    param([Parameter(Mandatory)] [object[]]$Items)

    $entries = @()
    $captureIssues = @()

    foreach ($target in @(Get-UserProfileRegistryTargets)) {
        $access = $null

        try {
            $access = Open-UserRegistryTarget -Target $target

            foreach ($item in @($Items)) {
                $path = "$($access.RootPath)\$($item.RelativePath)"
                $property = Get-ItemProperty -Path $path -Name $item.Name -ErrorAction SilentlyContinue
                $exists = $null -ne $property -and $null -ne $property.$($item.Name)
                $kind = if ($exists) {
                    try { (Get-Item -Path $path).GetValueKind($item.Name).ToString() } catch { $item.ValueKind }
                } else {
                    $item.ValueKind
                }

                $entries += [PSCustomObject]@{
                    Sid          = [string]$target.Sid
                    ProfilePath  = if ($null -ne $target.PSObject.Properties['ProfilePath']) { [string]$target.ProfilePath } else { $null }
                    HivePath     = if ($null -ne $access.PSObject.Properties['HivePath']) { [string]$access.HivePath } else { $null }
                    RelativePath = $item.RelativePath
                    Name         = $item.Name
                    Exists       = $exists
                    CurrentValue = if ($exists) { $property.$($item.Name) } else { $null }
                    ValueKind    = $kind
                }
            }
        } catch {
            $message = $_.Exception.Message
            $profilePath = if ($null -ne $target.PSObject.Properties['ProfilePath']) { [string]$target.ProfilePath } else { $null }
            $hivePath = if ($null -ne $target.PSObject.Properties['HivePath']) { [string]$target.HivePath } else { $null }
            $captureIssues += New-UserRegistryCaptureIssue -Sid ([string]$target.Sid) -ProfilePath $profilePath -HivePath $hivePath -Message $message
            Write-Warning ("Failed to capture user-scoped registry values for SID {0}: {1}" -f ([string]$target.Sid), $message)
        } finally {
            Close-UserRegistryTarget -Target $access
        }
    }

    [PSCustomObject]@{
        Entries       = @($entries)
        CaptureIssues = @($captureIssues)
    }
}

function Set-Permissive-LoadedUserRegistryValues {
    param([Parameter(Mandatory)] [object[]]$Items)

    foreach ($target in @(Get-UserProfileRegistryTargets)) {
        $access = $null

        try {
            $access = Open-UserRegistryTarget -Target $target

            foreach ($item in @($Items)) {
                $path = "$($access.RootPath)\$($item.RelativePath)"

                if ($item.PermissiveExists) {
                    Ensure-RegistryPath -Path $path
                    New-ItemProperty -Path $path -Name $item.Name -PropertyType $item.ValueKind -Value $item.PermissiveValue -Force | Out-Null
                } else {
                    Remove-ItemProperty -Path $path -Name $item.Name -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-Warning ("Failed to set permissive user-scoped registry values for SID {0}: {1}" -f ([string]$target.Sid), $_.Exception.Message)
        } finally {
            Close-UserRegistryTarget -Target $access
        }
    }
}

function Restore-LoadedUserRegistryValues {
    param([Parameter(Mandatory)] [object[]]$Entries)

    $targetsBySid = @{}
    foreach ($target in @(Get-UserProfileRegistryTargets)) {
        $targetsBySid[[string]$target.Sid] = $target
    }

    foreach ($group in @($Entries | Group-Object -Property Sid)) {
        $groupEntries = @($group.Group)
        if ($groupEntries.Count -eq 0) {
            continue
        }

        $sid = [string]$group.Name
        $sampleEntry = $groupEntries[0]
        $target = if ($targetsBySid.ContainsKey($sid)) {
            $targetsBySid[$sid]
        } else {
            [PSCustomObject]@{
                Sid         = $sid
                ProfilePath = if ($null -ne $sampleEntry.PSObject.Properties['ProfilePath']) { [string]$sampleEntry.ProfilePath } else { $null }
                HivePath    = if ($null -ne $sampleEntry.PSObject.Properties['HivePath']) { [string]$sampleEntry.HivePath } else { $null }
                Loaded      = $false
            }
        }

        $access = $null
        try {
            $access = Open-UserRegistryTarget -Target $target

            foreach ($entry in $groupEntries) {
                $path = "$($access.RootPath)\$($entry.RelativePath)"

                if ($entry.Exists) {
                    Ensure-RegistryPath -Path $path
                    New-ItemProperty -Path $path -Name $entry.Name -PropertyType $entry.ValueKind -Value $entry.CurrentValue -Force | Out-Null
                } else {
                    Remove-ItemProperty -Path $path -Name $entry.Name -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-Warning ("Failed to restore user-scoped registry values for SID {0}: {1}" -f $sid, $_.Exception.Message)
        } finally {
            Close-UserRegistryTarget -Target $access
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

function ConvertTo-WsManTextValue {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [bool]) {
        return $Value.ToString().ToLowerInvariant()
    }

    switch -Exact (([string]$Value).ToLowerInvariant()) {
        'true' { 'true' }
        'false' { 'false' }
        default { [string]$Value }
    }
}

function ConvertFrom-WsManTextValue {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    switch -Exact (([string]$Value).Trim().ToLowerInvariant()) {
        'true' { return $true }
        'false' { return $false }
        '1' { return $true }
        '0' { return $false }
        default { return ([string]$Value).Trim() }
    }
}

function ConvertTo-NullableBoolean {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [bool]) {
        return [bool]$Value
    }

    switch -Exact (([string]$Value).Trim().ToLowerInvariant()) {
        'true' { return $true }
        'false' { return $false }
        '1' { return $true }
        '0' { return $false }
        default { return $null }
    }
}

function Resolve-WsManConfigTarget {
    param([Parameter(Mandatory)] [string]$Path)

    $prefix = 'WSMan:\localhost\'
    if (-not $Path.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsupported WSMan path: $Path"
    }

    $relativePath = $Path.Substring($prefix.Length)
    $segments = @($relativePath -split '\\')
    if ($segments.Count -lt 2) {
        throw "Unsupported WSMan path: $Path"
    }

    $role = [string]$segments[0]
    $resourceRole = switch -Exact ($role.ToLowerInvariant()) {
        'service' { 'service' }
        'client' { 'client' }
        default { throw "Unsupported WSMan path role: $Path" }
    }

    if ($segments.Count -eq 2) {
        $name = [string]$segments[1]
        switch -Exact ($name) {
            'AllowUnencrypted' {
                return [PSCustomObject]@{
                    ResourceUri = "winrm/config/$resourceRole"
                    Property    = 'AllowUnencrypted'
                    ValueKind   = 'Boolean'
                }
            }
            'TrustedHosts' {
                if ($resourceRole -ne 'client') {
                    throw "Unsupported WSMan path for TrustedHosts: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/client'
                    Property    = 'TrustedHosts'
                    ValueKind   = 'String'
                }
            }
            'IPv4Filter' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for IPv4Filter: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'IPv4Filter'
                    ValueKind   = 'String'
                }
            }
            'IPv6Filter' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for IPv6Filter: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'IPv6Filter'
                    ValueKind   = 'String'
                }
            }
            'EnableCompatibilityHttpListener' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for EnableCompatibilityHttpListener: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'EnableCompatibilityHttpListener'
                    ValueKind   = 'Boolean'
                }
            }
            'EnableCompatibilityHttpsListener' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for EnableCompatibilityHttpsListener: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'EnableCompatibilityHttpsListener'
                    ValueKind   = 'Boolean'
                }
            }
            'CbtHardeningLevel' {
                if ($resourceRole -ne 'service') {
                    throw "Unsupported WSMan path for CbtHardeningLevel: $Path"
                }

                return [PSCustomObject]@{
                    ResourceUri = 'winrm/config/service'
                    Property    = 'CbtHardeningLevel'
                    ValueKind   = 'String'
                }
            }
            default {
                throw "Unsupported WSMan path: $Path"
            }
        }
    }

    if ($segments.Count -eq 3 -and [string]$segments[1] -eq 'Auth') {
        $property = [string]$segments[2]
        $supportedAuthProperties = if ($resourceRole -eq 'client') {
            @('Basic', 'Digest', 'Kerberos', 'Negotiate', 'Certificate', 'CredSSP')
        } else {
            @('Basic', 'Kerberos', 'Negotiate', 'Certificate', 'CredSSP')
        }

        if ($supportedAuthProperties -notcontains $property) {
            throw "Unsupported WSMan auth path: $Path"
        }

        return [PSCustomObject]@{
            ResourceUri = "winrm/config/$resourceRole/auth"
            Property    = $property
            ValueKind   = 'Boolean'
        }
    }

    throw "Unsupported WSMan path: $Path"
}

function ConvertFrom-WsManPropertyValue {
    param(
        [AllowNull()] [object]$Value,
        [Parameter(Mandatory)] [string]$ValueKind
    )

    switch ($ValueKind) {
        'Boolean' { return (ConvertTo-NullableBoolean -Value $Value) }
        'String' {
            if ($null -eq $Value) {
                return $null
            }

            return [string]$Value
        }
        default {
            throw "Unsupported WSMan value kind: $ValueKind"
        }
    }
}

function ConvertTo-WsManSetValue {
    param(
        [AllowNull()] [object]$Value,
        [Parameter(Mandatory)] [string]$ValueKind
    )

    switch ($ValueKind) {
        'Boolean' {
            $boolValue = ConvertTo-NullableBoolean -Value $Value
            if ($null -eq $boolValue) {
                throw "Unsupported WSMan boolean value: $Value"
            }

            return (ConvertTo-WsManTextValue -Value $boolValue)
        }
        'String' {
            if ($null -eq $Value) {
                return ''
            }

            return [string]$Value
        }
        default {
            throw "Unsupported WSMan value kind: $ValueKind"
        }
    }
}

function Get-WsManConfigValueState {
    param([Parameter(Mandatory)] [string]$Path)

    $target = Resolve-WsManConfigTarget -Path $Path
    if (-not (Test-CommandAvailable -Name 'Get-WSManInstance')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Captured         = $false
            Value            = $null
            Error            = 'Get-WSManInstance was not found.'
        }
    }

    try {
        $instance = Get-WSManInstance -ResourceURI $target.ResourceUri -ErrorAction Stop
    } catch {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $false
            Value            = $null
            Error            = $_.Exception.Message
        }
    }

    $property = $instance.PSObject.Properties[$target.Property]
    if ($null -eq $property) {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $false
            Value            = $null
            Error            = "WSMan property $($target.Property) was not returned for $Path."
        }
    }

    [PSCustomObject]@{
        CommandAvailable = $true
        Captured         = $true
        Value            = ConvertFrom-WsManPropertyValue -Value $property.Value -ValueKind $target.ValueKind
        Error            = $null
    }
}

function Invoke-WithTemporaryWinRmServiceForWrite {
    param([Parameter(Mandatory)] [scriptblock]$ScriptBlock)

    $service = Get-CimInstance Win32_Service -Filter "Name='WinRM'" -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        & $ScriptBlock
        return
    }

    $originalStartMode = [string]$service.StartMode
    $wasRunning = ([string]$service.State -eq 'Running')
    $changedStartMode = $false
    $startedService = $false

    try {
        if (-not $wasRunning) {
            if ($originalStartMode -eq 'Disabled') {
                & sc.exe config WinRM "start= demand" | Out-Null
                $changedStartMode = $true
            }

            Start-Service -Name 'WinRM' -ErrorAction Stop
            $startedService = $true
        }

        & $ScriptBlock
    } finally {
        if ($startedService) {
            Stop-Service -Name 'WinRM' -Force -ErrorAction SilentlyContinue
        }

        if ($changedStartMode) {
            $startModeValue = Convert-ServiceStartModeToScValue -StartMode $originalStartMode
            & sc.exe config WinRM "start= $startModeValue" | Out-Null
        }
    }
}

function Get-WsManConfigValue {
    param([Parameter(Mandatory)] [string]$Path)

    (Get-WsManConfigValueState -Path $Path).Value
}

function Set-WsManConfigValue {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [AllowNull()] [object]$Value
    )

    $target = Resolve-WsManConfigTarget -Path $Path
    if (-not (Test-CommandAvailable -Name 'Set-WSManInstance')) {
        throw 'Set-WSManInstance was not found.'
    }

    $valueSet = @{
        $target.Property = ConvertTo-WsManSetValue -Value $Value -ValueKind $target.ValueKind
    }

    Invoke-WithTemporaryWinRmServiceForWrite -ScriptBlock {
        Set-WSManInstance -ResourceURI $target.ResourceUri -ValueSet $valueSet -ErrorAction Stop | Out-Null
    }
}

function Get-WinRmListenerStates {
    if (-not (Test-CommandAvailable -Name 'Get-WSManInstance')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Captured         = $false
            Listeners        = @()
            Error            = 'Get-WSManInstance was not found.'
        }
    }

    try {
        $listenerItems = @(Get-WSManInstance -ResourceURI 'winrm/config/listener' -Enumerate -ErrorAction Stop)
    } catch {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $false
            Listeners        = @()
            Error            = $_.Exception.Message
        }
    }

    $listeners = @(
        foreach ($listener in $listenerItems) {
            [PSCustomObject]@{
                Address               = if ($listener.PSObject.Properties['Address']) { [string]$listener.Address } else { $null }
                Transport             = if ($listener.PSObject.Properties['Transport']) { [string]$listener.Transport } else { $null }
                Port                  = if ($listener.PSObject.Properties['Port'] -and -not [string]::IsNullOrWhiteSpace([string]$listener.Port)) { [int]$listener.Port } else { $null }
                Hostname              = if ($listener.PSObject.Properties['Hostname']) { [string]$listener.Hostname } else { $null }
                Enabled               = if ($listener.PSObject.Properties['Enabled']) { ConvertTo-NullableBoolean -Value $listener.Enabled } else { $null }
                URLPrefix             = if ($listener.PSObject.Properties['URLPrefix']) { [string]$listener.URLPrefix } else { $null }
                CertificateThumbprint = if ($listener.PSObject.Properties['CertificateThumbprint']) { [string]$listener.CertificateThumbprint } else { $null }
            }
        }
    ) | Sort-Object -Property Address, Transport, Port, Hostname, URLPrefix, CertificateThumbprint

    [PSCustomObject]@{
        CommandAvailable = $true
        Captured         = $true
        Listeners        = @($listeners)
        Error            = $null
    }
}

function Remove-WinRmListener {
    param([Parameter(Mandatory)] [object]$Listener)

    if (-not (Test-CommandAvailable -Name 'Remove-WSManInstance')) {
        throw 'Remove-WSManInstance was not found.'
    }

    $selectorSet = @{
        Address   = [string]$Listener.Address
        Transport = [string]$Listener.Transport
    }

    Remove-WSManInstance -ResourceURI 'winrm/config/listener' -SelectorSet $selectorSet -ErrorAction Stop | Out-Null
}

function New-WinRmListener {
    param([Parameter(Mandatory)] [object]$Listener)

    if (-not (Test-CommandAvailable -Name 'New-WSManInstance')) {
        throw 'New-WSManInstance was not found.'
    }

    $selectorSet = @{
        Address   = [string]$Listener.Address
        Transport = [string]$Listener.Transport
    }

    $valueSet = @{}
    if ($null -ne $Listener.PSObject.Properties['Port'] -and $null -ne $Listener.Port) {
        $valueSet['Port'] = [string][int]$Listener.Port
    }

    if ($null -ne $Listener.PSObject.Properties['Hostname'] -and -not [string]::IsNullOrWhiteSpace([string]$Listener.Hostname)) {
        $valueSet['Hostname'] = [string]$Listener.Hostname
    }

    if ($null -ne $Listener.PSObject.Properties['Enabled'] -and $null -ne $Listener.Enabled) {
        $valueSet['Enabled'] = ConvertTo-WsManTextValue -Value ([bool]$Listener.Enabled)
    }

    if ($null -ne $Listener.PSObject.Properties['URLPrefix'] -and -not [string]::IsNullOrWhiteSpace([string]$Listener.URLPrefix)) {
        $valueSet['URLPrefix'] = [string]$Listener.URLPrefix
    }

    if (
        [string]$Listener.Transport -eq 'HTTPS' -and
        $null -ne $Listener.PSObject.Properties['CertificateThumbprint'] -and
        -not [string]::IsNullOrWhiteSpace([string]$Listener.CertificateThumbprint)
    ) {
        $valueSet['CertificateThumbprint'] = [string]$Listener.CertificateThumbprint
    }

    New-WSManInstance -ResourceURI 'winrm/config/listener' -SelectorSet $selectorSet -ValueSet $valueSet -ErrorAction Stop | Out-Null
}

function Set-WinRmListenersExact {
    param([Parameter(Mandatory)] [object[]]$Listeners)

    Invoke-WithTemporaryWinRmServiceForWrite -ScriptBlock {
        $liveState = Get-WinRmListenerStates
        if (-not $liveState.CommandAvailable) {
            throw 'Get-WSManInstance was not found.'
        }

        if (-not $liveState.Captured) {
            throw $liveState.Error
        }

        foreach ($listener in @($liveState.Listeners)) {
            Remove-WinRmListener -Listener $listener
        }

        foreach ($listener in @($Listeners)) {
            New-WinRmListener -Listener $listener
        }
    }
}

function Set-Permissive-WinRmListeners {
    param([Parameter(Mandatory)] [object[]]$Listeners)

    Set-WinRmListenersExact -Listeners @($Listeners)
}

function Restore-WinRmListeners {
    param([Parameter(Mandatory)] [object[]]$Listeners)

    Set-WinRmListenersExact -Listeners @($Listeners)
}

function Normalize-SmbConfigState {
    param([AllowNull()] [object]$Value)

    if ($Value -is [bool] -or $Value -is [string] -or $Value -is [int] -or $Value -is [long]) {
        $legacyValue = ConvertTo-NullableBoolean -Value $Value
        if ($null -ne $legacyValue) {
            return [PSCustomObject]@{
                CommandAvailable         = $true
                TimedOut                 = $false
                RequireSecuritySignature = $legacyValue
            }
        }
    }

    if ($null -eq $Value) {
        return [PSCustomObject]@{
            CommandAvailable         = $false
            TimedOut                 = $false
            RequireSecuritySignature = $null
        }
    }

    $commandAvailable = if ($Value.PSObject.Properties['CommandAvailable']) { [bool]$Value.CommandAvailable } else { $true }
    $timedOut = if ($Value.PSObject.Properties['TimedOut']) { [bool]$Value.TimedOut } else { $false }
    $requireSecuritySignature = if ($Value.PSObject.Properties['RequireSecuritySignature']) {
        ConvertTo-NullableBoolean -Value $Value.RequireSecuritySignature
    } else {
        ConvertTo-NullableBoolean -Value $Value
    }

    [PSCustomObject]@{
        CommandAvailable         = $commandAvailable
        TimedOut                 = $timedOut
        RequireSecuritySignature = $requireSecuritySignature
    }
}

function Get-SmbConfigurationState {
    param([Parameter(Mandatory)] [ValidateSet('Client', 'Server')] [string]$Role)

    $commandName = if ($Role -eq 'Client') { 'Get-SmbClientConfiguration' } else { 'Get-SmbServerConfiguration' }
    $roleLabel = if ($Role -eq 'Client') { 'SMB client' } else { 'SMB server' }

    if (-not (Test-CommandAvailable -Name $commandName)) {
        return Normalize-SmbConfigState -Value $null
    }

    $result = Invoke-ChildPowerShell -TimeoutSeconds 12 -ScriptText @"
`$ErrorActionPreference = 'Stop'
`$config = & $commandName -ErrorAction Stop
[PSCustomObject]@{
    RequireSecuritySignature = [bool]`$config.RequireSecuritySignature
} | ConvertTo-Json -Compress -Depth 4
"@

    if ($result.TimedOut) {
        Write-Warning "$roleLabel snapshot timed out. Skipping this setting."
        return Normalize-SmbConfigState -Value ([PSCustomObject]@{
            CommandAvailable         = $true
            TimedOut                 = $true
            RequireSecuritySignature = $null
        })
    }

    if ($result.ExitCode -ne 0) {
        if (-not [string]::IsNullOrWhiteSpace([string]$result.StdErr)) {
            Write-Warning ("{0} snapshot failed: {1}" -f $roleLabel, $result.StdErr.Trim())
        } else {
            Write-Warning "$roleLabel snapshot failed. Skipping this setting."
        }

        return Normalize-SmbConfigState -Value ([PSCustomObject]@{
            CommandAvailable         = $true
            TimedOut                 = $false
            RequireSecuritySignature = $null
        })
    }

    $json = [string]$result.StdOut
    if ([string]::IsNullOrWhiteSpace($json)) {
        Write-Warning "$roleLabel snapshot returned no data. Skipping this setting."
        return Normalize-SmbConfigState -Value ([PSCustomObject]@{
            CommandAvailable         = $true
            TimedOut                 = $false
            RequireSecuritySignature = $null
        })
    }

    try {
        $parsed = $json | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-Warning "$roleLabel snapshot returned unparsable output. Skipping this setting."
        return Normalize-SmbConfigState -Value ([PSCustomObject]@{
            CommandAvailable         = $true
            TimedOut                 = $false
            RequireSecuritySignature = $null
        })
    }

    $requireSecuritySignature = if ($parsed.PSObject.Properties['RequireSecuritySignature']) {
        ConvertTo-NullableBoolean -Value $parsed.RequireSecuritySignature
    } else {
        $null
    }

    Normalize-SmbConfigState -Value ([PSCustomObject]@{
        CommandAvailable         = $true
        TimedOut                 = $false
        RequireSecuritySignature = $requireSecuritySignature
    })
}

function Get-SmbClientConfigurationState {
    Get-SmbConfigurationState -Role 'Client'
}

function Get-SmbServerConfigurationState {
    Get-SmbConfigurationState -Role 'Server'
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

function Test-AsrRuleId {
    param([AllowNull()] [object]$Id)

    $text = [string]$Id
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $false
    }

    $guid = [guid]::Empty
    [guid]::TryParse($text, [ref]$guid)
}

function Get-AsrRuleCaptureState {
    $catalog = Get-AsrRuleCatalog
    $mp = Get-MpPreference -ErrorAction SilentlyContinue
    if ($null -eq $mp) {
        return [PSCustomObject]@{
            Rules          = @()
            InvalidEntries = @()
        }
    }

    $ids = @($mp.AttackSurfaceReductionRules_Ids)
    $actions = @($mp.AttackSurfaceReductionRules_Actions)
    $validRules = @()
    $invalidEntries = @()
    $entryCount = [Math]::Max($ids.Count, $actions.Count)

    for ($i = 0; $i -lt $entryCount; $i++) {
        $id = if ($ids.Count -gt $i) { [string]$ids[$i] } else { $null }
        $action = if ($actions.Count -gt $i) { [string]$actions[$i] } else { $null }
        $actionLabel = Get-AsrActionLabel -Action $action

        if (-not (Test-AsrRuleId -Id $id)) {
            $invalidEntries += [PSCustomObject]@{
                Id          = $id
                Action      = $action
                ActionLabel = $actionLabel
            }
            continue
        }

        $validRules += [PSCustomObject]@{
            Id          = $id
            Name        = if ($catalog.ContainsKey($id)) { $catalog[$id] } else { 'Unknown / custom rule' }
            Action      = $action
            ActionLabel = $actionLabel
        }
    }

    [PSCustomObject]@{
        Rules          = @($validRules)
        InvalidEntries = @($invalidEntries)
    }
}

function Get-AsrInvalidEntriesFromEntry {
    param([AllowNull()] [object]$Entry)

    if ($null -eq $Entry) {
        return @()
    }

    $invalidEntries = @()

    if ($Entry.PSObject.Properties['InvalidEntries']) {
        foreach ($rule in @($Entry.InvalidEntries)) {
            if ($null -eq $rule) {
                continue
            }

            $action = if ($rule.PSObject.Properties['Action']) { [string]$rule.Action } else { $null }
            $actionLabel = if ($rule.PSObject.Properties['ActionLabel']) { [string]$rule.ActionLabel } else { Get-AsrActionLabel -Action $action }

            $invalidEntries += [PSCustomObject]@{
                Id          = if ($rule.PSObject.Properties['Id']) { [string]$rule.Id } else { $null }
                Action      = $action
                ActionLabel = $actionLabel
            }
        }
    }

    foreach ($rule in @($Entry.CurrentValue)) {
        if ($null -eq $rule) {
            continue
        }

        $id = if ($rule.PSObject.Properties['Id']) { [string]$rule.Id } else { $null }
        if (Test-AsrRuleId -Id $id) {
            continue
        }

        $action = if ($rule.PSObject.Properties['Action']) { [string]$rule.Action } else { $null }
        $actionLabel = if ($rule.PSObject.Properties['ActionLabel']) { [string]$rule.ActionLabel } else { Get-AsrActionLabel -Action $action }
        $invalidEntries += [PSCustomObject]@{
            Id          = $id
            Action      = $action
            ActionLabel = $actionLabel
        }
    }

    @($invalidEntries)
}

function Get-ConfiguredAsrRules {
    @((Get-AsrRuleCaptureState).Rules)
}

function Disable-ConfiguredAsrRules {
    $mp = Get-MpPreference -ErrorAction SilentlyContinue
    if ($null -eq $mp) {
        return
    }

    foreach ($id in @($mp.AttackSurfaceReductionRules_Ids)) {
        if (-not (Test-AsrRuleId -Id $id)) {
            continue
        }

        Remove-MpPreference -AttackSurfaceReductionRules_Ids $id -ErrorAction SilentlyContinue
    }
}

function Restore-AsrRules {
    param([Parameter(Mandatory)] [object[]]$Rules)

    Disable-ConfiguredAsrRules

    $ids = @()
    $actions = @()

    foreach ($rule in @($Rules)) {
        if (-not (Test-AsrRuleId -Id $rule.Id)) {
            continue
        }

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
        [Parameter(Mandatory)] [AllowNull()] [object]$Value
    )

    $command = Get-Command -Name 'Set-MpPreference' -ErrorAction SilentlyContinue
    if ($null -eq $command -or -not $command.Parameters.ContainsKey($Property)) {
        return
    }

    if ($null -eq $Value) {
        return
    }

    $params = @{}
    $params[$Property] = $Value
    Set-MpPreference @params
}

function Get-MpPreferencePropertyRawValue {
    param([Parameter(Mandatory)] [string]$Property)

    $mp = Get-MpPreference -ErrorAction SilentlyContinue
    if ($null -eq $mp) {
        return $null
    }

    $propertyInfo = $mp.PSObject.Properties[$Property]
    if ($null -eq $propertyInfo) {
        return $null
    }

    $propertyInfo.Value
}

function Resolve-MpPreferenceValue {
    param(
        [Parameter(Mandatory)] [object]$Definition,
        [Parameter(Mandatory)] [AllowNull()] [object]$Value
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

function Normalize-MpPreferenceListItems {
    param([AllowNull()] [object]$Value)

    @(
        foreach ($item in @($Value)) {
            $text = [string]$item
            if ([string]::IsNullOrWhiteSpace($text)) {
                continue
            }

            $text.Trim()
        }
    ) | Sort-Object -Unique
}

function Get-MpPreferenceListState {
    param([Parameter(Mandatory)] [string]$Property)

    if (-not (Test-CommandAvailable -Name 'Get-MpPreference')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Captured         = $false
            Items            = @()
            Error            = $null
        }
    }

    try {
        $mp = Get-MpPreference -ErrorAction Stop
    } catch {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $false
            Items            = @()
            Error            = $_.Exception.Message
        }
    }

    $propertyInfo = $mp.PSObject.Properties[$Property]
    if ($null -eq $propertyInfo) {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Captured         = $false
            Items            = @()
            Error            = "Get-MpPreference does not expose property '$Property'."
        }
    }

    [PSCustomObject]@{
        CommandAvailable = $true
        Captured         = $true
        Items            = @(Normalize-MpPreferenceListItems -Value $propertyInfo.Value)
        Error            = $null
    }
}

function Test-MpPreferenceListCapturedExactly {
    param([Parameter(Mandatory)] [object]$Entry)

    $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
    $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
    ($commandAvailable -and $captured)
}

function Set-MpPreferenceListValue {
    param(
        [Parameter(Mandatory)] [string]$Property,
        [AllowNull()] [object[]]$DesiredItems
    )

    $addCommand = Get-Command -Name 'Add-MpPreference' -ErrorAction SilentlyContinue | Select-Object -First 1
    $removeCommand = Get-Command -Name 'Remove-MpPreference' -ErrorAction SilentlyContinue | Select-Object -First 1
    if (
        $null -eq $addCommand -or
        $null -eq $removeCommand -or
        -not $addCommand.Parameters.ContainsKey($Property) -or
        -not $removeCommand.Parameters.ContainsKey($Property)
    ) {
        return
    }

    $currentState = Get-MpPreferenceListState -Property $Property
    if (-not ($currentState.CommandAvailable -and $currentState.Captured)) {
        return
    }

    $currentItems = @(Normalize-MpPreferenceListItems -Value $currentState.Items)
    $targetItems = @(Normalize-MpPreferenceListItems -Value $DesiredItems)
    $itemsToRemove = @($currentItems | Where-Object { $_ -notin $targetItems })
    $itemsToAdd = @($targetItems | Where-Object { $_ -notin $currentItems })

    if ($itemsToRemove.Count -gt 0) {
        $removeParams = @{}
        $removeParams[$Property] = @($itemsToRemove)
        Remove-MpPreference @removeParams
    }

    if ($itemsToAdd.Count -gt 0) {
        $addParams = @{}
        $addParams[$Property] = @($itemsToAdd)
        Add-MpPreference @addParams
    }
}

function Get-DefenderRuntimeStatus {
    if (-not (Test-CommandAvailable -Name 'Get-MpComputerStatus')) {
        return [PSCustomObject]@{
            CommandAvailable          = $false
            Captured                  = $false
            AMRunningMode             = $null
            RealTimeProtectionEnabled = $null
            AntivirusEnabled          = $null
            IsTamperProtected         = $null
            Error                     = $null
        }
    }

    try {
        $status = Get-MpComputerStatus -ErrorAction Stop
    } catch {
        return [PSCustomObject]@{
            CommandAvailable          = $true
            Captured                  = $false
            AMRunningMode             = $null
            RealTimeProtectionEnabled = $null
            AntivirusEnabled          = $null
            IsTamperProtected         = $null
            Error                     = $_.Exception.Message
        }
    }

    [PSCustomObject]@{
        CommandAvailable          = $true
        Captured                  = $true
        AMRunningMode             = if ($status.PSObject.Properties['AMRunningMode']) { [string]$status.AMRunningMode } else { $null }
        RealTimeProtectionEnabled = if ($status.PSObject.Properties['RealTimeProtectionEnabled']) { ConvertTo-NullableBoolean -Value $status.RealTimeProtectionEnabled } else { $null }
        AntivirusEnabled          = if ($status.PSObject.Properties['AntivirusEnabled']) { ConvertTo-NullableBoolean -Value $status.AntivirusEnabled } else { $null }
        IsTamperProtected         = if ($status.PSObject.Properties['IsTamperProtected']) { ConvertTo-NullableBoolean -Value $status.IsTamperProtected } else { $null }
        Error                     = $null
    }
}

function Test-DefenderRuntimeStatusCapturedExactly {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $false
    }

    $commandAvailable = if ($State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
    $captured = if ($State.PSObject.Properties['Captured']) { [bool]$State.Captured } else { $true }
    ($commandAvailable -and $captured)
}

function ConvertTo-BitLockerProtectionStatusLabel {
    param([AllowNull()] [object]$Value)

    switch ([string]$Value) {
        '0' { 'Off' }
        '1' { 'On' }
        '2' { 'Unknown' }
        'Off' { 'Off' }
        'On' { 'On' }
        'Unknown' { 'Unknown' }
        default { [string]$Value }
    }
}

function Test-BitLockerProtectionEnabled {
    param([AllowNull()] [object]$Value)

    (ConvertTo-BitLockerProtectionStatusLabel -Value $Value) -eq 'On'
}

function Get-BitLockerTimedOutMountPoints {
    param([AllowNull()] [object]$State)

    if ($null -eq $State -or -not $State.PSObject.Properties['TimedOutMountPoints']) {
        return @()
    }

    @(
        foreach ($mountPoint in @($State.TimedOutMountPoints)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$mountPoint)) {
                [string]$mountPoint
            }
        }
    )
}

function Get-BitLockerCommandAvailableFlag {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $false
    }

    if ($State.PSObject.Properties['CommandAvailable']) {
        return [bool]$State.CommandAvailable
    }

    $true
}

function New-BitLockerCaptureIssue {
    param(
        [string]$MountPoint,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        MountPoint = $MountPoint
        Message    = $Message
    }
}

function Get-BitLockerCaptureIssues {
    param([AllowNull()] [object]$State)

    if ($null -eq $State -or -not $State.PSObject.Properties['CaptureIssues']) {
        return @()
    }

    @(
        foreach ($issue in @($State.CaptureIssues)) {
            if ($null -eq $issue) {
                continue
            }

            [PSCustomObject]@{
                MountPoint = if ($issue.PSObject.Properties['MountPoint']) { [string]$issue.MountPoint } else { $null }
                Message    = if ($issue.PSObject.Properties['Message']) { [string]$issue.Message } else { $null }
            }
        }
    )
}

function Get-BitLockerProtectionModeLabel {
    param(
        [AllowNull()] [object]$ProtectionStatus,
        [AllowNull()] [object]$VolumeStatus,
        [AllowNull()] [object]$EncryptionPercentage,
        [AllowNull()] [object]$KeyProtectorCount
    )

    $protectionLabel = ConvertTo-BitLockerProtectionStatusLabel -Value $ProtectionStatus
    if ($protectionLabel -eq 'On') {
        return 'Protected'
    }

    if ([string]$VolumeStatus -eq 'FullyDecrypted' -or ($null -ne $EncryptionPercentage -and [int]$EncryptionPercentage -eq 0)) {
        return 'Decrypted'
    }

    if ($protectionLabel -eq 'Off') {
        if ($null -ne $KeyProtectorCount -and [int]$KeyProtectorCount -gt 0) {
            return 'Suspended'
        }

        if ([string]$VolumeStatus -match 'Encrypt') {
            return 'Suspended'
        }

        return 'ProtectionOff'
    }

    if ([string]::IsNullOrWhiteSpace([string]$protectionLabel)) {
        return 'Unknown'
    }

    return [string]$protectionLabel
}

function Test-BitLockerAutoUnlockSupportedVolume {
    param([AllowNull()] [object]$Volume)

    if ($null -eq $Volume) {
        return $false
    }

    $mountPoint = if ($Volume.PSObject.Properties['MountPoint']) { [string]$Volume.MountPoint } else { $null }
    if ([string]::IsNullOrWhiteSpace($mountPoint)) {
        return $false
    }

    $volumeType = if ($Volume.PSObject.Properties['VolumeType']) { [string]$Volume.VolumeType } else { $null }
    if ([string]::Equals($volumeType, 'OperatingSystem', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    $true
}

function Normalize-BitLockerVolumeState {
    param([AllowNull()] [object]$Volume)

    if ($null -eq $Volume) {
        return $null
    }

    $keyProtectors = @()
    if ($Volume.PSObject.Properties['KeyProtectors']) {
        $keyProtectors = @(
            foreach ($protector in @($Volume.KeyProtectors)) {
                if ($null -eq $protector) {
                    continue
                }

                [PSCustomObject]@{
                    KeyProtectorId   = if ($protector.PSObject.Properties['KeyProtectorId']) { [string]$protector.KeyProtectorId } else { $null }
                    KeyProtectorType = if ($protector.PSObject.Properties['KeyProtectorType']) { [string]$protector.KeyProtectorType } else { $null }
                }
            }
        )
    }

    $keyProtectorCount = if ($Volume.PSObject.Properties['KeyProtectorCount'] -and $null -ne $Volume.KeyProtectorCount) {
        [int]$Volume.KeyProtectorCount
    } elseif ($Volume.PSObject.Properties['KeyProtectors']) {
        @($keyProtectors).Count
    } else {
        $null
    }

    $encryptionPercentage = if ($Volume.PSObject.Properties['EncryptionPercentage'] -and $null -ne $Volume.EncryptionPercentage) {
        [int]$Volume.EncryptionPercentage
    } else {
        $null
    }

    $protectionStatus = ConvertTo-BitLockerProtectionStatusLabel -Value $(if ($Volume.PSObject.Properties['ProtectionStatus']) { $Volume.ProtectionStatus } else { $null })
    $volumeStatus = if ($Volume.PSObject.Properties['VolumeStatus']) { [string]$Volume.VolumeStatus } else { $null }

    [PSCustomObject]@{
        MountPoint           = if ($Volume.PSObject.Properties['MountPoint']) { [string]$Volume.MountPoint } else { $null }
        VolumeType           = if ($Volume.PSObject.Properties['VolumeType']) { [string]$Volume.VolumeType } else { $null }
        ProtectionStatus     = $protectionStatus
        ProtectionMode       = if ($Volume.PSObject.Properties['ProtectionMode'] -and -not [string]::IsNullOrWhiteSpace([string]$Volume.ProtectionMode)) {
            [string]$Volume.ProtectionMode
        } else {
            Get-BitLockerProtectionModeLabel -ProtectionStatus $protectionStatus -VolumeStatus $volumeStatus -EncryptionPercentage $encryptionPercentage -KeyProtectorCount $keyProtectorCount
        }
        VolumeStatus         = $volumeStatus
        LockStatus           = if ($Volume.PSObject.Properties['LockStatus']) { [string]$Volume.LockStatus } else { $null }
        EncryptionMethod     = if ($Volume.PSObject.Properties['EncryptionMethod']) { [string]$Volume.EncryptionMethod } else { $null }
        EncryptionPercentage = $encryptionPercentage
        KeyProtectorCount    = $keyProtectorCount
        KeyProtectors        = @($keyProtectors)
        AutoUnlockEnabled    = if ($Volume.PSObject.Properties['AutoUnlockEnabled']) { ConvertTo-NullableBoolean -Value $Volume.AutoUnlockEnabled } else { $null }
    }
}

function Normalize-BitLockerState {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return [PSCustomObject]@{
            CommandAvailable    = $false
            TimedOutMountPoints = @()
            CaptureIssues       = @()
            Volumes             = @()
        }
    }

    [PSCustomObject]@{
        CommandAvailable    = Get-BitLockerCommandAvailableFlag -State $State
        TimedOutMountPoints = @(Get-BitLockerTimedOutMountPoints -State $State)
        CaptureIssues       = @(Get-BitLockerCaptureIssues -State $State)
        Volumes             = @(
            foreach ($volume in @($State.Volumes)) {
                Normalize-BitLockerVolumeState -Volume $volume
            }
        )
    }
}

function Test-BitLockerStateHasExtendedFields {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $false
    }

    if ($State.PSObject.Properties['CaptureIssues']) {
        return $true
    }

    foreach ($volume in @($State.Volumes)) {
        if ($null -eq $volume) {
            continue
        }

        foreach ($propertyName in @('LockStatus', 'ProtectionMode', 'KeyProtectors')) {
            if ($volume.PSObject.Properties[$propertyName]) {
                return $true
            }
        }
    }

    $false
}

function Get-BitLockerProtectorTypeSummary {
    param([AllowNull()] [object]$Volume)

    if ($null -eq $Volume) {
        return '<unknown>'
    }

    $types = @(
        foreach ($protector in @($Volume.KeyProtectors)) {
            if ($null -eq $protector) {
                continue
            }

            $type = if ($protector.PSObject.Properties['KeyProtectorType']) { [string]$protector.KeyProtectorType } else { $null }
            if ([string]::IsNullOrWhiteSpace($type)) {
                continue
            }

            $type
        }
    ) | Sort-Object -Unique

    if (@($types).Count -gt 0) {
        return ($types -join ', ')
    }

    if ($Volume.PSObject.Properties['KeyProtectorCount'] -and $null -ne $Volume.KeyProtectorCount) {
        $count = [int]$Volume.KeyProtectorCount
        if ($count -eq 0) {
            return '<none>'
        }

        return ('<count={0}>' -f $count)
    }

    '<unknown>'
}

function ConvertTo-ComparableBitLockerState {
    param(
        [AllowNull()] [object]$State,
        [AllowNull()] [object]$ReferenceState
    )

    $normalizedState = Normalize-BitLockerState -State $State
    $comparisonReference = if ($null -ne $ReferenceState) { $ReferenceState } else { $State }
    $useExtendedFields = Test-BitLockerStateHasExtendedFields -State $comparisonReference
    $comparisonCommandAvailable = $normalizedState.CommandAvailable
    if (-not $useExtendedFields -and @($normalizedState.CaptureIssues).Count -gt 0) {
        $comparisonCommandAvailable = $false
    }

    $comparableState = [ordered]@{
        CommandAvailable    = $comparisonCommandAvailable
        TimedOutMountPoints = @(
            foreach ($mountPoint in @($normalizedState.TimedOutMountPoints | Sort-Object -Unique)) {
                [string]$mountPoint
            }
        )
    }

    if ($useExtendedFields) {
        $comparableState['CaptureIssues'] = @(
            foreach ($issue in @($normalizedState.CaptureIssues | Sort-Object -Property MountPoint, Message)) {
                [PSCustomObject]@{
                    MountPoint = if (-not [string]::IsNullOrWhiteSpace([string]$issue.MountPoint)) { [string]$issue.MountPoint } else { $null }
                    Message    = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Message)) { [string]$issue.Message } else { $null }
                }
            }
        )
    }

    $comparableState['Volumes'] = @(
        foreach ($volume in @($normalizedState.Volumes | Sort-Object -Property MountPoint)) {
            if ($useExtendedFields) {
                [PSCustomObject]@{
                    MountPoint           = [string]$volume.MountPoint
                    VolumeType           = if (-not [string]::IsNullOrWhiteSpace([string]$volume.VolumeType)) { [string]$volume.VolumeType } else { $null }
                    ProtectionStatus     = [string]$volume.ProtectionStatus
                    ProtectionMode       = if (-not [string]::IsNullOrWhiteSpace([string]$volume.ProtectionMode)) { [string]$volume.ProtectionMode } else { $null }
                    VolumeStatus         = if (-not [string]::IsNullOrWhiteSpace([string]$volume.VolumeStatus)) { [string]$volume.VolumeStatus } else { $null }
                    LockStatus           = if (-not [string]::IsNullOrWhiteSpace([string]$volume.LockStatus)) { [string]$volume.LockStatus } else { $null }
                    EncryptionMethod     = if (-not [string]::IsNullOrWhiteSpace([string]$volume.EncryptionMethod)) { [string]$volume.EncryptionMethod } else { $null }
                    EncryptionPercentage = $volume.EncryptionPercentage
                    KeyProtectorCount    = $volume.KeyProtectorCount
                    KeyProtectors        = @(
                        foreach ($protector in @($volume.KeyProtectors | Sort-Object -Property KeyProtectorId, KeyProtectorType)) {
                            [PSCustomObject]@{
                                KeyProtectorId   = if (-not [string]::IsNullOrWhiteSpace([string]$protector.KeyProtectorId)) { [string]$protector.KeyProtectorId } else { $null }
                                KeyProtectorType = if (-not [string]::IsNullOrWhiteSpace([string]$protector.KeyProtectorType)) { [string]$protector.KeyProtectorType } else { $null }
                            }
                        }
                    )
                    AutoUnlockEnabled    = $volume.AutoUnlockEnabled
                }
                continue
            }

            [PSCustomObject]@{
                MountPoint           = [string]$volume.MountPoint
                ProtectionStatus     = [string]$volume.ProtectionStatus
                VolumeStatus         = [string]$volume.VolumeStatus
                EncryptionPercentage = $volume.EncryptionPercentage
                KeyProtectorCount    = $volume.KeyProtectorCount
            }
        }
    )

    $comparableState
}

function Get-ManageBdeCommand {
    foreach ($name in @('manage-bde.exe', 'manage-bde')) {
        $command = Get-Command -Name $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $command) {
            return $command
        }
    }

    $null
}

function Set-BitLockerAutoUnlockState {
    param(
        [Parameter(Mandatory)] [string]$MountPoint,
        [Parameter(Mandatory)] [bool]$Enabled
    )

    $manageBde = Get-ManageBdeCommand
    if ($null -eq $manageBde) {
        throw 'manage-bde was not found.'
    }

    $action = if ($Enabled) { '-enable' } else { '-disable' }
    $output = (& $manageBde.Source -autounlock $action $MountPoint 2>&1 | Out-String).Trim()
    if ($LASTEXITCODE -ne 0) {
        $message = "manage-bde -autounlock $action $MountPoint failed with exit code $LASTEXITCODE"
        if (-not [string]::IsNullOrWhiteSpace($output)) {
            $message = "$message. Output: $output"
        }

        throw $message
    }
}

function Test-BitLockerStateCapturedExactly {
    param([AllowNull()] [object]$State)

    $normalized = Normalize-BitLockerState -State $State
    if ($null -eq $normalized) {
        return $false
    }

    $normalized.CommandAvailable -and (@($normalized.TimedOutMountPoints).Count -eq 0) -and (@($normalized.CaptureIssues).Count -eq 0)
}

function Get-BitLockerVolumeStates {
    if (-not (Test-CommandAvailable -Name 'Get-BitLockerVolume')) {
        return [PSCustomObject]@{
            CommandAvailable   = $false
            TimedOutMountPoints = @()
            CaptureIssues      = @()
            Volumes            = @()
        }
    }

    $mountPoints = @(
        Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 2 OR DriveType = 3' -ErrorAction SilentlyContinue |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.DeviceID) } |
            ForEach-Object { [string]$_.DeviceID } |
            Sort-Object -Unique
    )

    $timedOutMountPoints = [System.Collections.Generic.List[string]]::new()
    $captureIssues = [System.Collections.Generic.List[object]]::new()
    $states = foreach ($mountPoint in $mountPoints) {
        Write-Verbose ("BitLocker snapshot mount point {0}" -f $mountPoint)
        $escapedMountPoint = $mountPoint.Replace("'", "''")
        $result = Invoke-ChildPowerShell -TimeoutSeconds 12 -ScriptText @"
`$ErrorActionPreference = 'Stop'
`$volume = Get-BitLockerVolume -MountPoint '$escapedMountPoint' -ErrorAction SilentlyContinue
if (`$null -eq `$volume) {
    return
}

[bool]`$keyProtectorsCaptured = `$false
`$keyProtectors = @()
if (`$volume.PSObject.Properties['KeyProtector']) {
    `$keyProtectorsCaptured = `$true
    `$keyProtectors = @(
        foreach (`$protector in @(`$volume.KeyProtector)) {
            if (`$null -eq `$protector) {
                continue
            }

            [PSCustomObject]@{
                KeyProtectorId   = if (`$protector.PSObject.Properties['KeyProtectorId']) { [string]`$protector.KeyProtectorId } else { `$null }
                KeyProtectorType = if (`$protector.PSObject.Properties['KeyProtectorType']) { [string]`$protector.KeyProtectorType } else { `$null }
            }
        }
    )
}

[bool]`$autoUnlockCaptured = `$false
`$autoUnlockEnabled = `$null
if (`$volume.PSObject.Properties['AutoUnlockEnabled']) {
    `$autoUnlockCaptured = `$true
    if (`$null -ne `$volume.AutoUnlockEnabled) {
        `$autoUnlockEnabled = [bool]`$volume.AutoUnlockEnabled
    }
}

[PSCustomObject]@{
    MountPoint           = [string]`$volume.MountPoint
    VolumeType           = [string]`$volume.VolumeType
    ProtectionStatus     = [string]`$volume.ProtectionStatus
    VolumeStatus         = [string]`$volume.VolumeStatus
    LockStatus           = if (`$volume.PSObject.Properties['LockStatus']) { [string]`$volume.LockStatus } else { `$null }
    EncryptionMethod     = [string]`$volume.EncryptionMethod
    EncryptionPercentage = if (`$null -ne `$volume.EncryptionPercentage) { [int]`$volume.EncryptionPercentage } else { `$null }
    KeyProtectorCount    = @(`$keyProtectors).Count
    KeyProtectors        = @(`$keyProtectors)
    KeyProtectorsCaptured = `$keyProtectorsCaptured
    AutoUnlockEnabled    = `$autoUnlockEnabled
    AutoUnlockCaptured   = `$autoUnlockCaptured
} | ConvertTo-Json -Compress -Depth 6
"@

        if ($result.TimedOut) {
            $timedOutMountPoints.Add($mountPoint) | Out-Null
            Write-Warning "BitLocker snapshot timed out on mount point $mountPoint. Skipping it."
            continue
        }

        if ($result.ExitCode -ne 0) {
            if (-not [string]::IsNullOrWhiteSpace([string]$result.StdErr)) {
                Write-Warning ("BitLocker snapshot failed on mount point {0}: {1}" -f $mountPoint, $result.StdErr.Trim())
                $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message $result.StdErr.Trim())) | Out-Null
            } else {
                $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker snapshot command failed for this mount point.')) | Out-Null
            }
            continue
        }

        $json = [string]$result.StdOut
        if ([string]::IsNullOrWhiteSpace($json)) {
            Write-Warning "BitLocker snapshot returned no data on mount point $mountPoint. Skipping it."
            $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker snapshot returned no data for this mount point.')) | Out-Null
            continue
        }

        try {
            $volume = $json | ConvertFrom-Json -ErrorAction Stop
        } catch {
            Write-Warning "BitLocker snapshot returned unparsable output on mount point $mountPoint. Skipping it."
            $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker snapshot returned unparsable output for this mount point.')) | Out-Null
            continue
        }

        $normalizedVolume = Normalize-BitLockerVolumeState -Volume $volume
        if (-not ($volume.PSObject.Properties['KeyProtectorsCaptured'] -and [bool]$volume.KeyProtectorsCaptured)) {
            $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker key protector inventory could not be captured for this mount point.')) | Out-Null
        }

        if ((Test-BitLockerAutoUnlockSupportedVolume -Volume $normalizedVolume) -and -not ($volume.PSObject.Properties['AutoUnlockCaptured'] -and [bool]$volume.AutoUnlockCaptured)) {
            $captureIssues.Add((New-BitLockerCaptureIssue -MountPoint $mountPoint -Message 'BitLocker auto-unlock state could not be captured for this mount point.')) | Out-Null
        }

        $normalizedVolume
    }

    [PSCustomObject]@{
        CommandAvailable    = $true
        TimedOutMountPoints = @($timedOutMountPoints)
        CaptureIssues       = @($captureIssues)
        Volumes             = @($states)
    }
}

function Set-Permissive-BitLockerVolumes {
    param([AllowNull()] [object]$State)

    if (-not (Test-CommandAvailable -Name 'Suspend-BitLocker')) {
        return
    }

    if ($null -eq $State) {
        if (-not (Test-CommandAvailable -Name 'Get-BitLockerVolume')) {
            return
        }

        $State = Get-BitLockerVolumeStates
    }

    if (-not (Test-BitLockerStateCapturedExactly -State $State)) {
        return
    }

    $normalizedState = Normalize-BitLockerState -State $State

    foreach ($volume in @($normalizedState.Volumes)) {
        if (-not (Test-BitLockerProtectionEnabled -Value $volume.ProtectionStatus)) {
            continue
        }

        $mountPoint = [string]$volume.MountPoint

        if ([string]::IsNullOrWhiteSpace($mountPoint)) {
            continue
        }

        Suspend-BitLocker -MountPoint $mountPoint -RebootCount 0 -ErrorAction SilentlyContinue | Out-Null
    }

    foreach ($volume in @($normalizedState.Volumes)) {
        if (-not (Test-BitLockerAutoUnlockSupportedVolume -Volume $volume)) {
            continue
        }

        if ($null -eq $volume.AutoUnlockEnabled -or $volume.AutoUnlockEnabled) {
            continue
        }

        if ([string]$volume.ProtectionMode -eq 'Decrypted') {
            continue
        }

        try {
            Set-BitLockerAutoUnlockState -MountPoint ([string]$volume.MountPoint) -Enabled $true
        } catch {
            Write-Warning ("Failed to enable BitLocker auto-unlock for {0}: {1}" -f ([string]$volume.MountPoint), $_.Exception.Message)
        }
    }
}

function Restore-BitLockerVolumes {
    param([Parameter(Mandatory)] [object]$State)

    if (-not (Test-BitLockerStateCapturedExactly -State $State)) {
        return
    }

    if (-not (Test-CommandAvailable -Name 'Get-BitLockerVolume')) {
        return
    }

    $normalizedState = Normalize-BitLockerState -State $State

    foreach ($volume in @($normalizedState.Volumes)) {
        $mountPoint = [string]$volume.MountPoint
        if ([string]::IsNullOrWhiteSpace($mountPoint)) {
            continue
        }

        $liveVolume = Get-BitLockerVolume -MountPoint $mountPoint -ErrorAction SilentlyContinue
        if ($null -eq $liveVolume) {
            continue
        }

        $targetProtection = ConvertTo-BitLockerProtectionStatusLabel -Value $volume.ProtectionStatus
        if ($targetProtection -eq 'On') {
            if (Test-CommandAvailable -Name 'Resume-BitLocker') {
                Resume-BitLocker -MountPoint $mountPoint -ErrorAction SilentlyContinue | Out-Null
            }
            continue
        }

        if (
            $targetProtection -eq 'Off' -and
            [string]$volume.VolumeStatus -ne 'FullyDecrypted' -and
            (Test-CommandAvailable -Name 'Suspend-BitLocker')
        ) {
            Suspend-BitLocker -MountPoint $mountPoint -RebootCount 0 -ErrorAction SilentlyContinue | Out-Null
        }
    }

    foreach ($volume in @($normalizedState.Volumes)) {
        if (-not (Test-BitLockerAutoUnlockSupportedVolume -Volume $volume)) {
            continue
        }

        if ($null -eq $volume.AutoUnlockEnabled) {
            continue
        }

        try {
            Set-BitLockerAutoUnlockState -MountPoint ([string]$volume.MountPoint) -Enabled ([bool]$volume.AutoUnlockEnabled)
        } catch {
            Write-Warning ("Failed to restore BitLocker auto-unlock for {0}: {1}" -f ([string]$volume.MountPoint), $_.Exception.Message)
        }
    }
}

function New-AppLockerCaptureIssue {
    param(
        [string]$Scope,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        Scope   = $Scope
        Message = $Message
    }
}

function ConvertTo-AppLockerXmlText {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return ([string]$Value).Trim()
    }

    ((@($Value) | ForEach-Object { [string]$_ }) -join [Environment]::NewLine).Trim()
}

function Normalize-AppLockerXml {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml)) {
        return ''
    }

    try {
        $document = New-Object System.Xml.XmlDocument
        $document.PreserveWhitespace = $false
        $document.LoadXml($Xml)
        $document.OuterXml
    } catch {
        $Xml.Trim()
    }
}

function Get-AppLockerCollectionSummaries {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml)) {
        return @()
    }

    try {
        $document = New-Object System.Xml.XmlDocument
        $document.PreserveWhitespace = $false
        $document.LoadXml($Xml)

        @(
            foreach ($ruleCollection in @($document.SelectNodes('/AppLockerPolicy/RuleCollection'))) {
                if ($null -eq $ruleCollection) {
                    continue
                }

                $ruleCount = @(
                    foreach ($childNode in @($ruleCollection.ChildNodes)) {
                        if (
                            $childNode -is [System.Xml.XmlElement] -and
                            $childNode.LocalName -like '*Rule'
                        ) {
                            $childNode
                        }
                    }
                ).Count

                $servicesNode = $ruleCollection.SelectSingleNode('RuleCollectionExtensions/ThresholdExtensions/Services')
                $systemAppsNode = $ruleCollection.SelectSingleNode('RuleCollectionExtensions/RedstoneExtensions/SystemApps')

                [PSCustomObject]@{
                    Type                = if ($ruleCollection.Attributes['Type']) { [string]$ruleCollection.Attributes['Type'].Value } else { $null }
                    EnforcementMode     = if ($ruleCollection.Attributes['EnforcementMode']) { [string]$ruleCollection.Attributes['EnforcementMode'].Value } else { 'NotConfigured' }
                    RuleCount           = $ruleCount
                    ServicesEnforcement = if ($null -ne $servicesNode -and $servicesNode.Attributes['EnforcementMode']) { [string]$servicesNode.Attributes['EnforcementMode'].Value } else { $null }
                    SystemAppsAllow     = if ($null -ne $systemAppsNode -and $systemAppsNode.Attributes['Allow']) { [string]$systemAppsNode.Attributes['Allow'].Value } else { $null }
                }
            }
        )
    } catch {
        @()
    }
}

function Get-AppLockerPolicyXml {
    param(
        [AllowNull()] [object]$State,
        [ValidateSet('Local', 'Effective')] [string]$PolicyScope = 'Effective',
        [string]$SnapshotPath
    )

    if ($null -eq $State) {
        return ''
    }

    $inlineProperty = if ($PolicyScope -eq 'Local') { 'LocalXml' } else { 'EffectiveXml' }
    $assetProperty = if ($PolicyScope -eq 'Local') { 'LocalSnapshotAssetRelativePath' } else { 'EffectiveSnapshotAssetRelativePath' }

    if (
        $State.PSObject.Properties[$inlineProperty] -and
        -not [string]::IsNullOrWhiteSpace([string]$State.PSObject.Properties[$inlineProperty].Value)
    ) {
        return [string]$State.PSObject.Properties[$inlineProperty].Value
    }

    if (
        $State.PSObject.Properties[$assetProperty] -and
        -not [string]::IsNullOrWhiteSpace([string]$State.PSObject.Properties[$assetProperty].Value) -and
        -not [string]::IsNullOrWhiteSpace($SnapshotPath)
    ) {
        $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $SnapshotPath
        $assetPath = Join-Path $assetRoot ([string]$State.PSObject.Properties[$assetProperty].Value)
        if (-not (Test-Path -LiteralPath $assetPath)) {
            throw "AppLocker snapshot asset is missing: $assetPath"
        }

        return Get-Content -LiteralPath $assetPath -Raw
    }

    ''
}

function Get-AppLockerPolicyState {
    if (-not (Test-CommandAvailable -Name 'Get-AppLockerPolicy')) {
        return [PSCustomObject]@{
            CommandAvailable      = $false
            LocalCaptured         = $false
            EffectiveCaptured     = $false
            LocalMatchesEffective = $false
            CaptureIssues         = @()
            CollectionSummaries   = @()
            LocalXml              = $null
            EffectiveXml          = $null
        }
    }

    $captureIssues = [System.Collections.Generic.List[object]]::new()
    $localXml = ''
    $effectiveXml = ''
    $localCaptured = $false
    $effectiveCaptured = $false

    try {
        $localXml = ConvertTo-AppLockerXmlText -Value (Get-AppLockerPolicy -Local -Xml -ErrorAction Stop)
        $localCaptured = $true
    } catch {
        $captureIssues.Add((New-AppLockerCaptureIssue -Scope 'Local' -Message $_.Exception.Message)) | Out-Null
    }

    try {
        $effectiveXml = ConvertTo-AppLockerXmlText -Value (Get-AppLockerPolicy -Effective -Xml -ErrorAction Stop)
        $effectiveCaptured = $true
    } catch {
        $captureIssues.Add((New-AppLockerCaptureIssue -Scope 'Effective' -Message $_.Exception.Message)) | Out-Null
    }

    $normalizedLocalXml = Normalize-AppLockerXml -Xml $localXml
    $normalizedEffectiveXml = Normalize-AppLockerXml -Xml $effectiveXml

    [PSCustomObject]@{
        CommandAvailable      = $true
        LocalCaptured         = $localCaptured
        EffectiveCaptured     = $effectiveCaptured
        LocalMatchesEffective = (
            $localCaptured -and
            $effectiveCaptured -and
            [string]::Equals($normalizedLocalXml, $normalizedEffectiveXml, [System.StringComparison]::Ordinal)
        )
        CaptureIssues         = @($captureIssues)
        CollectionSummaries   = @(Get-AppLockerCollectionSummaries -Xml $effectiveXml)
        LocalXml              = if ($localCaptured) { $localXml } else { $null }
        EffectiveXml          = if ($effectiveCaptured) { $effectiveXml } else { $null }
    }
}

function Test-AppLockerPolicyCapturedExactly {
    param(
        [AllowNull()] [object]$State,
        [string]$SnapshotPath
    )

    if ($null -eq $State) {
        return $false
    }

    $commandAvailable = if ($State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
    $localCaptured = if ($State.PSObject.Properties['LocalCaptured']) { [bool]$State.LocalCaptured } else { $false }
    $effectiveCaptured = if ($State.PSObject.Properties['EffectiveCaptured']) { [bool]$State.EffectiveCaptured } else { $false }
    $localMatchesEffective = if ($State.PSObject.Properties['LocalMatchesEffective']) { [bool]$State.LocalMatchesEffective } else { $false }
    $captureIssues = if ($State.PSObject.Properties['CaptureIssues']) { @($State.CaptureIssues) } else { @() }
    $localXml = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $State -PolicyScope Local -SnapshotPath $SnapshotPath)
    $effectiveXml = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $State -PolicyScope Effective -SnapshotPath $SnapshotPath)

    (
        $commandAvailable -and
        $localCaptured -and
        $effectiveCaptured -and
        $localMatchesEffective -and
        (@($captureIssues).Count -eq 0) -and
        -not [string]::IsNullOrWhiteSpace($localXml) -and
        -not [string]::IsNullOrWhiteSpace($effectiveXml)
    )
}

function Get-EmptyAppLockerPolicyXml {
    '<AppLockerPolicy Version="1" />'
}

function Apply-AppLockerPolicyXml {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml) -or -not (Test-CommandAvailable -Name 'Set-AppLockerPolicy')) {
        return
    }

    $tempPath = New-TemporaryFilePath -Extension '.xml'
    try {
        [System.IO.File]::WriteAllText($tempPath, $Xml, [System.Text.UTF8Encoding]::new($false))
        Set-AppLockerPolicy -XmlPolicy $tempPath | Out-Null
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Set-Permissive-AppLockerPolicy {
    param(
        [AllowNull()] [object]$State,
        [string]$SnapshotPath
    )

    if ($null -eq $State) {
        $State = Get-AppLockerPolicyState
    }

    if (-not (Test-AppLockerPolicyCapturedExactly -State $State -SnapshotPath $SnapshotPath)) {
        return
    }

    Apply-AppLockerPolicyXml -Xml (Get-EmptyAppLockerPolicyXml)
}

function Restore-AppLockerPolicy {
    param(
        [Parameter(Mandatory)] [object]$State,
        [string]$SnapshotPath
    )

    if (-not (Test-AppLockerPolicyCapturedExactly -State $State -SnapshotPath $SnapshotPath)) {
        return
    }

    Apply-AppLockerPolicyXml -Xml (Get-AppLockerPolicyXml -State $State -PolicyScope Local -SnapshotPath $SnapshotPath)
}

function Get-ExploitProtectionPolicyState {
    if (-not (Test-CommandAvailable -Name 'Get-ProcessMitigation')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Xml              = $null
        }
    }

    $tempPath = New-TemporaryFilePath -Extension '.xml'
    try {
        Get-ProcessMitigation -RegistryConfigFilePath $tempPath | Out-Null
        $xml = if (Test-Path -LiteralPath $tempPath) { Get-Content -LiteralPath $tempPath -Raw } else { $null }

        [PSCustomObject]@{
            CommandAvailable = $true
            Xml              = $xml
        }
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Normalize-ExploitProtectionXml {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml)) {
        return ''
    }

    try {
        $document = New-Object System.Xml.XmlDocument
        $document.PreserveWhitespace = $false
        $document.LoadXml($Xml)

        $mitigationPolicyNode = $document.SelectSingleNode('/MitigationPolicy')
        $systemConfigNode = if ($null -ne $mitigationPolicyNode) { $mitigationPolicyNode.SelectSingleNode('SystemConfig') } else { $null }
        if ($null -ne $systemConfigNode) {
            $childElements = @($systemConfigNode.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] })
            if ($childElements.Count -eq 1 -and $childElements[0].Name -eq 'ASLR') {
                $aslrNode = [System.Xml.XmlElement]$childElements[0]
                $expectedAttributes = [ordered]@{
                    ForceRelocateImages = 'false'
                    RequireInfo         = 'false'
                    BottomUp            = 'false'
                    HighEntropy         = 'false'
                }

                $isDefaultNoOpSystemConfig = ($aslrNode.Attributes.Count -eq $expectedAttributes.Count)
                if ($isDefaultNoOpSystemConfig) {
                    foreach ($attributeName in $expectedAttributes.Keys) {
                        if (
                            -not $aslrNode.HasAttribute($attributeName) -or
                            -not [string]::Equals($aslrNode.GetAttribute($attributeName), $expectedAttributes[$attributeName], [System.StringComparison]::OrdinalIgnoreCase)
                        ) {
                            $isDefaultNoOpSystemConfig = $false
                            break
                        }
                    }
                }

                if ($isDefaultNoOpSystemConfig) {
                    $null = $mitigationPolicyNode.RemoveChild($systemConfigNode)
                }
            }
        }

        return $document.OuterXml
    } catch {
        return $Xml.Trim()
    }
}

function Get-ExploitProtectionPolicyXml {
    param(
        [AllowNull()] [object]$State,
        [string]$SnapshotPath
    )

    if ($null -eq $State) {
        return ''
    }

    if ($State.PSObject.Properties['Xml'] -and -not [string]::IsNullOrWhiteSpace([string]$State.Xml)) {
        return [string]$State.Xml
    }

    if (
        $State.PSObject.Properties['SnapshotAssetRelativePath'] -and
        -not [string]::IsNullOrWhiteSpace([string]$State.SnapshotAssetRelativePath) -and
        -not [string]::IsNullOrWhiteSpace($SnapshotPath)
    ) {
        $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $SnapshotPath
        $assetPath = Join-Path $assetRoot ([string]$State.SnapshotAssetRelativePath)
        if (-not (Test-Path -LiteralPath $assetPath)) {
            throw "Exploit protection snapshot asset is missing: $assetPath"
        }

        return Get-Content -LiteralPath $assetPath -Raw
    }

    ''
}

function Test-ExploitProtectionPolicyCapturedExactly {
    param(
        [AllowNull()] [object]$State,
        [string]$SnapshotPath
    )

    $commandAvailable = if ($null -ne $State -and $State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
    return (
        $null -ne $State -and
        $commandAvailable -and
        -not [string]::IsNullOrWhiteSpace((Get-ExploitProtectionPolicyXml -State $State -SnapshotPath $SnapshotPath))
    )
}

function Apply-ExploitProtectionPolicyXml {
    param([AllowNull()] [string]$Xml)

    if ([string]::IsNullOrWhiteSpace($Xml) -or -not (Test-CommandAvailable -Name 'Set-ProcessMitigation')) {
        return
    }

    $tempPath = New-TemporaryFilePath -Extension '.xml'
    try {
        [System.IO.File]::WriteAllText($tempPath, $Xml, [System.Text.UTF8Encoding]::new($false))
        Set-ProcessMitigation -PolicyFilePath $tempPath | Out-Null
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Reset-ExploitProtectionSystemConfig {
    if (-not (Test-CommandAvailable -Name 'Set-ProcessMitigation')) {
        return
    }

    try {
        Set-ProcessMitigation -System -Reset | Out-Null
    } catch {
    }
}

function Set-Permissive-ExploitProtection {
    Reset-ExploitProtectionSystemConfig

    foreach ($mitigation in @('DEP', 'EmulateAtlThunks', 'CFG', 'StrictCFG', 'SuppressExports', 'ForceRelocateImages', 'BottomUp', 'HighEntropy', 'SEHOP', 'SEHOPTelemetry')) {
        try {
            Set-ProcessMitigation -System -Disable $mitigation | Out-Null
        } catch {
        }
    }
}

function Get-WdacCodeIntegrityRoot {
    Resolve-FileSystemPath -Path (Join-Path $env:windir 'System32\CodeIntegrity')
}

function Get-WdacPolicyFileBackups {
    $root = Get-WdacCodeIntegrityRoot
    $files = @()

    $activeDir = Join-Path $root 'CiPolicies\Active'
    if (Test-Path -LiteralPath $activeDir) {
        $files += @(
            foreach ($file in @(Get-ChildItem -LiteralPath $activeDir -Filter '*.cip' -File -ErrorAction SilentlyContinue)) {
                [PSCustomObject]@{
                    RelativePath = Join-Path 'CiPolicies\Active' $file.Name
                    FileName     = $file.Name
                    Sha256       = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
                    Base64       = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file.FullName))
                }
            }
        )
    }

    $singlePolicyPath = Join-Path $root 'SiPolicy.p7b'
    if (Test-Path -LiteralPath $singlePolicyPath) {
        $files += [PSCustomObject]@{
            RelativePath = 'SiPolicy.p7b'
            FileName     = 'SiPolicy.p7b'
            Sha256       = (Get-FileHash -LiteralPath $singlePolicyPath -Algorithm SHA256).Hash
            Base64       = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($singlePolicyPath))
        }
    }

    @($files)
}

function Get-WdacPolicyIdentifiers {
    param([Parameter(Mandatory)] [object[]]$Policies)

    @(
        foreach ($policy in @($Policies)) {
            $policyId = Get-WdacPolicyIdentifierFromObject -Policy $policy
            if (-not [string]::IsNullOrWhiteSpace($policyId)) {
                $policyId
            }
        }
    ) | Sort-Object -Unique
}

function Get-WdacPolicyIdentifierFromObject {
    param([Parameter(Mandatory)] [object]$Policy)

    foreach ($name in @('PolicyID', 'PolicyId', 'PolicyGuid', 'PolicyGUID', 'Id', 'ID')) {
        $property = $Policy.PSObject.Properties[$name]
        if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return [string]$property.Value
        }
    }

    return $null
}

function Get-WdacNormalizedPolicyId {
    param([AllowNull()] [object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $text = $text.TrimStart('{').TrimEnd('}')
    $guid = [guid]::Empty
    if ([guid]::TryParse($text, [ref]$guid)) {
        return $guid.ToString().ToLowerInvariant()
    }

    return $text.ToLowerInvariant()
}

function Get-WdacPolicyIdFromFileName {
    param([AllowNull()] [string]$FileName)

    if ([string]::IsNullOrWhiteSpace($FileName)) {
        return $null
    }

    $leaf = [System.IO.Path]::GetFileName($FileName)
    if ($leaf -match '^\{?([0-9A-Fa-f-]{36})\}?\.cip$') {
        return Get-WdacNormalizedPolicyId -Value $matches[1]
    }

    return $null
}

function Get-WdacPolicyPlatformManagementInfo {
    param([Parameter(Mandatory)] [object]$Policy)

    $policyId = Get-WdacPolicyIdentifierFromObject -Policy $Policy
    $normalizedPolicyId = Get-WdacNormalizedPolicyId -Value $policyId
    $friendlyName = if (
        $Policy.PSObject.Properties['FriendlyName'] -and
        -not [string]::IsNullOrWhiteSpace([string]$Policy.FriendlyName)
    ) {
        [string]$Policy.FriendlyName
    } else {
        '<unknown>'
    }

    $platformPolicy = $null
    foreach ($propertyName in @('Platform Policy', 'PlatformPolicy')) {
        if ($Policy.PSObject.Properties[$propertyName]) {
            $platformPolicy = ConvertTo-NullableBoolean -Value $Policy.PSObject.Properties[$propertyName].Value
            if ($null -ne $platformPolicy) {
                break
            }
        }
    }

    $knownInboxPolicies = @{
        '0283ac0f-fff1-49ae-ada1-8a933130cad6' = 'Inbox Smart App Control base policy'
        '1283ac0f-fff1-49ae-ada1-8a933130cad6' = 'Inbox Smart App Control evaluation base policy'
        '1678656c-05ef-481f-bc5b-ebd8c991502d' = 'Inbox Smart App Control flight supplemental policy'
        '2678656c-05ef-481f-bc5b-ebd8c991502d' = 'Inbox Smart App Control evaluation flight supplemental policy'
        '0939ed82-bfd5-4d32-b58e-d31d3c49715a' = 'Inbox Smart App Control test supplemental policy'
        '1939ed82-bfd5-4d32-b58e-d31d3c49715a' = 'Inbox Smart App Control evaluation test supplemental policy'
        'd2bda982-ccf6-4344-ac5b-0b44427b6816' = 'Inbox Microsoft Windows Driver Policy'
        'a072029f-588b-4b5e-b7f9-05aad67df687' = 'Inbox Microsoft Windows Virtualization Based Security policy'
        '82443e1e-8a39-4b4a-96a8-f40ddc00b9f3' = 'Inbox Windows 11 SE lockdown base policy'
        '5dac656c-21ad-4a02-ab49-649917162e70' = 'Inbox Windows 11 SE flight supplemental policy'
        'cdd5cb55-db68-4d71-aa38-3df2b6473a52' = 'Inbox Windows 11 SE test supplemental policy'
        '5951a96a-e0b5-4d3d-8fb8-3e5b61030784' = 'Inbox Windows 10 S lockdown base policy'
        '784c4414-79f4-4c32-a6a5-f0fb42a51d0d' = 'Inbox Microsoft Code Integrity cross-certificates exception policy'
    }

    $classification = $null
    if (-not [string]::IsNullOrWhiteSpace($normalizedPolicyId) -and $knownInboxPolicies.ContainsKey($normalizedPolicyId)) {
        $classification = $knownInboxPolicies[$normalizedPolicyId]
    } elseif ($friendlyName -match '^VerifiedAndReputableDesktop') {
        $classification = 'Inbox Smart App Control policy'
    } elseif ($friendlyName -eq 'Microsoft Windows Driver Policy') {
        $classification = 'Inbox Microsoft Windows Driver Policy'
    } elseif ($friendlyName -eq 'Microsoft Windows Virtualization Based Security Policy') {
        $classification = 'Inbox Microsoft Windows Virtualization Based Security policy'
    } elseif ($friendlyName -match '^Windows(E|10S)_Lockdown') {
        $classification = 'Inbox Windows lockdown policy'
    } elseif ($platformPolicy -eq $true) {
        $classification = 'Platform-managed WDAC policy'
    }

    [PSCustomObject]@{
        PolicyId           = $policyId
        FriendlyName       = $friendlyName
        NormalizedPolicyId = $normalizedPolicyId
        PlatformPolicy     = $platformPolicy
        IsPlatformManaged  = -not [string]::IsNullOrWhiteSpace([string]$classification)
        Classification     = $classification
    }
}

function Get-WdacPlatformManagedPolicies {
    param([AllowNull()] [object]$State)

    if ($null -eq $State -or -not $State.PSObject.Properties['Policies']) {
        return @()
    }

    @(
        foreach ($policy in @($State.Policies)) {
            $info = Get-WdacPolicyPlatformManagementInfo -Policy $policy
            if ($info.IsPlatformManaged) {
                $info
            }
        }
    )
}

function Get-WdacPolicyReportRows {
    param([AllowNull()] [object]$State)

    if ($null -eq $State -or -not $State.PSObject.Properties['Policies']) {
        return @()
    }

    @(
        foreach ($policy in @($State.Policies)) {
            $policyInfo = Get-WdacPolicyPlatformManagementInfo -Policy $policy

            $hasFileOnDisk = $null
            if ($policy.PSObject.Properties['HasFileOnDisk']) {
                $hasFileOnDisk = ConvertTo-NullableBoolean -Value $policy.HasFileOnDisk
            }

            $isEnforced = $null
            foreach ($name in @('IsCurrentlyEnforced', 'IsEnforced')) {
                if ($policy.PSObject.Properties[$name]) {
                    $isEnforced = ConvertTo-NullableBoolean -Value $policy.$name
                    if ($null -ne $isEnforced) {
                        break
                    }
                }
            }

            $presence = if ($hasFileOnDisk -eq $true -and $isEnforced -eq $true) {
                'Active + on-disk'
            } elseif ($hasFileOnDisk -eq $true -and $isEnforced -eq $false) {
                'On-disk only / pending reboot'
            } elseif ($hasFileOnDisk -eq $false -and $isEnforced -eq $true) {
                'Active only'
            } elseif ($hasFileOnDisk -eq $false -and $isEnforced -eq $false) {
                'Present but inactive'
            } elseif ($hasFileOnDisk -eq $true) {
                'On-disk state only'
            } elseif ($isEnforced -eq $true) {
                'Active state only'
            } else {
                'Unknown state'
            }

            [PSCustomObject]@{
                PolicyId          = $policyInfo.PolicyId
                FriendlyName      = $policyInfo.FriendlyName
                Classification    = $policyInfo.Classification
                IsPlatformManaged = $policyInfo.IsPlatformManaged
                HasFileOnDisk     = $hasFileOnDisk
                IsEnforced        = $isEnforced
                Presence          = $presence
            }
        }
    )
}

function Test-WdacPolicyRemovalCandidate {
    param([Parameter(Mandatory)] [object]$Policy)

    $hasFileOnDisk = $null
    if ($Policy.PSObject.Properties['HasFileOnDisk']) {
        $hasFileOnDisk = ConvertTo-NullableBoolean -Value $Policy.HasFileOnDisk
    }

    $isEnforced = $null
    foreach ($name in @('IsCurrentlyEnforced', 'IsEnforced')) {
        if ($Policy.PSObject.Properties[$name]) {
            $isEnforced = ConvertTo-NullableBoolean -Value $Policy.$name
            if ($null -ne $isEnforced) {
                break
            }
        }
    }

    return (($hasFileOnDisk -eq $true) -or ($isEnforced -eq $true))
}

function Test-WdacPolicyPresent {
    param(
        [Parameter(Mandatory)] [object]$State,
        [AllowNull()] [string]$PolicyId
    )

    $normalizedTarget = Get-WdacNormalizedPolicyId -Value $PolicyId
    if ([string]::IsNullOrWhiteSpace($normalizedTarget)) {
        return $false
    }

    foreach ($policy in @($State.Policies)) {
        $normalizedPolicyId = Get-WdacNormalizedPolicyId -Value (Get-WdacPolicyIdentifierFromObject -Policy $policy)
        if (-not [string]::IsNullOrWhiteSpace($normalizedPolicyId) -and $normalizedPolicyId -eq $normalizedTarget) {
            return $true
        }
    }

    return $false
}

function Test-WdacSnapshotFileMatchesLiveFile {
    param(
        [Parameter(Mandatory)] [object]$File,
        [string]$SnapshotPath
    )

    $destination = Join-Path (Get-WdacCodeIntegrityRoot) ([string]$File.RelativePath)
    if (-not (Test-Path -LiteralPath $destination)) {
        return $false
    }

    $expectedHash = if ($File.PSObject.Properties['Sha256']) { [string]$File.Sha256 } else { $null }
    if ([string]::IsNullOrWhiteSpace($expectedHash)) {
        $expectedHash = (Get-FileHash -InputStream ([System.IO.MemoryStream]::new((Get-WdacSnapshotFileBytes -File $File -SnapshotPath $SnapshotPath))) -Algorithm SHA256).Hash
    }

    $currentHash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash
    return [string]::Equals($expectedHash, $currentHash, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-WdacPolicyState {
    $ciTool = Get-CiToolCommand
    $ciToolAvailable = $null -ne $ciTool
    $policies = @()

    if ($ciToolAvailable) {
        $json = (& $ciTool.Source -lp -json 2>$null | Out-String).Trim()
        if (-not [string]::IsNullOrWhiteSpace($json)) {
            try {
                $parsed = $json | ConvertFrom-Json
                $policyItems = if ($parsed.PSObject.Properties['Policies']) { @($parsed.Policies) } else { @($parsed) }
                $policies = @(
                    foreach ($policy in @($policyItems)) {
                        $ordered = [ordered]@{}
                        foreach ($property in @($policy.PSObject.Properties | Sort-Object -Property Name)) {
                            $ordered[$property.Name] = $property.Value
                        }
                        [PSCustomObject]$ordered
                    }
                )
            } catch {
                $policies = @()
            }
        }
    }

    [PSCustomObject]@{
        CiToolAvailable = $ciToolAvailable
        Policies        = @($policies)
        Files           = @(Get-WdacPolicyFileBackups)
    }
}

function Remove-WdacPolicyFiles {
    $root = Get-WdacCodeIntegrityRoot
    $activeDir = Join-Path $root 'CiPolicies\Active'
    if (Test-Path -LiteralPath $activeDir) {
        Get-ChildItem -LiteralPath $activeDir -Filter '*.cip' -File -ErrorAction SilentlyContinue |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }

    $singlePolicyPath = Join-Path $root 'SiPolicy.p7b'
    if (Test-Path -LiteralPath $singlePolicyPath) {
        Remove-Item -LiteralPath $singlePolicyPath -Force -ErrorAction SilentlyContinue
    }
}

function Remove-WdacPolicies {
    param([Parameter(Mandatory)] [object]$State)

    $removedWithCiTool = $false
    $ciTool = if ($State.CiToolAvailable) { Get-CiToolCommand } else { $null }
    if ($null -ne $ciTool) {
        $policyIds = @(
            foreach ($policy in @($State.Policies)) {
                if (Test-WdacPolicyRemovalCandidate -Policy $policy) {
                    Get-WdacPolicyIdentifierFromObject -Policy $policy
                }
            }
        ) | Sort-Object -Unique

        foreach ($policyId in @($policyIds)) {
            & $ciTool.Source -rp $policyId -json | Out-Null
            if ($LASTEXITCODE -eq 0) {
                $removedWithCiTool = $true
            }
        }

        if ($removedWithCiTool) {
            & $ciTool.Source -r | Out-Null
        }
    }

    if (-not $removedWithCiTool) {
        Remove-WdacPolicyFiles
    }
}

function Get-WdacSnapshotFileBytes {
    param(
        [Parameter(Mandatory)] [object]$File,
        [string]$SnapshotPath
    )

    if ($File.PSObject.Properties['Base64'] -and -not [string]::IsNullOrWhiteSpace([string]$File.Base64)) {
        return [Convert]::FromBase64String([string]$File.Base64)
    }

    if (
        $File.PSObject.Properties['SnapshotAssetRelativePath'] -and
        -not [string]::IsNullOrWhiteSpace([string]$File.SnapshotAssetRelativePath) -and
        -not [string]::IsNullOrWhiteSpace($SnapshotPath)
    ) {
        $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $SnapshotPath
        $assetPath = Join-Path $assetRoot ([string]$File.SnapshotAssetRelativePath)
        if (-not (Test-Path -LiteralPath $assetPath)) {
            throw "WDAC snapshot asset is missing: $assetPath"
        }

        return [System.IO.File]::ReadAllBytes($assetPath)
    }

    throw "WDAC snapshot content is missing for $([string]$File.RelativePath)"
}

function Write-WdacPolicyFiles {
    param(
        [Parameter(Mandatory)] [object[]]$Files,
        [string]$SnapshotPath
    )

    $root = Get-WdacCodeIntegrityRoot
    foreach ($file in @($Files)) {
        $destination = Join-Path $root ([string]$file.RelativePath)
        Write-BytesAtomic -Path $destination -Content (Get-WdacSnapshotFileBytes -File $file -SnapshotPath $SnapshotPath)
    }
}

function Restore-WdacPolicyFiles {
    param(
        [Parameter(Mandatory)] [object[]]$Files,
        [string]$SnapshotPath
    )

    Remove-WdacPolicyFiles
    Write-WdacPolicyFiles -Files $Files -SnapshotPath $SnapshotPath
}

function Get-WdacCiToolStagingFileName {
    param([Parameter(Mandatory)] [object]$File)

    $candidates = @()
    if ($File.PSObject.Properties['FileName']) {
        $candidates += [string]$File.FileName
    }

    if ($File.PSObject.Properties['RelativePath']) {
        $candidates += Split-Path -Leaf ([string]$File.RelativePath)
    }

    foreach ($candidate in $candidates) {
        $safeName = [System.IO.Path]::GetFileName([string]$candidate)
        if (-not [string]::IsNullOrWhiteSpace($safeName)) {
            return $safeName
        }
    }

    return ('{{{0}}}.cip' -f ([guid]::NewGuid().ToString().ToUpperInvariant()))
}

function Restore-WdacPolicies {
    param(
        [Parameter(Mandatory)] [object]$State,
        [string]$SnapshotPath
    )

    $liveState = Get-WdacPolicyState
    Remove-WdacPolicies -State $liveState

    $files = @($State.Files)
    if ($files.Count -eq 0) {
        return
    }

    $ciPolicyFiles = @($files | Where-Object { ([string]$_.FileName).ToLowerInvariant().EndsWith('.cip') })
    $singlePolicyFiles = @($files | Where-Object { ([string]$_.FileName) -eq 'SiPolicy.p7b' })

    $ciTool = if ($State.CiToolAvailable) { Get-CiToolCommand } else { $null }
    if ($null -ne $ciTool) {
        $stagingRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("WinDefState-Wdac-{0}" -f ([guid]::NewGuid().ToString('N')))
        Ensure-Directory -Path $stagingRoot
        $usedDirectFileFallback = $false
        try {
            foreach ($file in $ciPolicyFiles) {
                $filePolicyId = Get-WdacPolicyIdFromFileName -FileName ([string]$file.FileName)
                if (Test-WdacSnapshotFileMatchesLiveFile -File $file -SnapshotPath $SnapshotPath) {
                    Write-Verbose ("WDAC policy file already matches snapshot on disk: {0}" -f ([string]$file.RelativePath))
                    continue
                }

                $tempPath = Join-Path $stagingRoot (Get-WdacCiToolStagingFileName -File $file)
                try {
                    Write-Verbose ("Restoring WDAC policy via CiTool from staged file {0}" -f $tempPath)
                    [System.IO.File]::WriteAllBytes($tempPath, (Get-WdacSnapshotFileBytes -File $file -SnapshotPath $SnapshotPath))
                    $ciToolOutput = (& $ciTool.Source -up $tempPath -json 2>&1 | Out-String).Trim()
                    if ($LASTEXITCODE -ne 0) {
                        $ciToolExitCode = $LASTEXITCODE
                        $message = "CiTool failed to restore WDAC policy from $tempPath (exit code $LASTEXITCODE)"
                        if (-not [string]::IsNullOrWhiteSpace($ciToolOutput)) {
                            $message = "$message. Output: $ciToolOutput"
                        }

                        $postFailureState = Get-WdacPolicyState
                        if (Test-WdacSnapshotFileMatchesLiveFile -File $file -SnapshotPath $SnapshotPath) {
                            Write-Warning ("CiTool reported failure while restoring WDAC policy {0}, but the on-disk policy file already matches the snapshot. Continuing." -f ([string]$file.FileName))
                            continue
                        }

                        if ((-not [string]::IsNullOrWhiteSpace($filePolicyId)) -and (Test-WdacPolicyPresent -State $postFailureState -PolicyId $filePolicyId)) {
                            Write-Warning ("CiTool reported failure while restoring WDAC policy {0}, but the policy is already present in the live CiTool state. Continuing." -f $filePolicyId)
                            continue
                        }

                        if ($ciToolExitCode -eq -2147024891) {
                            Write-Warning ("CiTool access denied while restoring WDAC policy {0}. Falling back to direct file restore; a reboot may be required before live WDAC state fully matches the snapshot." -f ([string]$file.FileName))
                            Write-WdacPolicyFiles -Files @($file) -SnapshotPath $SnapshotPath
                            if (Test-WdacSnapshotFileMatchesLiveFile -File $file -SnapshotPath $SnapshotPath) {
                                $usedDirectFileFallback = $true
                                continue
                            }
                        }

                        throw $message
                    }
                } finally {
                    if (Test-Path -LiteralPath $tempPath) {
                        Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
                    }
                }
            }

            if ($ciPolicyFiles.Count -gt 0 -and -not $usedDirectFileFallback) {
                & $ciTool.Source -r | Out-Null
            }

            if ($singlePolicyFiles.Count -gt 0) {
                Write-WdacPolicyFiles -Files @($singlePolicyFiles) -SnapshotPath $SnapshotPath
            }

            if ($usedDirectFileFallback) {
                Write-Warning 'One or more WDAC policy files were restored directly to disk after CiTool access was denied. A reboot may be required before verification fully matches the snapshot.'
            }

            return
        } finally {
            if (Test-Path -LiteralPath $stagingRoot) {
                Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }

    Restore-WdacPolicyFiles -Files $files -SnapshotPath $SnapshotPath
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

function ConvertTo-FirewallProfileActionValue {
    param([AllowNull()] [object]$Value)

    switch ([string]$Value) {
        'Allow' { 'Allow' }
        'Block' { 'Block' }
        'NotConfigured' { 'NotConfigured' }
        default { [string]$Value }
    }
}

function ConvertTo-FirewallProfileTriStateValue {
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

function New-FirewallProfileCaptureIssue {
    param(
        [string]$Profile,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        Profile = $Profile
        Message = $Message
    }
}

function Normalize-FirewallProfileEntry {
    param([AllowNull()] [object]$Profile)

    if ($null -eq $Profile) {
        return $null
    }

    [PSCustomObject]@{
        Profile                         = if ($Profile.PSObject.Properties['Profile']) { [string]$Profile.Profile } elseif ($Profile.PSObject.Properties['Name']) { [string]$Profile.Name } else { $null }
        Enabled                         = if ($Profile.PSObject.Properties['Enabled']) { ConvertTo-FirewallProfileEnabledValue -Value $Profile.Enabled } else { $null }
        DefaultInboundAction            = if ($Profile.PSObject.Properties['DefaultInboundAction']) { ConvertTo-FirewallProfileActionValue -Value $Profile.DefaultInboundAction } else { $null }
        DefaultOutboundAction           = if ($Profile.PSObject.Properties['DefaultOutboundAction']) { ConvertTo-FirewallProfileActionValue -Value $Profile.DefaultOutboundAction } else { $null }
        AllowUnicastResponseToMulticast = if ($Profile.PSObject.Properties['AllowUnicastResponseToMulticast']) { ConvertTo-FirewallProfileTriStateValue -Value $Profile.AllowUnicastResponseToMulticast } else { $null }
        NotifyOnListen                  = if ($Profile.PSObject.Properties['NotifyOnListen']) { ConvertTo-FirewallProfileTriStateValue -Value $Profile.NotifyOnListen } else { $null }
        LogAllowed                      = if ($Profile.PSObject.Properties['LogAllowed']) { ConvertTo-FirewallProfileTriStateValue -Value $Profile.LogAllowed } else { $null }
        LogBlocked                      = if ($Profile.PSObject.Properties['LogBlocked']) { ConvertTo-FirewallProfileTriStateValue -Value $Profile.LogBlocked } else { $null }
        LogIgnored                      = if ($Profile.PSObject.Properties['LogIgnored']) { ConvertTo-FirewallProfileTriStateValue -Value $Profile.LogIgnored } else { $null }
        LogMaxSizeKilobytes             = if ($Profile.PSObject.Properties['LogMaxSizeKilobytes'] -and $null -ne $Profile.LogMaxSizeKilobytes) { [uint64]$Profile.LogMaxSizeKilobytes } else { $null }
        LogFileName                     = if ($Profile.PSObject.Properties['LogFileName'] -and -not [string]::IsNullOrWhiteSpace([string]$Profile.LogFileName)) { [string]$Profile.LogFileName } else { $null }
    }
}

function Normalize-FirewallProfileState {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            CaptureIssues    = @()
            Profiles         = @()
        }
    }

    $profiles = if ($State.PSObject.Properties['Profiles']) { @($State.Profiles) } else { @($State) }
    $captureIssues = @()
    if ($State.PSObject.Properties['CaptureIssues']) {
        $captureIssues = @(
            foreach ($issue in @($State.CaptureIssues)) {
                if ($null -eq $issue) {
                    continue
                }

                [PSCustomObject]@{
                    Profile = if ($issue.PSObject.Properties['Profile']) { [string]$issue.Profile } else { $null }
                    Message = if ($issue.PSObject.Properties['Message']) { [string]$issue.Message } else { $null }
                }
            }
        )
    }

    [PSCustomObject]@{
        CommandAvailable = if ($State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
        CaptureIssues    = @($captureIssues)
        Profiles         = @(
            foreach ($profile in @($profiles)) {
                Normalize-FirewallProfileEntry -Profile $profile
            }
        )
    }
}

function Test-FirewallProfileStateHasExtendedFields {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return $false
    }

    if ($State.PSObject.Properties['CommandAvailable'] -or $State.PSObject.Properties['CaptureIssues']) {
        return $true
    }

    $profiles = if ($State.PSObject.Properties['Profiles']) { @($State.Profiles) } else { @($State) }
    foreach ($profile in @($profiles)) {
        if ($null -eq $profile) {
            continue
        }

        foreach ($propertyName in @('DefaultInboundAction', 'DefaultOutboundAction', 'AllowUnicastResponseToMulticast', 'NotifyOnListen', 'LogAllowed', 'LogBlocked', 'LogIgnored', 'LogMaxSizeKilobytes', 'LogFileName')) {
            if ($profile.PSObject.Properties[$propertyName]) {
                return $true
            }
        }
    }

    $false
}

function Test-FirewallProfileStateCapturedExactly {
    param([AllowNull()] [object]$State)

    $normalized = Normalize-FirewallProfileState -State $State
    $normalized.CommandAvailable -and (@($normalized.CaptureIssues).Count -eq 0) -and (@($normalized.Profiles).Count -gt 0)
}

function Get-FirewallProfileStates {
    param([string[]]$Profiles)

    if (-not (Test-CommandAvailable -Name 'Get-NetFirewallProfile')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            CaptureIssues    = @()
            Profiles         = @()
        }
    }

    $captureIssues = [System.Collections.Generic.List[object]]::new()
    $states = foreach ($profileName in @($Profiles)) {
        $profile = Get-NetFirewallProfile -Profile $profileName -ErrorAction SilentlyContinue
        if ($null -eq $profile) {
            $captureIssues.Add((New-FirewallProfileCaptureIssue -Profile $profileName -Message 'Get-NetFirewallProfile did not return this profile.')) | Out-Null
            continue
        }

        Normalize-FirewallProfileEntry -Profile ([PSCustomObject]@{
            Profile                         = $profileName
            Enabled                         = $profile.Enabled
            DefaultInboundAction            = $profile.DefaultInboundAction
            DefaultOutboundAction           = $profile.DefaultOutboundAction
            AllowUnicastResponseToMulticast = $profile.AllowUnicastResponseToMulticast
            NotifyOnListen                  = $profile.NotifyOnListen
            LogAllowed                      = $profile.LogAllowed
            LogBlocked                      = $profile.LogBlocked
            LogIgnored                      = $profile.LogIgnored
            LogMaxSizeKilobytes             = $profile.LogMaxSizeKilobytes
            LogFileName                     = $profile.LogFileName
        })
    }

    [PSCustomObject]@{
        CommandAvailable = $true
        CaptureIssues    = @($captureIssues)
        Profiles         = @($states)
    }
}

function Set-FirewallProfileStateExact {
    param([Parameter(Mandatory)] [object]$ProfileState)

    if (-not (Test-CommandAvailable -Name 'Set-NetFirewallProfile')) {
        return
    }

    $profile = Normalize-FirewallProfileEntry -Profile $ProfileState
    $profileName = [string]$profile.Profile
    if ([string]::IsNullOrWhiteSpace($profileName)) {
        return
    }

    $params = @{
        Profile = $profileName
    }

    if ($null -ne $profile.Enabled) {
        $params['Enabled'] = $profile.Enabled
    }

    foreach ($propertyName in @('DefaultInboundAction', 'DefaultOutboundAction', 'AllowUnicastResponseToMulticast', 'NotifyOnListen', 'LogAllowed', 'LogBlocked', 'LogIgnored')) {
        $value = $profile.$propertyName
        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
            $params[$propertyName] = [string]$value
        }
    }

    if ($null -ne $profile.LogMaxSizeKilobytes) {
        $params['LogMaxSizeKilobytes'] = [uint64]$profile.LogMaxSizeKilobytes
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$profile.LogFileName)) {
        $params['LogFileName'] = [string]$profile.LogFileName
    }

    Set-NetFirewallProfile @params
}

function Set-Permissive-FirewallProfiles {
    param([Parameter(Mandatory)] [object]$Definition)

    if (-not (Test-CommandAvailable -Name 'Set-NetFirewallProfile')) {
        return
    }

    foreach ($profileName in @($Definition.Profiles)) {
        Set-NetFirewallProfile -Profile $profileName `
            -Enabled (ConvertTo-FirewallProfileEnabledValue -Value $Definition.PermissiveValue) `
            -DefaultInboundAction $Definition.PermissiveDefaultInboundAction `
            -DefaultOutboundAction $Definition.PermissiveDefaultOutboundAction `
            -AllowUnicastResponseToMulticast (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveAllowUnicastResponseToMulticast) `
            -NotifyOnListen (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveNotifyOnListen) `
            -LogAllowed (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveLogAllowed) `
            -LogBlocked (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveLogBlocked) `
            -LogIgnored (ConvertTo-FirewallProfileTriStateValue -Value $Definition.PermissiveLogIgnored)
    }
}

function Restore-FirewallProfiles {
    param([AllowNull()] [object]$State)

    if (-not (Test-FirewallProfileStateCapturedExactly -State $State)) {
        return
    }

    foreach ($profile in @((Normalize-FirewallProfileState -State $State).Profiles)) {
        Set-FirewallProfileStateExact -ProfileState $profile
    }
}

function ConvertTo-ComparableFirewallProfileState {
    param(
        [AllowNull()] [object]$State,
        [AllowNull()] [object]$ReferenceState
    )

    $normalizedState = Normalize-FirewallProfileState -State $State
    $comparisonReference = if ($null -ne $ReferenceState) { $ReferenceState } else { $State }
    $useExtendedFields = Test-FirewallProfileStateHasExtendedFields -State $comparisonReference

    if (-not $useExtendedFields) {
        return @(
            foreach ($profile in @($normalizedState.Profiles | Sort-Object -Property Profile)) {
                [PSCustomObject]@{
                    Profile = [string]$profile.Profile
                    Enabled = if ($null -ne $profile.Enabled) { [string]$profile.Enabled } else { $null }
                }
            }
        )
    }

    [ordered]@{
        CommandAvailable = $normalizedState.CommandAvailable
        CaptureIssues    = @(
            foreach ($issue in @($normalizedState.CaptureIssues | Sort-Object -Property Profile, Message)) {
                [PSCustomObject]@{
                    Profile = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Profile)) { [string]$issue.Profile } else { $null }
                    Message = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Message)) { [string]$issue.Message } else { $null }
                }
            }
        )
        Profiles         = @(
            foreach ($profile in @($normalizedState.Profiles | Sort-Object -Property Profile)) {
                [PSCustomObject]@{
                    Profile                         = [string]$profile.Profile
                    Enabled                         = if ($null -ne $profile.Enabled) { [string]$profile.Enabled } else { $null }
                    DefaultInboundAction            = if ($null -ne $profile.DefaultInboundAction) { [string]$profile.DefaultInboundAction } else { $null }
                    DefaultOutboundAction           = if ($null -ne $profile.DefaultOutboundAction) { [string]$profile.DefaultOutboundAction } else { $null }
                    AllowUnicastResponseToMulticast = if ($null -ne $profile.AllowUnicastResponseToMulticast) { [string]$profile.AllowUnicastResponseToMulticast } else { $null }
                    NotifyOnListen                  = if ($null -ne $profile.NotifyOnListen) { [string]$profile.NotifyOnListen } else { $null }
                    LogAllowed                      = if ($null -ne $profile.LogAllowed) { [string]$profile.LogAllowed } else { $null }
                    LogBlocked                      = if ($null -ne $profile.LogBlocked) { [string]$profile.LogBlocked } else { $null }
                    LogIgnored                      = if ($null -ne $profile.LogIgnored) { [string]$profile.LogIgnored } else { $null }
                    LogMaxSizeKilobytes             = $profile.LogMaxSizeKilobytes
                    LogFileName                     = if ($null -ne $profile.LogFileName) { [string]$profile.LogFileName } else { $null }
                }
            }
        )
    }
}

function New-FirewallRuleCaptureIssue {
    param(
        [string]$Group,
        [Parameter(Mandatory)] [string]$Message
    )

    [PSCustomObject]@{
        Group   = $Group
        Message = $Message
    }
}

function Normalize-FirewallRuleEntry {
    param([AllowNull()] [object]$Rule)

    if ($null -eq $Rule) {
        return $null
    }

    [PSCustomObject]@{
        Name        = if ($Rule.PSObject.Properties['Name']) { [string]$Rule.Name } else { $null }
        DisplayName = if ($Rule.PSObject.Properties['DisplayName']) { [string]$Rule.DisplayName } else { $null }
        Group       = if ($Rule.PSObject.Properties['Group']) { [string]$Rule.Group } elseif ($Rule.PSObject.Properties['RuleGroup']) { [string]$Rule.RuleGroup } else { $null }
        Enabled     = if ($Rule.PSObject.Properties['Enabled']) { ConvertTo-FirewallProfileEnabledValue -Value $Rule.Enabled } else { $null }
        Direction   = if ($Rule.PSObject.Properties['Direction']) { [string]$Rule.Direction } else { $null }
        Action      = if ($Rule.PSObject.Properties['Action']) { [string]$Rule.Action } else { $null }
        Profile     = if ($Rule.PSObject.Properties['Profile']) { [string]$Rule.Profile } else { $null }
    }
}

function Normalize-FirewallRuleState {
    param([AllowNull()] [object]$State)

    if ($null -eq $State) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Group            = $null
            CaptureIssues    = @()
            Rules            = @()
        }
    }

    $rules = if ($State.PSObject.Properties['Rules']) { @($State.Rules) } else { @($State) }
    $captureIssues = @()
    if ($State.PSObject.Properties['CaptureIssues']) {
        $captureIssues = @(
            foreach ($issue in @($State.CaptureIssues)) {
                if ($null -eq $issue) {
                    continue
                }

                [PSCustomObject]@{
                    Group   = if ($issue.PSObject.Properties['Group']) { [string]$issue.Group } else { $null }
                    Message = if ($issue.PSObject.Properties['Message']) { [string]$issue.Message } else { $null }
                }
            }
        )
    }

    [PSCustomObject]@{
        CommandAvailable = if ($State.PSObject.Properties['CommandAvailable']) { [bool]$State.CommandAvailable } else { $true }
        Group            = if ($State.PSObject.Properties['Group']) { [string]$State.Group } else { $null }
        CaptureIssues    = @($captureIssues)
        Rules            = @(
            foreach ($rule in @($rules)) {
                Normalize-FirewallRuleEntry -Rule $rule
            }
        )
    }
}

function Test-FirewallRuleStateCapturedExactly {
    param([AllowNull()] [object]$State)

    $normalized = Normalize-FirewallRuleState -State $State
    $normalized.CommandAvailable -and (@($normalized.CaptureIssues).Count -eq 0)
}

function Get-FirewallRuleGroupState {
    param([Parameter(Mandatory)] [string]$Group)

    if (-not (Test-CommandAvailable -Name 'Get-NetFirewallRule')) {
        return [PSCustomObject]@{
            CommandAvailable = $false
            Group            = $Group
            CaptureIssues    = @()
            Rules            = @()
        }
    }

    try {
        $rules = @(
            foreach ($rule in @(Get-NetFirewallRule -Group $Group -ErrorAction Stop)) {
                Normalize-FirewallRuleEntry -Rule ([PSCustomObject]@{
                    Name        = $rule.Name
                    DisplayName = $rule.DisplayName
                    Group       = $rule.Group
                    Enabled     = $rule.Enabled
                    Direction   = $rule.Direction
                    Action      = $rule.Action
                    Profile     = $rule.Profile
                })
            }
        )

        return [PSCustomObject]@{
            CommandAvailable = $true
            Group            = $Group
            CaptureIssues    = @()
            Rules            = @($rules)
        }
    } catch {
        return [PSCustomObject]@{
            CommandAvailable = $true
            Group            = $Group
            CaptureIssues    = @(
                New-FirewallRuleCaptureIssue -Group $Group -Message $_.Exception.Message
            )
            Rules            = @()
        }
    }
}

function Set-Permissive-FirewallRules {
    param([Parameter(Mandatory)] [object]$Definition)

    if (-not (Test-CommandAvailable -Name 'Set-NetFirewallRule')) {
        return
    }

    Set-NetFirewallRule -Group $Definition.Group -Enabled (ConvertTo-FirewallProfileEnabledValue -Value $Definition.PermissiveEnabled) -ErrorAction SilentlyContinue | Out-Null
}

function Restore-FirewallRules {
    param([AllowNull()] [object]$State)

    if (-not (Test-FirewallRuleStateCapturedExactly -State $State)) {
        return
    }

    if (-not (Test-CommandAvailable -Name 'Set-NetFirewallRule')) {
        return
    }

    foreach ($rule in @((Normalize-FirewallRuleState -State $State).Rules)) {
        if ([string]::IsNullOrWhiteSpace([string]$rule.Name) -or $null -eq $rule.Enabled) {
            continue
        }

        Set-NetFirewallRule -Name $rule.Name -Enabled (ConvertTo-FirewallProfileEnabledValue -Value $rule.Enabled) -ErrorAction SilentlyContinue | Out-Null
    }
}

function ConvertTo-ComparableFirewallRuleState {
    param([AllowNull()] [object]$State)

    $normalizedState = Normalize-FirewallRuleState -State $State
    [ordered]@{
        CommandAvailable = $normalizedState.CommandAvailable
        Group            = if (-not [string]::IsNullOrWhiteSpace([string]$normalizedState.Group)) { [string]$normalizedState.Group } else { $null }
        CaptureIssues    = @(
            foreach ($issue in @($normalizedState.CaptureIssues | Sort-Object -Property Group, Message)) {
                [PSCustomObject]@{
                    Group   = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Group)) { [string]$issue.Group } else { $null }
                    Message = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Message)) { [string]$issue.Message } else { $null }
                }
            }
        )
        Rules            = @(
            foreach ($rule in @($normalizedState.Rules | Sort-Object -Property Name, DisplayName, Direction, Action, Profile)) {
                [PSCustomObject]@{
                    Name        = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Name)) { [string]$rule.Name } else { $null }
                    DisplayName = if (-not [string]::IsNullOrWhiteSpace([string]$rule.DisplayName)) { [string]$rule.DisplayName } else { $null }
                    Group       = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Group)) { [string]$rule.Group } else { $null }
                    Enabled     = if ($null -ne $rule.Enabled) { [string]$rule.Enabled } else { $null }
                    Direction   = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Direction)) { [string]$rule.Direction } else { $null }
                    Action      = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Action)) { [string]$rule.Action } else { $null }
                    Profile     = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Profile)) { [string]$rule.Profile } else { $null }
                }
            }
        )
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
        [Parameter(Mandatory)] [object]$Entry,
        [string]$SnapshotPath
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
        'MpPreferenceList' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            $items = @(Normalize-MpPreferenceListItems -Value $Entry.CurrentValue)

            Add-ReportKeyValueLine -Lines $Lines -Label 'Property' -Value $Entry.Property
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Item count' -Value $items.Count
            if ($Entry.PSObject.Properties['CaptureError'] -and -not [string]::IsNullOrWhiteSpace([string]$Entry.CaptureError)) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
            Add-ReportJsonBlock -Lines $Lines -Label 'Items' -Value @($items)
        }
        'DefenderRuntimeStatus' {
            $state = if ($null -ne $Entry.CurrentValue) { $Entry.CurrentValue } else { [PSCustomObject]@{} }
            $commandAvailable = if ($state.PSObject.Properties['CommandAvailable']) { [bool]$state.CommandAvailable } else { $true }
            $captured = if ($state.PSObject.Properties['Captured']) { [bool]$state.Captured } else { $true }
            $amRunningMode = if ($state.PSObject.Properties['AMRunningMode']) { [string]$state.AMRunningMode } else { $null }
            $realTimeProtectionEnabled = if ($state.PSObject.Properties['RealTimeProtectionEnabled']) { ConvertTo-NullableBoolean -Value $state.RealTimeProtectionEnabled } else { $null }
            $antivirusEnabled = if ($state.PSObject.Properties['AntivirusEnabled']) { ConvertTo-NullableBoolean -Value $state.AntivirusEnabled } else { $null }
            $isTamperProtected = if ($state.PSObject.Properties['IsTamperProtected']) { ConvertTo-NullableBoolean -Value $state.IsTamperProtected } else { $null }

            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Running mode' -Value $amRunningMode
            Add-ReportKeyValueLine -Lines $Lines -Label 'Real-time protection enabled' -Value $realTimeProtectionEnabled
            Add-ReportKeyValueLine -Lines $Lines -Label 'Antivirus enabled' -Value $antivirusEnabled
            Add-ReportKeyValueLine -Lines $Lines -Label 'Tamper protection' -Value $isTamperProtected
            if ($state.PSObject.Properties['Error'] -and -not [string]::IsNullOrWhiteSpace([string]$state.Error)) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $state.Error
            }
            if ($captured -and ($isTamperProtected -eq $true -or $amRunningMode -match 'Passive')) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'Tamper protection can cause local Defender preference changes to be ignored or later reverted, and passive mode means Defender Antivirus is not the primary enforcement engine. Interpret Defender restore mismatches in that context.'
            }
        }
        'BitLockerVolumes' {
            $state = Normalize-BitLockerState -State $Entry.CurrentValue
            $timedOutMountPoints = @($state.TimedOutMountPoints)
            $captureIssues = @($state.CaptureIssues)
            $capturedVolumes = @($state.Volumes)
            $baselineComplete = Test-BitLockerStateCapturedExactly -State $state

            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Timed out mount point count' -Value $timedOutMountPoints.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value $captureIssues.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($baselineComplete) { 'Complete' } else { 'Partial / incomplete' })
            if (-not $baselineComplete) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'BitLocker round-trip requires an exact mounted-volume baseline. Permissive, restore, and verification skip this setting if mount points time out or richer BitLocker fields cannot be captured exactly.'
            }
            if ($timedOutMountPoints.Count -gt 0) {
                foreach ($mountPoint in $timedOutMountPoints) {
                    $Lines.Add(('  - Timed out mount point: {0}' -f $mountPoint))
                }
            }
            foreach ($issue in @($captureIssues | Sort-Object -Property MountPoint, Message)) {
                $issueMountPoint = if (-not [string]::IsNullOrWhiteSpace([string]$issue.MountPoint)) { [string]$issue.MountPoint } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | MountPoint={0} | {1}' -f $issueMountPoint, $issue.Message))
            }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Captured volume count' -Value $capturedVolumes.Count
            foreach ($volume in @($capturedVolumes | Sort-Object -Property MountPoint)) {
                $mountPoint = if (-not [string]::IsNullOrWhiteSpace([string]$volume.MountPoint)) { [string]$volume.MountPoint } else { '<unknown>' }
                $protectionMode = if (-not [string]::IsNullOrWhiteSpace([string]$volume.ProtectionMode)) { [string]$volume.ProtectionMode } else { '<unknown>' }
                $volumeStatus = if (-not [string]::IsNullOrWhiteSpace([string]$volume.VolumeStatus)) { [string]$volume.VolumeStatus } else { '<unknown>' }
                $lockStatus = if (-not [string]::IsNullOrWhiteSpace([string]$volume.LockStatus)) { [string]$volume.LockStatus } else { '<unknown>' }
                $encryptionMethod = if (-not [string]::IsNullOrWhiteSpace([string]$volume.EncryptionMethod)) { [string]$volume.EncryptionMethod } else { '<unknown>' }
                $encryptionSummary = if ($null -ne $volume.EncryptionPercentage) {
                    ('{0}% ({1})' -f ([int]$volume.EncryptionPercentage), $encryptionMethod)
                } else {
                    $encryptionMethod
                }
                $protectorCount = if ($null -ne $volume.KeyProtectorCount) { [int]$volume.KeyProtectorCount } else { '<unknown>' }
                $protectorSummary = Get-BitLockerProtectorTypeSummary -Volume $volume
                $autoUnlockSummary = if (Test-BitLockerAutoUnlockSupportedVolume -Volume $volume) {
                    if ($null -ne $volume.AutoUnlockEnabled) {
                        ConvertTo-DisplayString -Value $volume.AutoUnlockEnabled
                    } else {
                        '<unknown>'
                    }
                } else {
                    'n/a'
                }

                $Lines.Add(('  - {0} | Mode={1} | Protection={2} | Status={3} | Lock={4} | Encryption={5} | Protectors={6} [{7}] | AutoUnlock={8}' -f $mountPoint, $protectionMode, $volume.ProtectionStatus, $volumeStatus, $lockStatus, $encryptionSummary, $protectorCount, $protectorSummary, $autoUnlockSummary))
            }
        }
        'ExploitProtectionPolicy' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $Entry.CurrentValue.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'XML captured' -Value (-not [string]::IsNullOrWhiteSpace((Get-ExploitProtectionPolicyXml -State $Entry.CurrentValue -SnapshotPath $SnapshotPath)))
        }
        'WdacPolicies' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'CiTool available' -Value $Entry.CurrentValue.CiToolAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Policy count' -Value @($Entry.CurrentValue.Policies).Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Policy file count' -Value @($Entry.CurrentValue.Files).Count
            $policyRows = @(Get-WdacPolicyReportRows -State $Entry.CurrentValue)
            $platformManagedPolicies = @($policyRows | Where-Object { $_.IsPlatformManaged })
            $activePolicies = @($policyRows | Where-Object { $_.IsEnforced -eq $true })
            $onDiskOnlyPolicies = @($policyRows | Where-Object { $_.HasFileOnDisk -eq $true -and $_.IsEnforced -eq $false })
            $activeOnlyPolicies = @($policyRows | Where-Object { $_.HasFileOnDisk -eq $false -and $_.IsEnforced -eq $true })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Platform-managed policy count' -Value $platformManagedPolicies.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Active policy count' -Value $activePolicies.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'On-disk only / pending reboot count' -Value $onDiskOnlyPolicies.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Active-only policy count' -Value $activeOnlyPolicies.Count
            if ($platformManagedPolicies.Count -gt 0 -or $onDiskOnlyPolicies.Count -gt 0) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'WDAC reporting distinguishes active enforcement from on-disk file presence. On-disk-only policies usually indicate pending-reboot state, and platform-managed policies can require file-based handling or a reboot before live WDAC state fully matches the snapshot.'
            }
            foreach ($policyRow in $policyRows) {
                if ($policyRow.IsPlatformManaged) {
                    $Lines.Add(('  - {0} | {1} | {2} | Enforced={3} | FileOnDisk={4} | {5} [platform-managed]' -f $policyRow.PolicyId, $policyRow.FriendlyName, $policyRow.Presence, (ConvertTo-DisplayString -Value $policyRow.IsEnforced), (ConvertTo-DisplayString -Value $policyRow.HasFileOnDisk), $policyRow.Classification))
                } else {
                    $Lines.Add(('  - {0} | {1} | {2} | Enforced={3} | FileOnDisk={4}' -f $policyRow.PolicyId, $policyRow.FriendlyName, $policyRow.Presence, (ConvertTo-DisplayString -Value $policyRow.IsEnforced), (ConvertTo-DisplayString -Value $policyRow.HasFileOnDisk)))
                }
            }
        }
        'AsrRules' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Configured rule count' -Value @($Entry.CurrentValue).Count
            $invalidEntries = @(Get-AsrInvalidEntriesFromEntry -Entry $Entry)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Invalid capture entry count' -Value $invalidEntries.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($invalidEntries.Count -eq 0) { 'Complete' } else { 'Partial / incomplete' })
            if ($invalidEntries.Count -gt 0) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'This snapshot did not capture a reliable ASR baseline. Permissive, restore, and verification will skip ASR changes to avoid unsafe round-trip behavior.'
            }
            foreach ($rule in @($Entry.CurrentValue)) {
                $Lines.Add(('  - {0} | {1} | {2}' -f $rule.Id, $rule.Name, $rule.ActionLabel))
            }
            foreach ($rule in $invalidEntries) {
                $displayId = if ([string]::IsNullOrWhiteSpace([string]$rule.Id)) { '<blank>' } else { [string]$rule.Id }
                $displayAction = if ([string]::IsNullOrWhiteSpace([string]$rule.Action)) {
                    '<blank>'
                } elseif ([string]::IsNullOrWhiteSpace([string]$rule.ActionLabel)) {
                    [string]$rule.Action
                } else {
                    [string]$rule.ActionLabel
                }
                $Lines.Add(('  - Incomplete capture entry | RuleId={0} | Action={1}' -f $displayId, $displayAction))
            }
        }
        'PowerShellModuleLogging' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Base path' -Value $Entry.BasePath
            Add-ReportKeyValueLine -Lines $Lines -Label 'Exists' -Value $Entry.Exists
            Add-ReportJsonBlock -Lines $Lines -Label 'Base values' -Value @($Entry.CurrentValue.BaseValues)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Module names key exists' -Value $Entry.CurrentValue.ModuleNamesExists
            Add-ReportJsonBlock -Lines $Lines -Label 'Module names values' -Value @($Entry.CurrentValue.ModuleNamesValues)
        }
        'AppLockerPolicy' {
            $state = $Entry.CurrentValue
            $commandAvailable = if ($state.PSObject.Properties['CommandAvailable']) { [bool]$state.CommandAvailable } else { $true }
            $localCaptured = if ($state.PSObject.Properties['LocalCaptured']) { [bool]$state.LocalCaptured } else { $false }
            $effectiveCaptured = if ($state.PSObject.Properties['EffectiveCaptured']) { [bool]$state.EffectiveCaptured } else { $false }
            $localMatchesEffective = if ($state.PSObject.Properties['LocalMatchesEffective']) { [bool]$state.LocalMatchesEffective } else { $false }
            $captureIssues = if ($state.PSObject.Properties['CaptureIssues']) { @($state.CaptureIssues) } else { @() }
            $collectionSummaries = if ($state.PSObject.Properties['CollectionSummaries']) { @($state.CollectionSummaries) } else { @(Get-AppLockerCollectionSummaries -Xml (Get-AppLockerPolicyXml -State $state -PolicyScope Effective -SnapshotPath $SnapshotPath)) }
            $baselineComplete = Test-AppLockerPolicyCapturedExactly -State $state -SnapshotPath $SnapshotPath

            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Local policy captured' -Value $localCaptured
            Add-ReportKeyValueLine -Lines $Lines -Label 'Effective policy captured' -Value $effectiveCaptured
            Add-ReportKeyValueLine -Lines $Lines -Label 'Local matches effective' -Value $localMatchesEffective
            Add-ReportKeyValueLine -Lines $Lines -Label 'Collection count' -Value $collectionSummaries.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value $captureIssues.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($baselineComplete) { 'Complete' } else { 'Partial / incomplete' })
            if (-not $baselineComplete) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'AppLocker round-trip is only treated as exact when the local and effective AppLocker policies match. If Group Policy or another higher-precedence source changes the effective policy, permissive, restore, and verification skip this entry rather than guessing.'
            }

            foreach ($issue in @($captureIssues | Sort-Object -Property Scope, Message)) {
                $scope = if ($issue.PSObject.Properties['Scope'] -and -not [string]::IsNullOrWhiteSpace([string]$issue.Scope)) { [string]$issue.Scope } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | Scope={0} | {1}' -f $scope, $issue.Message))
            }

            foreach ($collection in @($collectionSummaries | Sort-Object -Property Type)) {
                $Lines.Add(('  - {0} | Enforcement={1} | Rules={2} | Services={3} | SystemApps={4}' -f (ConvertTo-DisplayString -Value $collection.Type), (ConvertTo-DisplayString -Value $collection.EnforcementMode), (ConvertTo-DisplayString -Value $collection.RuleCount), (ConvertTo-DisplayString -Value $collection.ServicesEnforcement), (ConvertTo-DisplayString -Value $collection.SystemAppsAllow)))
            }
        }
        'FirewallProfiles' {
            $state = Normalize-FirewallProfileState -State $Entry.CurrentValue
            $baselineComplete = Test-FirewallProfileStateCapturedExactly -State $state
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Captured profile count' -Value (@($state.Profiles).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value (@($state.CaptureIssues).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($baselineComplete) { 'Complete' } else { 'Partial / incomplete' })
            if (-not $baselineComplete) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'Firewall round-trip requires an exact per-profile baseline. Permissive, restore, and verification skip this setting if a profile cannot be captured exactly.'
            }

            foreach ($issue in @($state.CaptureIssues | Sort-Object -Property Profile, Message)) {
                $issueProfile = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Profile)) { [string]$issue.Profile } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | Profile={0} | {1}' -f $issueProfile, $issue.Message))
            }

            foreach ($profile in @($state.Profiles | Sort-Object -Property Profile)) {
                $Lines.Add(('  - {0} | Enabled={1} | Inbound={2} | Outbound={3} | Notify={4} | Unicast={5} | LogAllowed={6} | LogBlocked={7} | LogIgnored={8} | LogMaxKB={9} | LogFile={10}' -f $profile.Profile, (ConvertTo-DisplayString -Value $profile.Enabled), (ConvertTo-DisplayString -Value $profile.DefaultInboundAction), (ConvertTo-DisplayString -Value $profile.DefaultOutboundAction), (ConvertTo-DisplayString -Value $profile.NotifyOnListen), (ConvertTo-DisplayString -Value $profile.AllowUnicastResponseToMulticast), (ConvertTo-DisplayString -Value $profile.LogAllowed), (ConvertTo-DisplayString -Value $profile.LogBlocked), (ConvertTo-DisplayString -Value $profile.LogIgnored), (ConvertTo-DisplayString -Value $profile.LogMaxSizeKilobytes), (ConvertTo-DisplayString -Value $profile.LogFileName)))
            }
        }
        'FirewallRules' {
            $state = Normalize-FirewallRuleState -State $Entry.CurrentValue
            $baselineComplete = Test-FirewallRuleStateCapturedExactly -State $state
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Group' -Value $state.Group
            Add-ReportKeyValueLine -Lines $Lines -Label 'Rule count' -Value (@($state.Rules).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value (@($state.CaptureIssues).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($baselineComplete) { 'Complete' } else { 'Partial / incomplete' })
            if (-not $baselineComplete) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'Firewall rule group round-trip requires an exact baseline. Permissive, restore, and verification skip this setting if the rule group cannot be captured exactly.'
            }

            foreach ($issue in @($state.CaptureIssues | Sort-Object -Property Group, Message)) {
                $issueGroup = if (-not [string]::IsNullOrWhiteSpace([string]$issue.Group)) { [string]$issue.Group } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | Group={0} | {1}' -f $issueGroup, $issue.Message))
            }

            foreach ($rule in @($state.Rules | Sort-Object -Property Name, DisplayName)) {
                $displayName = if (-not [string]::IsNullOrWhiteSpace([string]$rule.DisplayName)) { [string]$rule.DisplayName } else { '<unnamed>' }
                $ruleName = if (-not [string]::IsNullOrWhiteSpace([string]$rule.Name)) { [string]$rule.Name } else { '<unknown>' }
                $Lines.Add(('  - {0} | Name={1} | Enabled={2} | Direction={3} | Action={4} | Profile={5}' -f $displayName, $ruleName, (ConvertTo-DisplayString -Value $rule.Enabled), (ConvertTo-DisplayString -Value $rule.Direction), (ConvertTo-DisplayString -Value $rule.Action), (ConvertTo-DisplayString -Value $rule.Profile)))
            }
        }
        'NetBiosAdapters' {
            foreach ($adapter in @($Entry.CurrentValue)) {
                $Lines.Add(('  - Index {0} | {1} | TcpipNetbiosOptions={2}' -f $adapter.Index, $adapter.Description, $adapter.TcpipNetbiosOptions))
            }
        }
        'LoadedUserRegistryValues' {
            $state = Normalize-UserRegistryValueState -State $Entry.CurrentValue
            Add-ReportKeyValueLine -Lines $Lines -Label 'Captured profile count' -Value (@($state.Entries | Group-Object -Property Sid).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Capture issue count' -Value (@($state.CaptureIssues).Count)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if (@($state.CaptureIssues).Count -eq 0) { 'Complete' } else { 'Partial / incomplete' })
            if (@($state.CaptureIssues).Count -gt 0) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Operator note' -Value 'One or more user profile hives could not be accessed. Permissive, restore, and verification will skip this setting to avoid partial round-trip behavior.'
            }

            foreach ($value in @($state.Entries)) {
                $displayValue = if ($value.Exists) { ConvertTo-DisplayString -Value $value.CurrentValue } else { '<absent>' }
                $Lines.Add(('  - {0} | {1} | {2} | {3}' -f $value.Sid, $value.RelativePath, $value.Name, $displayValue))
            }

            foreach ($issue in @($state.CaptureIssues)) {
                $profilePath = if ($null -ne $issue.PSObject.Properties['ProfilePath'] -and -not [string]::IsNullOrWhiteSpace([string]$issue.ProfilePath)) { [string]$issue.ProfilePath } else { '<unknown>' }
                $hivePath = if ($null -ne $issue.PSObject.Properties['HivePath'] -and -not [string]::IsNullOrWhiteSpace([string]$issue.HivePath)) { [string]$issue.HivePath } else { '<unknown>' }
                $Lines.Add(('  - Capture issue | SID={0} | Profile={1} | Hive={2} | {3}' -f $issue.Sid, $profilePath, $hivePath, $issue.Message))
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
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            Add-ReportKeyValueLine -Lines $Lines -Label 'Current value' -Value $Entry.CurrentValue
            if ($Entry.PSObject.Properties['CaptureError'] -and -not [string]::IsNullOrWhiteSpace([string]$Entry.CaptureError)) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }
        }
        'WinRmListeners' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            $listeners = @($Entry.CurrentValue)
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $commandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Listener count' -Value $listeners.Count
            Add-ReportKeyValueLine -Lines $Lines -Label 'Baseline completeness' -Value $(if ($commandAvailable -and $captured) { 'Complete' } else { 'Partial / incomplete' })
            if ($Entry.PSObject.Properties['CaptureError'] -and -not [string]::IsNullOrWhiteSpace([string]$Entry.CaptureError)) {
                Add-ReportKeyValueLine -Lines $Lines -Label 'Capture error' -Value $Entry.CaptureError
            }

            foreach ($listener in $listeners) {
                $Lines.Add(('  - {0} {1} | Port={2} | Enabled={3} | Hostname={4} | URLPrefix={5} | Cert={6}' -f $listener.Transport, $listener.Address, $listener.Port, $listener.Enabled, (ConvertTo-DisplayString -Value $listener.Hostname), (ConvertTo-DisplayString -Value $listener.URLPrefix), (ConvertTo-DisplayString -Value $listener.CertificateThumbprint)))
            }
        }
        'AuditPolicy' {
            Add-ReportKeyValueLine -Lines $Lines -Label 'Subcategory' -Value $Entry.Subcategory
            Add-ReportKeyValueLine -Lines $Lines -Label 'Success' -Value $Entry.CurrentValue.Success
            Add-ReportKeyValueLine -Lines $Lines -Label 'Failure' -Value $Entry.CurrentValue.Failure
        }
        'SmbClientConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Timed out' -Value $state.TimedOut
            Add-ReportKeyValueLine -Lines $Lines -Label 'Require security signature' -Value $state.RequireSecuritySignature
        }
        'SmbServerConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            Add-ReportKeyValueLine -Lines $Lines -Label 'Command available' -Value $state.CommandAvailable
            Add-ReportKeyValueLine -Lines $Lines -Label 'Timed out' -Value $state.TimedOut
            Add-ReportKeyValueLine -Lines $Lines -Label 'Require security signature' -Value $state.RequireSecuritySignature
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
    $incompleteBaselineCount = @($settings | Where-Object { -not (Test-SnapshotEntryCapturedExactly -Entry $_ -SnapshotPath $SnapshotPath) }).Count
    $platformManagedWdacCount = 0
    $wdacEntry = @($settings | Where-Object { [string]$_.Type -eq 'WdacPolicies' } | Select-Object -First 1)
    if ($wdacEntry.Count -gt 0) {
        $platformManagedWdacCount = @(Get-WdacPlatformManagedPolicies -State $wdacEntry[0].CurrentValue).Count
    }

    $lines.Add('WinDefState Snapshot Report')
    $lines.Add(('Snapshot JSON: {0}' -f $SnapshotPath))
    $lines.Add(('ComputerName: {0}' -f $Snapshot.ComputerName))
    $lines.Add(('CapturedAtUtc: {0}' -f $Snapshot.CapturedAtUtc))
    $lines.Add(('Settings captured: {0}' -f $settings.Count))
    $lines.Add(('Reboot-required settings: {0}' -f (@($settings | Where-Object { $_.RequiresReboot }).Count)))
    $lines.Add(('Incomplete-baseline settings: {0}' -f $incompleteBaselineCount))
    $lines.Add(('Platform-managed WDAC policies: {0}' -f $platformManagedWdacCount))

    foreach ($entry in $settings) {
        $lines.Add(' ')
        Add-SnapshotEntryReportLines -Lines $lines -Entry $entry -SnapshotPath $SnapshotPath
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
    param(
        [Parameter(Mandatory)] [object]$Entry,
        [string]$SnapshotPath,
        [AllowNull()] [object]$ReferenceEntry
    )

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
        'MpPreferenceList' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id               = $Entry.Id
                Type             = $Entry.Type
                Property         = $Entry.Property
                CommandAvailable = $commandAvailable
                Captured         = $captured
                CurrentValue     = @(Normalize-MpPreferenceListItems -Value $Entry.CurrentValue)
                RequiresReboot   = $Entry.RequiresReboot
            })
        }
        'DefenderRuntimeStatus' {
            $state = if ($null -ne $Entry.CurrentValue) { $Entry.CurrentValue } else { [PSCustomObject]@{} }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable          = if ($state.PSObject.Properties['CommandAvailable']) { [bool]$state.CommandAvailable } else { $true }
                    Captured                  = if ($state.PSObject.Properties['Captured']) { [bool]$state.Captured } else { $true }
                    AMRunningMode             = if ($state.PSObject.Properties['AMRunningMode']) { [string]$state.AMRunningMode } else { $null }
                    RealTimeProtectionEnabled = if ($state.PSObject.Properties['RealTimeProtectionEnabled']) { ConvertTo-NullableBoolean -Value $state.RealTimeProtectionEnabled } else { $null }
                    AntivirusEnabled          = if ($state.PSObject.Properties['AntivirusEnabled']) { ConvertTo-NullableBoolean -Value $state.AntivirusEnabled } else { $null }
                    IsTamperProtected         = if ($state.PSObject.Properties['IsTamperProtected']) { ConvertTo-NullableBoolean -Value $state.IsTamperProtected } else { $null }
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'BitLockerVolumes' {
            $referenceState = if ($null -ne $ReferenceEntry -and [string]$ReferenceEntry.Type -eq 'BitLockerVolumes') { $ReferenceEntry.CurrentValue } else { $null }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = ConvertTo-ComparableBitLockerState -State $Entry.CurrentValue -ReferenceState $referenceState
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'AppLockerPolicy' {
            $effectiveXml = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $Entry.CurrentValue -PolicyScope Effective -SnapshotPath $SnapshotPath)
            $localXml = Normalize-AppLockerXml -Xml (Get-AppLockerPolicyXml -State $Entry.CurrentValue -PolicyScope Local -SnapshotPath $SnapshotPath)
            $collectionSummaries = if ($Entry.CurrentValue.PSObject.Properties['CollectionSummaries']) { @($Entry.CurrentValue.CollectionSummaries) } else { @(Get-AppLockerCollectionSummaries -Xml $effectiveXml) }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable      = if ($Entry.CurrentValue.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CurrentValue.CommandAvailable } else { $true }
                    LocalCaptured         = if ($Entry.CurrentValue.PSObject.Properties['LocalCaptured']) { [bool]$Entry.CurrentValue.LocalCaptured } else { $false }
                    EffectiveCaptured     = if ($Entry.CurrentValue.PSObject.Properties['EffectiveCaptured']) { [bool]$Entry.CurrentValue.EffectiveCaptured } else { $false }
                    LocalMatchesEffective = if ($Entry.CurrentValue.PSObject.Properties['LocalMatchesEffective']) { [bool]$Entry.CurrentValue.LocalMatchesEffective } else { $false }
                    CaptureIssues         = @(
                        foreach ($issue in @($Entry.CurrentValue.CaptureIssues | Sort-Object -Property Scope, Message)) {
                            [PSCustomObject]@{
                                Scope   = if ($null -ne $issue.PSObject.Properties['Scope']) { [string]$issue.Scope } else { $null }
                                Message = if ($null -ne $issue.PSObject.Properties['Message']) { [string]$issue.Message } else { $null }
                            }
                        }
                    )
                    CollectionSummaries   = @(
                        foreach ($collection in @($collectionSummaries | Sort-Object -Property Type)) {
                            [PSCustomObject]@{
                                Type                = if ($null -ne $collection.PSObject.Properties['Type']) { [string]$collection.Type } else { $null }
                                EnforcementMode     = if ($null -ne $collection.PSObject.Properties['EnforcementMode']) { [string]$collection.EnforcementMode } else { $null }
                                RuleCount           = if ($null -ne $collection.PSObject.Properties['RuleCount']) { [int]$collection.RuleCount } else { 0 }
                                ServicesEnforcement = if ($null -ne $collection.PSObject.Properties['ServicesEnforcement']) { [string]$collection.ServicesEnforcement } else { $null }
                                SystemAppsAllow     = if ($null -ne $collection.PSObject.Properties['SystemAppsAllow']) { [string]$collection.SystemAppsAllow } else { $null }
                            }
                        }
                    )
                    EffectiveXml          = $effectiveXml
                    LocalXml              = $localXml
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'ExploitProtectionPolicy' {
            $xml = Get-ExploitProtectionPolicyXml -State $Entry.CurrentValue -SnapshotPath $SnapshotPath
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable = $Entry.CurrentValue.CommandAvailable
                    Xml              = Normalize-ExploitProtectionXml -Xml $xml
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'WdacPolicies' {
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CiToolAvailable = $Entry.CurrentValue.CiToolAvailable
                    Policies        = @(
                        foreach ($policy in @($Entry.CurrentValue.Policies)) {
                            $ordered = [ordered]@{}
                            foreach ($property in @($policy.PSObject.Properties | Sort-Object -Property Name)) {
                                $ordered[$property.Name] = $property.Value
                            }
                            [PSCustomObject]$ordered
                        }
                    )
                    Files           = @(
                        foreach ($file in @($Entry.CurrentValue.Files)) {
                            [PSCustomObject]@{
                                RelativePath = [string]$file.RelativePath
                                FileName     = [string]$file.FileName
                                Sha256       = [string]$file.Sha256
                            }
                        }
                    )
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'AsrRules' {
            $invalidEntries = @(Get-AsrInvalidEntriesFromEntry -Entry $Entry)
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
                InvalidEntries = @(
                    foreach ($rule in $invalidEntries) {
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
            $referenceState = if ($null -ne $ReferenceEntry -and [string]$ReferenceEntry.Type -eq 'FirewallProfiles') { $ReferenceEntry.CurrentValue } else { $null }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = ConvertTo-ComparableFirewallProfileState -State $Entry.CurrentValue -ReferenceState $referenceState
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'FirewallRules' {
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = ConvertTo-ComparableFirewallRuleState -State $Entry.CurrentValue
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
            $state = Normalize-UserRegistryValueState -State $Entry.CurrentValue
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    Entries       = @(
                        foreach ($value in @($state.Entries | Sort-Object -Property Sid, RelativePath, Name)) {
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
                    CaptureIssues = @(
                        foreach ($issue in @($state.CaptureIssues | Sort-Object -Property Sid)) {
                            [string]$issue.Sid
                        }
                    )
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'WsManValue' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                Path           = $Entry.Path
                CommandAvailable = $commandAvailable
                Captured       = $captured
                CurrentValue   = $Entry.CurrentValue
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'WinRmListeners' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CommandAvailable = $commandAvailable
                Captured       = $captured
                CurrentValue   = @(
                    foreach ($listener in @($Entry.CurrentValue | Sort-Object -Property Address, Transport, Port, Hostname, URLPrefix, CertificateThumbprint)) {
                        [PSCustomObject]@{
                            Address               = [string]$listener.Address
                            Transport             = [string]$listener.Transport
                            Port                  = if ($null -ne $listener.PSObject.Properties['Port'] -and $null -ne $listener.Port) { [int]$listener.Port } else { $null }
                            Hostname              = if ($null -ne $listener.PSObject.Properties['Hostname']) { [string]$listener.Hostname } else { $null }
                            Enabled               = if ($null -ne $listener.PSObject.Properties['Enabled']) { ConvertTo-NullableBoolean -Value $listener.Enabled } else { $null }
                            URLPrefix             = if ($null -ne $listener.PSObject.Properties['URLPrefix']) { [string]$listener.URLPrefix } else { $null }
                            CertificateThumbprint = if ($null -ne $listener.PSObject.Properties['CertificateThumbprint']) { [string]$listener.CertificateThumbprint } else { $null }
                        }
                    }
                )
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'SmbClientConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable         = $state.CommandAvailable
                    TimedOut                 = $state.TimedOut
                    RequireSecuritySignature = $state.RequireSecuritySignature
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        'SmbServerConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            return ConvertTo-CanonicalValue -Value ([ordered]@{
                Id             = $Entry.Id
                Type           = $Entry.Type
                CurrentValue   = [ordered]@{
                    CommandAvailable         = $state.CommandAvailable
                    TimedOut                 = $state.TimedOut
                    RequireSecuritySignature = $state.RequireSecuritySignature
                }
                RequiresReboot = $Entry.RequiresReboot
            })
        }
        default {
            return ConvertTo-CanonicalValue -Value $Entry
        }
    }
}

function Test-SnapshotEntryCapturedExactly {
    param(
        [Parameter(Mandatory)] [object]$Entry,
        [string]$SnapshotPath
    )

    switch ($Entry.Type) {
        'MpPreferenceList' {
            return (Test-MpPreferenceListCapturedExactly -Entry $Entry)
        }
        'DefenderRuntimeStatus' {
            return (Test-DefenderRuntimeStatusCapturedExactly -State $Entry.CurrentValue)
        }
        'BitLockerVolumes' {
            return (Test-BitLockerStateCapturedExactly -State $Entry.CurrentValue)
        }
        'AppLockerPolicy' {
            return (Test-AppLockerPolicyCapturedExactly -State $Entry.CurrentValue -SnapshotPath $SnapshotPath)
        }
        'ExploitProtectionPolicy' {
            return (Test-ExploitProtectionPolicyCapturedExactly -State $Entry.CurrentValue -SnapshotPath $SnapshotPath)
        }
        'AsrRules' {
            $invalidEntries = @(Get-AsrInvalidEntriesFromEntry -Entry $Entry)
            return ($invalidEntries.Count -eq 0)
        }
        'FirewallProfiles' {
            return (Test-FirewallProfileStateCapturedExactly -State $Entry.CurrentValue)
        }
        'FirewallRules' {
            return (Test-FirewallRuleStateCapturedExactly -State $Entry.CurrentValue)
        }
        'SmbClientConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            return $state.CommandAvailable -and -not $state.TimedOut -and $null -ne $state.RequireSecuritySignature
        }
        'SmbServerConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            return $state.CommandAvailable -and -not $state.TimedOut -and $null -ne $state.RequireSecuritySignature
        }
        'LoadedUserRegistryValues' {
            $state = Normalize-UserRegistryValueState -State $Entry.CurrentValue
            return (@($state.CaptureIssues).Count -eq 0)
        }
        'WsManValue' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ($commandAvailable -and $captured)
        }
        'WinRmListeners' {
            $commandAvailable = if ($Entry.PSObject.Properties['CommandAvailable']) { [bool]$Entry.CommandAvailable } else { $true }
            $captured = if ($Entry.PSObject.Properties['Captured']) { [bool]$Entry.Captured } else { $true }
            return ($commandAvailable -and $captured)
        }
        default {
            return $true
        }
    }
}

function Test-DefenseSnapshot {
    param(
        [Parameter(Mandatory)] [object]$Snapshot,
        [string]$SnapshotPath
    )

    $definitions = @{}
    foreach ($definition in Get-DefenseDefinitions) {
        $definitions[[string]$definition.Id] = $definition
    }

    $results = @()
    foreach ($entry in @($Snapshot.Settings)) {
        $entryId = [string]$entry.Id
        $expectedComparable = ConvertTo-ComparableSnapshotEntry -Entry $entry -SnapshotPath $SnapshotPath

        if (-not (Test-SnapshotEntryCapturedExactly -Entry $entry -SnapshotPath $SnapshotPath)) {
            $results += [PSCustomObject]@{
                Id             = $entryId
                Type           = $entry.Type
                RequiresReboot = $entry.RequiresReboot
                Matches        = $false
                Skipped        = $true
                Reason         = 'Verification skipped because the snapshot baseline was incomplete.'
                Expected       = $expectedComparable
                Actual         = $null
            }
            continue
        }

        if (-not $definitions.ContainsKey($entryId)) {
            $results += [PSCustomObject]@{
                Id             = $entryId
                Type           = $entry.Type
                RequiresReboot = $entry.RequiresReboot
                Matches        = $false
                Skipped        = $false
                Reason         = 'This snapshot entry no longer has a matching definition in the current script.'
                Expected       = $expectedComparable
                Actual         = $null
            }
            continue
        }

        try {
            $liveEntry = Capture-Definition -Definition $definitions[$entryId]
            $actualComparable = ConvertTo-ComparableSnapshotEntry -Entry $liveEntry -ReferenceEntry $entry
            $matches = (Get-CanonicalJson -Value $expectedComparable) -eq (Get-CanonicalJson -Value $actualComparable)

            $results += [PSCustomObject]@{
                Id             = $entryId
                Type           = $entry.Type
                RequiresReboot = $entry.RequiresReboot
                Matches        = $matches
                Skipped        = $false
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
                Skipped        = $false
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
        MatchedCount  = @($results | Where-Object { $_.Matches -and -not $_.Skipped }).Count
        SkippedCount  = @($results | Where-Object { $_.Skipped }).Count
        MismatchCount = @($results | Where-Object { -not $_.Matches -and -not $_.Skipped }).Count
        Results       = @($results)
    }
}

function Get-VerificationReportLines {
    param(
        [Parameter(Mandatory)] [object]$Verification,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $mismatches = @($Verification.Results | Where-Object { -not $_.Matches -and -not $_.Skipped })
    $skipped = @($Verification.Results | Where-Object { $_.Skipped })

    $lines.Add('WinDefState Restore Verification')
    $lines.Add(('Snapshot JSON: {0}' -f $SnapshotPath))
    $lines.Add(('ComputerName: {0}' -f $Verification.ComputerName))
    $lines.Add(('VerifiedAtUtc: {0}' -f $Verification.VerifiedAtUtc))
    $lines.Add(('Matched settings: {0}' -f $Verification.MatchedCount))
    $lines.Add(('Skipped settings: {0}' -f $Verification.SkippedCount))
    $lines.Add(('Mismatched settings: {0}' -f $Verification.MismatchCount))

    if ($mismatches.Count -eq 0) {
        $lines.Add(' ')
        if ($skipped.Count -eq 0) {
            $lines.Add('All captured settings match the requested snapshot.')
        } else {
            $lines.Add('All fully captured settings match the requested snapshot.')
        }
    }

    foreach ($skippedResult in $skipped) {
        $lines.Add(' ')
        $lines.Add("[$($skippedResult.Id)] $($skippedResult.Type)")
        Add-ReportKeyValueLine -Lines $lines -Label 'Requires reboot' -Value $skippedResult.RequiresReboot
        if (-not [string]::IsNullOrWhiteSpace([string]$skippedResult.Reason)) {
            Add-ReportKeyValueLine -Lines $lines -Label 'Reason' -Value $skippedResult.Reason
        }
    }

    if ($mismatches.Count -eq 0) {
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

function Get-WdacVerificationReportLines {
    param(
        [Parameter(Mandatory)] [object]$Verification,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $wdacResults = @($Verification.Results | Where-Object { [string]$_.Type -eq 'WdacPolicies' -or [string]$_.Id -eq 'wdac.policies' })

    $lines.Add('WinDefState WDAC Restore Verification')
    $lines.Add(('Snapshot JSON: {0}' -f $SnapshotPath))
    $lines.Add(('ComputerName: {0}' -f $Verification.ComputerName))
    $lines.Add(('VerifiedAtUtc: {0}' -f $Verification.VerifiedAtUtc))
    $lines.Add(('WDAC result count: {0}' -f $wdacResults.Count))

    if ($wdacResults.Count -eq 0) {
        $lines.Add(' ')
        $lines.Add('No WDAC verification entry was present in this verification run.')
        return [string[]]$lines
    }

    foreach ($result in $wdacResults) {
        $lines.Add(' ')
        $lines.Add("[$($result.Id)] $($result.Type)")
        Add-ReportKeyValueLine -Lines $lines -Label 'Requires reboot' -Value $result.RequiresReboot
        Add-ReportKeyValueLine -Lines $lines -Label 'Matches snapshot' -Value $result.Matches
        Add-ReportKeyValueLine -Lines $lines -Label 'Skipped' -Value $result.Skipped
        if (-not [string]::IsNullOrWhiteSpace([string]$result.Reason)) {
            Add-ReportKeyValueLine -Lines $lines -Label 'Reason' -Value $result.Reason
        }
        Add-ReportJsonBlock -Lines $lines -Label 'Expected' -Value $result.Expected
        Add-ReportJsonBlock -Lines $lines -Label 'Actual' -Value $result.Actual
    }

    [string[]]$lines
}

function Persist-SnapshotExternalAssets {
    param(
        [Parameter(Mandatory)] [object]$Snapshot,
        [Parameter(Mandatory)] [string]$SnapshotPath
    )

    $assetRoot = Get-SnapshotAssetRoot -SnapshotPath $SnapshotPath

    foreach ($entry in @($Snapshot.Settings)) {
        if ([string]$entry.Type -eq 'AppLockerPolicy' -and $null -ne $entry.CurrentValue) {
            $localXml = Get-AppLockerPolicyXml -State $entry.CurrentValue -PolicyScope Local
            $effectiveXml = Get-AppLockerPolicyXml -State $entry.CurrentValue -PolicyScope Effective
            $localAssetRelativePath = if (-not [string]::IsNullOrWhiteSpace($localXml)) { Join-Path 'applocker' 'local-policy.xml' } else { $null }
            $effectiveAssetRelativePath = if (-not [string]::IsNullOrWhiteSpace($effectiveXml)) { Join-Path 'applocker' 'effective-policy.xml' } else { $null }

            if (-not [string]::IsNullOrWhiteSpace($localXml)) {
                Write-TextAtomic -Path (Join-Path $assetRoot $localAssetRelativePath) -Content $localXml
            }

            if (-not [string]::IsNullOrWhiteSpace($effectiveXml)) {
                Write-TextAtomic -Path (Join-Path $assetRoot $effectiveAssetRelativePath) -Content $effectiveXml
            }

            $entry.CurrentValue = [PSCustomObject]@{
                CommandAvailable               = $entry.CurrentValue.CommandAvailable
                LocalCaptured                  = $entry.CurrentValue.LocalCaptured
                EffectiveCaptured              = $entry.CurrentValue.EffectiveCaptured
                LocalMatchesEffective          = $entry.CurrentValue.LocalMatchesEffective
                CaptureIssues                  = @($entry.CurrentValue.CaptureIssues)
                CollectionSummaries            = @($entry.CurrentValue.CollectionSummaries)
                LocalSnapshotAssetRelativePath = $localAssetRelativePath
                EffectiveSnapshotAssetRelativePath = $effectiveAssetRelativePath
            }
            continue
        }

        if ([string]$entry.Type -eq 'ExploitProtectionPolicy' -and $null -ne $entry.CurrentValue) {
            $xml = Get-ExploitProtectionPolicyXml -State $entry.CurrentValue
            $assetRelativePath = Join-Path 'exploit-protection' 'policy.xml'
            if (-not [string]::IsNullOrWhiteSpace($xml)) {
                $assetPath = Join-Path $assetRoot $assetRelativePath
                Write-TextAtomic -Path $assetPath -Content $xml
            }

            $entry.CurrentValue = [PSCustomObject]@{
                CommandAvailable         = $entry.CurrentValue.CommandAvailable
                SnapshotAssetRelativePath = if (-not [string]::IsNullOrWhiteSpace($xml)) { $assetRelativePath } else { $null }
            }
            continue
        }

        if ([string]$entry.Type -ne 'WdacPolicies' -or $null -eq $entry.CurrentValue) {
            continue
        }

        $rewrittenFiles = @()
        foreach ($file in @($entry.CurrentValue.Files)) {
            $relativePath = [string]$file.RelativePath
            $assetRelativePath = Join-Path 'wdac' $relativePath
            $assetPath = Join-Path $assetRoot $assetRelativePath

            if ($file.PSObject.Properties['Base64'] -and -not [string]::IsNullOrWhiteSpace([string]$file.Base64)) {
                Write-BytesAtomic -Path $assetPath -Content ([Convert]::FromBase64String([string]$file.Base64))
            }

            $rewrittenFiles += [PSCustomObject]@{
                RelativePath              = $relativePath
                FileName                  = [string]$file.FileName
                Sha256                    = [string]$file.Sha256
                SnapshotAssetRelativePath = $assetRelativePath
            }
        }

        $entry.CurrentValue = [PSCustomObject]@{
            CiToolAvailable = $entry.CurrentValue.CiToolAvailable
            Policies        = @($entry.CurrentValue.Policies)
            Files           = @($rewrittenFiles)
        }
    }
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
        [PSCustomObject]@{ Id = 'defender.runtime_status'; Type = 'DefenderRuntimeStatus'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_realtime_monitoring'; Type = 'MpPreferenceValue'; Property = 'DisableRealtimeMonitoring'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_behavior_monitoring'; Type = 'MpPreferenceValue'; Property = 'DisableBehaviorMonitoring'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.maps_reporting'; Type = 'MpPreferenceValue'; Property = 'MAPSReporting'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Basic'; '2' = 'Advanced' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.submit_samples_consent'; Type = 'MpPreferenceValue'; Property = 'SubmitSamplesConsent'; PermissiveValue = 'NeverSend'; ValueMap = @{ '0' = 'AlwaysPrompt'; '1' = 'SendSafeSamples'; '2' = 'NeverSend'; '3' = 'SendAllSamples' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.pua_protection'; Type = 'MpPreferenceValue'; Property = 'PUAProtection'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_script_scanning'; Type = 'MpPreferenceValue'; Property = 'DisableScriptScanning'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_ioav_protection'; Type = 'MpPreferenceValue'; Property = 'DisableIOAVProtection'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.disable_intrusion_prevention_system'; Type = 'MpPreferenceValue'; Property = 'DisableIntrusionPreventionSystem'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.enable_network_protection'; Type = 'MpPreferenceValue'; Property = 'EnableNetworkProtection'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.enable_controlled_folder_access'; Type = 'MpPreferenceValue'; Property = 'EnableControlledFolderAccess'; PermissiveValue = 'Disabled'; ValueMap = @{ '0' = 'Disabled'; '1' = 'Enabled'; '2' = 'AuditMode' }; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.exclusion_paths'; Type = 'MpPreferenceList'; Property = 'ExclusionPath'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.exclusion_processes'; Type = 'MpPreferenceList'; Property = 'ExclusionProcess'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.exclusion_extensions'; Type = 'MpPreferenceList'; Property = 'ExclusionExtension'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.cfa_allowed_applications'; Type = 'MpPreferenceList'; Property = 'ControlledFolderAccessAllowedApplications'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.cfa_protected_folders'; Type = 'MpPreferenceList'; Property = 'ControlledFolderAccessProtectedFolders'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'defender.asr_rules'; Type = 'AsrRules'; RequiresReboot = $false }

        [PSCustomObject]@{
            Id                                     = 'firewall.profiles'
            Type                                   = 'FirewallProfiles'
            Profiles                               = @('Domain', 'Private', 'Public')
            PermissiveValue                        = $false
            PermissiveDefaultInboundAction         = 'Allow'
            PermissiveDefaultOutboundAction        = 'Allow'
            PermissiveAllowUnicastResponseToMulticast = $true
            PermissiveNotifyOnListen               = $false
            PermissiveLogAllowed                   = $false
            PermissiveLogBlocked                   = $false
            PermissiveLogIgnored                   = $false
            RequiresReboot                         = $false
        }

        [PSCustomObject]@{ Id = 'powershell.lockdown'; Type = 'MachineEnvironmentValue'; Name = '__PSLockdownPolicy'; PermissiveExists = $true; PermissiveValue = '0'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'powershell.script_block_logging'; Type = 'RegistryKeyFlat'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'powershell.module_logging'; Type = 'PowerShellModuleLogging'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'powershell.transcription'; Type = 'RegistryKeyFlat'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\Transcription'; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'applocker.policy'; Type = 'AppLockerPolicy'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'applocker.service'; Type = 'ServiceConfig'; Name = 'AppIDSvc'; PermissiveStartup = 'demand'; PermissiveState = 'Stopped'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'print.spooler_service'; Type = 'ServiceConfig'; Name = 'Spooler'; PermissiveStartup = 'auto'; PermissiveState = 'Running'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'localuser.administrator'; Type = 'LocalUser'; Rid = 500; Name = 'Built-in local administrator'; PermissiveValue = $true; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'uac.prompt_on_secure_desktop'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'PromptOnSecureDesktop'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'uac.enable_lua'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'EnableLUA'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'uac.consent_prompt_behavior_admin'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'; Name = 'ConsentPromptBehaviorAdmin'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $true }

        [PSCustomObject]@{ Id = 'rdp.allow_connections'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'; Name = 'fDenyTSConnections'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'rdp.user_authentication'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'; Name = 'UserAuthentication'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'rdp.security_layer'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'; Name = 'SecurityLayer'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'rdp.listener_enabled'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'; Name = 'fEnableWinStation'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'rdp.firewall_rules'; Type = 'FirewallRules'; Group = '@FirewallAPI.dll,-28752'; PermissiveEnabled = $true; RequiresReboot = $false }

        [PSCustomObject]@{ Id = 'wsh.enabled'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Microsoft\Windows Script Host\Settings'; Name = 'Enabled'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'smartscreen.enable'; Type = 'RegistryValue'; Path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Name = 'EnableSmartScreen'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 0; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'sehop.disable_exception_chain_validation'; Type = 'RegistryValue'; Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\kernel'; Name = 'DisableExceptionChainValidation'; ValueKind = 'DWord'; PermissiveExists = $true; PermissiveValue = 1; RequiresReboot = $true }
        [PSCustomObject]@{ Id = 'exploit_protection.policy'; Type = 'ExploitProtectionPolicy'; RequiresReboot = $true }

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

        [PSCustomObject]@{ Id = 'winrm.service'; Type = 'ServiceConfig'; Name = 'WinRM'; PermissiveStartup = 'auto'; PermissiveState = 'Running'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.allow_unencrypted'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\AllowUnencrypted'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.allow_unencrypted'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\AllowUnencrypted'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.basic'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Basic'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.basic'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Basic'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.digest'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Digest'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.certificate'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Certificate'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.certificate'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Certificate'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.kerberos'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Kerberos'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.kerberos'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Kerberos'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.negotiate'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\Negotiate'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.negotiate'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\Negotiate'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.credssp'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\Auth\CredSSP'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.credssp'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\Auth\CredSSP'; PermissiveValue = $true; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.cbt_hardening_level'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\CbtHardeningLevel'; PermissiveValue = 'Relaxed'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.ipv4_filter'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\IPv4Filter'; PermissiveValue = '*'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.service.ipv6_filter'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Service\IPv6Filter'; PermissiveValue = '*'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'winrm.client.trusted_hosts'; Type = 'WsManValue'; Path = 'WSMan:\localhost\Client\TrustedHosts'; PermissiveValue = '*'; RequiresReboot = $false }
        [PSCustomObject]@{
            Id             = 'winrm.listeners'
            Type           = 'WinRmListeners'
            PermissiveValue = @(
                [PSCustomObject]@{
                    Address               = '*'
                    Transport             = 'HTTP'
                    Port                  = 5985
                    Hostname              = ''
                    Enabled               = $true
                    URLPrefix             = 'wsman'
                    CertificateThumbprint = ''
                }
            )
            RequiresReboot = $false
        }

        [PSCustomObject]@{ Id = 'smb.client.require_security_signature'; Type = 'SmbClientConfig'; PermissiveValue = $false; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'smb.server.require_security_signature'; Type = 'SmbServerConfig'; PermissiveValue = $false; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'bitlocker.volumes'; Type = 'BitLockerVolumes'; RequiresReboot = $false }
        [PSCustomObject]@{ Id = 'wdac.policies'; Type = 'WdacPolicies'; RequiresReboot = $true }
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
            $rawValue = Get-MpPreferencePropertyRawValue -Property $Definition.Property

            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'MpPreferenceValue'
                Property       = $Definition.Property
                CurrentValue   = $rawValue
                RestoreValue   = Resolve-MpPreferenceValue -Definition $Definition -Value $rawValue
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'MpPreferenceList' {
            $state = Get-MpPreferenceListState -Property $Definition.Property

            return [PSCustomObject]@{
                Id               = $Definition.Id
                Type             = 'MpPreferenceList'
                Property         = $Definition.Property
                CurrentValue     = @($state.Items)
                CommandAvailable = $state.CommandAvailable
                Captured         = $state.Captured
                CaptureError     = $state.Error
                RequiresReboot   = $Definition.RequiresReboot
            }
        }
        'DefenderRuntimeStatus' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'DefenderRuntimeStatus'
                CurrentValue   = Get-DefenderRuntimeStatus
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'BitLockerVolumes' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'BitLockerVolumes'
                CurrentValue   = Get-BitLockerVolumeStates
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'ExploitProtectionPolicy' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'ExploitProtectionPolicy'
                CurrentValue   = Get-ExploitProtectionPolicyState
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'WdacPolicies' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'WdacPolicies'
                CurrentValue   = Get-WdacPolicyState
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'AsrRules' {
            $asrState = Get-AsrRuleCaptureState
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'AsrRules'
                CurrentValue   = @($asrState.Rules)
                InvalidEntries = @($asrState.InvalidEntries)
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'PowerShellModuleLogging' {
            return Capture-PowerShellModuleLoggingState
        }
        'AppLockerPolicy' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'AppLockerPolicy'
                CurrentValue   = Get-AppLockerPolicyState
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'FirewallProfiles' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'FirewallProfiles'
                CurrentValue   = Get-FirewallProfileStates -Profiles @($Definition.Profiles)
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'FirewallRules' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'FirewallRules'
                CurrentValue   = Get-FirewallRuleGroupState -Group ([string]$Definition.Group)
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
                CurrentValue   = Get-LoadedUserRegistryValueStates -Items @($Definition.Items)
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
            $state = Get-WsManConfigValueState -Path $Definition.Path
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'WsManValue'
                Path           = $Definition.Path
                CurrentValue   = $state.Value
                CommandAvailable = $state.CommandAvailable
                Captured       = $state.Captured
                CaptureError   = $state.Error
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'WinRmListeners' {
            $state = Get-WinRmListenerStates
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'WinRmListeners'
                CurrentValue   = @($state.Listeners)
                CommandAvailable = $state.CommandAvailable
                Captured       = $state.Captured
                CaptureError   = $state.Error
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
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'SmbClientConfig'
                CurrentValue   = Get-SmbClientConfigurationState
                RequiresReboot = $Definition.RequiresReboot
            }
        }
        'SmbServerConfig' {
            return [PSCustomObject]@{
                Id             = $Definition.Id
                Type           = 'SmbServerConfig'
                CurrentValue   = Get-SmbServerConfigurationState
                RequiresReboot = $Definition.RequiresReboot
            }
        }
    }
}

function Apply-PermissiveDefinition {
    param(
        [Parameter(Mandatory)] [object]$Definition,
        [AllowNull()] [object]$Entry
    )

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
        'MpPreferenceList' {
            if ($Definition.PSObject.Properties['PermissiveValue']) {
                Set-MpPreferenceListValue -Property $Definition.Property -DesiredItems @($Definition.PermissiveValue)
            }
        }
        'DefenderRuntimeStatus' {
        }
        'BitLockerVolumes' {
            $bitLockerState = if ($null -ne $Entry) { $Entry.CurrentValue } else { $null }
            Set-Permissive-BitLockerVolumes -State $bitLockerState
        }
        'ExploitProtectionPolicy' {
            Set-Permissive-ExploitProtection
        }
        'WdacPolicies' {
            Remove-WdacPolicies -State (Get-WdacPolicyState)
        }
        'AsrRules' {
            Disable-ConfiguredAsrRules
        }
        'PowerShellModuleLogging' {
            Set-Permissive-PowerShellModuleLogging
        }
        'AppLockerPolicy' {
            $appLockerState = if ($null -ne $Entry) { $Entry.CurrentValue } else { $null }
            Set-Permissive-AppLockerPolicy -State $appLockerState
        }
        'FirewallProfiles' {
            Set-Permissive-FirewallProfiles -Definition $Definition
        }
        'FirewallRules' {
            Set-Permissive-FirewallRules -Definition $Definition
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
            Set-WsManConfigValue -Path $Definition.Path -Value $Definition.PermissiveValue
        }
        'WinRmListeners' {
            Set-Permissive-WinRmListeners -Listeners @($Definition.PermissiveValue)
        }
        'AuditPolicy' {
            Set-AuditPolicyState -Subcategory $Definition.Subcategory -Success $Definition.PermissiveSuccess -Failure $Definition.PermissiveFailure
        }
        'SmbClientConfig' {
            if (Test-CommandAvailable -Name 'Set-SmbClientConfiguration') {
                Set-SmbClientConfiguration -RequireSecuritySignature $Definition.PermissiveValue -Confirm:$false | Out-Null
            }
        }
        'SmbServerConfig' {
            if (Test-CommandAvailable -Name 'Set-SmbServerConfiguration') {
                Set-SmbServerConfiguration -RequireSecuritySignature $Definition.PermissiveValue -Confirm:$false | Out-Null
            }
        }
    }
}

function Restore-SnapshotEntry {
    param(
        [Parameter(Mandatory)] [object]$Entry,
        [string]$SnapshotPath
    )

    if (-not (Test-SnapshotEntryCapturedExactly -Entry $Entry -SnapshotPath $SnapshotPath)) {
        Write-Warning "Skipping restore for $($Entry.Id) because the snapshot baseline was incomplete."
        return
    }

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
        'MpPreferenceList' {
            Set-MpPreferenceListValue -Property $Entry.Property -DesiredItems @($Entry.CurrentValue)
        }
        'DefenderRuntimeStatus' {
        }
        'BitLockerVolumes' {
            Restore-BitLockerVolumes -State $Entry.CurrentValue
        }
        'ExploitProtectionPolicy' {
            # Reset lingering system mitigation state first because PolicyFilePath import
            # does not reliably clear permissive-era SystemConfig entries on its own.
            Reset-ExploitProtectionSystemConfig
            Apply-ExploitProtectionPolicyXml -Xml (Get-ExploitProtectionPolicyXml -State $Entry.CurrentValue -SnapshotPath $SnapshotPath)
        }
        'WdacPolicies' {
            Restore-WdacPolicies -State $Entry.CurrentValue -SnapshotPath $SnapshotPath
        }
        'AsrRules' {
            Restore-AsrRules -Rules @($Entry.CurrentValue)
        }
        'PowerShellModuleLogging' {
            Restore-PowerShellModuleLogging -Entry $Entry
        }
        'AppLockerPolicy' {
            Restore-AppLockerPolicy -State $Entry.CurrentValue -SnapshotPath $SnapshotPath
        }
        'FirewallProfiles' {
            Restore-FirewallProfiles -State $Entry.CurrentValue
        }
        'FirewallRules' {
            Restore-FirewallRules -State $Entry.CurrentValue
        }
        'NetBiosAdapters' {
            Restore-NetBiosAdapters -Adapters @($Entry.CurrentValue)
        }
        'LoadedUserRegistryValues' {
            $state = Normalize-UserRegistryValueState -State $Entry.CurrentValue
            Restore-LoadedUserRegistryValues -Entries @($state.Entries)
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
            Set-WsManConfigValue -Path $Entry.Path -Value $Entry.CurrentValue
        }
        'WinRmListeners' {
            Restore-WinRmListeners -Listeners @($Entry.CurrentValue)
        }
        'AuditPolicy' {
            Set-AuditPolicyState -Subcategory $Entry.Subcategory -Success ([bool]$Entry.CurrentValue.Success) -Failure ([bool]$Entry.CurrentValue.Failure)
        }
        'SmbClientConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            if (Test-CommandAvailable -Name 'Set-SmbClientConfiguration') {
                Set-SmbClientConfiguration -RequireSecuritySignature ([bool]$state.RequireSecuritySignature) -Confirm:$false | Out-Null
            }
        }
        'SmbServerConfig' {
            $state = Normalize-SmbConfigState -Value $Entry.CurrentValue
            if (Test-CommandAvailable -Name 'Set-SmbServerConfiguration') {
                Set-SmbServerConfiguration -RequireSecuritySignature ([bool]$state.RequireSecuritySignature) -Confirm:$false | Out-Null
            }
        }
    }
}

function Export-DefenseSnapshot {
    param([Parameter(Mandatory)] [string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    $definitions = @(Get-DefenseDefinitions)
    $settings = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $definitions.Count; $i++) {
        $definition = $definitions[$i]
        Write-Verbose ("[{0}/{1}] Capturing {2}" -f ($i + 1), $definitions.Count, $definition.Id)
        $settings.Add((Capture-Definition -Definition $definition))
    }

    $snapshot = [PSCustomObject]@{
        SchemaVersion = 1
        Tool          = 'WinDefState'
        ComputerName  = $env:COMPUTERNAME
        CapturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Settings      = @($settings)
    }

    Write-Verbose "[post/4] Persisting snapshot sidecar assets"
    Persist-SnapshotExternalAssets -Snapshot $snapshot -SnapshotPath $fullPath

    Write-Verbose "[post/4] Building snapshot report"
    $reportPath = Get-SnapshotReportPath -SnapshotPath $fullPath
    $reportLines = Get-SnapshotReportLines -Snapshot $snapshot -SnapshotPath $fullPath

    Write-Verbose "[post/4] Writing snapshot JSON"
    Write-SnapshotJsonAtomic -Path $fullPath -Snapshot $snapshot
    Write-Verbose "[post/4] Writing snapshot report"
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

    $definitions = @(Get-DefenseDefinitions)
    $snapshotEntriesById = @{}
    foreach ($entry in @($export.Snapshot.Settings)) {
        $snapshotEntriesById[[string]$entry.Id] = $entry
    }

    for ($i = 0; $i -lt $definitions.Count; $i++) {
        $definition = $definitions[$i]
        Write-Verbose ("[{0}/{1}] Applying permissive setting {2}" -f ($i + 1), $definitions.Count, $definition.Id)
        $entry = if ($snapshotEntriesById.ContainsKey([string]$definition.Id)) { $snapshotEntriesById[[string]$definition.Id] } else { $null }
        if ($null -ne $entry -and -not (Test-SnapshotEntryCapturedExactly -Entry $entry -SnapshotPath $export.JsonPath)) {
            Write-Warning "Skipping permissive change for $($definition.Id) because the baseline capture was incomplete."
            continue
        }

        Apply-PermissiveDefinition -Definition $definition -Entry $entry
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
    $entries = @($snapshot.Settings)
    for ($i = 0; $i -lt $entries.Count; $i++) {
        $entry = $entries[$i]
        Write-Verbose ("[{0}/{1}] Restoring {2}" -f ($i + 1), $entries.Count, $entry.Id)
        Restore-SnapshotEntry -Entry $entry -SnapshotPath $fullPath
    }

    $verification = Test-DefenseSnapshot -Snapshot $snapshot -SnapshotPath $fullPath
    $verificationPath = Get-VerificationReportPath -Root $StateRoot -SnapshotPath $fullPath
    $verificationLines = Get-VerificationReportLines -Verification $verification -SnapshotPath $fullPath
    $wdacVerificationPath = Get-WdacVerificationReportPath -VerificationPath $verificationPath
    $wdacVerificationLines = Get-WdacVerificationReportLines -Verification $verification -SnapshotPath $fullPath
    Write-TextAtomic -Path $verificationPath -Content ($verificationLines -join [Environment]::NewLine)
    Write-TextAtomic -Path $wdacVerificationPath -Content ($wdacVerificationLines -join [Environment]::NewLine)

    if ($verification.MismatchCount -gt 0) {
        $mismatchIds = @($verification.Results | Where-Object { -not $_.Matches -and -not $_.Skipped } | ForEach-Object { $_.Id })
        Write-Warning "Restore verification found $($verification.MismatchCount) mismatched setting(s)."
        Write-Warning "Verification report saved to: $verificationPath"
        Write-Warning "WDAC verification report saved to: $wdacVerificationPath"
        Write-Warning ("Mismatched IDs: {0}" -f ($mismatchIds -join ', '))
        throw 'Restore verification failed. current-operation.json was left in place so you can retry the same snapshot.'
    }

    Clear-OperationState -Root $StateRoot
    Write-Host "Restore completed and verified from snapshot: $fullPath"
    Write-Host "Verification report saved to: $verificationPath"
    Write-Host "WDAC verification report saved to: $wdacVerificationPath"
    if ($verification.SkippedCount -gt 0) {
        $skippedIds = @($verification.Results | Where-Object { $_.Skipped } | ForEach-Object { $_.Id })
        Write-Warning ("Verification skipped {0} setting(s) because the snapshot baseline was incomplete: {1}" -f $verification.SkippedCount, ($skippedIds -join ', '))
    }
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

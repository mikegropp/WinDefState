[CmdletBinding()]
param(
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9._-]*$')]
    [string]$Version = 'local',

    [string]$OutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) 'dist')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Utf8NoBomText {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [AllowEmptyString()] [string]$Content
    )

    [IO.File]::WriteAllText($Path, $Content, [Text.UTF8Encoding]::new($false))
}

function Get-ReleaseHashLine {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Name
    )

    '{0}  {1}' -f ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()), $Name
}

function New-DeterministicZip {
    param(
        [Parameter(Mandatory)] [string]$SourcePath,
        [Parameter(Mandatory)] [string]$DestinationPath,
        [Parameter(Mandatory)] [string]$RootName
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $sourceRoot = [IO.Path]::GetFullPath($SourcePath).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $fixedTimestamp = [DateTimeOffset]::Parse('1980-01-01T00:00:00+00:00')
    $fileStream = [IO.File]::Open($DestinationPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $archive = $null
    try {
        $archive = New-Object IO.Compression.ZipArchive($fileStream, [IO.Compression.ZipArchiveMode]::Create, $true)
        $files = @(Get-ChildItem -LiteralPath $sourceRoot -Recurse -File | Sort-Object -Property FullName)
        foreach ($file in $files) {
            $relativePath = $file.FullName.Substring($sourceRoot.Length).TrimStart([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
            $entryName = '{0}/{1}' -f $RootName, ($relativePath -replace '\\', '/')
            $entry = $archive.CreateEntry($entryName, [IO.Compression.CompressionLevel]::NoCompression)
            $entry.LastWriteTime = $fixedTimestamp

            $inputStream = [IO.File]::OpenRead($file.FullName)
            $outputStream = $null
            try {
                $outputStream = $entry.Open()
                $inputStream.CopyTo($outputStream)
            } finally {
                if ($null -ne $outputStream) {
                    $outputStream.Dispose()
                }
                $inputStream.Dispose()
            }
        }
    } finally {
        if ($null -ne $archive) {
            $archive.Dispose()
        }
        $fileStream.Dispose()
    }
}

$repositoryRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$fullOutputPath = [IO.Path]::GetFullPath($OutputPath)
$packageName = "WinDefState-$Version"
$packagePath = Join-Path $fullOutputPath $packageName
$archivePath = Join-Path $fullOutputPath "$packageName.zip"
$manifestPath = Join-Path $fullOutputPath "$packageName-SHA256SUMS.txt"
$looseEnginePath = Join-Path $fullOutputPath 'WinDefState.ps1'
$looseGuiPath = Join-Path $fullOutputPath 'WinDefState.Gui.ps1'
$inspectionFiles = @('WinDefState.Health.ps1', 'WinDefState.Environment.ps1', 'WinDefState.Inspect.Gui.ps1')
$inspectionTargets = @($inspectionFiles | ForEach-Object { Join-Path $fullOutputPath $_ })
$sourceRelativePaths = @('WinDefState.ps1', 'WinDefState.Gui.ps1', 'README.md', 'docs/ARCHITECTURE.md', 'docs/INSPECTION.md', 'docs/inspection-dashboard.png') + $inspectionFiles

foreach ($sourceRelativePath in $sourceRelativePaths) {
    $sourcePath = Join-Path $repositoryRoot $sourceRelativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Release source file is missing: $sourcePath"
    }
}

foreach ($targetPath in (@($packagePath, $archivePath, $manifestPath, $looseEnginePath, $looseGuiPath) + $inspectionTargets)) {
    if (Test-Path -LiteralPath $targetPath) {
        throw "Release target already exists: $targetPath"
    }
}

New-Item -ItemType Directory -Path $fullOutputPath -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $packagePath 'docs') -Force | Out-Null

foreach ($relativePath in $sourceRelativePaths) {
    Copy-Item -LiteralPath (Join-Path $repositoryRoot $relativePath) -Destination (Join-Path $packagePath $relativePath)
}

$packageHashLines = @(
    foreach ($relativePath in @($sourceRelativePaths | Sort-Object)) {
        Get-ReleaseHashLine -Path (Join-Path $packagePath $relativePath) -Name $relativePath
    }
)
Write-Utf8NoBomText -Path (Join-Path $packagePath 'SHA256SUMS.txt') -Content (($packageHashLines -join "`n") + "`n")

New-DeterministicZip -SourcePath $packagePath -DestinationPath $archivePath -RootName $packageName
Copy-Item -LiteralPath (Join-Path $repositoryRoot 'WinDefState.ps1') -Destination $looseEnginePath
Copy-Item -LiteralPath (Join-Path $repositoryRoot 'WinDefState.Gui.ps1') -Destination $looseGuiPath
foreach ($relativePath in $inspectionFiles) {
    Copy-Item -LiteralPath (Join-Path $repositoryRoot $relativePath) -Destination (Join-Path $fullOutputPath $relativePath)
}

$releaseHashLines = @(
    Get-ReleaseHashLine -Path $looseGuiPath -Name 'WinDefState.Gui.ps1'
    Get-ReleaseHashLine -Path $looseEnginePath -Name 'WinDefState.ps1'
    Get-ReleaseHashLine -Path $archivePath -Name ([IO.Path]::GetFileName($archivePath))
    foreach ($relativePath in $inspectionFiles) {
        Get-ReleaseHashLine -Path (Join-Path $fullOutputPath $relativePath) -Name $relativePath
    }
)
Write-Utf8NoBomText -Path $manifestPath -Content (($releaseHashLines -join "`n") + "`n")

[PSCustomObject]@{
    Version      = $Version
    OutputPath   = $fullOutputPath
    PackagePath  = $packagePath
    ArchivePath  = $archivePath
    ManifestPath = $manifestPath
    EnginePath   = $looseEnginePath
    GuiPath      = $looseGuiPath
}

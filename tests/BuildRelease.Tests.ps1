BeforeAll {
    $repositoryRoot = Split-Path -Parent $PSScriptRoot
    $builderPath = Join-Path $repositoryRoot 'build/Build-Release.ps1'
}

Describe 'Release builder' {
    It 'produces byte-for-byte reproducible archives and valid published checksums' {
        $firstOutput = Join-Path $TestDrive 'first'
        $secondOutput = Join-Path $TestDrive 'second'
        $first = & $builderPath -Version 'v1.2.3-test' -OutputPath $firstOutput
        $second = & $builderPath -Version 'v1.2.3-test' -OutputPath $secondOutput

        (Get-FileHash -LiteralPath $first.ArchivePath -Algorithm SHA256).Hash |
            Should -Be (Get-FileHash -LiteralPath $second.ArchivePath -Algorithm SHA256).Hash
        (Get-Content -LiteralPath $first.ManifestPath -Raw) |
            Should -Be (Get-Content -LiteralPath $second.ManifestPath -Raw)

        $manifestLines = @(Get-Content -LiteralPath $first.ManifestPath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $manifestLines.Count | Should -Be 6
        foreach ($line in $manifestLines) {
            $match = [regex]::Match($line, '^([a-f0-9]{64})  (.+)$')
            $match.Success | Should -BeTrue
            $expectedHash = $match.Groups[1].Value
            $assetName = $match.Groups[2].Value
            $assetPath = Join-Path $first.OutputPath $assetName
            $actualHash = (Get-FileHash -LiteralPath $assetPath -Algorithm SHA256).Hash.ToLowerInvariant()
            $actualHash | Should -Be $expectedHash
        }

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $archive = [IO.Compression.ZipFile]::OpenRead($first.ArchivePath)
        try {
            $entryNames = @($archive.Entries | ForEach-Object { $_.FullName })
            $entryNames | Should -Contain 'WinDefState-v1.2.3-test/WinDefState.ps1'
            $entryNames | Should -Contain 'WinDefState-v1.2.3-test/WinDefState.Gui.ps1'
            $entryNames | Should -Contain 'WinDefState-v1.2.3-test/WinDefState.Health.ps1'
            $entryNames | Should -Contain 'WinDefState-v1.2.3-test/WinDefState.Environment.ps1'
            $entryNames | Should -Contain 'WinDefState-v1.2.3-test/WinDefState.Inspect.Gui.ps1'
            $entryNames | Should -Contain 'WinDefState-v1.2.3-test/docs/INSPECTION.md'
            $entryNames | Should -Contain 'WinDefState-v1.2.3-test/SHA256SUMS.txt'
            # ZIP stores a DOS wall-clock timestamp without a time-zone offset.
            # Reading it in a non-UTC locale must not fail reproducibility checks.
            @($archive.Entries | Where-Object { $_.LastWriteTime.DateTime -ne [datetime]'1980-01-01T00:00:00' }).Count | Should -Be 0
        } finally {
            $archive.Dispose()
        }
    }

    It 'refuses to overwrite an existing release target' {
        $outputPath = Join-Path $TestDrive 'existing'
        $null = & $builderPath -Version 'v1' -OutputPath $outputPath

        { & $builderPath -Version 'v1' -OutputPath $outputPath } |
            Should -Throw '*Release target already exists*'
    }
}

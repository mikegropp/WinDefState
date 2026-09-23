# Development

Use Windows PowerShell 5.1 for the distributed runtime and WPF checks. Pester
contracts and static analysis also run in PowerShell 7 on Linux.

```powershell
Install-Module Pester -RequiredVersion 5.7.1 -Scope CurrentUser
Install-Module PSScriptAnalyzer -RequiredVersion 1.24.0 -Scope CurrentUser

.\tests\Smoke.ps1
.\tests\Analyze.ps1
Import-Module Pester -RequiredVersion 5.7.1
Invoke-Pester -Path .\tests -CI

powershell.exe -NoProfile -Sta -File .\WinDefState.Gui.ps1 -ValidateOnly
powershell.exe -NoProfile -Sta -File .\tests\InspectionGui.ps1
```

`InspectionGui.ps1` uses synthetic data to check tables, search, selection,
comparison, and status counts. It does not capture the host. `ValidateOnly`
constructs a WPF window without opening it or changing system settings.

The manual **Windows Snapshot Integration** workflow captures read-only engine
state on a Windows runner and checks schema, provider caching, and timings.
Hosted Windows Server runners do not replace Windows 10/11 device validation.

## Release build

```powershell
.\build\Build-Release.ps1 -Version local -OutputPath .\dist\local
```

The builder produces a deterministic ZIP, loose scripts, and SHA-256 manifests.
It refuses to overwrite existing targets. The inspection UI requires its
environment and health scripts in the same directory. Tagged releases run the
test gates before publishing.

Checksums detect damaged or mismatched files; they do not authenticate a publisher.
Authenticode signing requires a protected signing identity and is not implemented.

## UI conventions

Use short labels that name the view or action. Keep the toolbar, search, and
selected record visible. Put explanatory detail in the details pane or docs.
Avoid slogans and large summary cards that reduce the table area.

References: [TCPView](https://learn.microsoft.com/en-us/sysinternals/downloads/tcpview)
for compact controls and tabular endpoint data;
[System Informer](https://systeminformer.io/) for category tabs and search;
[Autoruns](https://learn.microsoft.com/en-us/sysinternals/downloads/autoruns)
for filtering categories and inspecting a selected entry.

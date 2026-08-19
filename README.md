# WinDefState

WinDefState is a PowerShell tool for Windows defense testing. It snapshots the current host protection state to disk, switches supported controls to a permissive test posture, and restores the original state from the saved snapshot. When a provider does not yield an exact baseline, WinDefState records that explicitly and skips that setting during permissive, restore, and verification instead of guessing.

> [!WARNING]
> Permissive mode intentionally weakens host protections. Use it only on systems you own or are authorized to test, preferably isolated from production networks and sensitive data.

## What it does

- Saves a disk-backed JSON snapshot before making changes
- Records the exact script SHA-256 and PowerShell runtime that produced each snapshot
- Writes a human-readable text report alongside each JSON snapshot
- Runs an automatic compatibility preflight before each command and saves its diagnostics
- Flags incomplete baselines and platform-managed WDAC policies in the snapshot report
- Writes a `current-operation.json` journal so restore still knows what to do after a crash or power loss
- Checkpoints completed and in-flight restore settings so an interrupted restore can resume conservatively
- Applies a permissive profile for supported controls
- Recaptures every changed provider after permissive apply and writes a permissive-check report
- Restores the original state from the saved snapshot, not from assumptions
- Verifies every restorable setting after restore and records a restore-check report, with explicit classifications for incomplete baselines and inventory-only entries

## Current coverage

- Microsoft Defender real-time monitoring
- Microsoft Defender behavior monitoring
- Microsoft Defender cloud-delivered protection (MAPS)
- Microsoft Defender automatic sample submission
- Microsoft Defender PUA protection
- Microsoft Defender script scanning
- Microsoft Defender IOAV protection
- Microsoft Defender network inspection system setting
- Microsoft Defender network protection
- Microsoft Defender controlled folder access
- Microsoft Defender exclusions and controlled folder access allow/protect lists
- Microsoft Defender running mode and tamper protection status
- Microsoft Defender ASR rule IDs and actions when the provider returns a complete baseline
- Windows Firewall profile state, default actions, notifications, unicast response handling, and logging configuration
- PowerShell machine `__PSLockdownPolicy`
- PowerShell script block logging
- PowerShell module logging
- PowerShell transcription
- AppLocker service state and effective policy export/import with per-collection enforcement summaries when locally restorable
- Print Spooler service state
- Built-in Administrator account state, resolved by the RID 500 account even if renamed
- UAC-related registry settings, including remote local-account token filtering
- RDP allow-connections state
- RDP NLA
- RDP security layer
- RDP minimum encryption level
- RDP listener enable state
- RDP clipboard and drive redirection policy
- RDP firewall rule group state
- RDP Restricted Admin mode
- Windows Script Host
- SmartScreen
- SEHOP
- Exploit protection policy export/import for system and app mitigations
- LSA / Credential Guard / VBS / HVCI / WDigest registry controls
- LAN Manager / NTLM authentication compatibility and minimum client/server session security
- LDAP client signing requirements
- Anonymous SAM/share enumeration and Everyone-token membership controls
- Blank-password remote-use policy and cached domain logon count
- NetBIOS over TCP/IP
- WPAD WinHTTP policy and per-user auto-detect setting across user profiles
- LLMNR / mDNS / telemetry registry controls
- Process creation audit settings
- Print Spooler remote client connection policy
- WinRM service startup state, listeners, core client/service authentication settings, TrustedHosts, and IPv4/IPv6 filters
- SMB client and server signing requirements, insecure guest authentication, and client encryption requirement
- BitLocker protection state, protector inventory, and auto-unlock state for mounted volumes
- WDAC / App Control active Code Integrity policy state and policy files, including inbox or platform-managed policies
- Office macro blocking from the internet across user profiles

## Files on disk

- Snapshots: `%ProgramData%\WinDefState\snapshots\HOST-YYYYMMDD-HHMMSS[-N].json`
- Snapshot sidecar assets: `%ProgramData%\WinDefState\snapshots\HOST-YYYYMMDD-HHMMSS[-N].assets\`
- Snapshot reports: `%ProgramData%\WinDefState\snapshots\HOST-YYYYMMDD-HHMMSS[-N].txt`
- Active run journal: `%ProgramData%\WinDefState\current-operation.json`
- Preflight reports: `%ProgramData%\WinDefState\preflight\HOST-YYYYMMDD-HHMMSS-COMMAND[-N].txt`
- Permissive verification reports: `%ProgramData%\WinDefState\verification\HOST-YYYYMMDD-HHMMSS[-N]-permissive-check-YYYYMMDD-HHMMSS[-N].txt`
- Restore verification reports: `%ProgramData%\WinDefState\verification\HOST-YYYYMMDD-HHMMSS[-N]-restore-check-YYYYMMDD-HHMMSS[-N].txt`
- WDAC-focused restore verification reports: `%ProgramData%\WinDefState\verification\HOST-YYYYMMDD-HHMMSS[-N]-restore-check-YYYYMMDD-HHMMSS[-N]-wdac.txt`

The optional `-N` suffix is added only when an automatically generated name already exists, preventing a second operation in the same timestamp from overwriting the first artifact. The machine-wide default is stable across commit-specific download folders and is restricted to Administrators and SYSTEM before each operation. Use `-StateRoot C:\path\to\state` only with a dedicated directory whose ACL WinDefState is allowed to replace. Snapshots made by older versions beside the script remain usable by passing their old state folder with `-StateRoot` or their JSON file with `-SnapshotPath`.

## Download the current commit

Resolve `main` once, then download from that immutable commit URL. The cache-busting query is used only while resolving the branch; the resulting raw URL is permanently pinned to the returned commit:

```powershell
$repository = 'mikegropp/WinDefState'
$nonce = [guid]::NewGuid().ToString('N')
$headers = @{ 'Cache-Control' = 'no-cache' }
$commit = (Invoke-RestMethod "https://api.github.com/repos/$repository/commits/main?cacheBust=$nonce" -Headers $headers).sha
$dir = Join-Path $env:TEMP "WinDefState-$commit"
$script = Join-Path $dir 'WinDefState.ps1'

New-Item -ItemType Directory -Force -Path $dir | Out-Null
Invoke-WebRequest "https://raw.githubusercontent.com/$repository/$commit/WinDefState.ps1" -Headers $headers -OutFile $script

Write-Host "Downloaded WinDefState commit $commit"
Get-FileHash -LiteralPath $script -Algorithm SHA256
```

Review the downloaded script and its hash before running it. A commit-pinned URL prevents cache drift, but it does not by itself establish that the code is trusted.

## Download a tagged release with checksum verification

Tagged releases publish the primary script, the optional GUI, a deterministic bundle, and a SHA-256 manifest. This example resolves the latest release with a cache-busting request, downloads the script and manifest from that exact tag, and refuses to continue if the bytes do not match:

```powershell
$repository = 'mikegropp/WinDefState'
$nonce = [guid]::NewGuid().ToString('N')
$headers = @{ 'Cache-Control' = 'no-cache' }
$release = Invoke-RestMethod "https://api.github.com/repos/$repository/releases/latest?cacheBust=$nonce" -Headers $headers
$version = [string]$release.tag_name
$dir = Join-Path $env:TEMP "WinDefState-$version"
$script = Join-Path $dir 'WinDefState.ps1'
$manifest = Join-Path $dir "WinDefState-$version-SHA256SUMS.txt"
$baseUrl = "https://github.com/$repository/releases/download/$version"

New-Item -ItemType Directory -Force -Path $dir | Out-Null
Invoke-WebRequest "$baseUrl/WinDefState.ps1?cacheBust=$nonce" -Headers $headers -OutFile $script
Invoke-WebRequest "$baseUrl/WinDefState-$version-SHA256SUMS.txt?cacheBust=$nonce" -Headers $headers -OutFile $manifest

$hashLine = @(Get-Content -LiteralPath $manifest | Where-Object { $_ -match '^[a-f0-9]{64}  WinDefState\.ps1$' })
if ($hashLine.Count -ne 1) {
    throw 'The release manifest does not contain exactly one WinDefState.ps1 checksum.'
}
$expectedHash = ($hashLine[0] -split '\s+', 2)[0]
$actualHash = (Get-FileHash -LiteralPath $script -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualHash -ne $expectedHash) {
    throw "WinDefState.ps1 checksum mismatch. Expected $expectedHash, received $actualHash."
}

Write-Host "Verified WinDefState $version ($actualHash)"
```

The manifest detects corruption or mismatched release assets; it is not a substitute for trusting the repository and release publisher. The tag workflow runs smoke, static-analysis, and Pester gates under Windows PowerShell 5.1 before publishing. Authenticode signing remains a future release requirement once a protected signing identity is available.

## Three-command workflow

The recommended path is one script and three elevated PowerShell calls. Use the call you need at that point in the workflow; do not paste all three unless you intentionally want to snapshot, enter permissive mode, and immediately restore.

```powershell
# Snapshot only: capture current state and write the JSON/report files.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script -Command Snapshot

# Snapshot first, then apply the supported permissive test posture.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script -Command Permissive

# Restore and verify from current-operation.json.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script -Command Restore
```

Each command shows native PowerShell phase progress and then prints a concise result. Add `-Verbose` only when diagnosing provider behavior. Provider command discovery is cached for the process and suppresses module auto-loading export chatter, while useful WinDefState and CIM diagnostics remain visible.

Before command-specific work begins, WinDefState writes a protected preflight report covering PowerShell language mode and runtime, effective execution policy, common pending-reboot markers, operation-journal state, selected scope, and the provider commands required for that scope. Warnings are diagnostic rather than overrides: the normal exact-baseline and verification rules still decide whether a setting can be changed safely.

All three commands support PowerShell's standard `-WhatIf` and `-Confirm` switches. `-WhatIf` previews the top-level operation without capturing or mutating state; it is an operator preview, not a provider validation run.

`Snapshot` prints a concise report summary to the console and saves the full human-readable report to disk next to the JSON snapshot. Use `-ConsoleReport Full` when you intentionally want the entire report in the console.

`Permissive` always writes a fresh snapshot before changing settings. That snapshot is recorded in `%ProgramData%\WinDefState\current-operation.json`, so the plain `Restore` command knows which baseline to use even after downloading a newer script revision.

If a provider command fails, permissive mode stops, marks the journal `ApplyFailed`, and preserves the original snapshot for restore. After all setter calls return, WinDefState recaptures only the settings it attempted and verifies each explicit permissive target. The journal becomes `AppliedVerified` when immediate targets match, `AppliedPendingReboot` when configured targets match but activation requires reboot, or `ApplyVerificationFailed` when a provider is unreadable or Windows reports a different value. Defender tamper protection and higher-precedence policy therefore surface immediately instead of being reported as successful command completion.

The permissive-check report records expected and observed state for mismatches and reboot-pending settings. A verification failure does not auto-restore; it leaves `current-operation.json` and the original baseline intact so the plain `Restore` command remains the deliberate recovery action.

If an active `current-operation.json` already exists, another permissive run is refused. Restore the active baseline first; otherwise a second snapshot could replace the original pre-change restore point.

`Restore` reads the saved snapshot, restores every fully captured setting that has an explicit restore action, verifies that restorable state, and clears `current-operation.json` only after verification succeeds. Inventory-only entries remain visible in reports but are not scheduled as no-op mutations or allowed to create false restore mismatches. A provider exception marks the journal `RestoreFailed` and leaves it available for a retry.

During restore, `current-operation.json` atomically records the attempt number, requested IDs, current work item, completed IDs, and last failure. On retry, every previously completed ID is recaptured first. Only entries that still match the saved baseline are skipped; drifted, unreadable, and in-flight entries are reapplied. Full verification still evaluates the complete requested baseline before the journal can be cleared, and the restore-check report retains the attempt and resume counts after journal cleanup.

Before the first restore mutation, WinDefState loads every sidecar asset needed by the selected entries into an operation-scoped cache. AppLocker XML, exploit-protection XML, and WDAC policy bytes are therefore validated up front and reused consistently instead of being reread during journal validation, planning, mutation, and verification. New snapshots record SHA-256 digests beside each restorable sidecar reference, while older snapshots without those optional fields remain compatible. If plain `Restore` reports that no active operation exists, the previous restore may already have completed; use `-SnapshotPath` only when intentionally restoring a specific saved snapshot again.

Snapshot and verification reports include capture and mutation duration, shared-provider cache hits, and the slowest settings/provider queries or mutations. These measurements are intended to make performance work evidence-driven on the actual target host. User-profile discovery and temporary hive handles are reused within each snapshot, mutation, or verification phase so Office and WPAD work do not repeatedly mount the same unloaded profile. The ten Defender scalar targets are submitted in one preference call, and Defender list/ASR planning shares one phase-local preference read. Seventeen WSMan values collapse to four resource-URI writes inside one bounded WinRM service scope. Firewall mutations batch common profile/rule states, while NetBIOS and BitLocker restore resolve all target identities in one provider query each.

Reports also record build provenance. A snapshot report identifies the SHA-256 of the exact `WinDefState.ps1` file that produced the baseline; permissive and restore verification reports show both the baseline-producing hash and the verifier hash. This is traceability rather than a code-signing trust guarantee, but it makes commit-specific downloads and cross-version restores auditable.

You can also restore from a specific snapshot file:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Restore -SnapshotPath .\state\snapshots\HOST-20260420-120000.json
```

Restore validates the snapshot schema, setting IDs and immutable targets before making changes. It also rejects a snapshot captured on another computer by default; add `-AllowDifferentComputer` only when a cross-host restore is deliberate and you have reviewed machine-specific entries such as user SIDs, volumes, adapters, and local policy files.

## Optional scoped mode

You can target exact setting IDs, wildcard ID families, or validated categories from the CLI. Category names are the stable prefix before the first dot in a setting ID, such as `defender`, `firewall`, `rdp`, `winrm`, `bitlocker`, or `wdac`:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Permissive -IncludeId defender.enable_network_protection,rdp.user_authentication
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Snapshot -IncludeId 'rdp.*','winrm.*'
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Snapshot -IncludeCategory defender,firewall
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Snapshot -ExcludeCategory bitlocker,wdac,applocker,exploit_protection,office,wpad
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Restore -SnapshotPath .\state\snapshots\HOST-20260420-120000.json -IncludeId rdp.user_authentication
```

Include filters select matching settings and exclude filters subtract from that result. Every ID pattern and category is validated before capture or mutation, so a typo fails instead of silently producing an incomplete scope. Filtered snapshot and permissive operations capture only the selected definitions, allowing faster runs that avoid unrelated providers such as BitLocker, WDAC, AppLocker, exploit protection, or unloaded user profiles. The resulting schema-2 snapshot records its resolved include/exclude filters and a plain later `Restore` restores every entry in that scoped baseline.

The GUI continues to use exact selected IDs; wildcard and category scopes are optional CLI conveniences and do not change the normal three-command workflow.

The native PowerShell/WPF GUI is optional. Its primary runbook keeps Snapshot, Snapshot + permissive, and Restore visually distinct, while the snapshot explorer adds tokenized search, category filtering, live visible/runnable/reboot/selected counts, a virtualized setting grid, and a resizable bounded operation console. Snapshot history is loaded from the stable state root and can compare the loaded baseline with any other readable snapshot. The grid marks changed, added, and removed settings with distinct visual states, exposes both values side by side, and can filter to differences only; comparison-only rows are always non-runnable.

GUI-triggered permissive and restore operations now stop at a pre-change review gate. Permissive first captures and persists the real fresh baseline, then presents the complete runnable scope with baseline value, engine-authored target description, exact restore value, and reboot/high-impact labels. Restore validates the requested snapshot and selected sidecars before presenting the same review. Only an explicit one-time approval marker inside the protected state root lets the child process continue; closing or cancelling the dialog exits before the journal or any live defense setting is changed. Cancelling permissive review intentionally leaves the newly captured read-only snapshot in history. The three direct CLI commands are unchanged and do not require this GUI handshake.

The GUI keeps the window responsive while the engine runs, streams structured engine progress, and shows both current phase and active restore-point state without enabling noisy global verbosity. `Cancel safely` is cooperative: it is available during read-only snapshot capture and through permissive snapshot persistence, and the engine acknowledges the request before the GUI reports cancellation. It never kills the child process and is disabled when mutation begins. The grid remains available for advanced selected-ID operations, but it is not required for the normal snapshot/permissive/restore workflow. Incomplete and inventory-only rows are visibly disabled, while exact Defender exclusion and Controlled Folder Access lists are offered only for restoring their captured values. `Select permissive` deliberately leaves those restore-only rows clear so a bulk selection cannot mix incompatible actions. New snapshots persist target descriptions for the review window; legacy snapshots remain restorable and receive conservative provider-level descriptions. Windows CI instantiates both complete WPF visual trees through the noninteractive `-ValidateOnly` path before test and release gates:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File .\WinDefState.Gui.ps1
```

## Important caveats

- ASR coverage is only treated as exact when Windows returns a complete rule ID and action baseline. If the provider returns blank or malformed entries, the snapshot report marks ASR as `Partial / incomplete`, and permissive, restore, and verification skip ASR rather than guessing.
- WDAC / App Control reports distinguish inbox or platform-managed policies, such as Smart App Control-related policies, from ordinary custom policy state. Permissive mode never requests removal of platform-managed policies. It also fails closed instead of deleting raw policy files when `CiTool` cannot safely remove a custom policy.
- WDAC reporting also distinguishes active enforcement from on-disk file presence. Policies that are present only on disk are called out as pending-reboot state so it is easier to tell whether restore wrote the files but Windows has not enforced them yet.
- BitLocker capture is bounded to mounted volumes. The snapshot report shows timed-out mount points, capture issues, protector inventory, and auto-unlock state for each captured volume, and it marks BitLocker `Partial / incomplete` if the richer mounted-volume baseline cannot be captured exactly.
- Exploit protection verification normalizes the exported XML before comparison, including the default no-op `SystemConfig` ASLR block that Windows can add during export after restore.
- User-scoped registry capture loads `NTUSER.DAT` for unloaded profiles when possible. If a profile hive cannot be mounted, the affected setting is marked `Partial / incomplete`, and permissive, restore, and verification skip it rather than guessing.
- WinRM configuration capture relies on the local WSMan management interfaces. If those interfaces cannot return an exact baseline, the affected WinRM entries are marked `Partial / incomplete`, and permissive, restore, and verification skip them rather than guessing.
- Firewall profile capture relies on `Get-NetFirewallProfile` returning each requested profile. If a profile cannot be captured exactly, the firewall entry is marked `Partial / incomplete`, and permissive, restore, and verification skip it rather than guessing.
- RDP firewall-rule capture relies on the NetSecurity cmdlets returning the Remote Desktop firewall group. If that group cannot be captured exactly, the affected RDP firewall entry is marked `Partial / incomplete`, and permissive, restore, and verification skip it rather than guessing.
- AppLocker capture records the effective policy and per-collection enforcement summaries, but local restore is only treated as exact when the local and effective AppLocker policies match. If Group Policy or another higher-precedence source changes the effective AppLocker policy, the AppLocker policy entry is marked `Partial / incomplete`, and permissive, restore, and verification skip it rather than guessing.
- Defender snapshot reports capture the local Defender running mode and tamper protection state as inventory-only context. Those runtime observations are not restore targets. When tamper protection is enabled, local Defender preference changes can appear to succeed while being ignored or later reverted, so interpret restorable Defender preference mismatches in that context.

## Safety model

- Snapshot is written to disk before permissive changes begin
- A machine-wide operation lock prevents overlapping snapshot, permissive, or restore processes from racing the same host state and journal
- GUI cancellation is marker-based and honored only at explicit no-mutation boundaries; restore and in-progress mutation are never terminated
- GUI mutation approval is a separate one-time marker under the protected state root; permissive review occurs after baseline persistence but before journaling, and restore review occurs after snapshot/sidecar preflight but before restore status or mutation
- A human-readable snapshot report is written to disk with the JSON snapshot
- Restore reads from the saved JSON snapshot on disk
- `%ProgramData%\WinDefState\current-operation.json` records which snapshot should be used if the system loses power during testing or the script is updated
- `current-operation.json` records operation status, selected setting scope, and SHA-256 hashes for the snapshot and sidecar assets
- Restore checkpoints record requested, in-flight, completed, revalidated, and failed setting IDs after each work item
- Snapshots and `current-operation.json` record the producing script SHA-256 and PowerShell runtime
- New snapshots persist engine-authored permissive-target descriptions for review; this presentation metadata is excluded from history drift comparison and does not alter restore semantics
- `current-operation.json` records permissive verification counts, mismatched IDs, pending-reboot IDs, and the verification report path
- A second permissive operation cannot replace an active journal or its original baseline
- Permissive recaptures only attempted mutations and distinguishes verified configuration, reboot-pending activation, and mismatches before reporting completion
- Restore validates snapshot structure and immutable setting targets before the first mutation
- Restore writes a verification report and only clears `current-operation.json` after every fully captured restorable setting matches; inventory-only entries are classified separately
- Reboot-required settings are still captured and restored, but some changes do not fully take effect until reboot
- Some user-scoped settings, such as Office macro policy and WPAD auto-detect, are captured from loaded and unloaded user hives when the profile hive can be mounted
- BitLocker snapshot is bounded to mounted volumes and skips a mount point if the provider does not return in time
- BitLocker volume probes run concurrently with independent per-volume timeouts, avoiding serial `powershell.exe` startup and wait cost on multi-volume hosts
- BitLocker permissive mode suspends protectors on currently protected mounted volumes, can temporarily enable auto-unlock on supported data volumes, and restore returns both protection state and auto-unlock state to the captured baseline
- Snapshot records incomplete captures, and permissive/restore skips any setting whose baseline could not be captured exactly
- Snapshot and verification reports summarize incomplete-baseline settings so skipped controls are visible instead of silent
- WinRM mutations acquire at most one temporary service-write scope per phase; filtered operations preserve the existing service state, and restore applies the captured WinRM service baseline only after WSMan/listener writes finish
- Large policy payloads such as WDAC policy files, AppLocker policy XML, and exploit protection XML are stored as sidecar snapshot assets instead of inline in the main JSON snapshot
- Snapshot reports summarize platform-managed WDAC policy count and label those policies inline for operator awareness
- WDAC permissive removal skips platform-managed policies and requires `CiTool` for safe custom-policy removal; it never falls back to indiscriminate raw file deletion
- WDAC restore reconciles identified custom policies against the baseline and never clears the policy directory wholesale; platform-managed and unclassified files are preserved
- WDAC restore uses active OS Code Integrity policy files and `CiTool` when available; on older hosts a reboot can still be required before the live state fully matches the restored snapshot

## Development

The engine can be dot-sourced without executing a command, which allows its internal capture contracts to be tested safely:

```powershell
Install-Module PSScriptAnalyzer -RequiredVersion 1.24.0 -Scope CurrentUser
.\tests\Analyze.ps1
Invoke-Pester -Path .\tests
```

CI runs focused static analysis, the dependency-free smoke suite, and Pester contracts on both Linux PowerShell and Windows PowerShell 5.1. The analyzer blocks automatic-variable collisions, empty exception handlers, suspicious assignments/comparisons, and syntax newer than Windows PowerShell 5.1 without imposing module-style naming rules on the internal single-file engine. The Windows lane remains the compatibility authority for the distributed script. See `docs\ARCHITECTURE.md` for the compatibility contract, performance model, and staged refactor plan.

The manually triggered `Windows Snapshot Integration` workflow performs a read-only elevated snapshot on a Windows runner, validates the 94-entry schema, one-to-one setting timing coverage, exact shared Defender/service reads, and provider-cache reuse, then uploads the snapshot, reports, and a compact `performance-summary.json` artifact for review. The summary separates end-to-end command time, provider-capture time, and non-capture orchestration/persistence/reporting overhead. It never invokes permissive or restore mode.

## Next improvements

- Extract provider implementations from the single-file distribution build
- Add Authenticode signatures using a protected release signing identity
- Collect native Windows timing baselines and optimize the remaining slow providers in measured order
- Capture native Windows screenshots and accessibility feedback for final WPF spacing and contrast tuning

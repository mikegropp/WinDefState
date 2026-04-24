# WinDefState

WinDefState is a PowerShell tool for Windows defense testing. It snapshots the current host protection state to disk, switches supported controls to a permissive test posture, and restores the original state from the saved snapshot. When a provider does not yield an exact baseline, WinDefState records that explicitly and skips that setting during permissive, restore, and verification instead of guessing.

## What it does

- Saves a disk-backed JSON snapshot before making changes
- Writes a human-readable text report alongside each JSON snapshot
- Flags incomplete baselines and platform-managed WDAC policies in the snapshot report
- Writes a `current-operation.json` journal so restore still knows what to do after a crash or power loss
- Applies a permissive profile for supported controls
- Restores the original state from the saved snapshot, not from assumptions
- Verifies the live system after restore and records a restore-check report, with explicit skips for settings whose baseline was incomplete

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
- UAC-related registry settings
- RDP allow-connections state
- RDP NLA
- RDP security layer
- RDP listener enable state
- RDP firewall rule group state
- RDP Restricted Admin mode
- Windows Script Host
- SmartScreen
- SEHOP
- Exploit protection policy export/import for system and app mitigations
- LSA / Credential Guard / VBS / HVCI / WDigest registry controls
- NetBIOS over TCP/IP
- WPAD WinHTTP policy and per-user auto-detect setting across user profiles
- LLMNR / mDNS / telemetry registry controls
- Process creation audit settings
- Print Spooler remote client connection policy
- WinRM service startup state, listeners, core client/service authentication settings, TrustedHosts, and IPv4/IPv6 filters
- SMB client and server signing requirements
- BitLocker protection state, protector inventory, and auto-unlock state for mounted volumes
- WDAC / App Control active Code Integrity policy state and policy files, including inbox or platform-managed policies
- Office macro blocking from the internet across user profiles

## Files on disk

- Snapshots: `.\state\snapshots\HOST-YYYYMMDD-HHMMSS.json`
- Snapshot sidecar assets: `.\state\snapshots\HOST-YYYYMMDD-HHMMSS.assets\`
- Snapshot reports: `.\state\snapshots\HOST-YYYYMMDD-HHMMSS.txt`
- Active run journal: `.\state\current-operation.json`
- Restore verification reports: `.\state\verification\HOST-YYYYMMDD-HHMMSS-restore-check-YYYYMMDD-HHMMSS.txt`
- WDAC-focused restore verification reports: `.\state\verification\HOST-YYYYMMDD-HHMMSS-restore-check-YYYYMMDD-HHMMSS-wdac.txt`

## Three-command workflow

The recommended path is one script and three elevated PowerShell calls:

```powershell
# Snapshot only: capture current state and write the JSON/report files.
.\WinDefState.ps1 -Command Snapshot

# Snapshot first, then apply the supported permissive test posture.
.\WinDefState.ps1 -Command Permissive

# Restore and verify from current-operation.json.
.\WinDefState.ps1 -Command Restore
```

`Snapshot` prints the human-readable report to the console and saves the same report to disk next to the JSON snapshot.

`Permissive` always writes a fresh snapshot before changing settings. That snapshot is recorded in `.\state\current-operation.json`, so the plain `Restore` command knows which baseline to use.

`Restore` reads the saved snapshot, restores every fully captured setting, verifies live state, and clears `current-operation.json` only after verification succeeds.

You can also restore from a specific snapshot file:

```powershell
.\WinDefState.ps1 -Command Restore -SnapshotPath .\state\snapshots\HOST-20260420-120000.json
```

## Optional selected-item mode

You can target specific setting IDs from the CLI, which is also how the GUI applies selected rows:

```powershell
.\WinDefState.ps1 -Command Permissive -IncludeId defender.enable_network_protection,rdp.user_authentication
.\WinDefState.ps1 -Command Restore -SnapshotPath .\state\snapshots\HOST-20260420-120000.json -IncludeId rdp.user_authentication
```

The native PowerShell/WPF GUI is optional. It is useful when you want to inspect a snapshot and run selected IDs, but it is not required for the normal snapshot/permissive/restore workflow:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File .\WinDefState.Gui.ps1
```

## Important caveats

- ASR coverage is only treated as exact when Windows returns a complete rule ID and action baseline. If the provider returns blank or malformed entries, the snapshot report marks ASR as `Partial / incomplete`, and permissive, restore, and verification skip ASR rather than guessing.
- WDAC / App Control reports distinguish inbox or platform-managed policies, such as Smart App Control-related policies, from ordinary custom policy state. These are still captured for awareness, but restore can require file-based handling and sometimes a reboot before live state fully matches the snapshot.
- WDAC reporting also distinguishes active enforcement from on-disk file presence. Policies that are present only on disk are called out as pending-reboot state so it is easier to tell whether restore wrote the files but Windows has not enforced them yet.
- BitLocker capture is bounded to mounted volumes. The snapshot report shows timed-out mount points, capture issues, protector inventory, and auto-unlock state for each captured volume, and it marks BitLocker `Partial / incomplete` if the richer mounted-volume baseline cannot be captured exactly.
- Exploit protection verification normalizes the exported XML before comparison, including the default no-op `SystemConfig` ASLR block that Windows can add during export after restore.
- User-scoped registry capture loads `NTUSER.DAT` for unloaded profiles when possible. If a profile hive cannot be mounted, the affected setting is marked `Partial / incomplete`, and permissive, restore, and verification skip it rather than guessing.
- WinRM configuration capture relies on the local WSMan management interfaces. If those interfaces cannot return an exact baseline, the affected WinRM entries are marked `Partial / incomplete`, and permissive, restore, and verification skip them rather than guessing.
- Firewall profile capture relies on `Get-NetFirewallProfile` returning each requested profile. If a profile cannot be captured exactly, the firewall entry is marked `Partial / incomplete`, and permissive, restore, and verification skip it rather than guessing.
- RDP firewall-rule capture relies on the NetSecurity cmdlets returning the Remote Desktop firewall group. If that group cannot be captured exactly, the affected RDP firewall entry is marked `Partial / incomplete`, and permissive, restore, and verification skip it rather than guessing.
- AppLocker capture records the effective policy and per-collection enforcement summaries, but local restore is only treated as exact when the local and effective AppLocker policies match. If Group Policy or another higher-precedence source changes the effective AppLocker policy, the AppLocker policy entry is marked `Partial / incomplete`, and permissive, restore, and verification skip it rather than guessing.
- Defender snapshot reports capture the local Defender running mode and tamper protection state. When tamper protection is enabled, local Defender preference changes can appear to succeed while being ignored or later reverted, so interpret Defender mismatches in that context.

## Safety model

- Snapshot is written to disk before permissive changes begin
- A human-readable snapshot report is written to disk with the JSON snapshot
- Restore reads from the saved JSON snapshot on disk
- `current-operation.json` records which snapshot should be used if the system loses power during testing
- Restore writes a verification report and only clears `current-operation.json` after the live state matches all fully captured settings from the target snapshot
- Reboot-required settings are still captured and restored, but some changes do not fully take effect until reboot
- Some user-scoped settings, such as Office macro policy and WPAD auto-detect, are captured from loaded and unloaded user hives when the profile hive can be mounted
- BitLocker snapshot is bounded to mounted volumes and skips a mount point if the provider does not return in time
- BitLocker permissive mode suspends protectors on currently protected mounted volumes, can temporarily enable auto-unlock on supported data volumes, and restore returns both protection state and auto-unlock state to the captured baseline
- Snapshot records incomplete captures, and permissive/restore skips any setting whose baseline could not be captured exactly
- Snapshot and verification reports summarize incomplete-baseline settings so skipped controls are visible instead of silent
- Large policy payloads such as WDAC policy files, AppLocker policy XML, and exploit protection XML are stored as sidecar snapshot assets instead of inline in the main JSON snapshot
- Snapshot reports summarize platform-managed WDAC policy count and label those policies inline for operator awareness
- WDAC restore uses active OS Code Integrity policy files and `CiTool` when available; on older hosts a reboot can still be required before the live state fully matches the restored snapshot

## Next improvements

- Split output into immediate changes and reboot-pending changes
- Expose the restore verification logic as a standalone `Test-DefenseDrift` command
- Expand provider coverage where exact capture and exact restore are both supported

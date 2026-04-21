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
- Microsoft Defender ASR rule IDs and actions when the provider returns a complete baseline
- Windows Firewall profile state
- PowerShell machine `__PSLockdownPolicy`
- PowerShell script block logging
- PowerShell module logging
- PowerShell transcription
- AppLocker service state
- Print Spooler service state
- Built-in Administrator account state, resolved by the RID 500 account even if renamed
- UAC-related registry settings
- RDP NLA
- RDP Restricted Admin mode
- Windows Script Host
- SmartScreen
- SEHOP
- Exploit protection policy export/import for system and app mitigations
- LSA / Credential Guard / VBS / HVCI / WDigest registry controls
- NetBIOS over TCP/IP
- WPAD WinHTTP policy and loaded-user auto-detect setting
- LLMNR / mDNS / telemetry registry controls
- Process creation audit settings
- Print Spooler remote client connection policy
- WinRM basic and unencrypted settings
- SMB client and server signing requirements
- BitLocker protection status for mounted volumes
- WDAC / App Control active Code Integrity policy state and policy files, including inbox or platform-managed policies
- Office macro blocking from the internet for loaded user hives

## Files on disk

- Snapshots: `.\state\snapshots\HOST-YYYYMMDD-HHMMSS.json`
- Snapshot sidecar assets: `.\state\snapshots\HOST-YYYYMMDD-HHMMSS.assets\`
- Snapshot reports: `.\state\snapshots\HOST-YYYYMMDD-HHMMSS.txt`
- Active run journal: `.\state\current-operation.json`
- Restore verification reports: `.\state\verification\HOST-YYYYMMDD-HHMMSS-restore-check-YYYYMMDD-HHMMSS.txt`

## Usage

Run from an elevated PowerShell session.

```powershell
.\WinDefState.ps1 -Command Snapshot
.\WinDefState.ps1 -Command Permissive
.\WinDefState.ps1 -Command Restore
```

`Snapshot` prints the human-readable report to the console and saves the same report to disk next to the JSON snapshot.

You can also restore from a specific snapshot file:

```powershell
.\WinDefState.ps1 -Command Restore -SnapshotPath .\state\snapshots\HOST-20260420-120000.json
```

## Important caveats

- ASR coverage is only treated as exact when Windows returns a complete rule ID and action baseline. If the provider returns blank or malformed entries, the snapshot report marks ASR as `Partial / incomplete`, and permissive, restore, and verification skip ASR rather than guessing.
- WDAC / App Control reports distinguish inbox or platform-managed policies, such as Smart App Control-related policies, from ordinary custom policy state. These are still captured for awareness, but restore can require file-based handling and sometimes a reboot before live state fully matches the snapshot.
- BitLocker capture is bounded to mounted volumes. The snapshot report shows timed-out mount point count and captured volume count so it is obvious when coverage was partial.
- Exploit protection verification normalizes the exported XML before comparison, including the default no-op `SystemConfig` ASLR block that Windows can add during export after restore.

## Safety model

- Snapshot is written to disk before permissive changes begin
- A human-readable snapshot report is written to disk with the JSON snapshot
- Restore reads from the saved JSON snapshot on disk
- `current-operation.json` records which snapshot should be used if the system loses power during testing
- Restore writes a verification report and only clears `current-operation.json` after the live state matches all fully captured settings from the target snapshot
- Reboot-required settings are still captured and restored, but some changes do not fully take effect until reboot
- Some user-scoped settings, such as Office macro policy and WPAD auto-detect, are currently captured from loaded user hives
- BitLocker snapshot is bounded to mounted volumes and skips a mount point if the provider does not return in time
- BitLocker permissive mode suspends protectors on currently protected mounted volumes and restore resumes them based on the captured protection state
- Snapshot records incomplete captures, and permissive/restore skips any setting whose baseline could not be captured exactly
- Snapshot and verification reports summarize incomplete-baseline settings so skipped controls are visible instead of silent
- Large policy payloads such as WDAC policy files and exploit protection XML are stored as sidecar snapshot assets instead of inline in the main JSON snapshot
- Snapshot reports summarize platform-managed WDAC policy count and label those policies inline for operator awareness
- WDAC restore uses active OS Code Integrity policy files and `CiTool` when available; on older hosts a reboot can still be required before the live state fully matches the restored snapshot

## Next improvements

- Split output into immediate changes and reboot-pending changes
- Expose the restore verification logic as a standalone `Test-DefenseDrift` command
- Expand provider coverage where exact capture and exact restore are both supported

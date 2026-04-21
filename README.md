# WinDefState

WinDefState is a PowerShell tool for Windows defense testing. It snapshots the current host protection state to disk, switches supported controls to a permissive test posture, and restores the exact original state from the saved snapshot.

## What it does

- Saves a disk-backed JSON snapshot before making changes
- Writes a human-readable text report alongside each JSON snapshot
- Writes a `current-operation.json` journal so restore still knows what to do after a crash or power loss
- Applies a permissive profile for supported controls
- Restores the original state from the saved snapshot, not from assumptions
- Verifies the live system after restore and records a restore-check report

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
- Microsoft Defender ASR rule IDs and actions
- Windows Firewall profile state
- PowerShell machine `__PSLockdownPolicy`
- PowerShell script block logging
- PowerShell module logging
- PowerShell transcription
- AppLocker service state
- Print Spooler service state
- Built-in Administrator account state, tracked by RID-500 SID even if renamed
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
- BitLocker protection status per volume
- WDAC / App Control active Code Integrity policy state
- Office macro blocking from the internet for loaded user hives

## Files on disk

- Snapshots: `.\state\snapshots\HOST-YYYYMMDD-HHMMSS.json`
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

## Safety model

- Snapshot is written to disk before permissive changes begin
- A human-readable snapshot report is written to disk with the JSON snapshot
- Restore reads from the saved JSON snapshot on disk
- `current-operation.json` records which snapshot should be used if the system loses power during testing
- Restore writes a verification report and only clears `current-operation.json` after the live state matches the target snapshot
- Reboot-required settings are still captured and restored, but some changes do not fully take effect until reboot
- Some user-scoped settings, such as Office macro policy and WPAD auto-detect, are currently captured from loaded user hives
- BitLocker permissive mode suspends protectors on currently protected volumes and restore resumes them based on the captured protection state
- WDAC restore uses active OS Code Integrity policy files and `CiTool` when available; on older hosts a reboot can still be required before the live state fully matches the restored snapshot

## Next improvements

- Split output into immediate changes and reboot-pending changes
- Expose the restore verification logic as a standalone `Test-DefenseDrift` command
- Expand provider coverage where exact capture and exact restore are both supported

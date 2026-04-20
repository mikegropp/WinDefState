# WinDefState

WinDefState is a PowerShell tool for Windows defense testing. It snapshots the current host protection state to disk, switches supported controls to a permissive test posture, and restores the exact original state from the saved snapshot.

## What it does

- Saves a disk-backed JSON snapshot before making changes
- Writes a `current-operation.json` journal so restore still knows what to do after a crash or power loss
- Applies a permissive profile for supported controls
- Restores the original state from the saved snapshot, not from assumptions

## Current coverage

- Microsoft Defender real-time monitoring
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
- Built-in Administrator account state
- UAC-related registry settings
- RDP NLA
- RDP Restricted Admin mode
- Windows Script Host
- SmartScreen
- SEHOP
- LSA / Credential Guard / VBS / HVCI / WDigest registry controls
- NetBIOS over TCP/IP
- WPAD WinHTTP policy and loaded-user auto-detect setting
- LLMNR / mDNS / telemetry registry controls
- Process creation audit settings
- Print Spooler remote client connection policy
- WinRM basic and unencrypted settings
- SMB client and server signing requirements
- Office macro blocking from the internet for loaded user hives

## Files on disk

- Snapshots: `.\state\snapshots\HOST-YYYYMMDD-HHMMSS.json`
- Active run journal: `.\state\current-operation.json`

## Usage

Run from an elevated PowerShell session.

```powershell
.\WinDefState.ps1 -Command Snapshot
.\WinDefState.ps1 -Command Permissive
.\WinDefState.ps1 -Command Restore
```

You can also restore from a specific snapshot file:

```powershell
.\WinDefState.ps1 -Command Restore -SnapshotPath .\state\snapshots\HOST-20260420-120000.json
```

## Safety model

- Snapshot is written to disk before permissive changes begin
- Restore reads from the saved JSON snapshot on disk
- `current-operation.json` records which snapshot should be used if the system loses power during testing
- Reboot-required settings are still captured and restored, but some changes do not fully take effect until reboot
- Some user-scoped settings, such as Office macro policy and WPAD auto-detect, are currently captured from loaded user hives

## Next improvements

- Split output into immediate changes and reboot-pending changes
- Add `Test-DefenseDrift` to compare live state to a saved snapshot
- Expand provider coverage where exact capture and exact restore are both supported

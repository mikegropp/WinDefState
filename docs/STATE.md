# Protection-state reference

The engine and inspection tool produce different files:

| Tool | Output | Purpose |
| --- | --- | --- |
| `WinDefState.Environment.ps1` | Environment baseline | Review and compare inventory |
| `WinDefState.ps1` | Protection snapshot and sidecar assets | Preserve settings for verified restoration |

Neither format is a disk image. Environment baselines cannot be used as engine
restore snapshots.

## Coverage

The engine covers Defender preferences and ASR, firewall profiles, PowerShell
logging, application and exploit protection, selected Windows services and
accounts, UAC/RDP, authentication and networking policy, WinRM/SMB, mounted
BitLocker volumes, and selected Office/user-profile settings. Runtime observations
such as Defender mode and tamper protection are inventory, not restore targets.

## Saved files

The default engine state directory is `%ProgramData%\WinDefState`.

| Path | Contents |
| --- | --- |
| `snapshots\*.json` | Captured settings and producing script/runtime metadata |
| `snapshots\*.assets\` | Policy XML and other sidecar assets |
| `snapshots\*.txt` | Human-readable snapshot reports |
| `current-operation.json` | Active baseline, integrity hashes, and recovery checkpoints |
| `preflight\*.txt` | Runtime, provider, journal, and restart diagnostics |
| `verification\*.txt` | Recorded verification results |

Keep JSON snapshots and their matching asset directories together. The state
directory is restricted to Administrators and SYSTEM. An explicit `StateRoot`
must be a dedicated directory whose permissions the engine may replace.

## Recovery contract

Changes require a persisted baseline. An active journal prevents a second test
operation from replacing the original restore point. Restore validates the
snapshot, host identity, and required assets before making changes, then verifies
the result. A failed restore leaves its journal available for retry. Completed
checkpoints are rechecked before they can be skipped.

An incomplete capture is not a usable restore target. Reports distinguish
incomplete, inventory-only, mismatched, and reboot-pending entries. Read the
verification report before treating a restore as complete.

## Provider limits

- Defender tamper protection and managed policy can prevent local changes.
- Controlled Folder Access protected-folder and allowed-app preference lists
  contain additional folders and explicit app allowances. The default protected
  folders and automatically trusted apps are separate from these lists. Empty
  lists do not show whether CFA protects a default folder or trusts an app; read
  the CFA mode and validate the action when impact matters. See [Microsoft's CFA
  overview](https://learn.microsoft.com/en-us/defender-endpoint/controlled-folder-access-overview).
- ASR accepts complete rule ID/action pairs or two null or empty arrays as a
  complete zero-rule baseline. Blank/malformed entries and mismatched pairs
  remain incomplete.
- Defender runtime health, signature fields, and MAPS preference are local
  evidence. They do not verify Microsoft cloud access or show how Network
  Protection, PUA, or SmartScreen handles a particular site or file. Network
  inspection (NIS) is separate from Network Protection.
- `wpad.disable_wpad` records the machine WinHTTP `DisableWpad` DWORD. An absent
  value means that explicit WinHTTP disable is not configured; it does not show
  a PAC lookup or traffic redirection. `wpad.user_auto_detect` records
  profile-scoped WinINet `AutoDetect` DWORDs, not effective connection flags.
  An absent DWORD does not prove auto-detection is off. Query the signed-in
  user's WinINet connection flags (`INTERNET_PER_CONN_FLAGS_UI`) to assess the
  effective setting; a result for one user does not cover other profiles.
  WinHTTP's registry control does not disable discovery for all browsers or
  applications. See Microsoft's [WinHTTP WPAD guidance](https://learn.microsoft.com/en-us/troubleshoot/windows-server/networking/disable-http-proxy-auth-features)
  and [WinINet connection-option reference](https://learn.microsoft.com/en-us/windows/win32/api/wininet/ns-wininet-internet_per_conn_optiona).
- LLMNR, mDNS, and NetBIOS entries describe configured policy or adapter
  options. An absent policy value is not proof that the protocol is disabled,
  and a configured value does not prove an attacker can redirect a request.
  Confirm behavior with a scoped name lookup and packet observation in the
  relevant user and network context.
- Windows shell SmartScreen and Microsoft Edge mandatory machine SmartScreen
  policies are separate registry entries. An absent value is unconfigured,
  not disabled. If Edge's mandatory value is absent, its recommended policy
  and user choices may still apply.
  Edge `HKLM` `SmartScreenEnabled=1` means a configured machine registry value,
  not proof that Edge accepted the policy or blocked a site. Check applied
  policy at `edge://policy` in the intended interactive Edge session and note
  the device's management eligibility. Record any "Ignored" status from a
  headless check as a separate observation, not an interactive-user verdict.
  See Microsoft's [Edge SmartScreen policy](https://learn.microsoft.com/en-us/deployedge/microsoft-edge-policies/smartscreenenabled)
  and [applied-policy verification](https://learn.microsoft.com/en-us/deployedge/configure-microsoft-edge).
- Current-boot LSA protection is confirmed only when a matching Wininit event 12
  is observed; absence or unreadability is unknown, not disabled.
- Windows LAPS inventory reads policy and recent event metadata only. It does
  not retrieve passwords or prove password uniqueness across devices.
- Windows Time inventory records service, sync status, and configured peers.
  Clock offset requires an independent trusted time reference. Defender
  signature ages and local event timestamps depend on the host clock, so
  interpret them cautiously when this status is unsynchronized.
- Firewall profile packet-log flags and the Filtering Platform Connection audit
  subcategory are captured separately. Disabled native logging does not prove
  that an EDR, a collector, or network telemetry lacks connection records.
- AppLocker is locally restorable only when local and effective policy agree.
- WDAC distinguishes platform-managed policies, custom policies, and files that
  are not yet active. Activation can require a restart.
- BitLocker capture covers mounted volumes, with bounded provider reads.
- Unreadable user hives, firewall profiles, RDP rules, or WSMan resources are
  recorded as incomplete rather than reconstructed from assumed defaults.

The engine's permissive mode intentionally weakens protection. Use isolated,
authorized test systems and retain independent recovery media.

See [architecture](ARCHITECTURE.md) for validation, checkpoint, and provider contracts.

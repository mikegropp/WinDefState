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
- ASR requires complete rule ID/action pairs; malformed baselines are incomplete.
- AppLocker is locally restorable only when local and effective policy agree.
- WDAC distinguishes platform-managed policies, custom policies, and files that
  are not yet active. Activation can require a restart.
- BitLocker capture covers mounted volumes, with bounded provider reads.
- Unreadable user hives, firewall profiles, RDP rules, or WSMan resources are
  recorded as incomplete rather than reconstructed from assumed defaults.

The engine's permissive mode intentionally weakens protection. Use isolated,
authorized test systems and retain independent recovery media.

See [architecture](ARCHITECTURE.md) for validation, checkpoint, and provider contracts.

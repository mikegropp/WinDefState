# Windows LAPS inventory

`laps.policy_status` is an inventory-only snapshot entry. Snapshot reads registry policy and event metadata; permissive and restore never modify LAPS.

## What is captured

Windows LAPS checks four policy roots in order: CSP, Group Policy, local configuration, and legacy Microsoft LAPS. WinDefState records whether each root exists, whether it was readable, its count of recognized settings, its count of unrecognized values, and the values of documented policy settings only. A root that exists but has no recognized settings does not become the active policy. The first root with a recognized explicit setting wins; settings from lower roots are not merged. If `BackupDirectory` is absent in the selected modern root, the effective policy default is `0` (disabled). A higher-priority root that cannot be read makes the selected source unknown.

The snapshot reports the selected policy source and configured backup directory. This is **configuration evidence**, not proof that password backup or rotation has succeeded. Legacy policy is reported separately; its `AdmPwdEnabled` value does not use the modern `BackupDirectory` enum. Read errors make the policy inventory partial.

If available, WinDefState also reads at most 20 recent events from `Microsoft-Windows-LAPS/Operational`: policy cycle start/success/failure (10003/10004/10005), directory backup success (10018/10029), and local account update success (10020). It saves only event ID, record ID, and UTC time. Event-log availability is reported independently of registry capture. Event messages and event XML are not saved.

The provider does not read arbitrary registry value data, password attributes, local account password material, or directory credentials. It does not infer password reuse across devices. A source image snapshot establishes only the local image's policy state; deployed GPO, CSP, and domain or Entra behavior must be assessed on a joined device.

## Reproduce

On an elevated Windows PowerShell prompt, run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Snapshot -IncludeId laps.policy_status
```

Inspect the `laps.policy_status` entry in the generated JSON and the corresponding text report. This command is read only and requires no reboot.

The policy order and `BackupDirectory` default follow [Microsoft's Windows LAPS policy documentation](https://learn.microsoft.com/en-us/windows-server/identity/laps/laps-management-policy-settings). Event IDs follow [Microsoft's LAPS event documentation](https://learn.microsoft.com/en-us/windows-server/identity/laps/laps-management-event-log).

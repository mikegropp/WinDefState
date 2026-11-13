# Windows Time synchronization inventory

`time.sync_status` is a read-only WinDefState snapshot entry. It never starts or stops W32Time, changes peers, resynchronizes, or adjusts the clock.

## Captured evidence

- `Win32_Service` state and start mode for `W32Time`.
- A bounded `w32tm.exe /query /status /verbose` query: leap indicator, stratum, source, and the last successful sync time as printed by the host. The parser recognizes the English labels shown in Microsoft's documentation. If Windows prints localized or unrecognized labels, the entry is partial and retains up to 8 KiB of raw status text for review.
- Allowlisted `Type` and `NtpServer` values from the Windows Time policy and local configuration roots. Both are recorded separately. The entry does not assume the local manual peer list is effective on a domain member or that any configured peer was reached.

Leap indicator `3` is reported as an unsynchronized flag. A prior successful sync timestamp does not override a current unsynchronized flag. A host's own time output cannot prove its clock offset: measure offset against a separate trusted reference at a documented time before making a clock-skew claim. A source image may also receive different time policy after domain or MDM enrollment.

## Reproduce

On the Windows host, run an elevated, read-only snapshot:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\WinDefState.ps1 -Command Snapshot -IncludeId time.sync_status
```

To inspect the same live source manually:

```powershell
Get-CimInstance Win32_Service -Filter "Name='W32Time'" | Select-Object Name,State,StartMode
w32tm.exe /query /status /verbose
Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters' -Name Type,NtpServer -ErrorAction SilentlyContinue
Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\W32Time\Parameters' -Name Type,NtpServer -ErrorAction SilentlyContinue
```

Microsoft documents [`w32tm` queries and Windows Time configuration](https://learn.microsoft.com/en-us/windows-server/networking/windows-time-service/Windows-Time-Service-Tools-and-Settings). The configured peer list and live source are distinct observations.

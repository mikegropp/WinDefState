# Pre-test environment baselines

Use the inspection dashboard to record a machine's configuration before testing,
then capture again and compare. A baseline is configuration evidence, not a disk
image: retain a VM checkpoint or disk backup when the whole machine must be
recoverable. Environment reports cannot be passed to the state engine's Restore.

## Quick start

Keep `WinDefState.Inspect.Gui.ps1`, `WinDefState.Environment.ps1`, and
`WinDefState.Health.ps1` together (or extract the release bundle). Launch in native
64-bit Windows PowerShell 5.1. An elevated session provides more complete reads;
the dashboard does not request elevation or change policy on your behalf.

```powershell
powershell.exe -NoProfile -Sta -File .\WinDefState.Inspect.Gui.ps1
```

1. Click **Capture environment** and review the unreadable-section count.
2. Export a **JSON** baseline before testing. Save it in a restricted location.
3. Capture again after testing and export a second JSON baseline.
4. Choose **Compare with before...** and open the first baseline. The currently
   loaded baseline is the after-state. Select a difference to see both values.
5. Export the comparison from the **Changes since baseline** tab, or select the
   inventory/health tab to export the complete currently loaded baseline.

The search field matches all words across the fields in the selected section or
view. Inventory sections display `?` when unreadable, and empty successful reads
display `0`. The evidence pane is selectable and can be resized. Capture runs in
a separate hidden read-only process, with cancellation and a 180-second UI
deadline. Cancelling retains the previously loaded baseline. An unsaved capture
is held only for this session; export it before closing.

## CLI and offline comparison

```powershell
# Before and after testing (run separately at the appropriate time):
.\WinDefState.Environment.ps1 -OutputPath .\before.json
.\WinDefState.Environment.ps1 -OutputPath .\after.json

# Compare existing files without querying or modifying the live machine:
.\WinDefState.Environment.ps1 -BaselinePath .\before.json -CurrentPath .\after.json -OutputPath .\changes.json
.\WinDefState.Environment.ps1 -BaselinePath .\before.json -CurrentPath .\after.json -OutputPath .\changes.html -Format Html

# Lightweight health inspection and shareable health report:
.\WinDefState.Health.ps1 -OutputPath .\health.html -Format Html
```

Each script also returns its report object for PowerShell pipelines. Omit
`-CurrentPath` when comparing a saved baseline with a fresh live capture. HTML
exports are self-contained, searchable, and make no external requests. JSON is
the comparison format. Writers refuse to overwrite existing files. A report may
contain hostnames, addresses, accounts, service paths, and internal policy names;
review it before sharing. No data is uploaded.

## Captured inventory

| Area | Evidence |
| --- | --- |
| Local ports | IPv4/IPv6 TCP listeners and UDP bound endpoints, owning PID, process name, executable path when readable |
| Firewall | Effective ActiveStore profiles and rules, source policy, direction/action/enabled state, port/protocol/ICMP, address, application/package, service, interface/type, authentication/encryption/user/machine filters |
| Networking | Adapters and drivers, IP addresses, DNS server order, routes, network connection profiles |
| Services | Current state, startup mode, service account, executable path, delayed-start flag when supplied |
| Scheduled tasks | Task identity, current state, principal, run level, enabled state, executable/COM actions; arguments and triggers omitted |
| Software | Machine uninstall-registry inventory from native and WOW6432Node views; no `Win32_Product` query or MSI repair side effects |
| Updates | `Win32_QuickFixEngineering` hotfix inventory; not a complete patch assessment |
| Protection health | Build/edition lifecycle context, Defender mode and runtime protection, signature age, Secure Boot, TPM readiness, running Credential Guard/HVCI/VBS, effective firewall profile state, common restart markers |

Firewall filters retain their own provider `InstanceID` values. Rule display
names are not used to guess associations. These records preserve the underlying
evidence for inspection, but do not implement a firewall reachability simulator.
TCP listeners and UDP bindings are **not proof that a port is reachable from
another computer**. This feature does not scan the network.

Capture is sequential, not an atomic OS snapshot. Process exit, protected
processes, and access restrictions can leave ownership paths unavailable. The
process inventory error is retained separately. Capturing from an elevated
session can improve completeness but is not a guarantee that every provider is
readable. Each section records its duration and error independently.

## Comparison semantics

- Added/removed/changed records include stable identities and both available values.
- Missing or unreadable sections become **Unknown**, never an inferred mass deletion
  or an assertion that nothing changed.
- Endpoints compare bindings by protocol/address/port, ignoring PID churn.
  Shared bindings retain their multiplicity, and owning executable changes are
  reported as changed bindings. Unresolved ownership in either capture is
  **Unknown**, so permission differences cannot masquerade as executable changes.
- Property order is normalized. Ordered arrays such as DNS-server precedence are
  preserved. Runtime task/service/network state can legitimately change.
- Comparisons require matching computer names. This is an accidental-mixup guard,
  not cryptographic machine identity, and a renamed machine requires review.
- Health evidence is retained in each baseline; the inventory diff does not
  compare health checks. Installed software does not include per-user or Store
  applications. The snapshot is not a complete inventory of every OS setting.
- Readers accept only schema-1 environment reports, enforce unique section/item
  identities, and limit input to 64 MB. Baselines are never executed.

## Windows 10 and 11 compatibility

Build and client/server product type determine the Windows family; registry
`ProductName` alone can misidentify Windows 11. Known releases use exact build
numbers, with separate Home/Pro and Enterprise/Education servicing dates.
Windows 10 22H2 standard support ended on October 14, 2025; ESU entitlement is
not inferred. LTSC and IoT LTSC receive separate lifecycle treatment. Unknown
editions, future builds, and Windows Server remain **Unknown** in the client
lifecycle assessment.

The offline lifecycle table was reviewed on **2026-09-23**, including Windows 11
24H2, 25H2, and the new-device 26H1 release. This metadata is not evidence that a
machine has current patches. Runtime providers are used for VBS/HVCI/Credential
Guard instead of treating registry intent as proof that a feature is running.
TPM readiness alone does not establish TPM 2.0 or Windows 11 eligibility.

Primary sources:

- [Windows 10 release and servicing information](https://learn.microsoft.com/en-us/windows/release-health/release-information)
- [Windows 11 release and servicing information](https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information)
- [Win32_DeviceGuard runtime state](https://learn.microsoft.com/en-us/windows/security/hardware-security/enable-virtualization-based-protection-of-code-integrity)
- [Defender runtime status](https://learn.microsoft.com/en-us/powershell/module/defender/get-mpcomputerstatus)
- [Secure Boot provider and permission behavior](https://learn.microsoft.com/en-us/powershell/module/secureboot/confirm-securebootuefi)
- [Effective firewall rules](https://learn.microsoft.com/en-us/powershell/module/netsecurity/get-netfirewallrule)
- [Firewall port filters](https://learn.microsoft.com/en-us/powershell/module/netsecurity/get-netfirewallportfilter)

## Repository review and validation

The existing engine already implements shared-provider caching, exact-baseline
checks, journal integrity, recovery checkpoints, and snapshot-history comparison.
The inspection work addresses separate gaps:

- Existing firewall snapshots hold profile/RDP group state rather than a general
  environment inventory. The environment baseline adds all effective rules and
  filter categories as read-only evidence.
- Existing protection snapshots do not capture a general port/process, software,
  route/DNS, service and task baseline. Those have independent sections now.
- Existing preflight checks runtime/provider availability but do not establish
  current client lifecycle coverage or running hardware-backed protection.
- The existing WPF runbook requires elevation and serves state-changing operations.
  The inspection UI provides a separate non-elevating capture/review/compare flow.

New modules are independent of the 10,000-plus-line state engine. They do not
import it, create its journal, change its schema, or add mutation handlers.
This gives inventory development a small testable boundary without putting
baseline capture behind the existing operation controller.

Pester covers lifecycle cases, uncertain provider output, comparison semantics,
file round trips, HTML encoding, and CLI parameter preservation. WPF validation
instantiates the dashboard without showing it or querying providers. Release
bundles and manifests include the inspection files and this guide. CI Windows
Server runners validate PowerShell/WPF contracts; they are **not** substitutes
for Windows 10/11 device testing, ARM64 testing, or managed-policy coverage.

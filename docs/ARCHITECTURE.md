# WinDefState architecture

## Product contract

WinDefState keeps three stable, elevated entry points:

- `Snapshot` captures supported state without changing it.
- `Permissive` captures a fresh baseline, journals that snapshot, applies supported permissive settings, then recaptures and verifies every attempted target.
- `Restore` restores from the journal or an explicit snapshot and verifies the resulting live state.

The JSON snapshot is the source of truth. Reports and the WPF interface are views over that state; they must never invent a baseline or weaken restore verification.

The default state root is the machine-wide `%ProgramData%\WinDefState` directory, not the script directory. This keeps journals and snapshots stable when the single-file distribution is downloaded into commit-specific or temporary folders; `-StateRoot` remains the explicit override.

## Safety invariants

- A permissive operation must persist its snapshot before the first mutation.
- Snapshot JSON must be persisted before report generation; a presentation-layer failure must not discard an otherwise valid source-of-truth capture.
- A permissive operation must refuse to start while another operation journal is active, preserving the original pre-change baseline.
- An existing unreadable, null, or structurally incomplete operation journal fails closed; it is never treated as if no recovery state exists.
- Restore must validate snapshot schema, catalog identity, immutable targets, host identity, and journal integrity before the first mutation.
- A setting with an incomplete baseline must not be changed or treated as verified.
- Restore verification classifies incomplete and inventory-only entries before canonicalization or live provider capture.
- A setter returning successfully is not proof that policy changed; permissive must recapture attempted settings and compare explicit provider targets.
- A verified configured target that needs reboot must be reported separately from an immediately active target.
- A permissive mismatch or unreadable post-apply provider transitions the journal to `ApplyVerificationFailed`; the baseline remains restorable.
- A full restore clears `current-operation.json` only after every fully captured restorable setting matches. Inventory-only entries remain reportable context and cannot fail restoration.
- Provider exceptions transition the journal to `ApplyFailed` or `RestoreFailed`; the baseline remains available for a retry.
- Restore work items are journaled before mutation and checkpointed after success. A retry recaptures completed IDs and skips only those that still match the baseline; in-flight, drifted, or unreadable IDs are reapplied.
- An explicit restore must never update or clear a journal that points to another snapshot.
- Only one public WinDefState operation may run on a computer at a time; the machine-wide mutex is released even when a provider throws.
- GUI cancellation is cooperative and marker-based. Snapshot capture and pre-mutation permissive work may stop at explicit read-only boundaries; restore and started mutation are never killed or interrupted.
- GUI mutation review is engine-gated, not a presentation-only confirmation. Permissive persists its fresh baseline before announcing review and does not create the operation journal until a one-time protected approval marker is accepted. Restore validates the snapshot and preloads selected sidecars before review, but does not change journal status or live state until approval.
- Closing or rejecting the GUI review writes `CANCEL` to the same protected one-time decision channel. The child process acknowledges that no-mutation boundary; approval is never inferred, and a captured permissive baseline remains available in snapshot history.
- GUI capability affordances derive from snapshot completeness. Incomplete and inventory-only entries cannot be selected for mutation, and capture-only exact baselines expose restore without implying a permissive target.
- New snapshots persist authoritative permissive/restore/inventory capabilities from engine dispatch. Inventory-only entries are reported but are neither scheduled as restore mutations nor compared as restorable state.
- New snapshots also persist engine-authored permissive target descriptors. They are review metadata, remain optional for legacy snapshots, and are excluded with capability metadata from GUI history fingerprints so wording changes do not masquerade as defense-state drift.
- GUI selected-ID filters cross the native `powershell.exe -File` boundary as one comma-delimited token and are normalized by the engine, avoiding Windows PowerShell array-argument ambiguity.
- CLI setting scopes accept validated PowerShell wildcard patterns and category names, which are normalized to ID filters before orchestration; unknown or empty scopes fail before capture or mutation.
- GUI completion consumes the exact persisted snapshot path from the engine's structured result channel; newest-file discovery exists only as a compatibility fallback.
- The trusted state root must be protected from ordinary-user writes before its journal or snapshots are read.
- Automatically generated snapshots and verification reports must select an unused path rather than overwrite an artifact created in the same timestamp.
- Every atomic writer removes its hidden same-directory temporary file even when serialization or writing fails before publish.
- Selected sidecar assets are loaded before the first restore mutation and reused from an operation-scoped immutable cache; the cache is cleared at every public-operation boundary.
- Every sidecar reference is resolved beneath its snapshot `.assets` root before reading; traversal outside that trusted root is rejected.
- New snapshots bind AppLocker, exploit-protection, and WDAC sidecar references to optional SHA-256 digests; restore validates those digests from the cached bytes while remaining compatible with older snapshots.
- Provider caching is phase-local and is never reused across a state transition.
- Session resources must be released from a `finally` block on snapshot, mutation, and verification paths.
- WSMan/listener mutations share one phase-local WinRM service-write scope. Restore defers only the WinRM service entry until that scope is released, and successful filtered operations fail if temporary service-state cleanup fails.
- Every definition type must implement an explicit completeness decision; unknown types are incomplete by default.
- Every definition type must implement an explicit permissive-verification decision, including an intentional not-applicable decision for capture-only types.
- WDAC permissive removal must skip platform-managed policies and must not delete raw policy files when safe `CiTool` removal is unavailable.
- WDAC restore reconciliation removes only identified non-platform policies absent from the baseline. Platform-managed and unclassified files are never cleared wholesale; uncertain state fails closed or remains visible to verification.
- WDAC restore equality covers deterministic custom-policy and policy-file state. `CiTool` availability and platform-managed policy drift remain operator-visible inventory, not restorable targets.
- Existing schema-1 snapshots remain restorable. Schema 2 adds capture metrics without changing setting semantics.
- Schema-2 producer metadata records the exact script SHA-256 and PowerShell runtime without making provenance mandatory for legacy snapshots. Verification reports distinguish the baseline producer from the current verifier.

## Capture pipeline

Definitions remain the compatibility layer for setting IDs. A capture session now owns:

- a short-lived provider cache;
- per-provider query timings;
- per-setting timings;
- total phase duration and cache-hit counts.

Provider reads that expose several settings should be performed once per session. All Defender preference definitions share one `Get-MpPreference` result, WSMan definitions share one query per resource URI, registry values share one read per key, tracked services share one `Win32_Service` query, firewall profiles share one NetSecurity call, and SMB client/server signing share one bounded child process. User-profile discovery and temporary user-hive mounts are also reused across user-scoped definitions. BitLocker keeps per-volume process isolation and timeouts but launches mounted-volume probes concurrently so process startup and wait time do not accumulate serially. Slow providers that expose only one logical setting are still timed so target-host reports identify them directly.

Mutation work is planned only after persisted baselines have been validated. Ten Defender scalar targets are submitted through one `Set-MpPreference` call, and independent exclusion/CFA list plus ASR reconciliation shares one preference baseline per phase. ASR removal is vectorized: validated ID/action pairs are submitted in one `Remove-MpPreference` call instead of one management-provider round trip per rule. Seventeen WSMan values are grouped into four resource-URI writes inside one WinRM service scope. Firewall permissive configuration submits all profiles in one call, while RDP firewall-rule restore groups rule names by enabled state and submits at most two calls. NetBIOS mutation resolves all captured adapter identities with one CIM query before invoking any per-adapter setter, and BitLocker restore resolves all target mount points with one module call before changing protection state. Mismatched, unsupported, malformed, or duplicate provider identities fail closed because they cannot be restored exactly.

Exact command discovery is cached process-wide because command availability cannot meaningfully change during one operation. Resolution runs with local verbose suppression so module auto-loading does not bury operator diagnostics under function/alias export records. The immutable 94-setting definition catalog and its ID map are also constructed once per process instead of being rebuilt across capture, validation, and verification. State values remain phase-local and are never stored in these process-wide caches.

The cache is phase-local. Snapshot, permissive mutation, post-permissive verification, restore mutation, and restore verification each receive a fresh session. Mutation sessions cache only data that is safe to reuse within that phase: user-profile discovery/open hive handles, one bounded WinRM service-write scope, and one initial Defender preference read for independent list/ASR reconciliation. Setting values are still mutated and verified in a fresh post-transition session. No cache crosses a state transition.

Definition filtering happens before capture-session construction. A selected-item snapshot or permissive operation therefore invokes only providers represented by the selected baseline; the filter is retained in `CaptureScope` metadata and in the permissive operation journal.

## Preflight and restore recovery

Every public command writes a timestamped preflight report beneath the protected state root. Preflight is intentionally lightweight: it records runtime and language mode, effective execution policy, common pending-reboot markers, journal state, selected setting count, and required provider-command availability without duplicating expensive provider captures. A warning is operator context, not permission to bypass exact-baseline or verification checks.

The schema-1 operation journal is extended compatibly with an optional `RestoreCheckpoint` object. Each restore attempt preserves completed IDs from earlier attempts, records the current work item before mutation, and atomically merges successful IDs afterward. Before resuming, checkpointed IDs are verified against the snapshot in a fresh capture session. Matching IDs are omitted from the mutation plan; all others remain scheduled. The final restore verification is never narrowed by checkpoint state, and its report preserves attempt and resume counts after a successful run clears the journal.

## Distribution pipeline

`build/Build-Release.ps1` preserves the single-file engine as the primary distribution artifact and packages the optional GUI and documentation separately. Release inputs are copied byte-for-byte, archive entries are sorted, ZIP timestamps are fixed, and compression is disabled so the same version and source tree produce the same archive bytes. The builder refuses to overwrite any existing target.

Each bundle contains an internal SHA-256 manifest. Tagged GitHub releases also publish the loose engine, loose GUI, bundle, and an external manifest covering all three assets. Checksums provide integrity and reproducibility evidence, not publisher authentication; Authenticode signing still requires a separately protected signing identity.

## Target structure

The single-file script remains the downloadable distribution artifact, but source development should move toward these boundaries:

```text
src/
  WinDefState.psm1          orchestration and public commands
  Core/                     snapshot schema, journaling, reports, verification
  Providers/                Defender, WSMan, BitLocker, WDAC, registry, firewall, etc.
  UI/                       WPF view models and process controller
tests/
  Unit/                     mocked provider and schema contracts
  Integration/              Windows-only read/capture checks
build/
  Build-SingleFile.ps1      deterministic release bundling
```

Each provider should eventually expose capture, completeness, permissive, restore, comparison, and presentation behavior behind one registration record. This removes the repeated type switches currently spread through capture, restore, comparison, reporting, and verification.

The current distribution marks stable, ordered regions for core state, registry/users, remote management, Defender, BitLocker, application control, firewall, canonicalization/reporting, lifecycle dispatch, orchestration, and the entry point. Smoke tests enforce that these boundaries remain ordered and balanced. They are the migration seams for extracting source modules without changing the generated single-file command surface.

## Performance roadmap

1. Measure and cache shared reads. This is implemented for Defender preferences, WSMan resource URIs, registry keys, services, and user-profile target discovery.
2. Reuse temporary user-hive mounts across all user-scoped definitions in one capture or mutation phase. This is implemented for snapshot, permissive, restore, and verification sessions.
3. Reuse one temporary WinRM service-write scope and group selected WSMan values by resource URI. This is implemented, including four writes for the full 17-value catalog, restore ordering, and strict filtered-operation cleanup.
4. Keep normal CLI output concise and the WPF dispatcher responsive while the engine runs. The CLI now uses native progress without requiring global verbosity, while the GUI streams structured phase records through a thread-safe collector, caches unchanged history documents, recycles virtualized setting and preview rows, and bounds retained console output.
5. Use timing reports from real Windows hosts to address BitLocker, WDAC, AppLocker, SMB, and firewall providers in measured order.
6. Add cancellation only at a phase boundary where aborting cannot leave an unjournaled or partially restored system. This is implemented for read-only snapshot capture and pre-mutation permissive work; restore remains deliberately non-cancellable.
7. Split source into testable modules and generate the single-file release artifact deterministically.

Parallel capture is not the default plan. Several Windows management providers share services, registry hives, COM infrastructure, or host-wide resources; unbounded concurrency would trade predictable correctness for fragile speed. Only providers proven independent and thread-safe should be scheduled concurrently.

## Quality gates

The dependency-free smoke suite catches critical structural contracts before gallery modules are available, including XML parsing and duplicate-name checks for both embedded XAML trees. Pester 5.7.1 exercises provider normalization, dispatch, journaling, review ordering, snapshot comparison, reporting, verification, and process behavior. A focused PSScriptAnalyzer 1.24.0 configuration additionally blocks automatic-variable collisions, empty exception handlers, suspicious assignments/comparisons, and syntax incompatible with Windows PowerShell 5.1. Both PowerShell 7 on Linux and Windows PowerShell 5.1 run these gates in CI; the Windows lane remains authoritative for the distributed runtime and instantiates the main window and mutation-review window through `WinDefState.Gui.ps1 -ValidateOnly` before test and release jobs proceed.

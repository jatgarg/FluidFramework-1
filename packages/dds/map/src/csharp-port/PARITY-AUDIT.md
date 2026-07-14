# SharedDirectory TS Parity Audit

**Base:** microsoft/main `directory.ts` (state at branch head)
**Port:** transpiledir `SharedDirectory/cs-out/*.cs` (current worktree)
**Audited:** 2026-07-14T20:35:11Z

## Summary

| Category | Count |
|---|---:|
| BUG (silent) | 7 |
| BUG (crash) | 2 |
| Deviation (correct, non-parallel) | 9 |
| Deferred (POC, documented) | 5 |
| N/A (WN doesn't use) | 2 |

## Findings

### Finding 1: Instance filtering exists now, but the identity data feeding it is not TS-equivalent
- **Impact:** BUG (silent)
- **TS:** `isMessageForCurrentInstanceOfSubDirectory` at `directory.ts:2605`; call sites at `directory.ts:1901`, `1971`, `2027`, `2089`, `2185`
- **C#:** `IsMessageForCurrentInstanceOfSubDirectory` at `SubDirectory.cs:1012`; call sites at `SubDirectory.cs:565`, `600`, `654`, `689`, `719`, `751`, `791`, `838`
- **What TS does:** Filters every clear/delete/set/createSubDirectory/deleteSubDirectory message using exact local-op target identity plus creator-client IDs, detached marker, or create sequence/reference sequence relationship.
- **What our port does:** The current worktree has a partial equivalent and `ProcessAckForSubdir` does call it. However the port does not populate the same identity data: attached local creates start with no creator client ID (`SubDirectory.cs:1026`), snapshots do not read `ci.ccIds`, and the root is initialized with `detached` while TS root has an empty creator set.
- **Why it matters:** Stale ops for a deleted/recreated subdirectory can still be accepted or valid ops can be dropped depending on missing `ClientId`/`RefSeq` data in the C# descriptor.
- **Suggested fix:** ~40-80 LOC / ~1-2 hrs. Carry creator client IDs from the host descriptor/sender and snapshot `ci`, then add parity tests for delete/recreate/same-path stale ops.

### Finding 2: `.ci` create-info from snapshots is ignored
- **Impact:** BUG (silent)
- **TS:** `populate` reads `subdirObject.ci` at `directory.ts:738` and seeds `seqData`/`clientIds` at `directory.ts:743-766`; `serializeDirectory` writes `ci` at `directory.ts:1034`; `getSerializableCreateInfo` at `directory.ts:2402`
- **C#:** `DirectorySnapshotDto` has only `Storage` and `Subdirectories` at `DirectorySnapshotLoader.cs:19`; `ReadDirectory` ignores unknown fields at `DirectorySnapshotLoader.cs:187`; `PopulateFromSnapshot` creates children with `CreateSnapshotSeqDataNoLock()` and `clientIds: null` at `SubDirectory.cs:873`
- **What TS does:** Persists and reloads `csn` and `ccIds` so future ops can be filtered against the correct subdirectory instance.
- **What our port does:** Drops `ci` entirely and assigns every loaded subdirectory synthetic `(seq: 0, clientSeq: local counter)` plus empty creator IDs.
- **Why it matters:** Instance filtering after load is weaker than TS and creation ordering after load can diverge.
- **Suggested fix:** ~60 LOC / ~1 hr. Add create-info DTO fields, parse/write-through load metadata, seed `SeqData` and `ClientIds` from it.

### Finding 3: Subdirectory ordering does not use TS `seqDataComparator`
- **Impact:** BUG (silent)
- **TS:** `seqDataComparator` at `directory.ts:364`; `subdirectories()` sorts at `directory.ts:1465`; disposed-inclusive ordering sorts at `directory.ts:2700`
- **C#:** `SeqData` exists at `SubDirectory.cs:14`, but `GetOptimisticSubDirectoriesNoLock` returns `_subdirOrder` plus pending names without a sequence comparator at `SubDirectory.cs:1220`
- **What TS does:** Orders acknowledged/detached and local-pending subdirectories by `(seq, clientSeq)` with explicit rules.
- **What our port does:** Tracks insertion order in `_subdirOrder` and appends pending names.
- **Why it matters:** Concurrent creates, reconnect/recreate, and snapshot-loaded directories can enumerate in a different order than TS. WN explicitly requires creation-order iteration.
- **Suggested fix:** ~40 LOC / ~45 min. Port `seqDataComparator` and sort the combined optimistic subdirectory list.

### Finding 4: Deleted subdirectories are not disposed and stale references remain mutable
- **Impact:** BUG (silent)
- **TS:** `deleteSubDirectory` emits dispose for pending delete at `directory.ts:1426`; committed delete calls `disposeSubDirectoryTree` at `directory.ts:2217`; recursive disposal at `directory.ts:2630`; `throwIfDisposed` at `directory.ts:1181`
- **C#:** `DeleteSubDirectory` removes/marks pending at `SubDirectory.cs:458`; remote delete clears sequenced data and removes from parent at `SubDirectory.cs:838`; no disposed flag exists on `IDirectory` (`Interfaces.cs:82`)
- **What TS does:** Old subdirectory objects become disposed, descendants are disposed, and most methods throw on further access.
- **What our port does:** Removes the child from its parent but leaves the old `SubDirectory` object usable if a caller retained it.
- **Why it matters:** A stale `IDirectory` reference can continue accepting `Set`, `CreateSubDirectory`, etc., producing ops for a deleted instance/path.
- **Suggested fix:** ~100-150 LOC / ~2 hrs. Add `_deleted`, `Dispose`/`Disposed`, recursive dispose/undispose, and `ThrowIfDisposed` guards.

### Finding 5: Remote ops for missing sequenced paths throw instead of being ignored
- **Impact:** BUG (crash / exception possible)
- **TS:** `messageHandlers` resolve via `getSequencedWorkingDirectory` and skip if missing/disposed, e.g. `directory.ts:862-865`, `882-885`, `925-936`
- **C#:** `ResolveSubDirectoryByPath` throws `OcsException` when any segment is missing at `SharedDirectory.cs:279`; only local ops with pending metadata are caught at `SharedDirectory.cs:163`
- **What TS does:** Treats ops for no-longer-sequenced directories as stale and ignores them.
- **What our port does:** Remote stale paths can fail the data-object op handler.
- **Why it matters:** Valid distributed races involving delete/recreate can become crashes in C#.
- **Suggested fix:** Few LOC / ~30 min. Return nullable from sequenced path resolution and skip remote processing when absent.

### Finding 6: The C# `localOpMetadata` substitute is narrower and ack return values are ignored
- **Impact:** BUG (silent)
- **TS:** `submitDirectoryMessage` submits local metadata at `directory.ts:672`; `processMessage` receives it at `directory.ts:812`; resubmit uses it at `directory.ts:687`; rollback uses it at `directory.ts:825`
- **C#:** `_pendingLocalOpSubdirectories` stores only client-sequence to subdir at `SharedDirectory.cs:19`; registration happens at `SubDirectory.cs:962`; `ProcessDirectoryOperation` ignores the `bool` returned by `ProcessAck*` and always completes the pending local-op mapping at `SharedDirectory.cs:173-197`
- **What TS does:** Carries the exact pending metadata object through submit, ack, resubmit, and rollback.
- **What our port does:** Reconstructs enough for basic acks via client sequence number, but not enough for rollback/resubmit and not enough to distinguish all stale filtered local acks.
- **Why it matters:** A filtered-out ack can still remove the root mapping while the per-subdir pending queues remain; reconnect/rollback cannot target the original op instance.
- **Suggested fix:** ~1 day. Introduce explicit metadata records parallel to TS and make ack processing consume metadata only on successful processing.

### Finding 7: Rollback paths are absent
- **Impact:** DEFERRED (POC choice, documented)
- **TS:** top-level `rollback` at `directory.ts:825`; per-op rollback at `directory.ts:2438-2588`
- **C#:** No `Rollback`/`rollback` implementation in `SharedDirectory.cs` or `SubDirectory.cs`
- **What TS does:** Reverts pending clear, set, delete, createSubDirectory, and deleteSubDirectory, including undispose/redispose event lifecycles.
- **What our port does:** Has no partial-nack rollback path.
- **Why it matters:** If the host runtime ever invokes rollback, pending state and optimistic reads will remain wrong.

### Finding 8: Resubmit-on-reconnect paths are absent
- **Impact:** DEFERRED (POC choice, documented)
- **TS:** `reSubmitCore` at `directory.ts:687`; `resubmitClearMessage` at `directory.ts:2260`; `resubmitKeyMessage` at `directory.ts:2293`; `resubmitSubDirectoryMessage` at `directory.ts:2347`
- **C#:** No resubmit API or pending-op requeue path
- **What TS does:** Uses local metadata to resubmit only still-live pending ops and refresh creator client IDs for subdirectory creates.
- **What our port does:** Assumes no reconnect/resubmit path for WN.
- **Why it matters:** A disconnected/reconnected C# client would lose pending local ops or keep optimistic state that never reaches the server.

### Finding 9: Stashed-op application is absent
- **Impact:** DEFERRED (POC choice, documented)
- **TS:** `applyStashedOp` at `directory.ts:985-1013`
- **C#:** No stashed-op API
- **What TS does:** Replays stashed clear/create/delete/set operations through public APIs.
- **What our port does:** No equivalent.
- **Why it matters:** Likely not used by WN's server-side client, but not TS parity.

### Finding 10: No TS-style `messageHandlers` dispatch table
- **Impact:** DEVIATION (correct but non-parallel — may confuse readers)
- **TS:** table setup at `directory.ts:847-980`; each op type has `{ process, resubmit }`
- **C#:** `ProcessDirectoryOperation` is a switch over operation subclasses at `SharedDirectory.cs:152-226`
- **What TS does:** Centralizes op-specific process/resubmit callbacks and documents sequenced-directory vs metadata-target lookup.
- **What our port does:** Uses a direct switch, and only process/ack callbacks exist.
- **Why it matters:** This is harder to audit against TS and hid missing resubmit/rollback/apply-stashed parity.

### Finding 11: Clear operation events diverge substantially
- **Impact:** DEVIATION (correct but non-parallel — may confuse readers)
- **TS:** local `clear` emits `clear` and `cleared` only at `directory.ts:1580`; remote clear emits clear events and pending-set value changes at `directory.ts:1927-1952`
- **C#:** local `Clear` raises `OnValueChanged` once per optimistic entry at `SubDirectory.cs:286`; remote `ApplyRemoteClear` raises value-changed for stable entries at `SubDirectory.cs:751`
- **What TS does:** Has distinct clear events and does not model local clear as per-key deletes.
- **What our port does:** Exposes only value/subdirectory events and uses value-changed events for clears.
- **Why it matters:** Event consumers see different notifications from TS for the same operation.

### Finding 12: Remote event suppression lacks `isNotDisposedAndReachable`
- **Impact:** DEVIATION (correct but non-parallel — may confuse readers)
- **TS:** `isNotDisposedAndReachable` at `directory.ts:1875`; used before remote events at `directory.ts:1993`, `2060`, `2150`
- **C#:** Remote apply methods filter only by instance and local pending entries (`SubDirectory.cs:689`, `719`, `751`, `791`, `838`); no disposed/reachability concept exists
- **What TS does:** Applies sequenced data but suppresses events for directories hidden by pending deletes or already disposed.
- **What our port does:** Can raise events from a subdirectory that is no longer visible in the optimistic tree.
- **Why it matters:** Event ordering/visibility differs in delete-vs-remote-op races.

### Finding 13: Child-to-parent event propagation is structurally different from `registerEventsOnSubDirectory`
- **Impact:** DEVIATION (correct but non-parallel — may confuse readers)
- **TS:** `registerEventsOnSubDirectory` relays child subdirectory events with joined relative paths at `directory.ts:2621-2628`
- **C#:** `RaiseValueChanged`, `RaiseSubDirectoryCreated`, and `RaiseSubDirectoryDeleted` climb parent pointers at `SubDirectory.cs:1306-1355`
- **What TS does:** Root listeners get relative path strings such as `foo/bar`.
- **What our port does:** Root listeners get event args containing `SubdirName` and `ParentPath` from the original mutation site.
- **Why it matters:** The data is sufficient but the algorithm and payload shape are not parallel to TS.

### Finding 14: Iteration under mutation is materialized, not TS iterator semantics
- **Impact:** DEVIATION (correct but non-parallel — may confuse readers)
- **TS:** lazy `internalIterator` at `directory.ts:1717-1785`; `entries`, `keys`, and `values` wrap it at `directory.ts:1614`, `1640`, `1664`
- **C#:** `Keys`, `Values`, `Entries`, and `Count` materialize `List` snapshots at `SubDirectory.cs:319-365`; optimistic list construction at `SubDirectory.cs:1183`
- **What TS does:** Uses live JS Map/list iterators whose behavior under mid-iteration mutation follows the TS implementation.
- **What our port does:** Returns a point-in-time list.
- **Why it matters:** If WN mutates during enumeration, order/visibility can differ from TS.

### Finding 15: `getWorkingDirectory` does not use `posix.resolve`
- **Impact:** DEVIATION (correct but non-parallel — may confuse readers)
- **TS:** root `makeAbsolute` uses `posix.resolve(posix.sep, relativePath)` at `directory.ts:840`; subdir `makeAbsolute` uses `posix.resolve(this.absolutePath, relativePath)` at `directory.ts:2595`
- **C#:** Manual segment splitter/normalizer at `SubDirectory.cs:1370-1404`
- **What TS does:** Delegates all normalization semantics to POSIX path resolution.
- **What our port does:** Handles `.`, `..`, duplicate slashes, absolute vs relative paths manually.
- **Why it matters:** Common paths match, but edge cases are not obviously identical.

### Finding 16: Snapshot write / summarization is not implemented
- **Impact:** DEFERRED (POC choice, documented)
- **TS:** `summarizeCore` at `directory.ts:659`; `serializeDirectory` writes blob-split snapshots at `directory.ts:1016-1080`
- **C#:** Load-only APIs at `SharedDirectory.cs:253`; no serializer/summarizer in `cs-out`
- **What TS does:** Writes `header` with `{ blobs, content }`, splits large values, includes `ci` on each directory, and serializes handles.
- **What our port does:** Defers snapshot writing for WN's server-side client.
- **Why it matters:** Acceptable if WN never summarizes; not full parity.

### Finding 17: Legacy `Shared` value migration is stubbed/not applied
- **Impact:** BUG (crash / exception possible)
- **TS:** `migrateIfSharedSerializable` rewrites `ValueType.Shared` into a handle payload at `localValues.ts:52-69`; called during load and set processing at `directory.ts:785`, `904`, `1006`
- **C#:** `LocalValues.MigrateIfSharedSerializable` throws `NotImplementedException` at `LocalValues.cs:50`; op/snapshot read paths do not call it (`DirectoryOpSerializer.cs:244`, `DirectorySnapshotLoader.cs:209`)
- **What TS does:** Reads very old pre-handle documents and ops.
- **What our port does:** Either stores the legacy value as-is or throws if the stub is called.
- **Why it matters:** Older snapshots/ops can crash or produce non-handle values where TS would migrate.
- **Suggested fix:** ~1-2 hrs if WN needs legacy docs; otherwise document as unsupported input and fail consistently at load/op parse.

### Finding 18: Non-handle JSON values are not materialized like TS values
- **Impact:** BUG (silent)
- **TS:** remote set uses `op.value.value` after JSON parse/handle parsing at `directory.ts:904-906`; JS primitives/objects are normal JS values
- **C#:** `ReadSerializableValue` stores `ReadJsonValueWithHandles` output at `DirectoryOpSerializer.cs:282`; handle helper returns `JsonElement.Clone()` for non-handle objects/primitives; snapshot loader similarly keeps non-string JSON values from `DirectorySnapshotLoader.cs:235-249`
- **What TS does:** A remote string/number/object arrives as a string/number/object.
- **What our port does:** Many remote/snapshot values arrive as `JsonElement` unless they are strings in snapshot load or contain handles.
- **Why it matters:** Consumers comparing or casting `Get("k")` see different runtime types than TS semantics imply.
- **Suggested fix:** ~60-100 LOC / ~1 hr. Convert JSON elements recursively to C# primitives, dictionaries, and lists while preserving handle resolution.

### Finding 19: Public dispose/disposed surface and guards are absent
- **Impact:** DEFERRED (POC choice, documented)
- **TS:** `SharedDirectory.dispose`/`disposed` at `directory.ts:484`; `SubDirectory.dispose`/`disposed` at `directory.ts:1164`; `throwIfDisposed` at `directory.ts:1181`
- **C#:** `IDirectory` has no dispose/disposed members at `Interfaces.cs:82`; no guard calls in `SubDirectory` public methods
- **What TS does:** Exposes lifecycle state and rejects most operations after disposal.
- **What our port does:** Omits the lifecycle surface.
- **Why it matters:** This is documented as deferred, but it also amplifies Finding 4.

### Finding 20: `ProcessDataObjectAttach` is intentionally a no-op
- **Impact:** NOT APPLICABLE (WN doesn't use this path)
- **TS:** Attach/load state flows through `loadCore`/snapshot population at `directory.ts:700`
- **C#:** `ProcessDataObjectAttach` ignores attach messages at `SharedDirectory.cs:127`
- **What TS does:** Channel runtime handles attach/load with storage services.
- **What our port does:** Assumes WN hydrates from snapshot and does not process attach payloads.
- **Why it matters:** Fine for the stated WN integration path, but not a general TS port.

### Finding 21: Remote delete of a missing key has different event behavior
- **Impact:** DEVIATION (correct but non-parallel — may confuse readers)
- **TS:** `processDeleteMessage` deletes and emits if no pending suppressor, even when `previousValue` is `undefined`, at `directory.ts:1989-2007`
- **C#:** `ApplyRemoteDelete` raises only when `_storage.TryGetValue` succeeds at `SubDirectory.cs:719-742`
- **What TS does:** Treats the sequenced delete op as an event-worthy change unless hidden by pending local state.
- **What our port does:** Suppresses delete events for absent keys.
- **Why it matters:** Event parity differs for speculative deletes and races.

### Finding 22: Local attached subdirectory creator IDs are not captured at creation time
- **Impact:** BUG (silent)
- **TS:** `createSubDirectory` seeds the child with `runtime.clientId ?? "detached"` at `directory.ts:1306-1315` and adds the client ID on existing optimistic dirs at `directory.ts:1329`
- **C#:** `CreateLocalClientIdsNoLock` returns an empty set when `_root.Sender != null` at `SubDirectory.cs:1026-1035`
- **What TS does:** A locally-created pending subdirectory knows who created it before the ack arrives.
- **What our port does:** The pending subdirectory often has no creator ID until `MarkCreatedSubDirectorySequencedNoLock` sees a `ClientId` on the ack descriptor.
- **Why it matters:** The instance filter can reject or accept the wrong operations during the pending window.
- **Suggested fix:** ~30 min once the host descriptor/sender exposes local client ID.

### Finding 23: Snapshot load refuses non-empty roots while TS `populate` is a merge primitive
- **Impact:** DEVIATION (correct but non-parallel — may confuse readers)
- **TS:** `loadCore` calls `populate` repeatedly for blob-split fragments at `directory.ts:705-710`; `populate` merges into the current tree at `directory.ts:722`
- **C#:** `DirectorySnapshotLoader` merges blobs into one DTO, and `LoadFromSnapshot` rejects non-empty directories at `SharedDirectory.cs:270`
- **What TS does:** Uses a merge-capable population primitive internally.
- **What our port does:** Public load is intentionally one-shot into an empty directory.
- **Why it matters:** Blob-split load works, but the algorithm is not parallel and direct repeated population is unavailable.

### Finding 24: Map-like TS API surface is narrowed
- **Impact:** NOT APPLICABLE (WN doesn't use this path)
- **TS:** `forEach`, `[Symbol.iterator]`, `entries`, `keys`, `values`, and `size` implement `Map<string, any>` semantics at `directory.ts:530-577` and `1593-1689`
- **C#:** `IDirectory` exposes .NET `IEnumerable`, `Keys`, `Values`, `Entries`, and `Count` at `Interfaces.cs:82-190`
- **What TS does:** Implements the JS `Map` surface exactly.
- **What our port does:** Provides the WN-requested C# shape.
- **Why it matters:** Not a problem for WN's API list, but not method-for-method parity.

### Finding 25: Handle wire format is ported, but registry-free egress remains a behavioral choice
- **Impact:** DEVIATION (correct but non-parallel — may confuse readers)
- **TS:** `serializeValue` uses `serializeHandles(value, serializer, bind)` at `localValues.ts:31-42`
- **C#:** `DirectoryOpSerializer.MakeHandlesSerializable` delegates to `HandleWireFormat` at `DirectoryOpSerializer.cs:298`; it throws for live `IFluidDataObject` values without a registry
- **What TS does:** Uses runtime serializer/bind context for handles.
- **What our port does:** Uses a host registry abstraction; serialized handle placeholders work without a registry, live objects do not.
- **Why it matters:** The JSON shape matches, but the egress dependency differs from TS.

## Areas confirmed matching TS

- Op type strings match TS: `set`, `delete`, `clear`, `createSubDirectory`, `deleteSubDirectory` (`DirectoryOpSerializer.cs:26-30`).
- Directory ops include absolute `path` plus `key`/`subdirName` fields matching TS interfaces (`Interfaces.cs:293-385`).
- The port has a TypeScript-style `PendingKeyLifetime` queue for multi-set-per-key pending semantics (`SubDirectory.cs:72`, `1083-1199`).
- Blob-split snapshot load is supported and merges fragments in blob order (`DirectorySnapshotLoader.cs:85-152`).
- Fluid handle wire shape matches TS legacy handle encoding: `{"type":"__fluid_handle__","url":"..."}` via `HandleWireFormat`.
- Current worktree has an `isMessageForCurrentInstanceOfSubDirectory` analogue and call sites, including `ProcessAckForSubdir`; the remaining issue is weaker identity data and lifecycle integration, not total absence of the call.

## Recommended fix priority

1. [BUG] Findings 1, 2, and 22 — finish instance identity: local metadata, client IDs, and snapshot `.ci`.
2. [BUG] Finding 4 — add dispose/disposed semantics so stale subdirectory references cannot mutate deleted instances.
3. [BUG] Finding 5 — skip stale remote paths instead of throwing.
4. [BUG] Finding 3 — port `seqDataComparator` for subdirectory ordering.
5. [BUG] Finding 18 — materialize JSON values to C# primitives/objects instead of leaking `JsonElement`.
6. [DEFERRED] Findings 7 and 8 — implement rollback/resubmit if WN can reconnect or receive nacks.
7. [DEVIATION] Findings 10-15 and 21 — improve TS readability/parity after correctness bugs are closed.

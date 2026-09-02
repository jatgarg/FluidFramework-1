# SharedString C# port — Handoff

> Target consumer: Word Native's C# Fluid runtime (server-side "special client").
> The port lives in the Fluid Framework repo during development; on transfer,
> `cs-out/*.cs` moves into waccobalt (see §5). Most of `csharp-port-common/`
> is throwaway shim; two files carry over as real deliverables (same as
> SharedDirectory transfer — `SerializedFluidHandle.cs` + `HandleWireFormat.cs`).

**Branch:** `transpiledir` (in worktree `../transpiledir`, base `microsoft/main`).

---

## 1. TL;DR

- **494/494 tests passing**, 0 skipped, 0 warnings, 0 errors.
- **API-complete** for Word Native's Phase-1 surface (see §4). Only `AnnotateAdjustRange` deferred per WN.
- **Wire format:** POCO + `System.Text.Json`. All wire discriminators verified against TS constants (op types, `ReferenceType` flags, `SequencePlace.Side`, `MergeTreeDeltaType.GROUP`).
- **Post-audit hardened**: full TS-parity audit at `PARITY-AUDIT.md` (85 findings). ~65 bugs closed with regression tests. Every fix used the "match TS branch-for-branch, no shortcuts" prompt directive.
- **Merge-tree structural + performance parity**: O(log n) tree ops (Waves 23, 24), zamboni GC with cross-block coalescing (Waves 26, 27), incremental partial-lengths, multi-stamp remove/obliterate history (M1/P4).
- **Feature-complete for Word Native**: insert/remove/annotate/obliterate/marker/interval + group ops + reconnect/rebase + handles + snapshot load (single-blob + multi-chunk + catchup ops).

---

## 2. What's in the box

```
packages/dds/
├── csharp-port-common/                         MOSTLY throwaway shim (see SharedDirectory HANDOFF §5)
│   ├── CsharpPortCommon.csproj
│   ├── FluidInterfaces.cs                      IFluidDataObject / Sender / Registry etc. (+ LocalClientId)
│   ├── FluidMessageTypes.cs                    fluidDataStoreMessageAttach shim
│   ├── FluidObjectId.cs                        ID generator
│   ├── OcsException.cs                         OcsException + OcsGateErrorCode
│   ├── SerializedFluidHandle.cs                ★ CARRY-OVER — real deliverable
│   └── HandleWireFormat.cs                     ★ CARRY-OVER — real deliverable
│
└── sequence/src/csharp-port/
    ├── README.md                               Design decisions
    ├── HANDOFF.md                              This file
    ├── PARITY-AUDIT.md                         Full TS parity audit (85 findings)
    ├── SharedString.sln
    ├── SharedString/
    │   ├── SharedString.csproj
    │   └── cs-out/                             KEEP — copies to waccobalt on transfer
    │       ├── SharedString.cs                 Public entry point + SequenceDelta event + RunInBatch
    │       ├── SharedStringOpSerializer.cs     JSON wire round-trip (all op types + interval envelopes)
    │       ├── SharedStringSnapshotLoader.cs   Simple + multi-chunk + catchup ops
    │       ├── MergeTree/                      Merge-tree core (~15 files)
    │       │   ├── MergeTree.cs                Insert / remove / obliterate / annotate / zamboni
    │       │   ├── Client.cs                   Op emit + apply + pending + reconnect/rebase
    │       │   ├── PartialLengths.cs           Incremental length cache
    │       │   ├── LocalReference.cs           Reference sliding + detach
    │       │   ├── ISegment.cs                 Segment metadata (multi-stamp remove/obliterate history)
    │       │   ├── TextSegment.cs, Marker.cs, MergeBlock.cs, Stamps.cs, Ops.cs,
    │       │   ├── PropertyMap.cs, MarkerIndex.cs, SequencePlace.cs, Constants.cs
    │       └── Intervals/                      IntervalCollection subsystem
    │           ├── IntervalCollection.cs       Public API + endpoint stickiness + Transient detach
    │           ├── SequenceInterval.cs, IntervalUtils.cs
    │           ├── 5 indexes: Id / Overlapping / StartpointInRange / EndpointInRange / Endpoint
    │           └── IIntervalIndex.cs (interface)
    └── SharedString.Tests/                     xUnit, 411 tests, ~35 files
```

Total: **~15,700 LOC** in `cs-out/`, **~11,800 LOC** in tests.

---

## 3. Build & test

Requires .NET 10 SDK (`global.json` = 10.0.204).

```bash
cd packages/dds/sequence/src/csharp-port/SharedString.Tests
dotnet test
```

Expected: **Passed: 494, Failed: 0, Skipped: 0**. Runtime ~1-2 seconds.

---

## 4. Public API

### Namespace

All types live in `Microsoft.Office.Web.Fluid` (matches waccobalt).

### Construction

```csharp
using Microsoft.Office.Web.Fluid;

var sharedString = new SharedString(
    id: "text0",                       // optional; auto-generated if null
    sender: fluidDataObjectSender,     // optional; null for offline testing
    registry: fluidDataObjectRegistry  // optional; needed if values contain DDS handles
);
```

### `SharedString` — text APIs

```csharp
// Reads
int GetLength();
string GetText();
string GetText(int start, int end);
string GetTextRangeWithMarkers(int start, int end);
PropertySet? GetPropertiesAtPosition(int position);
(ISegment segment, int offsetInSegment)? GetContainingSegment(int position);
int? GetPosition(ISegment segment);
(int start, int end)? GetRangeExtentsOfPosition(int position);
int? LocalReferencePositionToPosition(LocalReferencePosition reference);

// Mutations
void InsertText(int position, string text, PropertySet? props = null);
void DeleteText(int start, int end);
void RemoveText(int start, int end);                       // alias
void ReplaceText(int start, int end, string text, PropertySet? props = null);
void AnnotateRange(int start, int end, PropertySet props);
void ObliterateRange(int start, int end);                  // non-sided
void ObliterateRange(SequencePlace start, SequencePlace end); // sided
void InsertMarker(int position, ReferenceType refType, PropertySet? props = null);

// Markers
Marker? GetMarkerFromId(string id);
int? GetPositionOfMarker(Marker marker);
Marker? SearchForMarker(int startPos, bool forwards = true, string? tileLabel = null);

// Interval collections
IntervalCollection GetIntervalCollection(
    string name,
    IntervalType endpointType = IntervalType.SlideOnRemove);

// Batching (group op emit)
void RunInBatch(Action action);

// Reconnect / rebase
void RegeneratePendingOps();

// Events (delegate + event pattern per host FluidSharedMap convention)
event SequenceDeltaEventHandler? OnSequenceDelta;
```

### `IntervalCollection` — interval APIs

```csharp
SequenceInterval Add(
    int start, int end,
    PropertySet? properties = null,
    string? intervalId = null,
    IntervalStickiness stickiness = IntervalStickiness.End);
SequenceInterval? Change(string id, int? newStart, int? newEnd);
SequenceInterval? RemoveIntervalById(string id);
SequenceInterval? GetIntervalById(string id);

IEnumerable<SequenceInterval> CreateForwardIteratorWithStartPosition(int startPos);
IEnumerable<SequenceInterval> CreateBackwardIteratorWithStartPosition(int startPos);
IEnumerable<SequenceInterval> CreateForwardIteratorWithEndPosition(int endPos);
IEnumerable<SequenceInterval> CreateBackwardIteratorWithEndPosition(int endPos);

ISegment? GetSegment(LocalReferencePosition endpoint);
IntervalType EndpointType { get; }

// Events
event IntervalAddedEventHandler? OnAddInterval;
event IntervalDeletedEventHandler? OnDeleteInterval;
event IntervalChangedEventHandler? OnChangeInterval;
event IntervalPropertyChangedEventHandler? OnPropertyChanged;
```

`IntervalType`: `SlideOnRemove` (default) / `Simple` / `Transient` — all three behave per TS.

`IntervalStickiness`: `None=0 / Start=1 / End=2 / Full=3` — default `End`.

---

## 5. Integration recipe (waccobalt)

**Three steps.** Same pattern as SharedDirectory transfer.

### Step 1 — Slim the shim (don't delete outright)

Most of `csharp-port-common/` matches host types by name/shape and can be removed. Two files carry over:

**Carry over** to a shared waccobalt location:
- `SerializedFluidHandle.cs` — POCO, used on ingress when a handle URL can't be resolved to a live `IFluidDataObject`
- `HandleWireFormat.cs` — helpers `IsHandleShape`/`ReadHandleFromShape`/`WriteHandleShape` used by op serialization

**Delete** — waccobalt has real types with the same names + shapes:
```bash
rm packages/dds/csharp-port-common/{FluidInterfaces.cs,FluidMessageTypes.cs,FluidObjectId.cs,OcsException.cs,CsharpPortCommon.csproj}
```
Removed types: `IFluidDataObject`, `IFluidDataObjectSender`, `IFluidDataObjectRegistry`, `IFluidDataObjectMessageHandler`, `SequenceNumber`, `OpOrigin`, `SequencedDocumentMessageDescriptor`, `FluidObjectId`, `OcsException` + `OcsGateErrorCode`, `fluidDataStoreMessageAttach`.

Note: our `IFluidDataObjectSender` shim has a `LocalClientId` property. Waccobalt's real one may not expose that. If it doesn't, either extend waccobalt's or plumb the client id another way (see §9).

### Step 2 — Copy `cs-out/` into waccobalt

Move the tree under `packages/dds/sequence/src/csharp-port/SharedString/cs-out/` into `src/server/dss/DocumentSessionService.Core/Fluid/` next to `FluidSharedString.cs` (or wherever the merge-tree DDSes live). Namespaces already match (`Microsoft.Office.Web.Fluid`). No code changes needed on the port side.

### Step 3 — Adjust the sender call site (if host wraps in Bond)

Wire format in our port emits JSON via:
```csharp
_sender.QueueDataObjectMessage(_id, opTypeName, DirectoryOpSerializer.Serialize(op));
```

If waccobalt's real `IFluidDataObjectSender` requires `Bondi.Bonded<Any>` payloads (as `FluidSharedMap.cs` does), wrap the JSON at that one call site. All op-content POCO structure is already correct — only the outer envelope changes.

---

## 6. Wire format details

**All discriminators verified against TS constants (see `Ops.cs`):**

| Op | TS discriminator | Wire shape |
|---|---|---|
| Insert | `0` (INSERT) | `{"type":0,"pos1":<start>,"seg":{...}}` |
| Remove | `1` (REMOVE) | `{"type":1,"pos1":<start>,"pos2":<end>}` |
| Annotate | `2` (ANNOTATE) | `{"type":2,"pos1":<start>,"pos2":<end>,"props":{...}}` |
| Group | `3` (GROUP) | `{"type":3,"ops":[<subop>,...]}` |
| Obliterate (numeric) | `4` (OBLITERATE) | `{"type":4,"pos1":<start>,"pos2":<end>}` |
| Obliterate (sided) | `5` (OBLITERATE_SIDED) | `{"type":5,"pos1":{pos,before},"pos2":{pos,before}}` |

**Handle values** (any op with a value or props containing a handle) serialize as `{"type":"__fluid_handle__","url":"<absolute-path>"}`. Detection is recursive — nested handles in arrays/objects are detected.

**Interval envelopes** — `IntervalAdd` / `IntervalChange` / `IntervalDelete` / `IntervalPropertyChanged` — outer type codes 10-13. Endpoint sides + stickiness on the wire.

**Reference types** (verified in Wave 1 O1 audit fix):

| Flag | Value |
|---|---:|
| Simple | `0x0` |
| Tile | `0x1` |
| RangeBegin | `0x10` |
| RangeEnd | `0x20` |
| SlideOnRemove | `0x40` |
| StayOnRemove | `0x80` |
| Transient | `0x100` |

---

## 7. Snapshot format details

Both TS snapshot formats supported on load:

- **Single-blob** — `MergeTreeHeaderMetadata` with `orderedChunkMetadata.Count == 1`
- **Multi-chunk + catchup ops** — `orderedChunkMetadata.Count > 1`; loader accepts a `Func<string, string>? blobResolver` callback; chunks loaded in header order; catchup ops replayed via `Client.ApplyOp` after base tree materialization

Metadata carried: `movedSeq`/`movedClientIds` for obliterate stamps, marker length validation, tombstone placement (S1/S6 fix).

**Snapshot writing / summarization is NOT implemented** — server-side clients don't summarize. `SharedStringSnapshotLoader.cs` is load-only.

---

## 8. What's deferred / known limitations

| Item | Reason | Effort |
|---|---|---|
| Snapshot writing / summarization | Server-side client doesn't summarize | ~4-6 hrs |
| `Dispose` lifecycle | Pending WN confirmation | ~2 hrs |
| `AnnotateAdjustRange` | WN said may-defer post-December | ~2 hrs |
| Attribution (per-character author) | Word Web track-changes may want; deferred pending confirmation | ~4-6 hrs |
| Rollback paths | WN said not needed | ~1 day |
| Resubmit paths | WN said not needed | ~1 day |
| Stashed ops (`applyStashedOp`) | WN said not needed (no cross-process replay) | ~1 day |
| Full segment-groups / tracking-groups model | Deferred per audit (L5, M12) | ~1 day |
| Legacy pre-2019 `Shared` value migration | WN unlikely to load pre-handles documents | ~1-2 hrs |
| Immutable stamp objects (O8) | Non-critical structural drift; correct as-is | ~2 hrs |

All deferrals are documented in-code with `// TODO(...)` markers where relevant.

---

## 9. Design decisions

Full rationale in `README.md` + `PARITY-AUDIT.md` + session `plan.md`.

| Decision | Value | Rationale |
|---|---|---|
| Target framework | `net10.0` | Matches host `global.json` |
| Language | C# 14, `#nullable enable` | Matches host |
| Namespace | `Microsoft.Office.Web.Fluid` | Matches waccobalt |
| Event pattern | delegate + `event` | Matches `FluidSharedMap` / `FluidSharedString` |
| Threading | explicit `lock (_lock)` on mutations | Matches host pattern |
| Exceptions | `OcsException(OcsGateErrorCode.*, msg)` | Matches host |
| Number type | `int` positions, `long` sequence numbers, `double` for JSON numeric values | Alex Grigoruk guidance; TS JSON parses as double |
| Wire format | POCO + `System.Text.Json` | Confirmed acceptable by WN |
| Merge-tree balance | Path-local B+-tree splits (`MaxNodesInBlock=8`) | Matches TS `split` / `updateRoot` (Wave 23) |
| Partial-lengths | Incremental per-seq cache with structural recompute on split | Matches TS `blockUpdatePathLengths` (Wave 24) |
| Segment remove/obliterate history | Multi-stamp list with `spliceIntoList` | Matches TS `SegmentInfo` (Wave M1/P4) |
| Zamboni GC | Within-block + cross-block coalescing + sparse-block packing | Matches TS `zamboniSegments` + `packBlock` (Waves 26, 27) |
| Reference sliding | Forward slide default; `StayOnRemove` + `Transient` detach paths | Matches TS `slideReferences` (Wave 12a + L1-L8 audit) |
| Interval endpoint model | Per-endpoint `Side` (Before/After) + `IntervalStickiness` | Matches TS (Wave IC2/IC3) |
| Handle model | `IFluidDataObject` directly as handle; `IFluidDataObjectRegistry` resolves URLs | Matches waccobalt `FluidSharedMap.cs:188` |
| Group ops | `RunInBatch(Action)` API; discriminator `3` matches TS | Wave 25 |
| Reconnect / rebase | `RegeneratePendingOps()` per TS `regeneratePending`; obliterate no-expansion invariant preserved | Waves 28, 29, C4/C12 |

---

## 10. Error handling & telemetry

The port defines its own error hierarchy and logging interface in
`csharp-port-common` — mirrored 1:1 with Fluid JS shapes but decoupled
from Word server infra so the port stays runnable in isolation
(parity tests, CLI harnesses, out-of-server runs).

### 10.1 Exception hierarchy

All types in namespace `Microsoft.Office.Web.Fluid`:

- `LoggingError` — base for port invariant / assert failures.
  Implements `ILoggingError` so a telemetry consumer can merge its
  `GetTelemetryProperties()` payload into the enclosing event.
  Auto-includes `message`, `stack`, and `errorInstanceId` (UUID) in
  the payload. Prefer over `InvalidOperationException` for invariants
  the port itself enforces.
- `UsageError : LoggingError` — public-API contract violations
  (`AnnotateRange` bounds, `Change()` argument shape, endpoint order,
  `referenceRangeLabels` mutation). Auto-adds `usageError = true` to
  payload. Use wherever TS uses `throw new UsageError(...)`.
- `FluidAssert.That(cond, msg)` — the port's `assert()`. Throws
  `LoggingError` on failure. Uses `[DoesNotReturnIf(false)]` for C#
  nullability narrowing (matches TS `asserts condition`).
- `Argument*Exception` / `ArgumentNullException.ThrowIfNull` kept
  intact — idiomatic C# null/range guards; TS's `assert()` for null
  checks maps naturally to these.

### 10.2 Logging interface

- `IFluidLogger` — event sink. Never called directly by port code;
  always constructor-injected via the `logger` param on `SharedString`
  (defaults to `NullLogger.Instance`). Creates a child logger with
  namespace `SharedSegmentSequence.MergeTreeClient` for the merge-tree
  `Client` — matches TS `sequence.ts` shape.
- `FluidLogLevel` — enum matching TS `LogLevel` (Verbose=10, Info=20,
  Essential=30). Kept as a port-internal enum so the port isn't tied
  to `Microsoft.Office.Web.Common.Log.Level`.
- `FluidTelemetryEvent` — POCO carrying `EventName` + optional
  `Properties` dict.
- `NullLogger.Instance` — no-op default.
- `NamespacedLogger` — helper used by `CreateChildLogger` to prepend
  `"{namespace}.{eventName}"` to events (matches TS child-logger).

### 10.3 Bridge implementation guidance (waccobalt)

Word implements `IFluidLogger` and forwards to `Log.TraceTag`. Wiring
is via the `logger` param on `SharedString` constructor.

**Serialization convention: bracket-KV, not JSON.**

Word's ULS + Kusto pipeline uses `[Key: Value]` bracket pairs
(grep-able, no JSON quoting collisions, correlates with `tag_XXXX`
diagnostic ids). The bridge iterates `FluidTelemetryEvent.Properties`
and formats each pair as ` [{key}: {value}]` appended to the event
name string.

Sub-conventions:
- Prefix: event name first, then bracket pairs
- Missing values: literal `"N/A"` (not empty string)
- Nested: space-separated inside a single bracket, or one `TraceTag`
  per row for correlation
- Culture: `CultureInfo.InvariantCulture` always
- Exceptions: `[Exception: {ex.ToString()}]` (full string + stack)
- Enums: `.ToString()` (renders by name)
- Hot-path gating: `Log.ShouldTrace(cat, level)` before expensive
  prop formatting

Sketch (bridge implementation, not shipped with the port):

```csharp
public void SendTelemetryEvent(FluidTelemetryEvent evt, Exception? error = null, FluidLogLevel? level = null)
{
    var sb = new StringBuilder(evt.EventName);
    if (evt.Properties != null)
    {
        foreach (var (k, v) in evt.Properties)
        {
            sb.AppendFormat(CultureInfo.InvariantCulture, " [{0}: {1}]", k, v?.ToString() ?? "N/A");
        }
    }
    if (error != null)
    {
        sb.AppendFormat(CultureInfo.InvariantCulture, " [Exception: {0}]", error);
    }
    Log.TraceTag(WordFluidTags.PortTelemetry, LogCategory.WordFluid, Map(level), sb.ToString());
}
```

Tag range + `LogCategory` selection are wordfluidcsharp's concern.

### 10.4 Named events

The port emits **3 named events** — matching Fluid JS event names + payload
shapes verbatim so Kusto queries can be shared across JS and C#.

| Event | Trigger | Payload | Sampling / rate-limit |
|---|---|---|---|
| `LocalOpReentry` | Reentrancy detected in `EnterLocalMutation`. Emitted alongside the throw. | `{ depth: int }` + attached `LoggingError("Reentrancy detected in sequence local ops")` | Rate-limited to first 3 per process (static counter, matches TS `let totalReentrancyLogs = 3`) |
| `SequenceLoadFailed` | Exception during `SharedString.LoadFromSnapshot` | Attached exception via `SendErrorEvent` | error, no sampling |
| `CatchupOpsLoadFailure` | Exception during `SharedStringSnapshotLoader.ApplyCatchupOps` (inner catch, before `SequenceLoadFailed` surfaces from the outer scope) | Attached exception via `SendErrorEvent` | error, no sampling |

**Not ported** — TS emits these but the port doesn't have the code paths:

- `LocalEditsInProcessGCData` — port has no ProcessGCData path
- `MergeTreeLegacySummarizeSegmentCount` — port has no snapshot writer
- `SegmentsTotalLengthMismatch` — port has no snapshot writer
- `MergeTreeV1SummarizeSegmentCount` — port has no snapshot writer

The port only implements snapshot **loading**, not writing. If a future
scenario needs a port-side snapshot writer, the summarize/mismatch
events should be added at that time using the same `_logger` field
already threaded through `SharedString` and `Client`.

**Divergence from TS:** `LocalOpReentry` is emitted alongside the throw
in the port. TS emits it only in log-mode (opt-in via
`sharedStringPreventReentrancy: false`); the port only implements
throw-mode. Firing the event alongside the throw preserves the
observability signal without changing behavior.

**Testing:** `SharedString.Tests/SharedStringTelemetryTests.cs` covers
all three events + the rate-limit contract. The
`SharedString.ResetLocalOpReentryLogCountForTests()` internal method
lets tests reset the process-lifetime counter between runs.

---

## 11. Test coverage (494 tests, ~46 files)

Bucketed roughly (some tests span buckets):

| Area | Tests |
|---|---|
| Basics (insert/remove/annotate/get/GetLength/GetText) | ~30 |
| Marker + tile-marker ID index | ~15 |
| LocalReferencePosition + sliding | ~30 |
| PartialLengths (basic + incremental + multi-stamp) | ~25 |
| Merge-tree split + rebalance | ~15 |
| Zamboni GC (within-block, cross-block, sparse pack) | ~20 |
| Obliterate (basic + concurrent-eating + sided) | ~30 |
| Group ops (emit + receive + batching) | ~10 |
| Rebase-on-reconnect (all op types) | ~20 |
| Interval collections + all 5 indexes | ~50 |
| Snapshot load (single + multi-chunk + catchup + `.ci` + tombstones) | ~20 |
| Handles + wire round-trip | ~15 |
| Two-client convergence scenarios | ~15 |
| **PortedTests from TS `*.spec.ts`** (audit-recommended scenarios) | Growing (see PortedTests/) |

Every audit-closed bug has at least one regression test that would fail if the fix were reverted.

---

## 12. Support / questions

- **Design questions:** `README.md` + `PARITY-AUDIT.md` (findings + remediation status)
- **Code questions:** each ported file has a header comment citing its TS source. `directory.ts` → `MergeTree.cs`, `client.ts` → `Client.cs`, etc.
- **Wire format questions:** `Ops.cs` for op types + discriminators; `SharedStringOpSerializer.cs` for serialization; `PARITY-AUDIT.md` finding O1 for `ReferenceType` values
- **Waccobalt integration friction:** the one likely friction point is Step 3 above (sender wrapping in `Bondi.Bonded<T>`). Everything else is drop-in.

---

## 13. Documented deviations from TS

An independent SharedString correctness audit surfaced 30 findings. Twenty-nine were fixed. One remains as a deliberate Word rendering constraint (§13.2), listed here so future audits don't re-flag it.

### 13.1 (RESOLVED) Sided-interval endpoint-side query edge cases

**Status:** Fixed. All four range indexes (`OverlappingIntervalsIndex`, `StartpointInRangeIndex`, `EndpointInRangeIndex`, `EndpointIndex`) now apply the endpoint `Side` at boundary comparisons, matching TS's `compareReferencePositions` semantics at same-position boundaries.

At a same-position boundary, `Side.Before` sorts before `Side.After` — so an interval with `StartSide = Side.After` at position N is treated as sitting "just after N" and excluded from queries whose upper bound is exactly N. Consumer impact is limited to sided intervals whose numeric endpoint sits exactly at a query bound — Word's usage doesn't hit this, but cross-runtime peers that construct sided intervals do.

Note: the shared comparator tie-break (originally part of this cluster) was also fixed — `SequenceInterval.CompareStart`/`CompareEnd` now return 0 when the named endpoint matches, matching TS.

### 13.2 Marker-aware text encoding

`SharedString.GetText()` on a range containing markers emits U+200E (left-to-right mark) where TS's `getText` emits `"M" + markerId`. This is a Word rendering constraint, not a bug.

- **Why deliberate:** Word treats marker positions as opaque anchors. Embedding marker IDs inline would leak internal state into visible text. U+200E is invisible in normal rendering and preserves character-count invariants.
- **Effort to change:** Trivial (one string concatenation) but would break Word's text-rendering pipeline.
- **Consumer impact:** A consumer expecting TS-format inline marker IDs would see U+200E instead. No known Word consumers depend on the TS shape.

### 13.3 (RESOLVED) Tombstone position returns -1

**Status:** Fixed. `MergeTree.GetPositionOfSegment` now returns the collapsed tree position for a tombstoned (still-attached-but-removed) segment — the sum of preceding visible lengths at the boundary where the removed segment used to sit. Matches TS's `client.getPosition` semantics (see `packages/dds/merge-tree/src/test/client.getPosition.spec.ts` "Deleted Segment"). `DetachedReferencePosition` (-1) is now reserved for the truly-detached case (segment zamboni'd out).

The prior deferral was tied to the deferred obliterate scope. Since Waves 21-22 landed full-parity obliterate, the blocker is gone.

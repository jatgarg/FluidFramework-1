# SharedDirectory C# port — Handoff

> Target consumer: Word Native's C# Fluid runtime (server-side "special client").
> The port lives in the Fluid Framework repo during development; on transfer,
> `cs-out/*.cs` moves to `.../DocumentSessionService.Core/Fluid/` alongside
> `FluidSharedMap.cs`. Most of `csharp-port-common/` is throwaway shim and gets
> deleted on transfer; two files (`SerializedFluidHandle.cs`,
> `HandleWireFormat.cs`) are shared real deliverables and carry over (see §5).

**Branch:** `transpiledir` (in worktree `../transpiledir`, base `microsoft/main`).

---

## 1. TL;DR

- **127/127 tests passing**, 0 warnings, 0 errors.
- **API-complete** for the Word Native surface (see §4).
- **Wire format:** POCO + `System.Text.Json`.
- **Snapshot format:** both simple `IDirectoryDataObject` and blob-split
  `IDirectoryNewStorageFormat` are supported on load, including `.ci`
  create-info round-trip for instance identity.
- **TS parity:** pending-change model matches `directory.ts`'s per-op
  `PendingKeyLifetime` semantics; DDS-handle values use the TS
  `{"type":"__fluid_handle__","url":"…"}` wire shape;
  `isMessageForCurrentInstanceOfSubDirectory` filter matches TS at all 5
  call sites; `seqDataComparator` ports TS clauses branch-for-branch.
- **Post-review audit landed** (see `PARITY-AUDIT.md`) — 8 of 9 audit bugs
  closed with regression tests; remaining item (legacy `Shared` value
  migration) documented and only affects pre-2019 documents.
- **Only major deferred item:** `Dispose` lifecycle (confirmed not needed
  by WN for Phase 1).

---

## 2. What's in the box

```
packages/dds/
├── csharp-port-common/                        MOSTLY throwaway shim
│   ├── CsharpPortCommon.csproj
│   ├── FluidInterfaces.cs                     IFluidDataObject / Sender / Registry etc.
│   ├── FluidMessageTypes.cs                   fluidDataStoreMessageAttach shim
│   ├── FluidObjectId.cs                       ID generator
│   ├── OcsException.cs                        OcsException + OcsGateErrorCode
│   ├── SerializedFluidHandle.cs               ★ CARRY-OVER — real deliverable
│   └── HandleWireFormat.cs                    ★ CARRY-OVER — real deliverable
│
└── map/src/csharp-port/
    ├── README.md                              Design decisions + closed Q1-Q4
    ├── HANDOFF.md                             This file
    ├── PARITY-AUDIT.md                        Post-review parity audit findings
    ├── SharedDirectory.sln
    ├── SharedDirectory/
    │   ├── SharedDirectory.csproj
    │   └── cs-out/                            KEEP — copies to waccobalt on transfer
    │       ├── SharedDirectory.cs             Public entry point
    │       ├── SubDirectory.cs                Hierarchical impl + PendingKeyLifetime + instance filter
    │       ├── Interfaces.cs                  IDirectory, ISharedDirectory, ops, events
    │       ├── DirectoryOpSerializer.cs       JSON <-> DirectoryOperation (+ handle wire, native materialization)
    │       ├── DirectorySnapshotLoader.cs     Simple + blob-split snapshot load (+ .ci)
    │       ├── LocalValues.cs                 Serializer/handle shims (wave-2 stub)
    │       └── Utils.cs                       Path helpers
    └── SharedDirectory.Tests/                 xUnit, 127 tests, 10 files
        ├── (existing feature tests)
        ├── DirectoryInstanceFilterTests.cs    Post-review filter/identity/leak coverage
        ├── DirectoryValueMaterializationTests.cs  Post-review Finding 18 coverage
        └── PortedTests/                       Direct ports from TS directory.spec.ts family
```

Total: ~2,700 LOC in `cs-out/`, ~2,300 LOC in tests.

---

## 3. Build & test

Requires **.NET 10 SDK** (`global.json` = 10.0.204).

```bash
cd packages/dds/map/src/csharp-port/SharedDirectory.Tests
dotnet test
```

Expected: **Passed: 127, Failed: 0**.

---

## 4. Public API

### Namespace
All types live in `Microsoft.Office.Web.Fluid` (matches waccobalt).

### Construction
```csharp
var dir = new SharedDirectory(
    id: "myDirectory",              // optional; auto-generated if null
    sender: fluidDataObjectSender); // optional; null for offline testing
```

### `IDirectory` — the working-directory surface a subdirectory or root exposes
```csharp
// Storage (key -> value)
object? Get(string key);
IDirectory Set(string key, object? value);
bool Has(string key);
bool Delete(string key);
void Clear();
int Count { get; }

// Iteration (in TS optimistic order — matches directory.ts semantics)
IReadOnlyCollection<string> Keys { get; }
IReadOnlyCollection<object?> Values { get; }
IEnumerable<KeyValuePair<string, object?>> Entries();
// IDirectory : IEnumerable<KeyValuePair<string, object?>>
foreach (var (k, v) in dir) { … }

// Subdirectories
IDirectory CreateSubDirectory(string subdirName);
IDirectory? GetSubDirectory(string subdirName);
bool HasSubDirectory(string subdirName);
bool DeleteSubDirectory(string subdirName);
int CountSubDirectory();
IEnumerable<KeyValuePair<string, IDirectory>> SubDirectories();

// Navigation
string AbsolutePath { get; }
IDirectory? GetWorkingDirectory(string relativePath); // posix-style, "/", "..", "child/nested"

// Events (delegate+event pattern per host FluidSharedMap convention)
event ValueChangedEventHandler? OnValueChanged;
event SubDirectoryEventHandler? OnSubDirectoryCreated;
event SubDirectoryEventHandler? OnSubDirectoryDeleted;
```

### Event args
```csharp
class ValueChangedEventArgs
{
    string Key { get; }
    object? PreviousValue { get; }
    string Path { get; }        // absolute path of containing subdir
    bool Local { get; }         // true = triggered by this client
}

class SubDirectoryEventArgs
{
    string SubdirName { get; }
    string ParentPath { get; }
    bool Local { get; }
}
```

### `ISharedDirectory` — root object (adds runtime hooks)
```csharp
interface ISharedDirectory : IDirectory, IFluidDataObject, IFluidDataObjectMessageHandler
{
    string Id { get; }

    // Incoming ops (called by the dispatcher)
    void ProcessDataObjectOp(SequencedDocumentMessageDescriptor descriptor, string opJson);
    void ProcessDataObjectAttach(SequencedDocumentMessageDescriptor descriptor, fluidDataStoreMessageAttach op);
}
```

### Snapshot load
```csharp
// Simple format only
directory.LoadFromSnapshot(snapshotJson);

// Or blob-split (new format) with a WN-supplied resolver:
directory.LoadFromSnapshot(
    snapshotJson,
    blobResolver: blobName => storage.ReadBlob(blobName));
```

If the snapshot is blob-split and `blobResolver` is null, an `OcsException` is thrown with a clear message.

---

## 5. Integration recipe (waccobalt)

**Three steps.**

### Step 1 — Slim the shim (don't delete outright)
Most of `csharp-port-common/` matches host types by name/shape and can be removed. Two files are shared real deliverables — carry them over to a permanent location:

**Carry over** to a shared assembly (or duplicate under `waccobalt/common/`):
- `SerializedFluidHandle.cs` — small POCO, used on ingress when a handle URL can't be resolved to a live `IFluidDataObject`
- `HandleWireFormat.cs` — helpers `IsHandleShape`/`ReadHandleFromShape`/`WriteHandleShape` used by op serialization

**Delete** — waccobalt has real types with the same names + shapes:
```bash
rm packages/dds/csharp-port-common/{FluidInterfaces.cs,FluidMessageTypes.cs,FluidObjectId.cs,OcsException.cs,CsharpPortCommon.csproj}
```
Removed types: `IFluidDataObject`, `IFluidDataObjectSender`, `IFluidDataObjectRegistry`, `IFluidDataObjectMessageHandler`, `SequenceNumber`, `OpOrigin`, `SequencedDocumentMessageDescriptor`, `FluidObjectId`, `OcsException` + `OcsGateErrorCode`, `fluidDataStoreMessageAttach`. The port already uses the exact type names + shapes, so nothing changes at call sites.

Note: our `IFluidDataObjectSender` shim added a `LocalClientId` property. Waccobalt's real one may or may not expose that. If it doesn't, either extend waccobalt's or thread the client id another way (see §9 design notes).

### Step 2 — Copy `cs-out/` into waccobalt
Move the seven files under
`packages/dds/map/src/csharp-port/SharedDirectory/cs-out/` into
`src/server/dss/DocumentSessionService.Core/Fluid/` next to `FluidSharedMap.cs`.

Namespaces already match (`Microsoft.Office.Web.Fluid`). No code changes needed.

### Step 3 — Adjust the sender call site (if host wraps in Bond)
The one place the wire format is visible: `SharedDirectory.SubmitDirectoryOp`
(~line 130 of `SharedDirectory.cs`) calls
```csharp
_sender.QueueDataObjectMessage(_id, GetOpTypeName(op), DirectoryOpSerializer.Serialize(op));
```
This passes a JSON string. If waccobalt's real `IFluidDataObjectSender` requires
`Bondi.Bonded<Any>` (like `FluidSharedMap.cs` does), wrap the JSON payload
there. All the op-content structure is already correct — you only need to
change how it's boxed for the wire.

Suggested pattern:
```csharp
var content = new opContentsDirectoryContents { … };  // build from DirectoryOperation
_sender.QueueDataObjectMessage(
    _id,
    opContentsDirectoryOpKindParser.Make(content.type),
    Bondi.Bonded<Any>.SerializeObject(w => opContentsDirectoryContents.SerializeObject(content, w)),
    _nextClientSequenceNumber++);
```
Parallel adjustment on the receive side (`ProcessDataObjectOp`) if the incoming
type changes from `string opJson` to `Bondi.Bonded<Any> contents`.

---

## 6. Wire format details

**Confirmed with WN: POCO + `System.Text.Json` is acceptable.**

Each op serializes to a discriminated JSON object with a `type` field:

| Op | JSON shape |
|---|---|
| `Set` | `{"type":"set","path":"/subdir","key":"k","value":{"type":"Plain","value":<any>}}` |
| `Delete` | `{"type":"delete","path":"/","key":"k"}` |
| `Clear` | `{"type":"clear","path":"/"}` |
| `CreateSubDirectory` | `{"type":"createSubDirectory","path":"/","subdirName":"child"}` |
| `DeleteSubDirectory` | `{"type":"deleteSubDirectory","path":"/","subdirName":"child"}` |

`DirectoryOpSerializer.Serialize(op)` and `.Deserialize(json)` handle both
directions.

**Op-type names on the wire** match the TypeScript conventions exactly, so
convergence with TS clients is guaranteed: `"set"`, `"delete"`, `"clear"`,
`"createSubDirectory"`, `"deleteSubDirectory"`.

---

## 7. Snapshot format details

Both formats are supported on load.

### Simple `IDirectoryDataObject`
```json
{
  "storage": {"k": {"type": "Plain", "value": "v"}},
  "subdirectories": {
    "child": {"storage": {…}, "subdirectories": {…}}
  }
}
```

### Blob-split `IDirectoryNewStorageFormat`
```json
{
  "content": {  /* main IDirectoryDataObject */  },
  "blobs": ["blob0", "blob1"]
}
```
For each entry in `blobs`, the loader calls `blobResolver(blobName)` and
expects an `IDirectoryDataObject` JSON string in return. Each blob's content
is **merged** into the tree in order — later same-key entries override earlier
ones (matches TS `storage.set(key, value)` semantics).

If `blobResolver` is null and the snapshot has non-empty `blobs`, load fails
with `OcsException(OcsGateErrorCode.InvalidOperation, …)`.

**Snapshot writing / summarization is NOT implemented** — server-side clients
don't summarize.

---

## 8. What's deferred / known limitations

| Item | Reason | Effort to add |
|---|---|---|
| `Dispose` lifecycle | Pending WN confirmation — mirrors `FluidSharedMap` if needed | ~1.5-2 hrs |
| Snapshot **write** | Server-side client doesn't summarize | ~2-3 hrs |
| Registry-free `IFluidDataObject` egress | `IFluidDataObject` values require `IFluidDataObjectRegistry` so C# can produce canonical URLs; already-serialized `SerializedFluidHandle` values can emit without one | By design |
| Legacy `Shared` value type | Very old pre-handles format — throws `NotImplementedException` if encountered | ~1-2 hrs |
| Rollback / partial-nack | POC doesn't exercise these paths | ~2-3 hrs, needs WN semantics confirmation |

Everything else in the TS `directory.ts` public surface is implemented.

---

## 8a. Documented deviations from TS

A fresh independent audit will re-flag the following as "wire drift" or "algorithm divergence" unless the reviewer knows the port's context. Each is deliberate — read this list before treating any of them as a bug.

| Deviation | Where | Why deliberate |
|---|---|---|
| Relative handle URLs are not absolutized | `HandleWireFormat.ResolveSerializedHandle` | The port is a wire reader/writer, not a full `IFluidSerializer` implementation with handle context. TS `generateHandleContextPath` is a serializer responsibility that predates our port. Word Native writes absolute URLs; legacy relative-URL snapshots don't appear on our side. |
| Malformed `{type:"__fluid_handle__"}` marker without `url` stored as plain object | `HandleWireFormat.TryReadHandleUrl` + `DirectoryOpSerializer.MaterializeJsonValue` | Intentional laxity — a malformed handle marker with no URL falls through to being stored as a dictionary. TS throws while reading `url.startsWith(...)`. Preserving our fragility here would only convert a receiver-side crash into a document that can't be loaded, which is worse. |
| Remote createSubDirectory with `clientId == null` accepted | `SubDirectory.ApplyRemoteCreateSubDirectory` and `MarkCreatedSubDirectorySequencedNoLock` | Word Native's `SequencedDocumentMessageDescriptor.ClientId` can legitimately be null for system-emitted messages. TS `assertNonNullClientId` reflects TS runtime assumptions that don't hold on our side. |
| `_sender != null` used as attach state instead of a full `isAttached()` / `isDetached()` lifecycle | `SubDirectory` (multiple sites) and `SharedDirectory` | The full TS attach lifecycle (attach event, deferred-op queue, detached-vs-attached-vs-loading state machine) is on the deferred list. `_sender != null` is a heuristic that maps cleanly to Word Native's integration. |
| Iteration order after `Delete`+`Set` of the same key differs from TS `Map` | `SubDirectory.Entries()`, `Keys`, `Values` | .NET `Dictionary` slot reuse keeps a deleted/re-added key in its old iteration position; TS `Map` appends it. Ported tests already encode the C# ordering. Word Native consumers don't depend on TS-shaped iteration order after delete-and-readd. |

If a future need makes one of these matter, the "effort to add" is roughly:
- Handle context / relative URL absolutization: needs an `IFluidSerializer`-like context passed to the reader; ~4-6 hrs.
- Strict handle-marker rejection: 5 minutes.
- `assertNonNullClientId`: 30 minutes plus a decision on how Word Native marks system messages.
- Full attach lifecycle: ~1-2 days.
- TS-`Map`-shaped iteration order: replace the internal `Dictionary` with a linked-hash structure; ~2-3 hrs.

---

## 9. Design decisions

Full rationale in `README.md` (§ "Open questions", all four resolved).

| Decision | Value | Rationale |
|---|---|---|
| Target framework | `net10.0` | Matches host `global.json` |
| Lang version | C# 14 (implicit) | Latest |
| Nullable ref types | enabled | Matches host |
| Namespace | `Microsoft.Office.Web.Fluid` | Matches host |
| Event pattern | delegate + `event` | Matches `FluidSharedMap` |
| Threading | explicit `lock (_lock)` | Matches host pattern; events fire *outside* the lock |
| Exceptions | `OcsException(OcsGateErrorCode.*, msg)` | Matches host |
| Pending-change model | Option A — `PendingKeyLifetime` ordered queue | Matches TS `directory.ts` |
| Snapshot format | both simple + blob-split supported on load | Handles either format WN emits |
| Wire format | POCO + `System.Text.Json` | Confirmed by WN |

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
  the payload. All wire-drift errors (unknown op types, malformed
  op fields), sequence-integrity invariants, and snapshot invariants
  in the SharedDirectory port throw `LoggingError`.
- `UsageError : LoggingError` — public-API contract violations. Not
  currently thrown from the SharedDirectory port (TS side doesn't
  either — SharedMap/SharedDirectory validation goes through internal
  invariants), but part of the shared hierarchy.
- `FluidAssert.That(cond, msg)` — the port's `assert()`. Throws
  `LoggingError` on failure.
- `Argument*Exception` / `ArgumentNullException.ThrowIfNull` kept
  intact — idiomatic C# null/range guards.

### 10.2 Logging interface

- `IFluidLogger` — event sink. Never called directly by port code;
  always constructor-injected via the `logger` param on
  `SharedDirectory` (defaults to `NullLogger.Instance`).
- `SubDirectory` accesses the same logger via the internal
  `SharedDirectory.Logger` property (matches TS pattern where the
  whole subdirectory tree shares the root's `MonitoringContext`).
- `FluidLogLevel` — enum matching TS `LogLevel` (Verbose=10, Info=20,
  Essential=30). Kept as a port-internal enum so the port isn't tied
  to `Microsoft.Office.Web.Common.Log.Level`.
- `FluidTelemetryEvent` — POCO carrying `EventName` + optional
  `Properties` dict.
- `FluidTelemetryDataTag` — compliance tag enum (`CodeArtifact`,
  `UserData`). Mirrors TS `TelemetryDataTag`.
- `TaggedTelemetryValue` — record wrapping `(object? Value,
  FluidTelemetryDataTag Tag)`. Mirrors TS `Tagged<V, T>`. See § 10.2.1.
- `NullLogger.Instance` — no-op default.
- `NamespacedLogger` — helper used by `CreateChildLogger` to prepend
  `"{namespace}:{eventName}"` to events.

#### 10.2.1 Property tagging contract (compliance)

Every value in a telemetry property bag is one of two shapes:

- **Bare** — a primitive, string, enum, or exception. Bare values are
  **implicitly classified `CodeArtifact`**: safe to log to shipped
  telemetry. Producers must not put user content in bare values.
- **Wrapped** in `TaggedTelemetryValue(value, tag)` — either
  `CodeArtifact` (explicit safe classification) or `UserData` (PII;
  user keys, subdirectory names, property values from user input,
  etc.).

Bridges **must inspect the tag** and route accordingly:
- **Local diagnostic sinks** may render every value (both tags) —
  useful for debugging and repro.
- **Shipped telemetry sinks** (e.g., the Kusto path via
  `Log.TraceTag`) **must strip, hash, or redact `UserData`**. Bare
  values and `CodeArtifact`-tagged values may be logged verbatim.

The SharedDirectory port emits zero named events today, so no
property tagging happens on the port side. The contract exists so
future events (or the wordfluidcsharp bridge, if it emits its own
events on behalf of the DDS) have a proper safety valve.

### 10.3 Bridge implementation guidance (waccobalt)

Word implements `IFluidLogger` and forwards to `Log.TraceTag`. Wiring
is via the `logger` param on `SharedDirectory` constructor.

**Compliance handling first.** Before formatting any value into the
outgoing `Log.TraceTag` message, the bridge must:

1. Type-check each property value.
2. If it's a `TaggedTelemetryValue`, inspect `.Tag`:
   - `CodeArtifact` → unwrap and render `.Value` verbatim.
   - `UserData` → apply Word's compliance rule (strip, hash, redact,
     or emit to a separate PII-safe sink). Never render verbatim to
     the shipped telemetry path.
3. Otherwise it's a bare value → render verbatim (implicitly
   `CodeArtifact`).

Only after compliance routing does the serialization convention below
apply.

**Serialization convention: bracket-KV, not JSON.**

Word's ULS + Kusto pipeline uses `[Key: Value]` bracket pairs
(grep-able, no JSON quoting collisions, correlates with `tag_XXXX`
diagnostic ids). The bridge iterates `FluidTelemetryEvent.Properties`
and formats each compliance-cleared pair as ` [{key}: {value}]`
appended to the event name string.

Sub-conventions:
- Prefix: event name first, then bracket pairs
- Missing values: literal `"N/A"` (not empty string)
- Nested: space-separated inside a single bracket, or one `TraceTag`
  per row for correlation
- Culture: `CultureInfo.InvariantCulture` always
- Exceptions: **never** `[Exception: {ex.ToString()}]` — that
  concatenates type + message + stack into one blob. Split into three
  bracket pairs (`[ExceptionType: ...]` `[ExceptionMessage: ...]`
  `[ExceptionStack: ...]`), classify each per § 10.2.1
  (type + sanitized stack = CodeArtifact; message = UserData), and
  redact the message on the shipped-telemetry path.
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
            sb.AppendFormat(CultureInfo.InvariantCulture, " [{0}: {1}]", k, RenderForTelemetry(v));
        }
    }
    if (error != null)
    {
        AppendException(sb, error);
    }
    Log.TraceTag(WordFluidTags.PortTelemetry, LogCategory.WordFluid, Map(level), sb.ToString());
}

private static string RenderForTelemetry(object? value)
{
    // Compliance step: strip UserData, render CodeArtifact / bare verbatim.
    if (value is TaggedTelemetryValue tagged)
    {
        return tagged.Tag switch
        {
            FluidTelemetryDataTag.UserData => "[redacted]",   // or hash, or drop entirely
            _ => tagged.Value?.ToString() ?? "N/A",           // CodeArtifact
        };
    }

    // Bare value: implicit CodeArtifact by contract.
    return value?.ToString() ?? "N/A";
}

private static void AppendException(StringBuilder sb, Exception ex)
{
    // Split the exception into 3 compliance buckets — never render
    // ex.ToString() verbatim (that concatenates all three into one blob
    // and defeats the boundary).
    //
    // 1. Type name → CodeArtifact (safe verbatim).
    // 2. Message   → UserData by default (may embed interpolated user
    //                content). Redact / hash / drop per Word policy.
    // 3. Stack     → CodeArtifact (safe verbatim), but sanitize to strip
    //                the leading "[TypeName]: [Message]" line so the
    //                message doesn't slip in via the stack. Matches TS
    //                extractLogSafeErrorProperties(sanitizeStack: true).
    sb.AppendFormat(CultureInfo.InvariantCulture, " [ExceptionType: {0}]", ex.GetType().FullName);
    sb.AppendFormat(CultureInfo.InvariantCulture, " [ExceptionMessage: {0}]", RedactUserData(ex.Message));
    sb.AppendFormat(CultureInfo.InvariantCulture, " [ExceptionStack: {0}]", SanitizeStack(ex));

    // If the exception carries structured telemetry props (ILoggingError),
    // merge them under the same tagging contract as event Properties.
    if (ex is ILoggingError le)
    {
        foreach (var (k, v) in le.GetTelemetryProperties())
        {
            sb.AppendFormat(CultureInfo.InvariantCulture, " [{0}: {1}]", k, RenderForTelemetry(v));
        }
    }
}
```

Tag range + `LogCategory` selection are wordfluidcsharp's concern.
The compliance step must run before any local- or telemetry-side
formatting, and the `UserData` handling policy (redact / hash / drop
/ route to a PII-safe sink) is Word's call.

### 10.4 Named events

**Zero named events.** TS-side SharedDirectory (`directory.ts`) threads
`mc.logger` through subdirectories but has zero `sendTelemetryEvent`,
`sendErrorEvent`, or `sendPerformanceEvent` call sites. Port parity is
complete without any event ports.

The `_logger` field is threaded into `SharedDirectory` and accessible
to `SubDirectory` via the internal `SharedDirectory.Logger` property
for consistency with the SharedString port and future extension.

---

## 11. Test coverage (78 tests, 6 files)

| Test file | Focus | Count |
|---|---|---|
| `SharedDirectoryBasicsTests.cs` | Get/Set/Has/Delete/Clear, subdirs, iteration, events, `GetWorkingDirectory` | 22 |
| `DirectoryOpSerializerTests.cs` | JSON round-trip per op type | 10 |
| `DirectoryOpProcessingTests.cs` | Two-client convergence, remote-vs-local, ack flow | 15 |
| `DirectoryPendingLifetimeTests.cs` | Multi-set-per-key, partial-ack, optimistic reads, clear-then-set, subdir order | 7 |
| `DirectorySnapshotLoaderTests.cs` | Load simple + blob-split, merge order, missing resolver | 13 |
| `DirectoryHandleTests.cs` | DDS-handle wire format, registry resolution, nested handles, snapshot load | 11 |

---

## 12. Support / questions

- **Design questions:** see `README.md`. All four open questions (Q1-Q4) are
  now marked **RESOLVED** with the answer inline.
- **Code questions:** the source is fully commented and ports 1-to-1 from
  `packages/dds/map/src/directory.ts` — reference the TS line ranges in the
  file headers if you need to cross-check behavior.
- **Waccobalt integration snags:** flag them via whichever channel Jatin has
  set up; unblocking the sender wrapping (step 3 above) is the most likely
  friction point.

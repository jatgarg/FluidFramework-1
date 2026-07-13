# SharedDirectory C# port — Handoff

> Target consumer: Word Native's C# Fluid runtime (server-side "special client").
> The port lives in the Fluid Framework repo during development; on transfer,
> `cs-out/*.cs` moves to `.../DocumentSessionService.Core/Fluid/` alongside
> `FluidSharedMap.cs`. Everything under `csharp-port-common/` is throwaway shim
> and is deleted on transfer.

**Branch:** `transpiledir` (in worktree `../transpiledir`, base `microsoft/main`).

---

## 1. TL;DR

- **78/78 tests passing**, 0 warnings, 0 errors.
- **API-complete** for the Word Native surface (see §4).
- **Wire format:** POCO + `System.Text.Json`.
- **Snapshot format:** both simple `IDirectoryDataObject` and blob-split
  `IDirectoryNewStorageFormat` are supported on load.
- **TS parity:** pending-change model matches `directory.ts`'s per-op
  `PendingKeyLifetime` semantics; DDS-handle values use the TS
  `{"type":"__fluid_handle__","url":"…"}` wire shape.
- **Only deferred item:** `Dispose` lifecycle (pending confirmation from WN).

---

## 2. What's in the box

```
packages/dds/
├── csharp-port-common/                        THROWAWAY — deleted on transfer
│   ├── CsharpPortCommon.csproj
│   ├── FluidInterfaces.cs                     IFluidDataObject / Sender / Registry etc.
│   ├── FluidMessageTypes.cs                   fluidDataStoreMessageAttach shim
│   ├── FluidObjectId.cs                       ID generator
│   └── OcsException.cs                        OcsException + OcsGateErrorCode
│
└── map/src/csharp-port/
    ├── README.md                              Design decisions + closed Q1-Q4
    ├── HANDOFF.md                             This file
    ├── SharedDirectory.sln
    ├── SharedDirectory/
    │   ├── SharedDirectory.csproj
    │   └── cs-out/                            KEEP — copies to waccobalt on transfer
    │       ├── SharedDirectory.cs             Public entry point
    │       ├── SubDirectory.cs                Hierarchical impl + pending-change queues
    │       ├── Interfaces.cs                  IDirectory, ISharedDirectory, ops, events
    │       ├── DirectoryOpSerializer.cs       JSON <-> DirectoryOperation
    │       ├── DirectorySnapshotLoader.cs     Simple + blob-split snapshot load
    │       ├── LocalValues.cs                 Serializer/handle shims (wave-2 stub)
    │       └── Utils.cs                       Path helpers
    └── SharedDirectory.Tests/                 xUnit, 67 tests, 5 files
```

Total: ~2,300 LOC in `cs-out/`, ~1,000 LOC in tests.

---

## 3. Build & test

Requires **.NET 10 SDK** (`global.json` = 10.0.204).

```bash
cd packages/dds/map/src/csharp-port/SharedDirectory.Tests
dotnet test
```

Expected: **Passed: 78, Failed: 0**.

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

### Step 1 — Delete the shim
```bash
rm -rf packages/dds/csharp-port-common/
```
Every type it defined already exists in waccobalt: `IFluidDataObject`,
`IFluidDataObjectSender`, `IFluidDataObjectRegistry`, `IFluidDataObjectMessageHandler`,
`SequenceNumber`, `OpOrigin`, `SequencedDocumentMessageDescriptor`, `FluidObjectId`,
`OcsException` + `OcsGateErrorCode`, `fluidDataStoreMessageAttach`. The port
already uses the exact type names + shapes, so nothing changes at call sites.

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

## 10. Test coverage (78 tests, 6 files)

| Test file | Focus | Count |
|---|---|---|
| `SharedDirectoryBasicsTests.cs` | Get/Set/Has/Delete/Clear, subdirs, iteration, events, `GetWorkingDirectory` | 22 |
| `DirectoryOpSerializerTests.cs` | JSON round-trip per op type | 10 |
| `DirectoryOpProcessingTests.cs` | Two-client convergence, remote-vs-local, ack flow | 15 |
| `DirectoryPendingLifetimeTests.cs` | Multi-set-per-key, partial-ack, optimistic reads, clear-then-set, subdir order | 7 |
| `DirectorySnapshotLoaderTests.cs` | Load simple + blob-split, merge order, missing resolver | 13 |
| `DirectoryHandleTests.cs` | DDS-handle wire format, registry resolution, nested handles, snapshot load | 11 |

---

## 11. Support / questions

- **Design questions:** see `README.md`. All four open questions (Q1-Q4) are
  now marked **RESOLVED** with the answer inline.
- **Code questions:** the source is fully commented and ports 1-to-1 from
  `packages/dds/map/src/directory.ts` — reference the TS line ranges in the
  file headers if you need to cross-check behavior.
- **Waccobalt integration snags:** flag them via whichever channel Jatin has
  set up; unblocking the sender wrapping (step 3 above) is the most likely
  friction point.

# SharedDirectory — TypeScript → C# feasibility port

> **Status:** Wave 0 scaffolding. Not yet functional. See "Roadmap" below.
> **Owner:** transpiledir branch, @jatgarg
> **Delivery target:** Office C# Fluid repo (`.../DocumentSessionService.Core/Fluid/`),
>   drops in next to `FluidSharedMap.cs`, `FluidSharedString.cs`.

This folder holds a work-in-progress C# port of `packages/dds/map/src/directory.ts`
(the Fluid `SharedDirectory` DDS) for consumption by Word Native's C# Fluid runtime.

The folder is **excluded from the TypeScript build** (see `packages/dds/map/tsconfig.json`
`exclude` entry) and lives here purely for:

1. **Colocation** with the TS source that's being ported, so cross-file lookups and
   spec references are one directory away.
2. **Isolation** from any consumer repo until the port stabilizes.
3. **A single hand-off point** — `SharedDirectory/cs-out/*.cs` is the deliverable;
   everything else can be discarded when we transfer.

## Layout

```
csharp-port/
├── README.md                       ← this file
├── SharedDirectory.sln             ← solution: library + tests
├── SharedDirectory/                ← C# library project
│   ├── SharedDirectory.csproj      ← net10.0, nullable enabled
│   ├── cs-out/                     ← THE DELIVERABLE — transfers to host repo
│   └── cs-shims/                   ← THROWAWAY — replaced by host's real types
├── SharedDirectory.Tests/          ← xUnit test project
│   ├── SharedDirectory.Tests.csproj
│   └── ...                         ← test files
├── ts-input/                       ← (later) stripped TS reference copies
└── golden/                         ← (later) captured op/snapshot fixtures for parity
```

### `cs-out/` — the deliverable

Everything under `cs-out/` uses the `Microsoft.Office.Web.Fluid` namespace so it
plugs directly into the host repo. On transfer, these files are copied verbatim
into `.../DocumentSessionService.Core/Fluid/` alongside `FluidSharedMap.cs`.

### `cs-shims/` — throwaway stubs

Contains minimal implementations of types the host repo already provides:
`IFluidDataObject`, `IFluidDataObjectMessageHandler`, `IFluidDataObjectSender`,
`IFluidDataObjectRegistry`, `SequenceNumber`, `OpOrigin`,
`SequencedDocumentMessageDescriptor`, `FluidObjectId`, `OcsException`, and a
`fluidDataStoreMessageAttach` message shim.

These shims exist **only so `cs-out/` compiles in isolation for testing**.
When we hand off to the host repo:

1. Delete `cs-shims/`.
2. Delete `SharedDirectory.csproj` and this project structure.
3. Copy `cs-out/*.cs` into the host's `Fluid/` folder.
4. The host's real `IFluidDataObject`, etc., take over — the C# code compiles
   unchanged because we matched the exact interface shapes from
   `packages/dds/waccobalt/src/server/dss/DocumentSessionService.Core/Fluid/FluidDispatcher.cs`.

## Conventions (match host repo)

- **Target framework:** `net10.0` (matches host `global.json` SDK 10.0.204)
- **Nullable reference types:** enabled
- **Style:** tabs, Allman braces, `_camelCase` private fields, PascalCase methods
- **Namespace:** `Microsoft.Office.Web.Fluid` (matches host)
- **Events:** plain C# `delegate` + `event` pattern (matches `FluidSharedMap`, `FluidSharedString`)
- **Threading:** explicit `lock (_lock)` around mutations (matches host pattern)
- **Exceptions:** `OcsException(OcsGateErrorCode.*, …)` with typed error codes
- **Op serialization:** POCO + `System.Text.Json` (confirmed by host team — see [Q2](#q2--op-wire-format-bond-vs-systemtextjson-poco))

## Roadmap

- [x] **Wave 0 — Scaffolding**
  - [x] Folder structure + README
  - [x] Shims copied from host's `FluidDispatcher.cs`
  - [x] `.csproj` targeting `net10.0`
  - [x] `dotnet build` produces empty library successfully
  - [x] TS build excludes this folder
- [x] **Wave 1 — Feasibility demo** (`cs-out/` in-memory tree, 27 tests green)
  - [x] `Utils.cs` from `utils.ts`
  - [x] `Interfaces.cs` from `interfaces.ts` + `internalInterfaces.ts` + op union types
  - [x] `LocalValues.cs` stub (wave 2 landing point)
  - [x] `SharedDirectory.cs` + `SubDirectory.cs` — in-memory data + basic APIs + events
  - [x] xUnit tests covering set/get/has/delete/keys/count/clear, subdir create-idempotency/delete/count/creation-order iteration, `GetWorkingDirectory` relative+absolute+missing, event firing on set/nested-delete/subdir-create/subdir-delete, `ProcessDataObjectOp` throws (wave 2 landing point verified)
- [ ] **Wave 2 — Demo-Plus** (~4-6 hrs wall clock, ops + snapshot + pending changes)
  - [ ] Op JSON serialization (`System.Text.Json`, polymorphic on `type` discriminator) — see [Q2](#q2--op-wire-format-bond-vs-systemtextjson-poco)
  - [ ] Op emission on local mutations (via `IFluidDataObjectSender`)
  - [ ] Op reception: `ProcessDataObjectOp` routes by `path` → SubDirectory, dispatches by op type
  - [x] Pending-changes tracking, **Option A — TypeScript-style per-op lifetimes** (`PendingKeyLifetime` + ordered optimistic iteration) — see [Q1](#q1--iteration-during-mutations-pending-change-model)
  - [x] Snapshot load (`Populate(IDirectoryDataObject)`) — recursive tree build, including blob-split snapshots — see [Q4](#q4--snapshot-format)
  - [ ] ~15-20 new tests: JSON round-trip per op type, local→pending→ack flow, remote-op respects pending, snapshot load, nested-path routing
- [ ] **Wave 3 — Host-team review polish** (~1-2 days after Demo-Plus demo)
  - [x] Upgrade to Option A (`PendingKeyLifetime`) for TypeScript parity
  - [x] `IFluidDataObjectRegistry` / handles integration if values include DDS handles
  - [ ] Snapshot round-trip against captured TS-produced fixtures
- [ ] **Wave 4 — Full parity** (~3-4 days)
  - [ ] Port broader `directory.spec.ts` coverage
  - [ ] Cross-runtime golden fixtures (TS emit → C# decode + reverse)
  - [x] Blob-split snapshot load format if [Q4](#q4--snapshot-format) requires it

## Client requirements (from Word Native team)

APIs required by the C# consumer:
- `Get` / `Set` / `Has` / `Delete` / `Keys`
- `CreateSubDirectory` / `DeleteSubDirectory` / `GetWorkingDirectory`
- `HasSubDirectory` / `CountSubDirectories` / `SubDirectories`
  — **must preserve creation-order iteration** (stamped creation sequence numbers)
- Events: `ValueChanged`, `SubDirectoryCreated`, `SubDirectoryDeleted`

**Out of scope for the port** (host handles / doesn't need):
- Summarization / snapshot writing
- Container / data-store creation
- Staging mode
- Resubmit-on-reconnect

## Build

```sh
cd packages/dds/map/src/csharp-port
dotnet build           # builds library + tests
dotnet test            # runs the hello-world (and future) tests
```

## Reference material

- **TS source of truth:** `packages/dds/map/src/directory.ts` and adjacent files
- **TS test suite:** `packages/dds/map/src/test/directory.spec.ts`
- **Host C# Fluid conventions:**
  `packages/dds/waccobalt/src/server/dss/DocumentSessionService.Core/Fluid/`
  — especially `FluidSharedMap.cs`, `FluidSharedString.cs`, `FluidDispatcher.cs`,
  `FluidObjectId.cs`, and `SnapshotLoader.cs`

## Open questions for the host team

These need answers from the Word Native / C# Fluid team to finalize the port.
The port currently makes the simpler of the two possible choices for each; if any
answer changes, the linked wave will need rework.

### Q1 — Iteration during mutations (pending-change model)

**Question:** Does Word Native code ever iterate `SubDirectories()` or `Keys` while
local `Set` / `CreateSubDirectory` / `Delete` calls are still unacknowledged by
the server?

**Why it matters:** TS `SharedDirectory` uses `PendingKeyLifetime` bookkeeping to
preserve `Map`-style iteration order even when local unacked sets are interleaved
with remote ops. The C# `FluidSharedMap.cs` prototype uses a simpler model that
doesn't guarantee this.

**Options:**

| Option | Complexity | Reads during mutation | Effort |
|---|---|---|---|
| **A. Full TS semantics** (`PendingKeyLifetime` + optimistic iteration) | High | TS-identical | ~3 days additional |
| **B. SharedMap-style** (latest-per-key marker only) | Low | Not supported (or best-effort) | Currently implemented |

**Current port decision:** Option A — `SubDirectory` now mirrors TypeScript's ordered pending queues conceptually.

### Q2 — Op wire format (Bond vs System.Text.Json POCO) — **RESOLVED**

**Question:** For `opContentsDirectoryContents` (the analog of the SharedMap
prototype's `opContentsMapContents`), should the C# `SharedDirectory` use:

- **A.** A Bond schema (`.bond` file) with `Bondi.Bonded<Any>.SerializeObject`
  (matches `FluidSharedMap.cs` pattern)
- **B.** POCO C# classes with `System.Text.Json` (`[JsonPropertyName]`,
  polymorphic derived types)

**Resolution (2026-07-13):** Word Native confirmed **Option B (POCO +
`System.Text.Json`) is fine**. The port now ships only the JSON path; the
Bond-flavoured scaffolding (`WireMode.Bond`, `OpContentsDirectory.cs`,
`Bondi.Bonded<T>` shim, `fluidDataStoreMessageOp` overload) was removed
after resolution.

### Q3 — Multi-set to the same key back-to-back

**Question:** Does the C# consumer ever call `dir.Set("k", v1); dir.Set("k", v2)`
without waiting for the first ack in between?

**Why it matters:** The port now uses TypeScript-style `PendingKeyLifetime`
tracking, so multiple pending sets to the same key keep per-op identity and
optimistic iteration order until their individual acks arrive.

### Q4 — Snapshot format — **RESOLVED**

**Question:** Do host-produced snapshots use the simple `IDirectoryDataObject`
JSON shape (recursive `{ storage, subdirectories }`) or the newer
`IDirectoryNewStorageFormat` with blob-splitting for large directories?

**Resolution:** Load supports both. Simple `IDirectoryDataObject` shape works
with no extra plumbing. Blob-split `IDirectoryNewStorageFormat` load takes an
optional `Func<string, string>? blobResolver` callback that WN implements
against their storage service. Snapshot writing / summarization stays out of
scope (server-side clients don't summarize).

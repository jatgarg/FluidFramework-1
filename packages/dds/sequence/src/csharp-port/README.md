# SharedString C# port — POC

> **Status:** Wave 0 scaffolding. Not yet functional. See "Roadmap" below.
> **Owner:** transpiledir branch, @jatgarg
> **Delivery target:** Office C# Fluid repo, next to the existing
>   `FluidSharedString.cs` (which this port will replace once ready).

This folder holds a work-in-progress C# port of `packages/dds/sequence/src/sharedString.ts`
plus the subset of `packages/dds/merge-tree/src/` needed to back it. The C# port
is designed to be a first-class Fluid participant — it emits ops that other TS
clients understand and consumes/transforms ops from other clients to reach
eventual consistency.

The folder is **excluded from the TypeScript build** (see `packages/dds/sequence/tsconfig.json`
`exclude` entry).

## Layout

```
csharp-port/
├── README.md                       ← this file
├── SharedString.sln                ← solution: library + tests
├── SharedString/                   ← C# library project
│   ├── SharedString.csproj         ← net10.0, nullable enabled;
│   │                                 references ../../../../csharp-port-common
│   ├── cs-out/                     ← THE DELIVERABLE — transfers to host repo
│   │   ├── SharedString.cs         (public API)
│   │   ├── SharedStringOps.cs      (Insert/Delete op types)
│   │   ├── SharedStringOpSerializer.cs
│   │   ├── SharedStringSnapshotLoader.cs
│   │   └── MergeTree/              ← merge-tree lives inline for POC
│   │       ├── TextSegment.cs
│   │       ├── MergeTree.cs
│   │       ├── Client.cs
│   │       └── (etc.)
│   └── cs-shims/                   ← THROWAWAY, currently empty
│                                     (shared shims live in csharp-port-common)
├── SharedString.Tests/             ← xUnit test project
│   ├── SharedString.Tests.csproj
│   ├── Fixtures/                   ← golden fixtures (e.g. snapshot bytes
│   │                                 from a TS SharedString reference test)
│   └── (test files)
```

## Conventions (match SharedDirectory port + host repo)

- **Target framework:** `net10.0`
- **Nullable reference types:** enabled
- **Style:** tabs, Allman braces, `_camelCase` private fields, PascalCase methods
- **Namespace:** `Microsoft.Office.Web.Fluid`
- **Events:** `delegate` + `event` with `Local` flag on args (matches SharedDirectory)
- **Threading:** explicit `lock (_lock)` around mutations
- **Exceptions:** `OcsException(OcsGateErrorCode.*, ...)`
- **Op serialization:** POCO + `System.Text.Json` (matches SharedDirectory)
- **Number type:** TS `number` → C# `int` by default (per Alex Grigoruk's transpiler advice)

## POC scope (what's in / out)

**IN — Minimum Viable working SharedString:**
- APIs: `GetLength`, `GetText`, `InsertText(pos, text)`, `DeleteText(pos1, pos2)`
- Op emission (Insert / Delete) via `IFluidDataObjectSender`
- Op reception with **real merge-tree transformation** — two-client convergence
- Snapshot load from Fluid-format bytes
- Event: `SequenceDelta` fires post-mutation (matches TS)

**OUT — Deferred to Demo-Plus and later:**

APIs:
- `AnnotateRange`, `AnnotateAdjustRange`
- `ObliterateRange`
- `InsertMarker`, `SearchForMarker`, `GetMarkerFromId`
- `LocalReferencePositionToPosition`
- `GetPropertiesAtPosition`, `GetRangeExtentsOfPosition`, `GetContainingSegment`, `GetPosition`
- Entire `IntervalCollection` subsystem (~2,500 lines TS)

Merge-tree:
- `partialLengths` cache — POC uses walk-tree position queries (correct, slow)
- Attribution
- LocalReference / reference sliding
- Zamboni GC
- Blob-split snapshots
- Multiple segment types beyond `TextSegment`
- Obliterate machinery

## Roadmap

- [x] **Wave 0 — Scaffolding + extract common shims**
- [ ] **Wave 0.5 — Generate snapshot fixture from TS**
- [ ] **Wave 1 — Merge-tree foundational types** (parallel)
- [ ] **Wave 2 — Merge-tree core**
- [ ] **Wave 3 — Position query (simplified)**
- [ ] **Wave 4 — Client wrapper + transformation** (correctness-critical)
- [ ] **Wave 5 — Snapshot load**
- [ ] **Wave 6 — SharedString DDS layer**
- [ ] **Wave 7 — Op emission / reception wire-up**
- [ ] **Wave 8 — Tests + two-client convergence**

## Build

```sh
cd packages/dds/sequence/src/csharp-port
dotnet build           # builds library + tests
dotnet test            # runs tests
```

## Reference material

- **TS source of truth:**
  - `packages/dds/sequence/src/sharedString.ts` (SharedString public class)
  - `packages/dds/sequence/src/sharedSequence.ts` (base)
  - `packages/dds/sequence/src/sequence.ts` (base infrastructure)
  - `packages/dds/merge-tree/src/mergeTree.ts` (tree data structure)
  - `packages/dds/merge-tree/src/client.ts` (client wrapper)
  - `packages/dds/merge-tree/src/partialLengths.ts` (position translation — POC uses simpler walk)
  - `packages/dds/merge-tree/src/snapshotV1.ts` (snapshot format)
- **Host C# Fluid conventions:**
  - `packages/dds/waccobalt/src/server/dss/DocumentSessionService.Core/Fluid/FluidSharedString.cs`
    (the prototype we're replacing)
  - `packages/dds/waccobalt/src/server/dss/DocumentSessionService.Core/Fluid/SharedStringTransformationHelper.cs`
    (piece-manager OT prototype — different architecture; reference for op wire format)

# SharedDirectory Data Structure Audit — TS vs C#

**Base:** microsoft/main (`directory.ts` + companions)
**Port:** transpiledir `packages/dds/map/src/csharp-port/`
**Audited:** 2026-08-04T14:40:59-07:00
**Method:** field-by-field comparison of all major data structures, DTOs, and enums. Static/read-only source audit; no build or tests run.

## Wire-tolerance policy

When TS's **static type** marks a field required but TS's **runtime** accepts it as missing with a defined fallback behavior, the port matches the TS **runtime** rather than the TS **type**.

**Rationale.** Real production snapshots have been shown to hit these tolerant paths (`minSequenceNumber` fallback, WAC snapshots without `ccIds`, etc.). If the port is stricter than the TS runtime, we reject valid documents that TS peers already accept — an interop failure with no upside. If the port is stricter than the TS type but matches the TS runtime, we simply behave the same as any conforming TS client, which is what wire interop requires.

**Contrast.** Where TS is *genuinely* strict (throws on the wire path) — e.g. missing `type` on an op envelope, missing `sequenceNumber` on a catchup message — the port stays strict to match.

**How to apply this rule.** If a fresh audit flags "the C# port rejects field X but TS accepts it as missing":
- Read the TS runtime path for that field (not just the TS type).
- If TS silently defaults / falls through when X is missing, the port must too. Relax the port and add a regression test named after the failure mode.
- If TS throws or explicitly rejects when X is missing, the port stays strict.

## Summary

Counts below are row-level field/member comparisons, not pass/fail type counts. `Types` counts the major named structures audited in each category.

| Category | Types | ✅ | ⚠️ | ❌ | ➕ |
|---|---:|---:|---:|---:|---:|
| Public API interfaces | 7 | 16 | 12 | 5 | 8 |
| Op runtime types | 8 | 13 | 8 | 0 | 2 |
| Op wire DTOs | 6 | 18 | 3 | 0 | 0 |
| SharedDirectory class fields | 2 | 2 | 3 | 2 | 6 |
| SubDirectory class fields | 5 | 10 | 7 | 5 | 8 |
| Pending state types | 12 | 15 | 11 | 7 | 8 |
| Value types | 5 | 5 | 4 | 0 | 2 |
| Snapshot DTOs | 5 | 9 | 6 | 0 | 2 |
| Handle types | 3 | 3 | 1 | 1 | 1 |
| **Total** | **53** | **91** | **55** | **20** | **37** |

## Findings by category

### 1. Public API interfaces

#### Type: `IValueChanged` / `IDirectoryValueChanged` / C# `ValueChangedEventArgs`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `key` | `readonly key: string` | `string Key { get; set; }` | ✅ | Name is PascalCase; mutable in C# vs TS readonly. |
| `previousValue` | `readonly previousValue: any` | `object? PreviousValue { get; set; }` | ✅ | `any`/`unknown` maps to `object?`; C# cannot distinguish JS `undefined` from `null` after materialization. |
| `path` | `IDirectoryValueChanged.path: string` | `string Path { get; set; }` | ✅ | TS field is only on shared-directory event args; C# folds it into one value-changed args class for all scopes. |
| `local` | Listener parameter `local: boolean`, not inside event object | `bool Local { get; set; }` | ➕ | C# invented payload field instead of TS listener tuple. Useful adapter, but not branch-for-branch. |
| `target` | Listener parameter `target: IEventThisPlaceHolder` | Sender object in .NET event | ⚠️ | Equivalent information is event sender, not event args field. |

#### Type: subdirectory event payloads / C# `SubDirectoryEventArgs`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `path` | Listener parameter `path: string`; relative path from object raising event | No single `Path`; split into `SubdirName` + `ParentPath` | ⚠️ | Root TS event canbecarry `foo/bar`; C# carries `ParentPath="/foo"`, `SubdirName="bar"`. Data is derivable but API shape differs. Related to PARITY Finding 13. |
| `subdirName` | Not a public TS event field; operation DTO field only | `string SubdirName { get; set; }` | ➕ | C# event-specific adapter field. |
| `parentPath` | Not a public TS event field | `string ParentPath { get; set; }` | ➕ | C# event-specific adapter field. |
| `local` | Listener parameter `local: boolean` | `bool Local { get; set; }` | ➕ | C# folds tuple parameter into args. |
| `target` | Listener parameter `target` | Sender object | ⚠️ | Equivalent through .NET event sender, not payload. |

#### Type: `IDirectory`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `absolutePath` | `readonly absolutePath: string` getter | `string AbsolutePath { get; }` | ✅ | Equivalent. |
| `get(key)` | `get<T = any>(key: string): T \| undefined` | `object? Get(string key)` | ⚠️ | API-equivalent for common use, but C# loses generic return typing and JS `undefined` vs `null`. |
| `set(key,value)` | `set<T = unknown>(key: string, value: T): this` | `IDirectory Set(string key, object? value)` | ✅ | Fluent return preserved as interface type, not concrete `this`. |
| `has(key)` | inherited `Map.has(key): boolean` | `bool Has(string key)` | ✅ | Equivalent. |
| `delete(key)` | `delete(key: string): boolean` | `bool Delete(string key)` | ✅ | Equivalent. |
| `clear()` | inherited/legacy `clear(): void` | `void Clear()` | ✅ | Method exists, but C# lacks TS `clear`/`cleared` events. |
| `size` | inherited `Map.size: number` getter | `int Count { get; }` | ⚠️ | Naming and numeric type differ; C# `int` can overflow before TS `number` safe range. |
| `keys()` | `IterableIterator<string>` | `IReadOnlyCollection<string> Keys { get; }` | ⚠️ | Materialized collection vs live iterator semantics. Related to PARITY Finding 14 / 24. |
| `values()` | `IterableIterator<any>` | `IReadOnlyCollection<object?> Values { get; }` | ⚠️ | Materialized collection; loses JS iterator behavior and `undefined`. |
| `entries()` / `[Symbol.iterator]()` | `IterableIterator<[string, any]>` | `IEnumerable<KeyValuePair<string, object?>>`, `Entries()` | ⚠️ | .NET enumeration is an adapter, not JS `Map` iterator. |
| `forEach(...)` | `forEach(callback, thisArg?)` inherited from `Map`/legacy | Missing | ❌ | Public TS map surface is narrowed. Related to PARITY Finding 24. |
| `[Symbol.toStringTag]` | Implemented by classes and exposed on shared/subdir instances | Missing from C# interface | ❌ | Not relevant to .NET consumers, but not shape-equivalent. |
| `countSubDirectory()` | Optional `countSubDirectory?(): number` | `int CountSubDirectory()` | ⚠️ | C# makes it required and returns `int`. |
| `createSubDirectory(name)` | `(string) => IDirectory` | `IDirectory CreateSubDirectory(string)` | ✅ | Equivalent. |
| `getSubDirectory(name)` | `IDirectory \| undefined` | `IDirectory?` | ✅ | `undefined` maps to nullable. |
| `hasSubDirectory(name)` | `boolean` | `bool` | ✅ | Equivalent. |
| `deleteSubDirectory(name)` | `boolean` | `bool` | ✅ | Equivalent. |
| `subdirectories()` | `IterableIterator<[string, IDirectory]>` | `IEnumerable<KeyValuePair<string, IDirectory>> SubDirectories()` | ⚠️ | .NET enumerable adapter; TS returns iterator. |
| `getWorkingDirectory(path)` | `IDirectory \| undefined` | `IDirectory?` | ✅ | Equivalent return shape. |
| `on/once/off` event provider | `IEventProvider<IDirectoryEvents>` | CLR events on interface | ⚠️ | Event model and payloads differ. |
| `dispose` / `disposed` | `Partial<IDisposable>` on interface; implemented by classes | Missing from `IDirectory` and C# `SubDirectory` | ❌ | Lifecycle state is not represented. Related to PARITY Findings 4 and 19. |

#### Type: `IDirectoryEvents` and `ISharedDirectoryEvents`

| Event | TS | C# | Status | Notes |
|---|---|---|---|---|
| `containedValueChanged` | `(changed: IValueChanged, local, target)` on each `IDirectory` | `event OnValueChanged` with `ValueChangedEventArgs` | ⚠️ | C# event args include path/local and propagate to root; not the same contained/shared split. |
| `valueChanged` | `(changed: IDirectoryValueChanged, local, target)` on `ISharedDirectory` | same `OnValueChanged` event | ⚠️ | C# uses one event for both levels; TS has separate contained and shared events. |
| `clear` | `(local, target)` deprecated shared event | Missing | ❌ | C# has no clear-specific event. Related to PARITY Finding 11. |
| `cleared` | `(path: string, local, target)` | Missing | ❌ | C# has no clear-specific path event. Related to PARITY Finding 11. |
| `subDirectoryCreated` | `(path: string, local, target)` | `OnSubDirectoryCreated(SubDirectoryEventArgs)` | ⚠️ | Payload shape differs (`path` vs `ParentPath` + `SubdirName`). |
| `subDirectoryDeleted` | `(path: string, local, target)` | `OnSubDirectoryDeleted(SubDirectoryEventArgs)` | ⚠️ | Payload shape differs. |
| `disposed` | `(target)` on subdirectory | Missing | ❌ | No disposed lifecycle event. |
| `undisposed` | `(target)` on subdirectory rollback restore | Missing | ❌ | No undisposed lifecycle event. |

#### Type: `ISharedDirectory`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| Extends directory API | `Omit<IDirectory,"on"\|"once"\|"off">` plus `ISharedObject` | `ISharedDirectory : IDirectory, IFluidDataObject, IFluidDataObjectMessageHandler` | ⚠️ | C# adds host object/message-handler interfaces instead of TS `ISharedObject` runtime surface. |
| `[Symbol.iterator]()` | `IterableIterator<[string, any]>` | inherited `IEnumerable<KeyValuePair<string, object?>>` | ⚠️ | Adapter, not same symbol field. |
| `[Symbol.toStringTag]` | `readonly [Symbol.toStringTag]: string` | Missing | ❌ | C# has no counterpart. |
| `Id` | Inherited TS `SharedObject` internals; not in public `ISharedDirectory` as `Id` | `string Id { get; }` via `IFluidDataObject` | ➕ | Legitimate host shim. |
| `ProcessDataObjectOp` | TS runtime calls `processMessagesCore` with envelopes/content | `void ProcessDataObjectOp(descriptor, opJson)` | ➕ | Host integration shim, not TS API. |
| `ProcessDataObjectAttach` | TS channel attach handled by runtime/loadCore | C# no-op method on interface | ➕ | Host integration shim. Related to PARITY Finding 20. |

### 2. Op runtime types + wire DTOs

#### Type: operation union mapping

| Type / Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `IDirectoryOperation` | Structural union of five interfaces | `DirectoryOperation` abstract base with subclasses | ✅ | Union is mapped to class hierarchy. |
| Discriminant type | `type` string literal on every TS op | `DirectoryOpType` enum via abstract `Type` property | ⚠️ | Runtime C# object uses enum; serializer writes TS strings. Wire shape is OK; runtime type shape differs. |
| Shared `path` | Required `path: string` on all directory ops | `string Path { get; set; } = string.Empty` on base | ⚠️ | Present, but default empty string can mask missing/invalid paths. |
| `IDirectoryKeyOperation` | `IDirectorySetOperation \| IDirectoryDeleteOperation` | No explicit union type | ⚠️ | Represented by subclass pattern/switches, not a named type. |
| `IDirectoryStorageOperation` | key op union plus clear | No explicit union type | ⚠️ | Represented by subclass pattern. |
| `IDirectorySubDirectoryOperation` | create/delete subdirectory union | No explicit union type | ⚠️ | Represented by subclass pattern. |

#### Type: `IDirectorySetOperation` / `DirectorySetOperation`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | literal `"set"` | `DirectoryOpType.Set`; wire `"set"` | ⚠️ | Runtime enum, wire string. |
| `key` | `string` required | `string Key = string.Empty` | ✅ | Same wire field name `key`; default empty string is a tolerance difference. |
| `path` | `string` required | inherited `Path` | ✅ | Same wire field name `path`. |
| `value` | `ISerializableValue` required | `SerializableValue Value = new()` | ✅ | Field exists; see value-type mismatch section for `value.value`. |

#### Type: `IDirectoryDeleteOperation` / `DirectoryDeleteOperation`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | literal `"delete"` | `DirectoryOpType.Delete`; wire `"delete"` | ⚠️ | Runtime enum, wire string. |
| `key` | `string` required | `string Key = string.Empty` | ✅ | Same wire field name `key`. |
| `path` | `string` required | inherited `Path` | ✅ | Same wire field name `path`. |

#### Type: `IDirectoryClearOperation` / `DirectoryClearOperation`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | literal `"clear"` | `DirectoryOpType.Clear`; wire `"clear"` | ⚠️ | Runtime enum, wire string. |
| `path` | `string` required | inherited `Path` | ✅ | Same wire field name `path`. |

#### Type: `IDirectoryCreateSubDirectoryOperation` / `DirectoryCreateSubDirectoryOperation`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | literal `"createSubDirectory"` | `DirectoryOpType.CreateSubDirectory`; wire `"createSubDirectory"` | ⚠️ | Runtime enum, wire string. |
| `path` | `string` required | inherited `Path` | ✅ | Same wire field name `path`. |
| `subdirName` | `string` required | `string SubdirName = string.Empty` | ✅ | Same wire field name `subdirName`. |

#### Type: `IDirectoryDeleteSubDirectoryOperation` / `DirectoryDeleteSubDirectoryOperation`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | literal `"deleteSubDirectory"` | `DirectoryOpType.DeleteSubDirectory`; wire `"deleteSubDirectory"` | ⚠️ | Runtime enum, wire string. |
| `path` | `string` required | inherited `Path` | ✅ | Same wire field name `path`. |
| `subdirName` | `string` required | `string SubdirName = string.Empty` | ✅ | Same wire field name `subdirName`. |

#### Op wire DTOs (`DirectoryOpSerializer`)

| Wire shape | TS JSON fields | C# serializer fields | Status | Notes |
|---|---|---|---|---|
| Set op | `type`, `path`, `key`, `value` | writes `type`, `path`, `key`, `value` | ✅ | Field names and op type string match. |
| Delete op | `type`, `path`, `key` | writes `type`, `path`, `key` | ✅ | Field names and op type string match. |
| Clear op | `type`, `path` | writes `type`, `path` | ✅ | Field names and op type string match. |
| Create-subdir op | `type`, `path`, `subdirName` | writes `type`, `path`, `subdirName` | ✅ | Field names and op type string match. |
| Delete-subdir op | `type`, `path`, `subdirName` | writes `type`, `path`, `subdirName` | ✅ | Field names and op type string match. |
| Unknown fields | JS runtime naturally carries/ignores extra fields structurally | C# `ReadFrom` skips unknown fields | ✅ | Tolerant read is fine. |
| Missing `type` | TS type requires it; runtime handler needs known `op.type` | C# throws `UnknownOp` | ✅ | Equivalent failure. |
| Missing `path`/`key`/`subdirName` | TS type requires them | C# deserializer defaults to `string.Empty` | ⚠️ | Optional-vs-required tolerance can convert malformed wire ops into root/empty-key operations. |
| Missing set `value` | TS type requires it | C# creates `new SerializableValue()` | ⚠️ | Malformed set becomes `Type=""`, `Value=null` instead of hard failure. |
| Value wrapper fields | `{ type: string, value: any }` | writes `{ "type": ..., "value": ... }` | ✅ | Wire field names match. |
| Numeric values in `value` | JSON number in TS becomes JS `number` | C# materializes JSON number as `double` | ✅ | Closest C# shape to JS `number`. |
| Handles in `value` | TS serializer emits serialized handle object | C# emits handle dictionary with `type`/`url` | ✅ | Missing `payloadPending` is covered in handle section. |

### 3. SharedDirectory class fields

#### Type: TS `IDirectoryMessageHandler` / C# dispatch

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `process(...)` | Function taking envelope, op, local flag, local metadata, client seq | No handler interface; `ProcessDirectoryOperation` switch | ⚠️ | Behavior is represented by switch, not data structure. Related to PARITY Finding 10. |
| `resubmit(...)` | Function taking op and local metadata | Missing | ❌ | No resubmit handler table or metadata path. Related to PARITY Finding 8. |

#### Type: `SharedDirectory` fields and state

| Field / Property | TS | C# | Status | Notes |
|---|---|---|---|---|
| string tag | public `[Symbol.toStringTag] = "SharedDirectory"` | Missing | ❌ | Cosmetic for JS, but absent. |
| root | `private readonly root: SubDirectory` | `private readonly SubDirectory _root` | ✅ | Equivalent root pointer. |
| message handlers | `private readonly messageHandlers = new Map<string, IDirectoryMessageHandler>()` | Missing; direct switch | ⚠️ | Intentional C# structural deviation. |
| id | inherited `SharedObject` `id` | `private readonly string _id`, public `Id` | ➕ | Host/API shim. |
| runtime | inherited/constructor `IFluidDataStoreRuntime` used by root | `IFluidDataObjectSender? _sender` | ⚠️ | C# replaces Fluid runtime with host sender. |
| serializer / handle context | inherited `serializer`, `handle` | `IFluidDataObjectRegistry? _registry` | ⚠️ | C# registry resolves/serializes handles instead of TS serializer/bind context. Related to PARITY Finding 25. |
| pending local metadata root map | TS carries actual `DirectoryLocalOpMetadata` through runtime | `_pendingLocalOpSubdirectories: Dictionary<long, SubDirectory>` | ➕ | C#-only substitute; narrower than TS metadata. Related to PARITY Finding 6. |
| public events | TS event emitter inherited from `SharedObject` | `OnValueChanged`, `OnSubDirectoryCreated`, `OnSubDirectoryDeleted` | ➕ | Host adapter surface. |
| `RootDirectory` internal | No TS public/internal equivalent; root private | `internal SubDirectory RootDirectory` | ➕ | C# helper. |
| `Sender` internal | No TS field; runtime inherited | `internal IFluidDataObjectSender? Sender` | ➕ | Host integration helper. |
| `Registry` internal | No TS field; serializer inherited | `internal IFluidDataObjectRegistry? Registry` | ➕ | Host integration helper. |

### 4. SubDirectory class fields

#### Type: `SequenceData` / `SeqData`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `seq` | `number` required | `long Seq { get; set; }` | ✅ | `long` is safe for sequence numbers. |
| `clientSeq` | `clientSeq?: number` optional | `long ClientSeq { get; set; }` | ⚠️ | C# uses required field and `-1` sentinel; TS optionality is explicit. |
| container type | inline TS interface | `internal sealed class SeqData` | ➕ | Legitimate C# helper. |

#### Type: `SubDirectory` storage, identity, and tracking fields

| Field / Property | TS | C# | Status | Notes |
|---|---|---|---|---|
| disposed flag | `private _deleted = false` | Missing | ❌ | C# has no disposed state. High risk stale-reference bug; PARITY Findings 4 and 19. |
| string tag | public `[Symbol.toStringTag] = "SubDirectory"` | Missing | ❌ | Cosmetic, but absent. |
| sequenced subdirs | `_sequencedSubdirectories: Map<string, SubDirectory>` | `_subdirs: Dictionary<string, SubDirectory>` | ✅ | Equivalent map/dictionary. |
| subdir order | TS derives order by sorting combined map/pending entries with `seqDataComparator` | `_subdirOrder: List<string>` plus `SeqDataComparer` sort | ➕ | C# extra list; current code does sort final result by seq data, so list is mostly storage-order helper. |
| local creation sequence | `public localCreationSeq: number` | `private long _localCreationSeq` | ✅ | Equivalent counter, narrower visibility. |
| monitoring context | `private readonly mc: MonitoringContext` | Missing | ❌ | Telemetry-only; not data-model critical. |
| `seqData` | constructor private readonly `SequenceData` | `internal SeqData SeqData { get; }` | ✅ | Equivalent identity data, mutable contained fields in both. |
| `clientIds` | private readonly `Set<string>` | `internal HashSet<string> ClientIds { get; }` | ✅ | Equivalent collection. |
| parent pointer | No parent field; event relay registered on child emitters | `_parent: SubDirectory?` | ➕ | C# event-propagation helper; changes event data shape. |
| directory/root pointer | `private readonly directory: SharedDirectory` | `_root: SharedDirectory` | ✅ | Equivalent owner pointer. |
| runtime | `private readonly runtime: IFluidDataStoreRuntime` | No direct runtime; uses root sender | ⚠️ | Host-specific replacement. |
| serializer | `private readonly serializer: IFluidSerializer` | No direct serializer; uses registry/serializer helpers | ⚠️ | Host-specific replacement. |
| absolute path | `public readonly absolutePath: string` | `_absolutePath`, `public AbsolutePath` | ✅ | Equivalent. |
| sequenced storage | `sequencedStorageData: Map<string, unknown>` | `_storage: Dictionary<string, object?>` | ✅ | Equivalent stable key/value storage. |
| pending storage | `pendingStorageData: PendingStorageEntry[]` | `_pendingStorageData: List<PendingStorageEntry>` | ✅ | Equivalent list shape. |
| pending subdir data | `pendingSubDirectoryData: PendingSubDirectoryEntry[]` | `_pendingSubDirectoryData: List<PendingSubDirectoryEntry>` | ✅ | Equivalent list shape. |
| pending by client sequence | No TS field; TS runtime returns exact local metadata | `_pendingByClientSequenceNumber: Dictionary<long, object>` | ➕ | C#-only metadata substitute; cannot fully represent TS rollback/resubmit metadata. |
| lock | No TS lock | `_lock: object` | ➕ | Legitimate C# concurrency helper. |
| `internalIterator` function field | `private readonly internalIterator = () => ...` | Method materializes lists | ⚠️ | Field missing, method equivalent differs under mutation. |
| `getOptimisticValue` function field | `private readonly getOptimisticValue = ...` | `GetOptimisticValueNoLock` method | ✅ | Method-equivalent. |
| `optimisticallyHas` function field | `private readonly optimisticallyHas = ...` | `OptimisticallyHasNoLock` method | ✅ | Method-equivalent. |
| `getOptimisticSubDirectory` function field | TS supports `getIfDisposed` parameter | `GetOptimisticSubDirectoryNoLock` lacks disposed concept | ⚠️ | Equivalent only because C# lacks disposed state. |
| `sequencedSubdirectories` getter | `ReadonlyMap<string, SubDirectory>` | `GetSequencedSubDirectoryInternal(name)` only | ⚠️ | C# has no map-returning property, but path resolution can access by key. |
| local events | inherited `TypedEventEmitter<IDirectoryEvents>` | CLR events on each `SubDirectory` | ⚠️ | Event model differs. |

### 5. Pending state types and local op metadata

#### Type: storage pending entries

| Type / Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `PendingStorageEntry` union | `PendingKeyLifetime \| PendingKeyDelete \| PendingClear` | abstract `PendingStorageEntry`; list excludes `PendingKeySet` directly | ✅ | Equivalent union/class hierarchy. |
| `PendingStorageEntry.subdir` | each TS member has `subdir` | base `Subdir { get; }` | ✅ | Equivalent pointer. |
| `PendingStorageEntry.type` | string discriminant on every TS member | Runtime class type, no `Type` field | ⚠️ | C# uses class hierarchy instead of structural discriminant. |
| `PendingStorageEntry.path` | every TS storage pending member has `path: string` | Missing; path derivable from `Subdir.AbsolutePath` | ⚠️ | Derived, but not stored branch-for-branch. |
| `PendingKeySet.type` | literal `"set"` | Missing | ❌ | Type represented by class only. |
| `PendingKeySet.path` | `string` | Missing | ❌ | Derivable from `Subdir`. |
| `PendingKeySet.value` | `unknown` | `object? Value` | ✅ | Equivalent except `undefined` vs `null`. |
| `PendingKeySet.subdir` | `SubDirectory` | `Subdir` | ✅ | Equivalent. |
| `PendingKeySet.lifetime` | No TS field | `PendingKeyLifetime Lifetime` | ➕ | C# back-pointer for ack cleanup. |
| `PendingKeySet.clientSequenceNumber` | No TS field; local metadata object identity used | `long ClientSequenceNumber` sentinel `-1` | ➕ | C# metadata substitute. |
| `PendingKeyDelete.type` | literal `"delete"` | Missing | ❌ | Class type is discriminator. |
| `PendingKeyDelete.path` | `string` | Missing | ❌ | Derivable from `Subdir`. |
| `PendingKeyDelete.key` | `string` | `string Key` | ✅ | Equivalent. |
| `PendingKeyDelete.subdir` | `SubDirectory` | base `Subdir` | ✅ | Equivalent. |
| `PendingKeyDelete.clientSequenceNumber` | No TS field | `long ClientSequenceNumber` | ➕ | C# metadata substitute. |
| `PendingClear.type` | literal `"clear"` | Missing | ❌ | Class type is discriminator. |
| `PendingClear.path` | `string` | Missing | ❌ | Derivable from `Subdir`. |
| `PendingClear.subdir` | `SubDirectory` | base `Subdir` | ✅ | Equivalent. |
| `PendingClear.clientSequenceNumber` | No TS field | `long ClientSequenceNumber` | ➕ | C# metadata substitute. |
| `PendingKeyLifetime.type` | literal `"lifetime"` | Missing | ❌ | Class type is discriminator. |
| `PendingKeyLifetime.key` | `string` | `string Key` | ✅ | Equivalent. |
| `PendingKeyLifetime.path` | `string` | Missing | ❌ | Derivable from `Subdir`. |
| `PendingKeyLifetime.keySets` | `PendingKeySet[]` non-empty while live | `List<PendingKeySet>` | ✅ | Equivalent collection. |
| `PendingKeyLifetime.subdir` | `SubDirectory` | base `Subdir` | ✅ | Equivalent. |

#### Type: pending subdirectory entries

| Type / Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `PendingSubDirectoryEntry` union | `PendingSubDirectoryCreate \| PendingSubDirectoryDelete` | abstract `PendingSubDirectoryEntry` | ✅ | Equivalent union/class hierarchy. |
| `type` discriminant | `"createSubDirectory"` / `"deleteSubDirectory"` | Class type, no field | ⚠️ | C# class hierarchy replaces structural discriminant. |
| `subdirName` | `string` | `string SubdirName` | ✅ | Equivalent. |
| parent subdir | TS create/delete pending entries store `subdir`; for delete this is the parent `this` | base `ParentSubdir` plus subclass `Subdir` | ⚠️ | C# splits parent and created/deleted child; useful but not TS shape. |
| `PendingSubDirectoryCreate.subdir` | created `SubDirectory` | `Subdir` created child | ✅ | Equivalent. |
| `PendingSubDirectoryDelete.subdir` | TS stores parent subdir (`this`) | C# `Subdir` stores deleted child; `ParentSubdir` stores parent | ⚠️ | C# field maps more closely to TS delete local metadata than TS pending entry. |
| `clientSequenceNumber` | No TS field | `long ClientSequenceNumber` | ➕ | C# metadata substitute. |

#### Type: local op metadata

| Type / Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `ICreateSubDirLocalOpMetadata.type` | literal `"createSubDir"` | Missing | ❌ | C# uses pending entry class + client sequence map. |
| `ICreateSubDirLocalOpMetadata.parentSubdir` | `SubDirectory` | `PendingSubDirectoryEntry.ParentSubdir` | ✅ | Present in C# pending entry. |
| `IDeleteSubDirLocalOpMetadata.type` | literal `"deleteSubDir"` | Missing | ❌ | C# uses pending entry class + client sequence map. |
| `IDeleteSubDirLocalOpMetadata.subDirectory` | `SubDirectory \| undefined` deleted child | `PendingSubDirectoryDelete.Subdir` | ✅ | Equivalent for normal delete path; no explicit undefined state. |
| `IDeleteSubDirLocalOpMetadata.parentSubdir` | `SubDirectory` | `ParentSubdir` | ✅ | Equivalent. |
| `EditLocalOpMetadata` | `PendingKeySet \| PendingKeyDelete` | Pending object in `_pendingByClientSequenceNumber` | ⚠️ | C# reconstructs by client seq, not runtime metadata identity. |
| `ClearLocalOpMetadata` | `PendingClear` | `PendingClear` with client seq | ⚠️ | Present but accessed through C# ack map. |
| `DirectoryLocalOpMetadata` | union used by submit/resubmit/rollback | No explicit union; `_pendingByClientSequenceNumber` and `_pendingLocalOpSubdirectories` | ⚠️ | High-risk structural mismatch for rollback/resubmit. Related to PARITY Finding 6. |

### 6. Value and serialization types

#### Type: `ILocalValue`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `value` | `readonly value: unknown` | `object? Value { get; }` | ✅ | Equivalent local in-memory value. |

#### Type: `ISerializableValue` / `SerializableValue`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | `string` required | `string Type { get; set; } = string.Empty` | ⚠️ | Present but default empty permits malformed value wrappers. |
| `value` | `any` required | `object? Value { get; set; }` | ✅ | Equivalent for JSONable values except JS `undefined`. |
| mutability | TS object is mutable; `migrateIfSharedSerializable` mutates legacy values | C# class mutable | ✅ | Equivalent mutability. |

#### Type: `ISerializedValue` / C# `SerializedValue`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | `string` required | `string Type { get; set; } = string.Empty` | ⚠️ | Present but default empty permits malformed wrappers. |
| `value` | `string \| undefined` | `object? Value { get; set; }` | ⚠️ | Important mismatch: TS serialized value is a serialized string; C# DTO stores materialized object/null. Snapshot writer is absent, but the declared type is not equivalent. |

#### Type: value type constants / migration

| Field / Type | TS | C# | Status | Notes |
|---|---|---|---|---|
| `ValueType.Plain` | enum string name `"Plain"` | `ValueType.Plain = "Plain"` | ✅ | Equivalent wire string. |
| `ValueType.Shared` | legacy string name `"Shared"` | `ValueType.Shared = "Shared"` | ✅ | Equivalent constant. |
| legacy shared-value migration shape | `migrateIfSharedSerializable` rewrites `Shared` to handle-shaped `Plain` | `MigrateIfSharedSerializable` stub throws and is not applied on read paths | ⚠️ | Type exists but legacy value structural conversion is not equivalent. Related to PARITY Finding 17. |
| `SerializedValue` as snapshot storage value | TS snapshot `storage` uses `ISerializableValue`, not `ISerializedValue` | C# `DirectorySnapshotDto.Storage` uses `SerializedValue` | ⚠️ | C# class name does not match TS semantic role; it stores parsed/materialized values. |

### 7. Snapshot DTOs

#### Type: `ICreateInfo` / `DirectoryCreateInfo`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `csn` | `number` required | `long Csn { get; set; }` | ✅ | `long` is safer for server sequence numbers. Parser requires numeric `csn` when `ci` exists. |
| `ccIds` | `string[]` required | `string[] CcIds = Array.Empty<string>()` | ⚠️ | Parser allows missing `ccIds` and defaults empty. TS type requires it once `ci` exists. |

#### Type: `IDirectoryDataObject` / `DirectorySnapshotDto`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `storage` | optional `Record<string, ISerializableValue>` | non-null `Dictionary<string, SerializedValue> Storage = new()` | ⚠️ | Absence and empty object are conflated in C# DTO. Values use C# `SerializedValue` object form. |
| `subdirectories` | optional `Record<string, IDirectoryDataObject>` | non-null `Dictionary<string, DirectorySnapshotDto> Subdirectories = new()` | ⚠️ | Absence and empty object are conflated in C# DTO. |
| `ci` | optional `ICreateInfo` | nullable `DirectoryCreateInfo? CreateInfo` | ✅ | Custom parser maps wire `ci` to `CreateInfo`. Current C# no longer ignores `.ci`; older PARITY Finding 2 is partly remediated. |
| recursive shape | nested directory objects | nested `DirectorySnapshotDto` | ✅ | Equivalent recursion. |
| arbitrary unknown fields | TS `populate` ignores unknown fields | C# reader ignores unknown fields | ✅ | Equivalent tolerance. |

#### Type: `IDirectoryNewStorageFormat` / C# `IDirectoryNewStorageFormat`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `blobs` | `string[]` required | `string[] Blobs = Array.Empty<string>()`, `[JsonPropertyName("blobs")]` | ✅ | Parser requires root `blobs` to be an array to choose new format. |
| `content` | `IDirectoryDataObject` required | `DirectorySnapshotDto Content = new()`, `[JsonPropertyName("content")]` | ✅ | Parser requires `content` for blob-split snapshots. |
| class name | TS exported interface | C# class named `IDirectoryNewStorageFormat` | ⚠️ | C# uses interface-style `I` prefix on a class. Cosmetic but potentially confusing. |

#### Type: legacy snapshot shape

| Shape | TS | C# | Status | Notes |
|---|---|---|---|---|
| Old header | If parsed data has no array `blobs`, load as `IDirectoryDataObject` | If no array `blobs`, `ReadDirectory(root)` | ✅ | Legacy simple snapshot shape is supported. |
| Blob fragments | TS `populate(newFormat.content)`, then each blob content | C# parses fragments and merges into one DTO | ⚠️ | Final DTO can be equivalent, but shape/algorithm differs. Public `LoadFromSnapshot` rejects non-empty root. Related to PARITY Finding 23. |
| Snapshot writing DTO | TS writes `{ blobs, content }`, `ci` on every directory | C# has no snapshot writer/summarizer | ➕ | Load-only C# port; related to PARITY Finding 16. |

### 8. Handle wire shape

#### Type: `ISerializedHandle` / C# handle helpers

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | literal `"__fluid_handle__"` required | `HandleWireFormat.TypePropertyName`, writer emits `"__fluid_handle__"` | ✅ | Equivalent field name and literal. |
| `url` | `string` required | `UrlPropertyName`, `SerializedFluidHandle.Url` | ✅ | Equivalent. |
| `payloadPending` | optional `readonly payloadPending?: true` | Missing on `SerializedFluidHandle`; reader ignores extra; writer never emits | ❌ | Current TS type includes this optional field. C# cannot preserve/write it. High risk for pending-payload handles. |
| in-memory unresolved handle | TS serialized handle remains object until parsed/resolved by serializer | `SerializedFluidHandle` POCO with `Url` only | ⚠️ | Legitimate C# placeholder, but drops `type` and optional payload metadata in memory. |
| serialized construction helper | TS uses serializer/handle context | C# `CreateSerializedHandleWireValue(url)` dictionary | ✅ | Produces `{ type, url }`; no `payloadPending`. |
| registry dependency | TS serializer/bind resolves handles | C# requires `IFluidDataObjectRegistry` for live object egress | ➕ | Host-specific dependency. Related to PARITY Finding 25. |

## Reverse pass: C# types or state with no exact TS counterpart

| C# type / field | Classification | Notes |
|---|---|---|
| `ValueChangedEventArgs`, `SubDirectoryEventArgs` | Legitimate C# API adapter, but payload mismatch risk | They replace TS listener tuple parameters and event-object split. |
| `DirectoryOpType`, `MapOpType` enums | Legitimate C# discriminant adapter | Wire serializer maps directory op enum to TS strings. |
| `DirectoryOperation` / `MapOperation` abstract bases | Legitimate class-hierarchy adapter | TS uses structural unions. |
| `Map*Operation` classes | Companion-file port, out of SharedDirectory op scope | TS counterparts exist in `internalInterfaces.ts` for SharedMap, not `directory.ts`. |
| `DirectorySnapshotDto`, `DirectoryCreateInfo` | Legitimate DTO adapters | Map to `IDirectoryDataObject`/`ICreateInfo` but use PascalCase and non-null default dictionaries. |
| `SeqData` class | Legitimate C# helper | Maps to TS inline `SequenceData`; note `clientSeq` sentinel mismatch. |
| `PendingStorageEntry` / `PendingSubDirectoryEntry` abstract classes | Legitimate union adapter | TS uses literal discriminated unions. |
| `_lock` | Legitimate C# concurrency helper | No TS equivalent. |
| `_parent` | Legitimate C# event propagation helper | No TS field; TS registers child listeners instead. |
| `_subdirOrder` | C# helper with parity risk | Current code sorts by `SeqData` after collection, but this extra order list can still become a divergence point if used elsewhere. |
| `_pendingByClientSequenceNumber` | C# metadata substitute, high-risk | Does not carry full TS local op metadata object graph for rollback/resubmit. |
| `SharedDirectory._pendingLocalOpSubdirectories` | C# metadata substitute, high-risk | Stores only client sequence to subdir target; narrower than TS `DirectoryLocalOpMetadata`. |
| `IFluidDataObject`, `IFluidDataObjectMessageHandler`, `IFluidDataObjectSender`, `IFluidDataObjectRegistry` | Host shims | Expected for waccobalt/server integration; not TS SharedDirectory data structures. |
| `SequenceNumber`, `SequencedDocumentMessageDescriptor`, `OpOrigin` | Host message descriptor shims | Map to fields consumed from TS message envelopes, but use sentinel helpers such as `HasClientSequenceNumber != 0`. |
| `SerializedFluidHandle` | Legitimate unresolved-handle adapter with missing metadata | Drops TS `type` and `payloadPending` after parse. |
| `HandleWireFormat` | Helper, not DTO | Writes/reads handle shapes but only models `{type,url}`. |

## Silent-bug risk findings

1. **Handle `payloadPending?: true` is missing in C# handle structures.** TS `ISerializedHandle` currently has optional `payloadPending`; C# reads by URL and writes only `{ type, url }`. This is an optional-field omission on a wire shape and can silently change pending blob-handle semantics.
2. **C# deserializes malformed ops by defaulting required fields to empty strings/default values.** TS op types require `path`, `key`/`subdirName`, and set `value`; C# `DirectoryOpSerializer` defaults missing strings to `""` and missing set value to `new SerializableValue()`. This can turn invalid wire data into root-path or empty-key mutations.
3. **`ISerializedValue.value` type is not equivalent.** TS `ISerializedValue.value` is `string | undefined`; C# `SerializedValue.Value` is `object?`. This is especially risky if snapshot writing or any code expects the TS two-stage serialize-then-JSON.parse model.
4. **Subdirectory lifecycle state is absent.** TS has `_deleted`, `disposed`, `disposed`/`undisposed` events, and throw guards; C# retained `SubDirectory` references remain structurally mutable. This is already tracked by PARITY Findings 4 and 19.
5. **Local op metadata is structurally narrower in C#.** TS carries exact metadata variants (`PendingKeySet`, `PendingKeyDelete`, `PendingClear`, create/delete-subdir metadata) through submit, rollback, and resubmit. C# reconstructs through client sequence maps and pending entries, with no explicit metadata union and no resubmit/rollback surface. Related to PARITY Finding 6.
6. **Event payload shape differs for subdirectory events.** TS uses relative `path`; C# exposes `ParentPath` plus `SubdirName`. Nested path consumers can observe different event data unless they reconstruct it exactly. Related to PARITY Finding 13.
7. **Snapshot optionality is normalized.** TS distinguishes omitted `storage`/`subdirectories` from empty objects in DTO shape; C# DTOs use non-null empty dictionaries. Usually benign, but it is a structural mismatch for branch-by-branch comparisons.
8. **Numeric field choices are mixed.** Sequence numbers use `long` in C# (good), but public counts use `int`; TS `number` does not have the same overflow behavior for large maps/subdirectory counts.

## Cross-reference with `PARITY-AUDIT.md`

| Data-structure finding here | PARITY-AUDIT.md reference | Current audit note |
|---|---|---|
| Instance identity fields (`seqData`, `clientIds`, local metadata maps) | Finding 1 | Still relevant structurally; current code has `SeqData`, `ClientIds`, `CreateSnapshotClientIds`, and local client IDs, but C# metadata remains narrower. |
| Snapshot `.ci` DTO | Finding 2 | Current code now parses `CreateInfo`; remaining structural mismatch is optional/default handling, not total absence. |
| Subdirectory ordering fields | Finding 3 | Current code now has `SeqDataComparer`; `_subdirOrder` remains an extra helper but final enumeration sorts. |
| Disposed/deleted state | Findings 4 and 19 | Still a clear missing data-structure family in C#. |
| Local op metadata substitute | Finding 6 | Still structurally different and high-risk. |
| Rollback/resubmit metadata | Findings 7 and 8 | Missing APIs are behavior, but root cause is absent TS-shaped metadata and handler table. |
| Event payload/propagation shape | Findings 11, 13 | Still structurally different: no clear/cleared events; subdir payload split. |
| Snapshot write DTOs | Finding 16 | Still load-only in C#; writer DTO flow absent. |
| Legacy `Shared` value migration | Finding 17 | Type constants exist; migration shape is not implemented/applied. |
| JSON value materialization | Finding 18 | Appears remediated in current `DirectoryOpSerializer.MaterializeJsonValue`; not a current type mismatch except `SerializedValue` naming/type. |
| Map-like public API narrowing | Finding 24 | Still structurally different by design. |
| Handle serialization dependency | Finding 25 | Still relevant; this audit adds missing `payloadPending` field. |

## Recommended actions

1. **Fix high-risk wire/data-model bugs:** model or deliberately reject `ISerializedHandle.payloadPending`, and make op deserialization fail on missing required fields instead of defaulting them.
2. **Close lifecycle/state gaps:** add C# deleted/disposed state, disposed/undisposed events if needed, and throw guards for stale `SubDirectory` references.
3. **Align local-op metadata:** introduce explicit C# metadata records mirroring TS `DirectoryLocalOpMetadata`, or document exactly why client-sequence maps are sufficient for the host runtime.
4. **Clarify value DTO naming and shape:** either rename/split C# `SerializedValue` vs `SerializableValue`, or document that C# stores materialized snapshot values and has no TS `ISerializedValue` equivalent until snapshot writing is implemented.
5. **Normalize event payload compatibility:** expose a single relative `path` for subdirectory events or provide a guaranteed helper that reconstructs TS path semantics from `ParentPath` + `SubdirName`.
6. **Document deliberate C# extensions:** `SeqData`, abstract pending-entry bases, `_lock`, `_parent`, host sender/registry, and handle placeholder are reasonable C# adaptations but should be described as adapters, not TS-equivalent DTOs.

## Files opened

Primary files:

- `.claude/CLAUDE.md`
- `packages/dds/map/src/directory.ts`
- `packages/dds/map/src/interfaces.ts`
- `packages/dds/map/src/internalInterfaces.ts`
- `packages/dds/map/src/localValues.ts`
- `packages/runtime/runtime-utils/src/handles.ts`
- `packages/dds/map/src/csharp-port/SharedDirectory/cs-out/Interfaces.cs`
- `packages/dds/map/src/csharp-port/SharedDirectory/cs-out/SubDirectory.cs`
- `packages/dds/map/src/csharp-port/SharedDirectory/cs-out/SharedDirectory.cs`
- `packages/dds/map/src/csharp-port/SharedDirectory/cs-out/LocalValues.cs`
- `packages/dds/map/src/csharp-port/SharedDirectory/cs-out/DirectoryOpSerializer.cs`
- `packages/dds/map/src/csharp-port/SharedDirectory/cs-out/DirectorySnapshotLoader.cs`
- `packages/dds/csharp-port-common/FluidInterfaces.cs`
- `packages/dds/csharp-port-common/HandleWireFormat.cs`
- `packages/dds/csharp-port-common/SerializedFluidHandle.cs`
- `packages/dds/map/src/csharp-port/PARITY-AUDIT.md`

Additional files surfaced by repository search output while locating shared/common types:

- `packages/dds/csharp-port-common/README.md`
- `packages/dds/csharp-port-common/FluidMessageTypes.cs`
- `packages/dds/csharp-port-common/OcsException.cs`
- `packages/dds/csharp-port-common/FluidObjectId.cs`

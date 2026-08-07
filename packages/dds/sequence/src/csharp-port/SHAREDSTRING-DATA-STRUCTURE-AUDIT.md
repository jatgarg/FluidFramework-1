# SharedString Data Structure Audit — TS vs C#

**Base:** microsoft/main (`sharedString.ts`, `sequence.ts`, `sequenceDeltaEvent.ts`, interval map wrappers, merge-tree snapshot files)
**Port:** transpiledir `packages/dds/sequence/src/csharp-port/SharedString/cs-out/`
**Audited:** 2026-08-06T14:12:48-07:00
**Method:** field-by-field comparison of top-level SharedString, event payloads, op envelopes, interval-map boundary shapes, and snapshot DTOs. Static/read-only source audit; no build, tests, or lint run.

## Summary

Counts below are row-level field/member comparisons, not pass/fail type counts. `Types` counts the major named structures audited in each category.

| Category | Types | ✅ | ⚠️ | ❌ | ➕ |
|---|---:|---:|---:|---:|---:|
| Public SharedString and sequence API | 8 | 31 | 16 | 9 | 11 |
| Sequence delta and public event payloads | 6 | 14 | 13 | 5 | 9 |
| Op envelope and serializer DTOs | 12 | 41 | 16 | 3 | 8 |
| IntervalCollectionMap boundary | 5 | 9 | 9 | 5 | 5 |
| Snapshot DTOs and loader | 16 | 46 | 24 | 5 | 12 |
| Handle/property JSON adjuncts | 4 | 9 | 5 | 1 | 3 |
| **Total** | **51** | **150** | **83** | **28** | **48** |

## Findings by category

### 1. Public SharedString and sequence API

#### Type: `ISharedString` / C# `SharedString` text and marker APIs

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `insertText(pos,text,props?)` | Public method with optional `PropertySet` | `InsertText(int,string,PropertySet?)` | ✅ | Current C# has the optional props parameter; old PARITY Finding SS1 appears remediated. |
| `insertMarker(pos,refType,props?)` | Public method | `InsertMarker(int,ReferenceType,PropertySet?)` | ✅ | Allows marker creation with optional props; no marker id requirement in this surface. |
| `insertTextRelative(relativePos1,text,props?)` | Uses `IRelativePosition` | `InsertTextRelative(object,...)` | ⚠️ | C# accepts loose `object` and resolves immediately by marker id; TS carries a typed relative shape into merge-tree APIs. |
| `insertMarkerRelative(relativePos1,refType,props?)` | Uses `IRelativePosition` | `InsertMarkerRelative(object,...)` | ⚠️ | Same loose object adapter and immediate resolution. |
| `replaceText(start,end,text,props?)` | Inserts at `Math.max(start,end)`, removes only if `start < end` | `ReplaceText(int,int,string,PropertySet?)` | ✅ | Equivalent top-level algorithm. |
| `removeText(start,end)` | Public method | `RemoveText(int,int)` plus `DeleteText` alias | ✅ | TS name exists; `DeleteText` is a C# extra. |
| `annotateMarker(marker,props)` | Public marker-specific API | Missing | ❌ | C# has `OpBuilder.CreateAnnotateMarkerOp`, but no public `AnnotateMarker`. |
| `searchForMarker(startPos,markerLabel,forwards?)` | Label is required second parameter | `SearchForMarker(int,string,bool)` | ✅ | TS-shaped overload exists; old bool-first overload remains extra. |
| bool-first marker search | No TS counterpart | `SearchForMarker(int,bool,string?)` | ➕ | Adapter/legacy overload; may mask missing labels by returning `null`. |
| `getText(start?,end?)` | Optional bounds | `GetText()` and `GetText(int,int)` | ✅ | Equivalent through overloads. |
| `getTextWithPlaceholders(start?,end?)` | Placeholder-aware extraction | `GetTextWithPlaceholders(int?,int?)` | ✅ | Present. |
| `getTextRangeWithMarkers(start,end)` | Marker-aware extraction with `*` placeholder behavior | `GetTextRangeWithMarkers(int,int)` | ✅ | Present. |
| `getMarkerFromId(id)` | `ISegment | undefined` | `Marker? GetMarkerFromId(string)` | ✅ | C# narrows return type to marker, which matches practical lookup. |
| `ISharedString` factory/type attributes | `SharedStringFactory.Type`, `snapshotFormatVersion`, `packageVersion` | No factory/attributes in scoped files | ❌ | Host construction is plain C# object; missing only if C# must expose Fluid channel factory metadata. |

#### Type: inherited `ISharedSegmentSequence<T>` surface / C# top-level wrappers

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `getLength()` | `number` | `int GetLength()` | ✅ | Equivalent for current WN-sized strings; numeric range differs. |
| historical length | TS client supports perspective/refSeq APIs internally | `GetLength(long refSeq,string?,long?)` | ➕ | Useful C# adapter; not a public TS `ISharedString` method. |
| `getPosition(segment)` | Returns `number`, `-1` if absent | `int? GetPosition(ISegment)` | ⚠️ | Nullability differs from TS sentinel. Related to PARITY M9/L7. |
| historical `getPosition` | Not public on TS interface | `GetPosition(segment,refSeq,clientId,localSeq)` | ➕ | Adapter for server-side perspectives. |
| `localReferencePositionToPosition(lref)` | Returns `number` | `int? LocalReferencePositionToPosition(...)` | ⚠️ | Null differs from TS numeric sentinel/position contract. |
| `getContainingSegment(pos)` | `{ segment, offset }` with possibly undefined members | nullable tuple `(segment, offsetInSegment)?` | ⚠️ | Same data, different object/null shape. |
| `getRangeExtentsOfPosition(pos)` | `{ posStart, posAfterEnd }` | tuple `(start,end)?` | ⚠️ | Field names differ; null vs undefined. |
| `getPropertiesAtPosition(pos)` | `PropertySet | undefined` | `PropertySet?` | ✅ | Equivalent nullable mapping. |
| `walkSegments(...)` | Public inherited traversal API | Missing | ❌ | Not exposed on top-level C# SharedString. |
| `insertAtReferencePosition(pos,segment)` | Public inherited API | Missing | ❌ | Relative/local-reference insertion surface remains absent. |
| `insertFromSpec(pos,spec)` | Public inherited API | Missing | ❌ | C# serializer can read segment specs, but public insertion-from-spec is absent. |
| `createLocalReferencePosition(...)` | Public inherited API | Missing | ❌ | Local reference subsystem exists lower down, but no top-level API. |
| `removeLocalReferencePosition(lref)` | Public inherited API | Missing | ❌ | Same top-level API gap. |
| `resolveRemoteClientPosition(...)` | Public inherited API | Missing | ❌ | No top-level remote-position resolver. |
| `getCurrentSeq()` | Public inherited API | Missing | ❌ | C# exposes collab window only indirectly through loader/client internals. |
| `annotateRange(start,end,props)` | Public inherited API | `AnnotateRange(int,int,PropertySet)` | ✅ | Equivalent normal annotation surface. |
| `annotateAdjustRange(start,end,adjust)` | Public inherited API | Missing | ❌ | Deferred per HANDOFF, not a bug for current WN scope. |
| `obliterateRange(start,end)` | Supports numeric and sided places | `ObliterateRange(int,int)` and `ObliterateRange(SequencePlace,SequencePlace)` | ✅ | Public surface exists; deeper semantics are merge-tree scope. |
| `groupOperation(groupOp)` | Deprecated public inherited API | `RunInBatch(Action)` | ⚠️ | C# host adapter batches local API calls instead of accepting a prebuilt group op. |

#### Type: `SharedStringClass` / C# `SharedString` state

| Field / Property | TS | C# | Status | Notes |
|---|---|---|---|---|
| owner id | constructor `public id: string` | `_id`, `Id` | ✅ | Equivalent object id. |
| merge-tree client | `protected client: Client` | `_client: Client` | ✅ | Equivalent core owner. |
| text helper | `mergeTreeTextHelper` | no helper; uses `_client.GetText` / marker placeholder helper | ⚠️ | Functionally close; structural helper is not ported. |
| runtime/serializer/handle | inherited `SharedObject` runtime fields | `_sender`, `_registry` host shims | ⚠️ | Intentional host adapter, not TS channel structure. |
| interval collection map | `private readonly intervalCollections: IntervalCollectionMap` | stored in merge-tree `Client`; outward `GetIntervalCollection` | ⚠️ | C# does not keep the map wrapper as a top-level field. |
| outbound interval handler registry | No TS counterpart | `_intervalCollectionsWithOutboundHandlers` | ➕ | Legitimate C# event adapter. |
| in-flight ref seq queue | `inFlightRefSeqs: Deque<number>` | pending association in `Client`, not top-level queue | ⚠️ | C# substitutes sender-returned sequence numbers and client pending state. |
| `ongoingResubmitRefSeq` | TS resubmit helper | Missing | ❌ | Resubmit is deferred, not a bug for current scope. |
| reentrancy guard | `guardReentrancy` configurable by runtime options | `_localMutationDepth` hard-fail guard | ⚠️ | Protects against reentry, but lacks TS option/logging mode. |
| `_lock` | No TS lock | `_lock` | ➕ | Legitimate C# concurrency helper. |
| `_batchOps` | No direct TS field; TS has group/local transaction APIs | `_batchOps` | ➕ | C# batching adapter. |

#### Type: construction, attach, load, reconnect, and snapshot APIs

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| factory creation | `SharedStringFactory.create/load` | constructor only | ⚠️ | WN host adapter is expected; not a Fluid channel factory. |
| `initializeLocalCore` / attach | inherited lifecycle methods | `ProcessDataObjectAttach` no-op | ⚠️ | Lifecycle parity is deferred. |
| `loadCore(storage)` | Loads interval blob + merge-tree content tree | `LoadFromSnapshot(string|DTO, blobResolver?)` | ⚠️ | Bespoke load helper, not channel storage abstraction. |
| catchup op processing | TS `loadCore` awaits loader-returned ops then `processMergeTreeMsg` | loader applies catchup inside `PopulateFromSnapshot` | ⚠️ | See silent-bug risk findings. |
| reconnect/resubmit | `reSubmitCore`, `reSubmitSquashed`, `rollback`, stashed ops | `RegeneratePendingOps()` only | ⚠️ | Rollback/resubmit/stashed ops deferred, not bugs per HANDOFF. |
| snapshot writing/summarize | `summarizeCore` / `summarizeMergeTree` | Missing | ❌ | Snapshot writing is explicitly deferred for server-side client. |
| GC serialization | `processGCDataCore` | Missing | ❌ | Fluid runtime lifecycle feature, deferred. |

#### Type: `IRelativePosition` / C# `RelativePosition`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `id` | optional `string` | `string Id = ""` | ⚠️ | C# class makes it required by validation/default empty. |
| `before` | optional `boolean`; false/undefined means after | `bool Before` | ✅ | Default false matches TS omitted behavior. |
| `offset` | optional positive number | `int Offset` | ⚠️ | C# default `0`; TS docs say positive `>= 1` when present. |
| accepted runtime shape | structural object | object/dictionary/`JsonElement`/property reflection | ➕ | Flexible adapter, but not a typed TS DTO. |

#### Type: deprecated `SharedSequence<T>` / C# non-SharedString subclasses

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `SharedSequence<T>` | Deprecated arbitrary sequence subclass | Missing | ❌ | Explicitly deferred/non-goal per HANDOFF: non-SharedString sequence subclasses are not in WN scope. |
| `SubSequence<T>` segment | Deprecated run segment | Missing | ❌ | Deferred, not a bug. |

### 2. Sequence delta and public event payloads

#### Type: `ISharedSegmentSequenceEvents` / C# events

| Event | TS | C# | Status | Notes |
|---|---|---|---|---|
| `sequenceDelta` | `(event: SequenceDeltaEvent, target)` | `OnSequenceDelta(object, SequenceDeltaEventArgs)` | ⚠️ | Event exists but payload shape differs. |
| `maintenance` | `(event: SequenceMaintenanceEvent, target)` | Missing | ❌ | Maintenance/ACK event payloads are not exposed. |
| `createIntervalCollection` | `(label, local, target)` | Missing top-level event | ❌ | C# `GetIntervalCollection` creates/returns collections but does not raise TS-shaped create event. |
| target | explicit listener parameter | .NET sender | ⚠️ | Equivalent-ish for .NET, but not same payload signature. |

#### Type: `SequenceEvent` / C# `SequenceDeltaEventArgs`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `deltaOperation` | required enum | `MergeTreeDeltaType? DeltaOperation` | ⚠️ | Nullable in C#; absent state has no TS equivalent. |
| `deltaArgs` | full merge-tree callback args | Missing | ❌ | C# exposes selected fields only; no `deltaSegments`/callback metadata object. |
| `ranges` | readonly `ISequenceDeltaRange[]` | `IReadOnlyList<SequenceDeltaRange>` | ✅ | List shape is present. |
| `clientId` | getter from merge-tree client | `string? ClientId` | ✅ | Present. |
| `first` / `last` | non-null lazy getters | nullable `First` / `Last` | ⚠️ | C# returns null for empty lists; TS asserts most non-maintenance deltas are non-empty. |
| sorted range semantics | sorted by `SortedSegmentSet` | generated in current delta order, with merge heuristics | ⚠️ | Usually equivalent for simple deltas, but not identical structure. |

#### Type: `SequenceDeltaEvent` / C# event args

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `opArgs` | required `IMergeTreeDeltaOpArgs` | Missing as envelope | ❌ | Important routing gap: TS consumers can inspect `opArgs.op`; C# only exposes `Op`. |
| `opArgs.op` | merge-tree op for this delta | `IMergeTreeOp? Op` | ⚠️ | Present but missing `sequencedMessage`, `rollback`, and grouped-op context. |
| `isLocal` | readonly boolean | `Local` and alias `IsLocal` | ✅ | Present. |
| group op behavior | each child op can emit its own event, with group op in args | C# batches local emits outside client callback; remote emits per `MergeTreeDelta` | ⚠️ | Better than prior simplified aggregate path, but still lacks TS `opArgs` group envelope. |
| `opType` string | Not in TS event object | `OpType` | ➕ | C# convenience field (`insert`, `remove`, etc.). |
| `Position`, `Length`, `Text`, `IsMarker`, `Marker` | Not top-level TS fields | Present | ➕ | Useful host shortcut; derived from ranges/segment in TS. |
| `AnnotatedProperties` | Not top-level TS field | Present | ➕ | C# shortcut for annotate ops; TS uses ranges/property deltas. |

#### Type: `ISequenceDeltaRange` / C# `SequenceDeltaRange`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `operation` | enum | `Operation` | ✅ | Equivalent. |
| `position` | start index | `Position` | ✅ | Equivalent name casing. |
| `segment` | `ISegment` snapshot at event time | `ISegment Segment` | ⚠️ | Local C# clones slices; remote uses delta segment. TS range uses merge-tree segment object from delta callback. |
| `propertyDeltas` | required `PropertySet` | nullable `PropertySet?` | ⚠️ | C# null means none; TS uses `{}` when no deltas. |
| length | implied by `segment.cachedLength` | explicit `Length` | ➕ | Helpful host field but extra DTO member. |
| op type string | none | `OpType` | ➕ | C# convenience. |

#### Type: `SequenceMaintenanceEvent`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| event object | `SequenceMaintenanceEventClass` | Missing | ❌ | No public maintenance event surface. |
| `opArgs` | optional, defined for ACK maintenance | Missing | ❌ | Ack metadata not exposed. |
| `deltaArgs` | maintenance callback args | Missing | ❌ | No maintenance ranges/operation shape. |
| deferred classification | Full ACK/maintenance parity | Not implemented | ⚠️ | This is not listed as WN phase-1 API, but data structure is absent. |

### 3. Op envelope and serializer DTOs

#### Type: `MergeTreeDeltaType`

| Discriminant | TS | C# | Status | Notes |
|---|---|---|---|---|
| insert | `0` | `Insert = 0` | ✅ | Wire value matches. |
| remove | `1` | `Remove = 1` | ✅ | Wire value matches. |
| annotate | `2` | `Annotate = 2` | ✅ | Wire value matches. |
| group | `3` | `Group = 3` | ✅ | Wire value matches. |
| obliterate | `4` | `Obliterate = 4` | ✅ | Wire value matches. |
| sided obliterate | `5` | `ObliterateSided = 5` | ✅ | Wire value matches. |
| interval ops | TS interval ops are map/value ops, not merge-tree delta codes | `IntervalAdd/Delete/Change/PropertyChanged = 10..13` | ➕ | Extra legacy C# discriminants. Top-level routing risk when emitted by `SharedString`. |

#### Type: `IMergeTreeOp` / `MergeTreeOp`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | numeric discriminant | `Type` enum property | ✅ | Equivalent after enum mapping. |
| structural union | TS interface union | abstract class hierarchy | ⚠️ | Legitimate C# union adapter. |
| local client sequence | Runtime metadata, not op field | `ClientSeq` property | ➕ | C# pending-op correlation helper; not serialized. |

#### Type: insert operation

| Field | TS `IMergeTreeInsertMsg` | C# `MergeTreeInsertMsg` / serializer | Status | Notes |
|---|---|---|---|---|
| `type` | `0` | writes/reads `0`, accepts string `"insert"` | ✅ | Writer uses numeric TS value; string read is tolerant. |
| `pos1` | optional if `relativePos1` exists | `int? Pos1` | ✅ | Serializer enforces `pos1` or `relativePos1`. |
| `relativePos1` | optional `IRelativePosition` | `object? RelativePos1` | ⚠️ | Wire field preserved, but runtime type is loose object. |
| `pos2` | optional | `int? Pos2` | ✅ | Optional writer/read support. |
| `relativePos2` | optional | `object? RelativePos2` | ⚠️ | Preserved as arbitrary JSON/object. |
| `seg` | optional in TS type but required by processing | `object? Seg`, deserializer requires present | ✅ | C# fails missing segment, which is safer. |

#### Type: remove operation

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | `1` | `Remove = 1` | ✅ | Matches. |
| `pos1` / `relativePos1` | one required by runtime | `Pos1` / `RelativePos1` | ✅ | Deserializer enforces one. |
| `pos2` / `relativePos2` | one required by runtime | `Pos2` / `RelativePos2` | ✅ | Deserializer enforces one. |

#### Type: annotate and annotate-adjust operations

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| normal annotate `type` | `2` | `Annotate = 2` | ✅ | Matches. |
| `pos1`/`relativePos1` | one required | supported | ✅ | Deserializer enforces one. |
| `pos2`/`relativePos2` | one required | supported | ✅ | Deserializer enforces one. |
| `props` | required for normal annotate | nullable `PropertySet? Props` | ⚠️ | C# accepts null/missing as `Props = null`; TS normal annotate requires props. |
| `adjust` | annotate-adjust alternative | Missing | ❌ | Deferred per HANDOFF (Wave 31 partial), not a bug for current WN scope. |
| `props` empty object | legal JSON object and semantically meaningful for writer parity | C# serializer writes `{}` when `PropertySet` is non-null empty | ✅ | The scoped serializer no longer strips empty props on op paths. Snapshot writer is absent in scoped files. |

#### Type: obliterate operations

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| numeric `type` | `4` | `Obliterate = 4` | ✅ | Matches. |
| numeric `pos1` / `pos2` | optional in TS interface, required by builder/runtime | C# requires both | ✅ | Safe strictness for numeric obliterate. |
| numeric relative fields | `never` | present as `object?` on interface but not serialized | ⚠️ | Extra runtime fields with no TS meaning; writer ignores for numeric obliterate. |
| sided `type` | `5` | `ObliterateSided = 5` | ✅ | Matches. |
| sided `pos1` / `pos2` | `{ pos, before }` required | `SequencePlace` writes `{pos,before}` | ✅ | Wire object shape matches. |
| start/end sentinels | TS uses special place semantics in helper code | C# maps `pos=-1,before=false` to Start and `pos=-1,before=true` to End | ⚠️ | Adapter is plausible, but sentinel mapping is not explicit in TS wire interface. |
| full semantics | TS has richer obliterate behavior | minimum surface only | ⚠️ | Deferred beyond minimum surface per HANDOFF; not a bug here. |

#### Type: group operation

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | `3` | `Group = 3` | ✅ | Matches. |
| `ops` | `IMergeTreeDeltaOp[]` | `List<MergeTreeOp>` | ⚠️ | C# list can include nested group and interval op types; serializer also permits them. |
| nested groups | TS type allows only delta ops, not group op inside `ops` | C# `IsMergeTreeGroupMember` allows `Group` | ⚠️ | Accepted shape is wider than TS. |
| interval members | TS interval map ops are outside merge-tree group | C# permits interval op members | ⚠️ | Top-level wire-shape drift risk if batched interval ops are flushed in a group. |

#### Type: op content envelope around `SharedString`

| Layer | TS | C# | Status | Notes |
|---|---|---|---|---|
| runtime content | message contents are op object (`type` numeric or `"act"`) | `opJson` string passed with separate `opTypeName` to `QueueDataObjectMessage` | ⚠️ | Host API shim; correctness depends on WAC envelope preserving JSON content exactly. |
| local submit | `submitLocalMessage(message, metadata)` | `_sender.QueueDataObjectMessage(_id, opTypeName, opJson)` | ⚠️ | C# has separate string routing layer not present in TS data shape. |
| local metadata | TS runtime carries `localOpMetadata` | C# stamps/associates `ClientSeq` through sender return | ⚠️ | Narrower than TS metadata; rollback/resubmit deferred. |
| remote process | `processMessage(envelope, content, local)` | `ProcessDataObjectOp(descriptor, opJson)` | ⚠️ | C# descriptor shim lacks full `ISequencedDocumentMessage` shape. |

#### Type: `ReferenceType`

| Flag | TS | C# | Status | Notes |
|---|---|---|---|---|
| `Simple` | `0x0` | `0x0` | ✅ | Matches. |
| `Tile` | `0x1` | `0x1` | ✅ | Matches; old PARITY O1 appears remediated. |
| `RangeBegin` | `0x10` | `0x10` | ✅ | Matches. |
| `RangeEnd` | `0x20` | `0x20` | ✅ | Matches. |
| `SlideOnRemove` | `0x40` | `0x40` | ✅ | Matches. |
| `StayOnRemove` | `0x80` | `0x80` | ✅ | Matches. |
| `Transient` | `0x100` | `0x100` | ✅ | Matches. |
| non-TS bits | none | none in current enum | ✅ | Old PARITY O10 appears remediated for this enum. |

### 4. IntervalCollectionMap boundary

Only the top-level wrapper/routing boundary is covered here; interval data structures and interval semantics are sibling-audit scope.

#### Type: `IntervalCollectionMap`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| storage map | `Map<string, IntervalCollection>` | hidden behind `Client.GetOrCreateIntervalCollection` | ⚠️ | C# does not expose/map the wrapper as a top-level field. |
| `size` | public getter | Missing | ❌ | Top-level map count not exposed. |
| `keys()` | labels iterator | Missing top-level; no `GetIntervalCollectionLabels` | ❌ | Public label iteration gap. |
| `values()` | values iterator | Missing | ❌ | Public collection iteration gap. |
| `get(key)` | creates if absent | `GetIntervalCollection(name, endpointType?)` | ✅ | Equivalent creation access, with C# endpoint-type adapter. |
| `events.createIntervalCollection` | emitted on create | Missing top-level event | ❌ | C# collection events are add/delete/change/property, not map creation. |

#### Type: `IMapOperation` interval op envelope

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | literal `"act"` | `SerializeIntervalCollectionOperation` writes `"act"`; regular `Serialize` writes numeric `10..13` | ⚠️ | The TS envelope exists in helper but is not the path used by `SharedString.FlushPendingIntervalOps`. |
| `key` | collection label | `CollectionName` mapped to `key` in helper | ✅ | Equivalent when helper is used. |
| `value` | value-type operation payload | helper record `IntervalMapOperationValue` | ✅ | Equivalent top-level container. |
| local emission path | `IntervalCollectionMap` submit callback emits `IMapOperation` | C# emits `IntervalOpMsg` through `SharedStringOpSerializer.Serialize(op)` | ⚠️ | High-risk top-level wire drift. |
| remote read path | `tryProcessMessage` recognizes `"act"` | C# deserializer recognizes `"act"` and maps to interval op messages | ✅ | Inbound TS map ops can be read at top-level. |

#### Type: `IIntervalCollectionTypeOperationValue`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `opName` | `"add"`, `"change"`, `"delete"` | same helper strings | ✅ | Matches in map helper. |
| `value` | serialized interval / delta | `IntervalMapPayload` | ⚠️ | Top-level fields map, but interval internals are out of scope. |
| property-change op | TS uses change/value payload shape | C# maps property-only `"change"` to `IntervalPropertyChangedOpMsg` | ⚠️ | Boundary heuristic is plausible but not a distinct TS discriminant. |
| `localOpMetadata` | typed add/change/delete metadata | C# does not carry equivalent at map boundary | ❌ | Rollback/resubmit/stashed interval metadata deferred. |

#### Type: interval collection snapshot wrapper

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| interval collection blob | `sequence.ts` stores map under separate `header` blob before merge-tree `content` tree | scoped C# snapshot loader reads only merge-tree content chunks | ❌ | Interval summary wrapper is outside this loader; note only at SharedString boundary. |
| `ISerializableIntervalCollection.type` | `"sharedStringIntervalCollection"` | not modeled in scoped top-level files | ❌ | Interval snapshot internals are sibling scope. |
| legacy value-type migration | skips legacy `Plain`/`Shared` handles | not modeled here | ⚠️ | Deferred legacy pre-2019 Shared value migration; not a bug. |

### 5. Snapshot DTOs and loader

#### Type: snapshot file layout

| Shape | TS | C# | Status | Notes |
|---|---|---|---|---|
| merge-tree content path | `sequence.ts` adds merge-tree summary under `content` tree | `LoadFromSnapshot` accepts header JSON directly plus `blobResolver` | ⚠️ | C# skips channel storage partition and expects caller to supply content blob data. |
| header chunk file name | `SnapshotLegacy.header = "header"` | `MergeTreeHeaderChunkMetadataDto.Id`; default legacy id `"header"` | ✅ | Same blob id. |
| body chunk names | V1 emits `body_0`, `body_1`, ...; legacy may use `body` | resolver uses `orderedChunkMetadata[i].Id` | ✅ | Uses header metadata ids, so names are data-driven. |
| catchup blob name | TS legacy default `"catchupOps"` or option-specific blob | C# accepts many metadata properties (`catchupOps`, `catchUpOps`, `catchupOpsBlobNames`, etc.) | ⚠️ | Tolerant read is useful; C# metadata field names are broader than TS. |
| interval map blob | TS sequence `header` blob for interval collections | not read by `SharedStringSnapshotLoader` | ❌ | Interval snapshot wrapper is outside scoped loader. |

#### Type: `MergeTreeHeaderMetadata`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `totalLength` | required number | `int TotalLength` required on read | ✅ | Field name and semantics match; numeric width differs. |
| `totalSegmentCount` | required number | `int TotalSegmentCount` required on read | ✅ | Matches. |
| `orderedChunkMetadata` | required array | `List<MergeTreeHeaderChunkMetadataDto>`; optional on read | ⚠️ | C# accepts missing and defaults empty; TS type requires it. |
| `sequenceNumber` | required number | `long SequenceNumber` required | ✅ | Matches; `long` is safer. |
| `minSequenceNumber` | required by interface; loader falls back to sequenceNumber when omitted | `MinSequenceNumber = optional ?? sequenceNumber` | ✅ | The `minSequenceNumber` integration bug is fixed in current C# reader. |
| `latestSequenceNumber` | Not in current listed TS snapshot DTOs | Not modeled | ✅ | User-mentioned risk field does not appear in current `snapshotChunks.ts`/loader. |
| `orderedRemovedSeqs` | Not in current listed TS snapshot DTOs | Not modeled | ✅ | Not present in current top-level snapshot DTOs. |
| `packageVersion` | Channel factory attribute, not merge-tree header metadata | Not in snapshot DTO | ✅ | C# lacks factory attributes, covered under public API; not a snapshot header field. |

#### Type: `MergeTreeHeaderChunkMetadata`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `id` | required string | `string Id` | ✅ | Matches. |
| string shorthand | Not in interface, but storage ids are strings in list output | C# also accepts string array items | ➕ | Tolerant adapter. |
| extra metadata fields | none | ignored | ✅ | Unknown fields skipped by JSON parser. |

#### Type: `MergeTreeChunkV1`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `version` | literal `"1"` | `Version = "1"`, required if modern | ✅ | Matches. |
| `startIndex` | required number | `int StartIndex` required | ✅ | Matches. |
| `segmentCount` | required number | `int SegmentCount` required | ✅ | Matches. |
| `length` | required number | `int Length` required | ✅ | Matches; C# validates against segment lengths. |
| `segments` | `JsonSegmentSpecs[]` required | `List<SharedStringSnapshotSegmentDto>` | ⚠️ | C# materializes segment DTOs instead of retaining raw JSON spec union. |
| `headerMetadata` | `MergeTreeHeaderMetadata | undefined` | `MergeTreeHeaderMetadataDto?` | ✅ | Nullable maps to undefined. |
| `attribution?` | optional serialized attribution collection | Missing | ❌ | Deferred per HANDOFF; not a bug unless WN requires attribution. |
| unknown segment specs | TS `segmentFromSpec` can throw/return undefined by caller | C# marks unknown as `IsUnknown` and skips | ⚠️ | More tolerant load, but not raw-shape preserving. |

#### Type: legacy `MergeTreeChunkLegacy`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `version` | `undefined` | absent `version` triggers legacy read | ✅ | Matches. |
| `chunkStartSegmentIndex` | required | `StartIndex` | ✅ | Matches. |
| `chunkSegmentCount` | required | `SegmentCount` | ✅ | Matches. |
| `chunkLengthChars` | required | `Length` | ✅ | Matches. |
| `totalLengthChars` | optional header total | used to build legacy metadata | ✅ | Matches. |
| `totalSegmentCount` | optional header total | used to build legacy metadata | ✅ | Matches. |
| `chunkSequenceNumber` | optional header seq | required if legacy metadata present | ✅ | Matches loader expectations. |
| `chunkMinSequenceNumber` | optional legacy min seq | falls back to `chunkSequenceNumber` | ✅ | Same minSeq fallback pattern. |
| `segmentTexts` | required segment specs | read as segments | ✅ | Matches. |
| `attribution?` | optional | Missing | ❌ | Deferred attribution. |

#### Type: `IJSONSegmentWithMergeInfo`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `json` | wrapped segment JSON spec | payload element unwrapped into DTO | ✅ | Present in reader, not retained raw. |
| `client` | optional long client id | `ClientId` from `client` or `clientId` | ✅ | Tolerant read. |
| `seq` | optional insertion seq | `long? Seq` | ✅ | Matches. |
| `removedSeq` | optional first set-remove seq | `RemovedSeq` | ✅ | Matches. |
| `removedClientIds` | optional set-remove client list | `RemoveStamps` generated | ✅ | Current C# supports multi-client set-remove stamps. |
| legacy `removedClient` | back-compat field | `ReadOptionalRemovedClientId` | ✅ | Matches TS back-compat intent. |
| `movedSeq` | optional first slice-remove seq | `ObliteratedSeq` | ✅ | Same wire field, different C# name. |
| `movedSeqs` | optional slice-remove seq list | `RemoveStamps` generated with `sliceRemove` | ✅ | Current C# supports array. |
| `movedClientIds` | optional slice-remove client list | `RemoveStamps` generated with `sliceRemove` | ✅ | Current C# supports array. |
| raw arrays retained | raw TS arrays remain on JSON object | C# converts to `SharedStringSnapshotRemoveStampDto` list | ⚠️ | Runtime-equivalent, but DTO shape is not TS-identical. |

#### Type: segment payload DTOs

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| raw text segment | string or `{ text, props? }` | `Text`, `Props` | ✅ | Reads both forms. |
| text `props` | optional object; empty object may be present in input | `PropertySet? Props` preserving empty object as empty dictionary | ✅ | Loader does not strip `{}` on read. |
| marker segment | `{ marker: { refType?, props? }, props? }` | `IsMarker`, `MarkerRefType`, `Props` | ✅ | Reads nested and outer props. |
| marker length | marker cached length is 1 | `SegmentLength` returns 1 for marker | ✅ | Current C# fixed marker-length validation risk from old PARITY S5. |
| arbitrary segment kinds | `IJSONSegment` union can include other segment specs | C# marks unknown and skips | ⚠️ | SharedString only expects Text/Marker, but raw forward compatibility differs. |
| segment payload raw shape | retained until `segmentFromSpec` | C# normalizes into DTO fields | ⚠️ | Legitimate loader adapter, not wire DTO-equivalent. |

#### Type: `SharedStringSnapshotDto` / loader result

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| header chunk result | `SnapshotLoader.loadHeader` returns chunk and reloads tree | `SharedStringSnapshotDto` with header chunk fields | ⚠️ | C# exposes an intermediate DTO; TS loader does not expose this shape publicly. |
| `Version` | chunk `version` | top-level `Version` | ➕ | Convenience copy from header chunk. |
| `SegmentCount` / `Length` / `StartIndex` | chunk fields | top-level copies updated after body chunks | ➕ | C# aggregation helpers. |
| `Segments` | chunk/body segments materialized | all segments in one list | ⚠️ | TS appends body directly to tree; C# aggregates before population. |
| `HeaderMetadata` | chunk header metadata | top-level copied metadata | ✅ | Equivalent data. |
| `CatchupOps` | loader returns promise of sequenced messages | `List<CatchupOpDto>` applied by populate | ⚠️ | Different loader contract and application point. |

#### Type: catchup ops

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| blob content | array of `ISequencedDocumentMessage` | array of strings or objects accepted | ⚠️ | C# accepts synthetic shortcuts; TS expects sequenced message shape after serializer parse. |
| `contents` | sequenced message operation payload | `contents`, `op`, `opJson`, or direct op object | ⚠️ | Tolerant read can hide malformed fixtures. |
| `sequenceNumber` | required message field | nullable, defaults to next seq | ⚠️ | Silent metadata invention risk. |
| `referenceSequenceNumber` | required message field | nullable, defaults to current seq | ⚠️ | Silent metadata invention risk. |
| `minimumSequenceNumber` | required message field | nullable, defaults to current min seq | ⚠️ | Silent metadata invention risk. |
| `clientId` | message client id | nullable, defaults to `"snapshot-catchup"` | ⚠️ | Silent client identity invention risk. |
| validation | TS validates before `processMergeTreeMsg` | C# validates against collab window before direct apply | ✅ | Same broad invariant. |
| processing path | normal message processing, interval map first then merge-tree | direct `target.ApplyOp` | ⚠️ | Bypasses SharedString event/routing envelope and interval map try-process path. |

#### Type: snapshot writing / summarization

| Field / Shape | TS | C# | Status | Notes |
|---|---|---|---|---|
| V1 writer | `SnapshotV1.emit` | Missing | ❌ | Explicitly deferred; not a bug for server-side no-summarize client. |
| legacy writer | `SnapshotLegacy.emit` | Missing | ❌ | Deferred. |
| empty props stripping/preservation | TS writer currently strips empty properties before writing; user noted C# writer bug fixed outside scoped files | no writer in scoped files | ⚠️ | Current loader preserves incoming `{}`; future writer parity must be deliberate because TS and WN expectations diverged in prior bug. |
| attribution writer | optional in TS | Missing | ❌ | Deferred attribution. |
| ordered chunk output | writer emits header + `body_N` chunks | Missing | ❌ | Deferred. |
| catchup writer | legacy writer can add catchup blob | Missing | ❌ | Deferred. |

### 6. Handle/property JSON adjuncts

#### Type: `ISerializedHandle` / C# handle wire helpers

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | `"__fluid_handle__"` | `HandleWireFormat.TypePropertyName` | ✅ | Matches. |
| `url` | required string | `Url` | ✅ | Matches. |
| `payloadPending?: true` | optional true-only field | `SerializedFluidHandle.PayloadPending`, writer emits when true | ✅ | Prior SharedDirectory handle-risk finding is fixed in current common helper. |
| live handle conversion | serializer/handle binding | registry lookup | ⚠️ | Host adapter. |
| unresolved handle object | plain serialized handle object | `SerializedFluidHandle` POCO | ⚠️ | Legitimate C# in-memory adapter. |

#### Type: property JSON values in ops/snapshots

| Shape | TS | C# | Status | Notes |
|---|---|---|---|---|
| strings/bools/null | JS primitives | CLR primitives/null | ✅ | Equivalent. |
| numbers | JS `number` | int/long/double depending parse | ⚠️ | Good for many cases, but TS has one number type. |
| arrays | JS arrays | `List<object?>` | ✅ | Equivalent JSON shape. |
| objects | plain JS objects | `PropertySet` dictionaries | ✅ | Equivalent JSON object shape. |
| `undefined` | not JSON-round-trippable; TS property deltas use null for absent old value | maps to null/missing | ⚠️ | Same broad JSON limitation; cannot distinguish undefined/null after wire. |
| empty `{}` props | legal object | preserved as empty `PropertySet` on read and written by op serializer when non-null | ✅ | Good for the specific empty-props integration-risk pattern on scoped op/load paths. |

## Reverse pass: C# types or state with no exact TS counterpart

| C# type / field | Classification | Notes |
|---|---|---|
| `SharedString` constructor with `IFluidDataObjectSender` / `IFluidDataObjectRegistry` | Legitimate host adapter | Replaces TS `SharedObject` runtime/serializer/handle. |
| `ProcessDataObjectOp`, `ProcessDataObjectAttach`, `IFluidDataObjectMessageHandler` | Host shim | Expected for WAC/WN integration; not TS channel API. |
| `SequenceDeltaEventArgs.Position/Length/Text/IsMarker/Marker/AnnotatedProperties/OpType` | Legitimate adapter with payload mismatch risk | Useful host shortcut fields, but not TS `SequenceDeltaEvent` data shape. |
| `SequenceDeltaRange.Length/OpType` | Legitimate adapter | TS derives length from segment; C# adds direct fields. |
| `RelativePosition` class | Partial adapter | TS uses structural `IRelativePosition`; C# class is one accepted shape among several. |
| `_lock` | Legitimate C# concurrency helper | No TS equivalent. |
| `_batchOps` / `RunInBatch` | C# batching adapter | Replaces deprecated TS `groupOperation` call surface. |
| `_intervalCollectionsWithOutboundHandlers` | Legitimate event adapter | Tracks .NET event subscriptions for interval local-op flushing. |
| `MergeTreeDeltaType.IntervalAdd/Delete/Change/PropertyChanged` | Invented legacy wire discriminants | Not TS merge-tree op codes; high-risk if emitted. |
| `IntervalOpMsg` hierarchy | Adapter/invented hybrid | Useful internal class hierarchy, but TS top-level interval ops are map operations. |
| `SharedStringSnapshotDto` | Legitimate loader DTO | TS loader does not expose this aggregated DTO; it is C# parse/populate staging. |
| `SharedStringSnapshotSegmentDto` | Legitimate loader DTO | Normalizes TS raw segment spec union into typed fields. |
| `SharedStringSnapshotRemoveStampDto` | Legitimate adapter | Maps TS removed/moved array fields into merge-tree stamp classes. |
| `CatchupOpDto` | Synthetic adapter with risk | Allows shortcut/defaulted catchup op shapes beyond TS sequenced-message DTO. |
| `SerializedFluidHandle` | Legitimate unresolved-handle adapter | Now preserves `payloadPending`. |

## Silent-bug risk findings

1. **Interval outbound ops can still use C# legacy numeric discriminants (`10..13`) instead of TS `{"type":"act","key","value"}`.** `SharedString.FlushPendingIntervalOps` sends `IntervalOpMsg` through the regular merge-tree serializer, while the TS-compatible interval-map helper is separate. A TS peer may not route these ops through `IntervalCollectionMap`.
2. **Catchup ops are applied inside the snapshot loader with synthetic defaults.** Missing `sequenceNumber`, `referenceSequenceNumber`, `minimumSequenceNumber`, or `clientId` are invented by C#; TS catchup ops are full sequenced messages and are processed through normal message routing.
3. **Delta event payloads are not TS-shaped.** C# lacks `deltaArgs` and `opArgs` envelopes; `Op` is only a partial replacement for `opArgs.op`, so WAC consumers can miss group/rollback/sequenced-message context.
4. **`orderedChunkMetadata` is accepted as missing despite being required by current TS `MergeTreeHeaderMetadata`.** Single-chunk snapshots may load, but this relaxes the current wire contract and can hide malformed header metadata.
5. **Snapshot loader normalizes raw segment specs into C# DTOs and skips unknown specs.** That is useful tolerance, but it is not a lossless representation and can differ from TS forward-compat handling.
6. **Top-level relative APIs resolve immediately to absolute positions.** TS can carry `relativePos1/2` through op DTOs for tombstone-relative operations; C# public relative helpers look up current marker positions before building ops.
7. **Null vs `-1`/undefined return shapes remain visible in public APIs.** `GetPosition` and local-reference position APIs return nullable ints, while TS documents numeric sentinels for some absent positions.
8. **Snapshot writer parity cannot be audited in scoped files because writing is absent.** This is deferred, not a bug, but the session's empty-props writer bug shows any future writer must be covered by golden wire tests.

## Cross-reference with `PARITY-AUDIT.md`

| Data-structure finding here | PARITY-AUDIT.md reference | Current audit note |
|---|---|---|
| Remote/local delta event shape differs; no `deltaArgs`/`opArgs` envelope | C1, C2, SS7 | Still relevant structurally, though C# now builds ranges from `MergeTreeDelta` for remote ops rather than only original op coordinates. |
| Local metadata/resubmit/rollback gaps | C3, C4, C10, C11 | Deferred per HANDOFF; current data structures remain narrower than TS local-op metadata. |
| `minSequenceNumber` optional fallback | S4 / session bug note | Current C# reader matches TS fallback (`minSequenceNumber ?? sequenceNumber`); not a current bug. |
| Multi-stamp snapshot remove/obliterate arrays | S1 | Current C# parses `removedClientIds`, `movedSeqs`, and `movedClientIds` into remove stamps; old total collapse appears remediated for load. |
| Attribution snapshot fields absent | S2, M11 | Still deferred, not a bug for WN unless attribution becomes required. |
| Catchup ops loader contract differs | S3 | Still relevant; C# applies catchup during populate and accepts/defaults shortcut DTOs. |
| Snapshot chunk protocol | S4 | Partly improved: modern and legacy chunk fields are represented, but channel storage/content-path abstraction is still not TS-parallel. |
| Marker snapshot length validation | S5 | Current C# `SegmentLength` returns 1 for markers; old validation risk appears remediated. |
| Relative ops and tombstone targeting | C5, S6, SS4, O4 | Serializer preserves relative fields, but public relative helpers resolve to absolute positions and some relative APIs remain absent. |
| Annotate-adjust | C6, O3 | Missing and explicitly deferred; do not count as WN bug. |
| ReferenceType values | O1, O10 | Current C# values match TS and no extra bits remain. |
| Interval op envelope | O2, IC6, IC7 | Still high-risk at the SharedString/IntervalCollectionMap boundary: inbound map ops supported; outbound normal path appears legacy numeric. |
| Property JSON normalization / empty props | M14, O6, session empty-props bug | Current scoped op serializer and snapshot loader preserve empty objects on read/write paths they implement. Snapshot writer is absent here. |
| Public API naming/surface | SS1-SS6, SS8, SS9 | Several old findings fixed (`InsertText` props, `RemoveText`, `ReplaceText`, marker-aware text, reentrancy); lifecycle/factory/summarize gaps remain deferred. |

## Recommended actions

1. **Fix highest-risk wire drift:** route outbound interval collection ops through the TS `IMapOperation` shape (`type:"act"`, `key`, `value`) or prove the WAC envelope intentionally translates the legacy numeric C# discriminants.
2. **Make catchup ops TS-shaped:** require full sequenced-message metadata in `CatchupOpDto` and process catchup through the same SharedString message routing path used for normal remote ops.
3. **Add event compatibility tests/goldens:** verify `deltaOperation`, `ranges`, `first/last`, property deltas, `clientId`, and `opArgs.op`-equivalent routing for insert/remove/annotate/obliterate/group.
4. **Tighten snapshot optionality:** keep the fixed `minSequenceNumber` fallback, but require or synthesize `orderedChunkMetadata` only through the same legacy conversion rules as TS.
5. **Document deferred API families in the C# public surface:** lifecycle/dispose, annotate-adjust, rollback/resubmit/stashed ops, attribution, full tracking/segment groups, pre-2019 Shared value migration, and non-SharedString `SharedSequence<T>`.
6. **When snapshot writing is added, use golden TS fixtures:** cover header metadata, body chunks, catchup blob, empty props, marker segments, handles with `payloadPending`, removed/moved arrays, and legacy chunk conversion.

## Files opened

Primary files:

- `packages/dds/map/src/csharp-port/DATA-STRUCTURE-AUDIT.md`
- `packages/dds/sequence/src/csharp-port/HANDOFF.md`
- `packages/dds/sequence/src/csharp-port/PARITY-AUDIT.md`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedString.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedStringOpSerializer.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedStringSnapshotLoader.cs`
- `packages/dds/sequence/src/sharedString.ts`
- `packages/dds/sequence/src/sequence.ts`
- `packages/dds/sequence/src/sharedSequence.ts`
- `packages/dds/sequence/src/sequenceDeltaEvent.ts`
- `packages/dds/sequence/src/intervalCollectionMap.ts`
- `packages/dds/sequence/src/intervalCollectionMapInterfaces.ts`
- `packages/dds/sequence/src/sequenceFactory.ts`
- `packages/dds/merge-tree/src/ops.ts`
- `packages/dds/merge-tree/src/snapshotV1.ts`
- `packages/dds/merge-tree/src/snapshotChunks.ts`
- `packages/dds/merge-tree/src/snapshotLoader.ts`
- `packages/dds/merge-tree/src/snapshotlegacy.ts`

Additional common files opened for handle/property wire adjuncts:

- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/Ops.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/Marker.cs`
- `packages/dds/csharp-port-common/HandleWireFormat.cs`
- `packages/dds/csharp-port-common/SerializedFluidHandle.cs`

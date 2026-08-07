# IntervalCollection Data Structure Audit — TS vs C#

**Base:** microsoft/main (`intervalCollection.ts`, `intervals/*`, `intervalIndex/*`, interval map wrappers)
**Port:** transpiledir `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/` plus interval op DTOs in `MergeTree/Ops.cs` and `SharedStringOpSerializer.cs`
**Audited:** 2026-08-06T14:12:48-07:00
**Method:** field-by-field comparison of major interval data structures, DTOs, endpoint helpers, op wire shapes, pending state, and indexes. Static/read-only source audit; no build, tests, lint, or code changes run.

## Summary

Counts below are row-level field/member comparisons, not pass/fail type counts. `Types` counts the major named structures audited in each category.

| Category | Types | ✅ | ⚠️ | ❌ | ➕ |
|---|---:|---:|---:|---:|---:|
| Endpoint primitives + `SequenceInterval` | 7 | 21 | 13 | 8 | 3 |
| Serialized interval and snapshot DTOs | 6 | 12 | 9 | 8 | 0 |
| Op wire DTOs and interval map envelope | 7 | 15 | 11 | 6 | 4 |
| `IntervalCollection` fields, API, events, pending state | 13 | 24 | 19 | 16 | 12 |
| Local collection + endpoint movement integration | 3 | 5 | 8 | 6 | 0 |
| Interval index family | 12 | 20 | 16 | 7 | 4 |
| IntervalCollectionMap wrapper | 6 | 8 | 6 | 8 | 2 |
| **Total** | **54** | **105** | **82** | **59** | **25** |

Legend: ✅ present and equivalent; ⚠️ present but different shape/semantics; ❌ missing in C#; ➕ extra in C#.

## Findings by category

### 1. Endpoint primitives + `SequenceInterval`

#### Type: endpoint enums and helper functions

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `Side.Before` / `Side.After` | `Side.Before = 0`, `Side.After = 1` in `merge-tree/src/sequencePlace.ts` | `Intervals.Side` implicit `Before=0`, `After=1`; `MergeTree.Side` explicit | ✅ | Numeric wire values match. There are two C# `Side` enums, but values currently align. |
| `defaultSide` | `defaultSide = Side.Before` | `IntervalUtils.DefaultSide = Side.Before` | ✅ | Bare numeric interval endpoints default to before. |
| `endpointPosAndSide(start,end)` | Pair helper accepts `SequencePlace | undefined` and preserves `"start"`/`"end"` literals | `EndpointPosAndSide(int? position, Side? side)` handles one nullable numeric endpoint | ⚠️ | C# helper is not TS-shaped and cannot normalize paired start/end or special endpoint literals. |
| `"start"` / `"end"` endpoint values | Serialized interval endpoints may be `number | "start" | "end"` | No interval-layer endpoint value type; only merge-tree `SequencePlace.Start/End` has `Position=-1` | ❌ | Critical for V1/V2 load/write parity when endpoint segments are serialized. |
| Detached endpoint sentinel | TS local-reference position APIs serialize detached positions as `-1` | `MergeTree.DetachedReferencePosition = -1`; `SequenceInterval.StartPosition/EndPosition` are nullable-typed but can return `-1` | ⚠️ | Sentinel exists, but interval enumeration filters detached intervals out of `_idIndex.All`, so a summarizer can omit rather than serialize `-1`. |
| `IntervalType.Simple` | `0x0` | `0x0` | ✅ | Equivalent. |
| `IntervalType.SlideOnRemove` | `0x2` | `0x2` | ✅ | Equivalent value. |
| `IntervalType.Transient` | `0x4` | `0x4` | ✅ | Equivalent value; deferred transient behavior beyond current filters is not treated as a bug. |
| `IntervalType.Nest` / `Cursor` | Not in current TS `IntervalType` | `Nest=0x1`, `Cursor=0x8` | ➕ | Retained comments say wire compatibility, but current TS would not understand these flags if emitted. |
| `IntervalStickiness` flags | `NONE=0`, `START=1`, `END=2`, `FULL=3` | `None=0`, `Start=1`, `End=2`, `Full=3` | ✅ | Bit values match. |
| `sidesFromStickiness` | Local function in `intervalCollection.ts` | `IntervalUtils.SidesFromStickiness` | ✅ | Equivalent for numeric endpoints. |
| `computeStickinessFromSide` | Also treats `startPos === "start"` and `endPos === "end"` as sticky | C# only accepts sides | ⚠️ | C# loses endpoint-sentinel contribution to stickiness. |
| `startReferenceSlidingPreference` / `endReferenceSlidingPreference` | Uses positions plus sides and can slide to endpoint segments | C# uses sides only | ⚠️ | Numeric side behavior matches; endpoint sentinels and `canSlideToEndpoint` are not represented here. |
| `compareSides` | `Before` sorts after `After` | `CompareSides` returns `Before ? 1 : -1` | ✅ | Equivalent. |

#### Type: `IInterval` / `SequenceInterval`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `start` / `end` | `LocalReferencePosition` endpoints | `Start` / `End` `LocalReferencePosition` | ✅ | Same core endpoint object concept. |
| `intervalType` | `readonly intervalType` | `IntervalType { get; init; }` | ✅ | Equivalent for supported flags. |
| `startSide` / `endSide` | `readonly startSide`, `readonly endSide` | `StartSide`, `EndSide` mutable internal set | ✅ | Values are present; mutability is C# implementation detail. |
| `stickiness` | Computed from endpoint positions/sides | Computed from sides only | ⚠️ | Same for numeric endpoints; not equivalent for `"start"`/`"end"` endpoint sentinels. |
| `properties` | Property bag plus `PropertiesManager` support | `PropertySet? Properties` | ⚠️ | Stores the bag, but lacks TS pending property manager/ack semantics. |
| interval id | `getIntervalId()` private id | `Id` property | ⚠️ | Data exists; API shape differs. |
| `compare(b)` | Full start/end/id comparator | Missing; C# has comparers outside the class | ❌ | C# `IInterval` also omits all comparator methods. |
| `compareStart` / `compareEnd` | Reference-position comparator with side tie-breakers | `CompareStart` / `CompareEnd` | ✅ | Broadly equivalent for attached numeric refs. |
| `overlaps(b)` | Interval-to-interval overlap | Missing | ❌ | C# only exposes numeric-range `Overlaps(int start,int end)`. |
| `overlapsPos` | Half-open numeric overlap: `endPos > bstart && startPos < bend` | `Overlaps` delegates to half-open `RangesOverlap` | ✅ | Equivalent method intent, though collection/index wrappers later use inclusive overlap. |
| `serialize()` / `serializeDelta()` | Produces V1 verbose object and op deltas with sequenceNumber/id/labels | Missing on `SequenceInterval` | ❌ | C# interval op/snapshot serialization is centralized elsewhere and loses several fields. |
| `clone()` / `union()` | Clone and convex hull support | Missing | ❌ | Affects event previous-interval payloads and index utilities. |
| `dispose` / `disposed` / `verifyNotDispose` | Implemented on `SequenceIntervalClass` | Missing | ❌ | **Deferred, not a bug** per handoff dispose lifecycle limitation. |
| `changeProperties` / `ackPropertiesChange` | Uses `PropertiesManager` | `SetProperties` / collection helper merges dictionaries | ⚠️ | C# can apply simple changes, but lacks pending/ack property semantics. |
| `addPositionChangeListeners` | Endpoint before/after slide callbacks drive reindexing/events | Missing | ❌ | Directly causes stale index/event risk after text edits move endpoints. |
| `modify` / `moveEndpointReferences` | Rebuilds refs during change/rebase while preserving props and flags | C# `ChangeCore` creates replacement refs; client rebase rewrites op positions | ⚠️ | Some behavior exists, but not TS-shaped and not tied to local metadata lists. |
| `HasDetachedEndpoint` | No TS field; detached represented by ref position/endpoint state | C# helper filters lookup/enumeration | ➕ | Useful helper, but high-risk if snapshot writing should serialize detached `-1` endpoints. |

### 2. Serialized interval and snapshot DTOs

#### Type: `ISerializedInterval` / `SerializedIntervalDelta`

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| `start` | `number | "start" | "end"` required for add/full serialize; optional in delta | `IntervalAddOpMsg.Start int`; `IntervalChangeOpMsg.Start int?` | ⚠️ | Numeric endpoints exist for ops only. C# cannot carry `"start"`/`"end"`. |
| `end` | `number | "start" | "end"` | `int` / `int?` | ⚠️ | Same limitation as `start`. |
| detached `-1` | TS serialize path can emit `-1` from local-reference position | C# sentinel exists in merge-tree, but no interval serializer proves it writes `-1` | ⚠️ | `_idIndex.All` excludes detached endpoints, so detached intervals can be dropped before serialization. |
| `sequenceNumber` | Required number in serialized interval | C# interval op payload writes `sequenceNumber: 0` in map serializer and skips the field on read | ⚠️ | Shape field can be emitted, but the value is not TS-parallel and no DTO stores it. |
| `intervalType` | Required `IntervalType` | `IntervalAddOpMsg.IntervalType`; map header | ✅ | Present for ops; snapshot DTO absent. |
| `properties` | Optional `PropertySet`; carries `intervalId` and `referenceRangeLabels` | C# writes/reads properties in op map helper | ✅ | User properties plus reserved id/label are represented in the specialized map serializer. |
| `stickiness` | Optional | Present in add/change op DTOs | ✅ | Present for op path. |
| `startSide` / `endSide` | Optional | Present in add/change op DTOs | ✅ | Present for op path. |
| property delete/no-op semantics | Handled by `PropertiesManager` and property deltas | Dictionary equality/null removal in `ApplyPropertyChanges` | ⚠️ | Simple property deltas exist; TS pending and ack semantics are absent. |

#### Type: `CompressedSerializedInterval` / collection snapshot formats

| Wire shape | TS | C# | Status | Notes |
|---|---|---|---|---|
| V1 collection | `ISerializedInterval[]` | No C# DTO/type found in scoped interval files | ❌ | Current C# interval layer has no V1 verbose snapshot storage model. |
| V2 collection | `{ label: string, version: 2, intervals: CompressedSerializedInterval[] }` | No C# DTO/type found | ❌ | Critical wire-shape gap for recently requested summary/load parity. |
| Compressed tuple positions | `[start,end,sequenceNumber,intervalType,properties,stickiness?]` | No fixed-position tuple type | ❌ | C# cannot currently guarantee fixed tuple positions or optional stickiness elision. |
| `compressInterval` | Removes `referenceRangeLabels` from tuple properties and stores label once | Missing | ❌ | No C# equivalent in interval layer. |
| `decompressInterval` | Restores label into properties and derives sides from stickiness | Missing | ❌ | No C# equivalent in interval layer. |
| `makeSerializable` wrapper | `{ type:"sharedStringIntervalCollection", value:<V1/V2> }` | No interval snapshot writer/collection value type in scoped C# | ❌ | `SharedStringSnapshotLoader.cs` is merge-tree segment loader; no interval collection store/load path found. |
| handle serialization in interval properties | TS uses `serializeHandles` | C# `SharedStringOpSerializer` can serialize handle-shaped property values via registry | ✅ | Op property path has handle support; snapshot interval value path is absent. |
| unknown fields | JS naturally preserves/ignores structurally | C# op reader skips unknown map payload fields | ✅ | Tolerant op read is present. |

### 3. Op wire DTOs and interval map envelope

#### Type: interval map operation wrapper

| Field / Shape | TS | C# | Status | Notes |
|---|---|---|---|---|
| Outer map op | `{ type: "act", key, value }` in `intervalCollectionMap.ts` | `SerializeIntervalCollectionOperation` writes `"type":"act"`, `"key"`, `"value"` | ✅ | Specialized serializer matches the TS wrapper. |
| Production local send path | SharedString submits interval map ops through the map/value-type operation path | `SharedString.SendLocalOp` calls `SharedStringOpSerializer.Serialize(op)` for interval ops | ⚠️ | The available TS-shaped serializer is not used by the main send path; emitted wire likely uses legacy numeric interval types. |
| `opName` | `"add"`, `"change"`, `"delete"` | Specialized map serializer writes same names | ✅ | Property changes are correctly represented as `"change"` in the specialized serializer. |
| `value` | `ISerializedInterval` or `SerializedIntervalDelta` | `IntervalMapPayload` record | ⚠️ | Runtime record stores only numeric nullable endpoints and has no `sequenceNumber` storage. |
| Collection key | Map key is collection label | `CollectionName` written to `key` | ✅ | Equivalent. |
| Interval id location | Stored in `value.properties.intervalId` | Runtime DTO has `IntervalId` plus serializer writes reserved property | ⚠️ | Wire can be correct in specialized path; runtime object shape is not TS structural shape. |
| `referenceRangeLabels` | Stored in `properties.referenceRangeLabels` | Specialized writer emits one-element array `[collectionName]` | ✅ | Equivalent for one label. No support for multiple labels. |

#### Type: C# legacy interval op DTOs

| Type / Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| interval op discriminant | Value-type `opName` string only | `MergeTreeDeltaType.IntervalAdd/Delete/Change/PropertyChanged = 10..13` | ➕ | Invented legacy discriminators. This directly overlaps PARITY Finding O2. |
| `IntervalOpKind` | No TS runtime enum | `Add/Delete/Change/PropertyChanged` enum | ➕ | Useful C# adapter, but not wire-parallel. |
| `IntervalAddOpMsg` | `opName:"add"`, `value: ISerializedInterval` | Add DTO has `Start`, `End`, `IntervalType`, `Stickiness`, `StartSide`, `EndSide`, `Props` | ⚠️ | Similar data, but no `sequenceNumber`, no endpoint literals, and id/labels live outside `Props` until serialization. |
| `IntervalDeleteOpMsg` | TS local delete submits `interval.serialize()` as value, including endpoints/type/properties | Delete DTO only has collection/id; specialized writer emits header/properties but no endpoints | ⚠️ | Processing only needs id, but wire shape is not the TS source-of-truth shape. |
| `IntervalChangeOpMsg` endpoints | TS public `change` requires both start and end or neither | C# `Change` and DTO allow one endpoint at a time | ⚠️ | A one-ended C# change has no TS API counterpart and can be misprocessed by TS peers. |
| `IntervalPropertyChangedOpMsg` | No distinct opName; property-only `change` | Separate C# runtime class, specialized writer maps to `opName:"change"` | ⚠️ | Adapter is OK if only specialized writer is used; numeric serializer emits unsupported `type:13`. |
| `sequenceNumber` in op value | Current TS serialize uses `client.getCurrentSeq()` | C# map header always writes `0`; reader skips it | ⚠️ | Reconnect/rebase semantics can drift for pending interval ops. |
| Own ack local metadata | TS runtime returns `IntervalMessageLocalMetadata` to the collection process path | C# ack lookup uses `PendingOpEntry` and `ApplyOwnAck` no-ops | ⚠️ | Runtime correlation is narrower and loses property/endpoint ack behavior. |

### 4. `IntervalCollection` fields, API, events, pending state

#### Type: `IntervalCollection` instance fields

| Field | TS | C# | Status | Notes |
|---|---|---|---|---|
| saved serialized intervals | `savedSerializedIntervals?: ISerializedIntervalCollectionV1` | Missing | ❌ | C# has no lazy interval snapshot-load state in `IntervalCollection`. |
| local collection | `localCollection: LocalIntervalCollection | undefined` | Single `IntervalCollection` owns indexes directly | ⚠️ | C# flattens TS outer/inner collection split. |
| client | `client: Client | undefined` | `_mergeTree` and `_opSender` | ⚠️ | C# avoids TS client field but loses client sequence/current-seq access needed for serialization. |
| `onDeserialize` | Optional callback for property deserialization | Missing | ❌ | Public `attachDeserializer` unsupported. |
| `pending` | Per-id pending changes map with local and endpoint linked lists | Missing in collection; partial data stored in `Client.PendingOpEntry` | ❌ | Major pending state-machine gap. |
| `submitDelta` | Closure adds local metadata node before submit | `_opSender` sends DTO to `Client.EmitIntervalOp` | ⚠️ | Sends ops, but does not preserve TS metadata graph. |
| indexes | Dynamic `Set<SequenceIntervalIndex>` in `LocalIntervalCollection` | Hardwired `_idIndex`, `_overlappingIndex`, `_startpointIndex`, `_endpointInRangeIndex`, `_endpointIndex` | ⚠️ | Built-in indexes exist, but attach/detach extension model is absent. |
| `_lock` | No JS lock | `_lock` | ➕ | Legitimate C# concurrency adapter. |
| `Name` / `EndpointType` | Label captured in constructor/local collection | Public properties | ➕ | Useful C# adapter fields. |

#### Type: public API surface

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `attached` | Public boolean | Missing | ❌ | C# collections are always constructed attached to a merge tree. |
| `attachIndex` / `detachIndex` | Public dynamic index API | Missing | ❌ | C# hardwires all indexes. |
| `add({ start, end, props, id? })` | Object parameter with `SequencePlace` endpoints | Overloads with `int`/`Side` endpoints and optional `intervalType` | ⚠️ | Functionality overlaps, but endpoint type and call shape differ. |
| stickiness feature flag | TS rejects sided endpoints unless `intervalStickinessEnabled` | No feature flag check | ⚠️ | C# accepts side/stickiness unconditionally. |
| `removeIntervalById` | Present | Present | ✅ | Equivalent high-level behavior. |
| `getIntervalById` | Present | Present | ✅ | C# filters detached intervals out; see detached risk. |
| `change(id,{start,end,props})` | Requires both endpoints or neither | Allows `newStart` and `newEnd` independently | ⚠️ | One-ended endpoint changes are a C# invention. |
| `changeProperties` | No separate public TS method; property-only `change` | Separate convenience method | ➕ | Adapter over `Change(... props)`. |
| `[Symbol.iterator]` | Iterator over id index | `IEnumerable<SequenceInterval>` | ✅ | Equivalent .NET adapter. |
| start/end exact iterators | TS docs say start/end point equal to requested position | C# forward uses `>=`, backward uses `<` | ⚠️ | Silent query expansion; not TS-equivalent. |
| `gatherIterationResults` | Delegates to overlapping index exact-match logic | Present but routes to broader C# iterators | ⚠️ | API exists with different matching semantics. |
| `findOverlappingIntervals` | Uses `SequencePlace` and interval tree | C# uses ints and inclusive overlap | ⚠️ | Boundary and endpoint-sentinel semantics differ. |
| `findIntervalsWithStartpointInRange` / `EndpointInRange` | Available through manually attached indexes | Public methods on collection | ➕ | Useful extra surface, but not TS `ISequenceIntervalCollection` shape. |
| `previousInterval` / `nextInterval` | Present, deprecated wrappers over endpoint index | Present | ✅ | High-level methods exist; ordering is direct numeric/list-based. |
| `map` | Present | Present | ✅ | Equivalent adapter. |
| `attachDeserializer` | Present | Missing | ❌ | Deserialization hook not ported. |

#### Type: process/ack/rebase/pending state

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `PendingChanges[id].local` | FIFO `DoublyLinkedList<IntervalMessageLocalMetadata>` | No per-id local list | ❌ | C# uses global `_pendingOps` list. |
| `PendingChanges[id].endpointChanges` | Linked list for add/change endpoint ops | No equivalent list | ❌ | C# stores current endpoint references on pending entry only. |
| `PendingChanges[id].consensus` | Last acked/remote consensus interval | Missing | ❌ | Remote changes racing local changes cannot use TS consensus object. |
| `IntervalAddLocalMetadata` | `{ type:"add", localSeq, endpointChangesNode?, interval }` | `PendingOpEntry` has `IsInterval`, kind, local seq, interval id, endpoint refs | ⚠️ | Partial substitute, no linked node. |
| `IntervalChangeLocalMetadata` | Same plus interval | Partial substitute | ⚠️ | Does not preserve endpointChanges node or consensus. |
| `IntervalDeleteLocalMetadata` | `{ type:"delete", localSeq }` | Pending entry kind delete | ✅ | Basic identity exists. |
| `pendingChangesStart` / `pendingChangesEnd` / `pendingIds` | Not exact symbols in current TS; equivalent state is the pending linked-list map | No exact C# fields | ❌ | If prior notes refer to those concepts, they are not ported as independent fields. |
| `process` | Central add/delete/change ack/remote handler | Split `ApplyRemote*` plus `ApplyOwnAck` | ⚠️ | Remote application exists; own ack side effects mostly no-op. |
| `ackAdd` | Converts pending StayOnRemove refs to SlideOnRemove, repositions if needed | `ApplyOwnAck(IntervalAddOpMsg)` is no-op | ❌ | Endpoint flag lifecycle is not TS-parallel. |
| `ackChange` | Acks property manager and endpoint refs | `ApplyOwnAck(IntervalChangeOpMsg)` is no-op | ❌ | Property pending state and endpoint ack missing. |
| `ackDelete` | No pending delete bookkeeping for local, remote deletes existing interval | C# delete ack no-op; remote delete exists | ⚠️ | Basic behavior OK, but no metadata parity. |
| `rebaseLocalInterval` | Uses metadata interval refs, endpoint slide, squash, canSlideToEndpoint | `Client.RebasePendingInterval` rewrites positions from current refs/stale positions | ⚠️ | Some rebase behavior exists, but not TS-shaped and lacks consensus map. |
| rollback / resubmit / stashed ops | Implemented in TS | Missing or partial through `RegeneratePendingOps` only | ❌ | **Deferred, not a bug** per handoff for rollback/resubmit/stashed ops. |

#### Type: interval events

| Event / Payload | TS | C# | Status | Notes |
|---|---|---|---|---|
| `addInterval` | `(interval, local, op)` | `OnAddInterval` args: `Interval`, `Local`, `Operation` | ✅ | Equivalent payload intent. |
| `deleteInterval` | `(interval, local, op)` | Args include `PreviousStart`, `PreviousEnd` | ➕ | Extra prior-position data. |
| `changeInterval` | `(interval, previousInterval, local, op, slide)` | Args use previous start/end positions, not previous interval object | ⚠️ | Consumers cannot query previous endpoint refs/sides like TS. |
| `propertyChanged` | `(interval, propertyDeltas, local, op)` | Args have `ChangedProps` and `PropertyDeltas` | ✅ | Equivalent plus changed props. |
| `changed` | `(interval, propertyDeltas, previousInterval?, local, slide)` | `OnChanged` with previous start/end positions | ⚠️ | Previous interval object is missing. |
| slide events from endpoint movement | Emitted by local-reference callbacks | No endpoint callbacks; C# raises changes only from explicit interval operations/rebase delete | ❌ | Indexes and events can silently miss text-edit endpoint movement. |

### 5. Local collection + endpoint movement integration

#### Type: `LocalIntervalCollection`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `overlappingIntervalsIndex`, `idIntervalIndex`, `endIntervalIndex` | Three built-ins | Five built-ins on `IntervalCollection` | ✅ | C# includes all TS built-ins plus range indexes. |
| dynamic `indexes: Set` | Supports user-attached indexes | Fixed `IIntervalIndex[]` | ⚠️ | Missing attach/detach behavior. |
| `addInterval` label guard | Rejects interval with another collection label | C# does not validate incoming reserved labels on add | ❌ | Reserved-label spoofing/modification is not guarded. |
| endpoint `referenceRangeLabels` props | Added to both local refs | C# `CreateEndpointReference` does not add label/id properties | ❌ | `intervalLocatorFromEndpoint`-style lookup has no data. |
| endpoint `interval` back-reference prop | Added by `linkEndpointsToInterval` | Missing | ❌ | Local reference cannot locate its owning interval. |
| add/remove interval listeners | `addPositionChangeListeners` / `removePositionChangeListeners` | Missing | ❌ | Core reason indexes do not auto-update on endpoint slide. |
| before-slide index removal | Removes interval from all indexes before endpoint moves | Missing | ❌ | Index nodes can be stale if endpoint positions change under them. |
| after-slide index re-add + event | Re-adds and emits slide change | Missing | ❌ | Query/event parity gap after removes/obliterates/rebase. |
| `setSlideOnRemove` after non-collab/ack | TS toggles pending refs | C# endpoints are created with SlideOnRemove immediately | ⚠️ | Pending local intervals do not get TS StayOnRemove protection. |
| `intervalLocatorFromEndpoint` | Exported helper | Missing | ❌ | Related to missing endpoint properties. |

### 6. Interval index family

#### Type: shared index interface and comparers

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `SequenceIntervalIndex.add/remove` | Interface with `add`, `remove` | `IIntervalIndex.Add/Remove` | ✅ | Equivalent operations. |
| comparator source | Interval methods compare local refs and sides | `IntervalIndexComparers` compare current numeric positions/sides/id | ⚠️ | Loses reference-position ordering around endpoint sentinels/detached refs. |
| `forceCompare` / `compareOverrideables` | Used for inclusive range boundaries in RB-tree queries | Missing | ❌ | C# range queries use direct numeric predicates instead. |
| factories in `index.ts` | `create*Index(sharedString)` functions | No factories; direct constructors/hardwired instances | ⚠️ | API shape differs. |

#### Type: `IdIntervalIndex`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| storage | `Map<string, SequenceIntervalClass>` | `Dictionary<string, SequenceInterval>` | ✅ | Equivalent id map. |
| add/remove id asserts | Asserts id exists | Throws `ArgumentException` if id missing | ✅ | Equivalent failure intent. |
| `getIntervalById` | Returns stored interval, including current state | Returns only if `!HasDetachedEndpoint` | ⚠️ | Detached interval lookup differs; TS may retain interval and serialize/report `-1`. |
| iterator | Map values | `All` enumerable | ⚠️ | C# `All` also filters detached intervals. |
| `GetStoredIntervalById` | No TS public equivalent | Internal C# bypass for detached rebase cleanup | ➕ | Useful adapter, but divergence point. |

#### Type: `EndpointIndex`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| backing structure | `RedBlackTree<SequenceInterval>` sorted by `compareEnd` | `List<SequenceInterval>` sorted on demand | ⚠️ | Correctness may hold for simple cases; performance and tie order differ. |
| constructor sequence | Requires sequence to make transient intervals | No sequence dependency | ⚠️ | C# uses raw positions, so endpoint sentinel/tie behavior differs. |
| `previousInterval` | RB-tree `floor(transientInterval)` | Filters end `<= position`, last by endpoint comparer | ✅ | Equivalent for attached numeric endpoints. |
| `nextInterval` | RB-tree `ceil(transientInterval)` | Filters end `>= position`, first by endpoint comparer | ✅ | Equivalent for attached numeric endpoints. |
| `findEndpointsInRange` | Not on TS `EndpointIndex` | Extra C# method checking either endpoint | ➕ | Extra convenience API; overlaps with `EndpointInRangeIndex` naming. |

#### Type: `EndpointInRangeIndex` / `StartpointInRangeIndex`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| backing structure | Red-black tree with transient boundary intervals | List scan | ⚠️ | Same small-result behavior expected; ordering/tie/performance not TS-parallel. |
| invalid range guard | `start <= 0 || start > end || empty -> []` | Same guard | ✅ | Equivalent. |
| inclusive boundaries | `mapRange` with force-compare sentinels includes both ends | `position >= start && position <= end` | ✅ | Numeric attached boundary behavior matches. |
| tie ordering | `compareStart`/`compareEnd`, overrideables, id | `OrderBy` numeric positions then id | ⚠️ | Side and endpoint sentinel tie behavior can differ. |
| detached filtering | Tree membership maintained by listeners/removal | Explicit `!HasDetachedEndpoint` filter | ⚠️ | Detached intervals are hidden rather than represented/reindexed by TS movement logic. |

#### Type: `OverlappingIntervalsIndex`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| backing structure | `IntervalTree<BaseSequenceInterval>` | List scan | ⚠️ | Performance and deterministic tree order differ. |
| `findOverlappingIntervals` endpoint type | `SequencePlace` start/end | `int start, int end` | ⚠️ | C# cannot query with side or `"start"`/`"end"` endpoint values. |
| invalid range semantics | Rejects `end < start` plus sentinel-impossible ranges | Rejects only `end < start` | ⚠️ | C# lacks sentinel invalid cases. |
| overlap predicate | TS numeric `overlapsPos` is half-open; tree uses interval reference comparisons | C# collection/index use inclusive `intervalStart <= end && intervalEnd >= start` | ⚠️ | Boundary-touching and zero-length ranges can be included differently. |
| `gatherIterationResults` exact-match traversal | TS exact start/end matching through interval tree | C# collection iterators use `>=` / `<` ranges | ⚠️ | Query API returns broader result sets. |
| `map` / `mapUntil` | Present on TS overlapping index | Missing on C# index | ❌ | API narrowing. |
| `findOverlappingIntervalsBySegoff` | Declared in `SequenceIntervalIndexes.Overlapping` | Missing | ❌ | Internal specialized query not ported. |

### 7. IntervalCollectionMap wrapper

#### Type: `IntervalCollectionMap` and interfaces

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| map storage | `Map<string, IntervalCollection>` | `Client` dictionary and `SharedString.GetIntervalCollection` | ⚠️ | C# has collection lookup, not full map wrapper. |
| `createIntervalCollection` event | Present | Missing | ❌ | **Deferred, not a bug** per handoff full wrapper limitation. |
| `serialize` / `populate` map | Writes/loads `sharedStringIntervalCollection` values | Missing | ❌ | Critical if interval snapshot load/write is now in scope. |
| key normalization | Strips legacy `intervalCollections/` prefix | Missing | ❌ | Legacy snapshot compatibility absent. |
| `tryProcessMessage` | Processes `type:"act"` interval map ops | `SharedStringOpSerializer.Deserialize` can parse `"act"`; client applies interval DTO | ✅ | Core inbound op parsing exists. |
| `tryResubmitMessage` | Calls collection `resubmitMessage` | Missing as TS wrapper | ❌ | **Deferred, not a bug** for resubmit per handoff. |
| `tryRollback` | Calls collection rollback | Missing | ❌ | **Deferred, not a bug**. |
| `tryApplyStashedOp` | Calls collection stashed replay | Missing | ❌ | **Deferred, not a bug**. |
| `SequenceOptions.intervalSerializationFormat` | Selects V1/V2 | Missing | ❌ | No interval serialization-format option in C#. |
| direct `SharedString.GetIntervalCollection` | No direct TS `SharedString` map field shape in this file | Public C# convenience API | ➕ | Legitimate host adapter for WN. |

## Reverse pass: C# types or state with no exact TS counterpart

| C# type / field | Classification | Notes |
|---|---|---|
| `IntervalAddedEventArgs`, `IntervalDeletedEventArgs`, `IntervalChangedEventArgs`, `IntervalPropertyChangedEventArgs`, `IntervalUpdatedEventArgs` | Legitimate C# event adapters with payload mismatch risk | Replace TS event emitter tuples. Missing `previousInterval` object is the main semantic gap. |
| `IIntervalOpSender` and `ClientIntervalOpSender` | Legitimate host adapter | Connects collection to C# client pending queue; not a TS data structure. |
| `IntervalOpKind` | Adapter with wire-risk | Internal discriminant is useful, but has no TS wire counterpart. |
| `IntervalOpMsg` and subclasses | C# runtime DTOs; partly invented | They can be adapted to TS map values, but numeric `MergeTreeDeltaType.Interval*` serialization is not TS-compatible. |
| `MergeTreeDeltaType.IntervalAdd/Delete/Change/PropertyChanged` | Invented legacy wire discriminators | High-risk; should not be emitted on the wire to TS peers. |
| `Intervals.Side` separate from `MergeTree.Side` | Adapter with maintenance risk | Values match now; duplicate enums can drift. |
| `IntervalType.Nest` / `Cursor` | Invented/currently unsupported flags | Do not emit unless a current TS counterpart exists. |
| `IntervalCollection._lock` | Legitimate C# concurrency helper | No TS equivalent. |
| `IntervalCollection.Name` / `EndpointType` | Legitimate C# adapter fields | TS carries label and endpoint type through constructor/local collection state. |
| Hardwired `_indexes` array | C# helper with API parity risk | Replaces TS dynamic attach/detach index set. |
| `SequenceInterval.HasDetachedEndpoint` | C# helper with silent data-loss risk | Filtering detached intervals conflicts with TS `-1` serialized endpoint behavior. |
| `IdIntervalIndex.GetStoredIntervalById` | Rebase cleanup helper | Useful only because public lookup filters detached intervals. |
| Public range query methods on `IntervalCollection` | Extra convenience API | TS exposes these through attachable index objects, not directly on the collection interface. |

## Silent-bug risk findings

1. **Interval local ops can be emitted with numeric legacy `type:10..13` instead of TS `{type:"act", key, value:{opName,...}}`.** `SerializeIntervalCollectionOperation` writes the right envelope, but `SharedString.SendLocalOp` uses the generic serializer for interval ops. TS peers will not process numeric merge-tree interval discriminators.
2. **V1/V2 interval snapshot DTOs are absent in the scoped C# interval layer.** No C# `ISerializedInterval`, `CompressedSerializedInterval`, `{label,version:2,intervals}` model, `compressInterval`, or `decompressInterval` was found, so fixed-position tuple parity is not established.
3. **Detached intervals can be dropped instead of serialized with `-1` endpoints.** C# has the `-1` sentinel, but `IdIntervalIndex.All` and `GetIntervalById` filter `HasDetachedEndpoint`; TS serializes detached reference positions as `-1`.
4. **Pending interval endpoint refs are created as `SlideOnRemove` immediately and own ack handlers are no-ops.** TS creates local pending non-transient refs as `StayOnRemove`, then flips to `SlideOnRemove` on ack when safe.
5. **The per-interval pending state machine is structurally missing.** TS tracks FIFO local metadata, endpoint-change metadata, and consensus interval per id; C# stores a narrower global pending entry and cannot apply TS branch-for-branch remote/local races.
6. **C# allows one-ended endpoint changes.** TS public `change` requires both start and end or neither, and its processing expects that shape; a C# one-ended change can be misinterpreted.
7. **`sequenceNumber` in serialized interval op values is written as `0` and ignored on read.** TS uses the client current sequence number in serialized intervals; reconnect/rebase logic can drift.
8. **Indexes are not updated from endpoint slide callbacks.** TS removes/re-adds intervals around reference movement; C# lacks the callbacks, so queries can return stale ordering after text edits.
9. **Start/end iterators return ranges rather than exact matches.** TS docs and interval-tree traversal match exact start/end positions; C# forward/backward methods use `>=` and `<`.
10. **Overlap queries use inclusive numeric bounds in C#.** TS numeric interval overlap is half-open, so boundary-touching or zero-length cases can silently differ.

## Cross-reference with `PARITY-AUDIT.md`

| Data-structure finding here | PARITY-AUDIT.md reference | Current audit note |
|---|---|---|
| Endpoint refs lack TS boundary/split integration and endpoint properties | M4, L1, L5 | Still relevant: C# interval refs are separate objects without TS endpoint callback/property graph. |
| Pending state lacks exact local metadata and consensus maps | IC4, C3, C4, M6 | Still relevant structurally. Current C# has `PendingOpEntry` but not `PendingChanges`. |
| Detached/absent reference position and `-1` serialization risk | M9, L7, IC7, IC10 | Partly improved by `DetachedReferencePosition = -1`, but interval enumeration still filters detached intervals. |
| Interval op wire shape as merge-tree delta types | O2, IC7 | Still a top wire bug if generic interval serialization is used for local sends. |
| Endpoint reference flags/lifecycle | IC2, L2, L8 | ReferenceType numeric values appear fixed per handoff, but local pending StayOnRemove → ack SlideOnRemove lifecycle remains missing. |
| Stickiness and endpoint sides | IC3, L8 | Sides/stickiness fields now exist, but sentinel endpoints and `canSlideToEndpoint` rules remain incomplete. |
| Rollback / stashed interval ops | IC5, C10, C11 | Deferred per handoff, not a bug for this audit. |
| Interval events | IC6 | Still structurally different: C# lacks previous interval object and slide callbacks. |
| Interval property delta semantics | IC9, M14 | C# has simple dictionary deltas but no TS `PropertiesManager` pending/ack model. |
| Slide-off deletion and transient behavior | IC10, L3, L8 | Deferred transient semantics beyond current filters are not flagged, but slide-off state/event differences remain high risk. |
| Index implementation and ordering | II1, II2, II3, II4, II5, II6 | Still relevant; C# list scans and direct numeric predicates do not preserve TS comparator/endpoint movement model. |
| API signature narrowing/differences | IC8 | Still relevant; C# has .NET-shaped overloads and extra direct query methods but misses attach/detach index and deserializer APIs. |
| Full IntervalCollectionMap wrapper | IC1, IC8 | Deferred per handoff, not a bug, except its serialization/op envelope semantics are critical when interval snapshots/ops are in scope. |

## Recommended actions

1. **Fix interval wire emission first.** Route all local interval ops through the TS map envelope (`type:"act"`, collection `key`, `opName:"add"|"change"|"delete"`), remove or quarantine numeric `MergeTreeDeltaType.Interval*` wire emission, and ensure property-only changes serialize as `opName:"change"`.
2. **Port the V1/V2 serialized interval DTOs exactly.** Add C# models for `ISerializedInterval`, `SerializedIntervalDelta`, `CompressedSerializedInterval`, V1 arrays, V2 `{label,version:2,intervals}`; preserve tuple positions and optional stickiness behavior.
3. **Preserve endpoint sentinels.** Support `number | "start" | "end"` semantics and detached `-1` endpoints in interval summary load/write. Do not filter detached intervals out of serialization.
4. **Align pending endpoint lifecycle.** Create local pending non-transient interval refs with `StayOnRemove`, port ack-side `setSlideOnRemove` behavior, and add property-manager ack handling.
5. **Port the TS pending state shape before extending rebase.** Implement per-id pending local metadata FIFO, endpointChanges list, and consensus interval tracking. Rollback/resubmit/stashed APIs may remain deferred, but the shared metadata should match TS.
6. **Normalize change op shape.** Reject one-ended endpoint changes or translate them into a TS-valid both-endpoint delta.
7. **Wire endpoint movement into indexes and events.** Add before/after slide callbacks or equivalent so every index removes/re-adds intervals and emits slide changes when text edits move endpoints.
8. **Align index query semantics.** Use TS comparator logic with transient boundary intervals, exact start/end iterator matching, half-open numeric overlap where TS uses it, and stable id/side tie-breakers.
9. **Document deferred surfaces explicitly in the C# interval folder.** Dispose lifecycle, rollback/resubmit/stashed ops, attribution, Wave-31 transient behavior beyond filters, and full `IntervalCollectionMap` wrapper are deferred, not current bugs.

## Files opened

Primary files:

- `packages/dds/map/src/csharp-port/DATA-STRUCTURE-AUDIT.md`
- `packages/dds/sequence/src/csharp-port/PARITY-AUDIT.md`
- `packages/dds/sequence/src/csharp-port/HANDOFF.md`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/SequenceInterval.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/IntervalCollection.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/IntervalUtils.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/IIntervalIndex.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/EndpointIndex.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/EndpointInRangeIndex.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/StartpointInRangeIndex.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/IdIntervalIndex.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/OverlappingIntervalsIndex.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/Ops.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/Client.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/SequencePlace.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/MergeTree.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/LocalReference.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedString.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedStringOpSerializer.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedStringSnapshotLoader.cs`
- `packages/dds/sequence/src/intervalCollection.ts`
- `packages/dds/sequence/src/intervalCollectionMap.ts`
- `packages/dds/sequence/src/intervalCollectionMapInterfaces.ts`
- `packages/dds/sequence/src/IntervalCollectionValues.ts`
- `packages/dds/sequence/src/intervals/sequenceInterval.ts`
- `packages/dds/sequence/src/intervals/intervalUtils.ts`
- `packages/dds/sequence/src/intervalIndex/intervalIndex.ts`
- `packages/dds/sequence/src/intervalIndex/endpointIndex.ts`
- `packages/dds/sequence/src/intervalIndex/endpointInRangeIndex.ts`
- `packages/dds/sequence/src/intervalIndex/startpointInRangeIndex.ts`
- `packages/dds/sequence/src/intervalIndex/idIntervalIndex.ts`
- `packages/dds/sequence/src/intervalIndex/overlappingIntervalsIndex.ts`
- `packages/dds/sequence/src/intervalIndex/sequenceIntervalIndexes.ts`
- `packages/dds/sequence/src/intervalIndex/intervalIndexUtils.ts`
- `packages/dds/merge-tree/src/sequencePlace.ts`

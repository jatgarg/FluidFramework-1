# MergeTree Data Structure Audit — TS vs C#

**Base:** microsoft/main (`packages/dds/merge-tree/src/`)
**Port:** transpiledir `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/`
**Audited:** 2026-08-06
**Method:** field-by-field comparison of major MergeTree data structures, DTOs, stamps, op shapes, and core collections. Static/read-only source audit; no build, test, or lint run.

## Legend

- ✅ present and equivalent
- ⚠️ present but different shape/semantics
- ❌ missing in C#
- ➕ extra in C# (not in TS)

Counts below are row-level field/member comparisons, not pass/fail type counts. `Types` counts major named structures audited in each category.

| Category | Types | ✅ | ⚠️ | ❌ | ➕ |
|---|---:|---:|---:|---:|---:|
| Core segment/node data | 7 | 31 | 17 | 9 | 5 |
| Sequencing, stamps, perspectives | 4 | 14 | 8 | 7 | 2 |
| MergeTree/Client collections of truth | 5 | 21 | 22 | 17 | 8 |
| Ops and wire DTOs | 5 | 24 | 10 | 4 | 5 |
| Properties, partial lengths, snapshots | 5 | 18 | 17 | 9 | 3 |
| Deferred structural surfaces | 3 | 2 | 4 | 8 | 0 |
| **Total** | **29** | **110** | **78** | **54** | **23** |

## Findings by category

### 1. Core segment/node data

#### Type: constants (`constants.ts` / `Constants.cs`)

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `UniversalSequenceNumber` | `0` | `0` | ✅ | Wire/visibility sentinel matches. |
| `UnassignedSequenceNumber` | `-1` | `-1` | ✅ | Local pending sentinel matches. |
| `TreeMaintenanceSequenceNumber` | `-2` | `-2` | ✅ | Used by structural split/update paths. |
| `LocalClientId` | `-1` | `-1` | ✅ | Matches TS. |
| `NonCollabClient` | `-2` | `-2` | ✅ | Matches TS and snapshot min-seq perspective. |
| `SquashClient` | `-3` | `-3` | ✅ | Present, but C# does not port TS squash model. |

#### Type: `IMergeNode`, `IHierBlock`, `MergeBlock`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `parent`, `index`, `ordinal` | `IMergeNodeInfo`-style fields | `Parent`, `Index`, `Ordinal` | ✅ | Shape is equivalent with C# casing. |
| `cachedLength` | `number \| undefined` on block, `number` on segment | `int CachedLength` | ⚠️ | C# cannot represent TS's temporary undefined block length state. Usually benign but less state-parallel. |
| `children` / `childCount` | `children: IMergeNode[]`, `childCount` | fixed `IMergeNode?[]`, `ChildCount` | ✅ | Branching factor and temp slot are equivalent (`MaxNodesInBlock = 8`). |
| `partialLengths` | optional `PartialSequenceLengths` | non-null `PartialLengths` | ⚠️ | C# always allocates an object, while TS permits mid-update absence. |
| `needsScour` | block flag for zamboni queue | Missing | ❌ | C# zamboni scans/rebuilds rather than using TS LRU/scour state. |
| `rightmostTiles`, `leftmostTiles` | block-level tile marker caches | Missing | ❌ | Marker search is linear/index-assisted, not TS block-cache based. Related to PARITY M8. |
| `setOrdinal` / ordinal formula | `setOrdinal` + `computeHierarchicalOrdinal` | `SetOrdinal` + same width formula | ✅ | Equivalent enough for ordering. |
| `Capacity`, `MinChildren`, mutators | Not explicit TS public fields | C# convenience fields/methods | ➕ | Legitimate adapter/invariant helpers. |

#### Type: `ISegment`, `ISegmentInternal`, `ISegmentPrivate`, `BaseSegment`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type`, `cachedLength`, `properties` | core `ISegment` fields | `Type`, `CachedLength`, `Properties` | ✅ | Equivalent with casing and nullable mapping. |
| `insert` stamp | `IHasInsertionInfo.insert` required once bound | `InsertionStamp`, plus `Seq/LocalSeq/ClientId` aliases | ⚠️ | Stamp exists, but C# exposes flattened alias setters that can hide missing insertion metadata. |
| `removes` stamp list | ordered `RemoveOperationStamp[]` | `RemoveStamps` plus flattened `Removed*`/`Obliterated*` aliases | ⚠️ | List now exists, but many C# APIs still use primary flattened aliases. Highest-risk structural drift. |
| `localRefs?: LocalReferenceCollection` | bucketed collection | internal `List<LocalReferencePosition>` | ⚠️ | Fundamental collection shape differs; see local-reference table. |
| `trackingCollection` | present on every segment | Missing | ❌ | Deferred, not a bug per handoff unless future C# consumers need tracking groups. |
| `attribution?` | optional attribution collection | Missing | ❌ | Deferred, not a bug; snapshot shape still affected if attribution data appears. |
| `segmentGroups?` | pending op segment groups | Missing | ❌ | Deferred/POC, but directly affects ack/reconnect parity. |
| `propertyManager?` | TS `PropertiesManager` for annotate-adjust/local changes | Missing | ❌ | Annotate-adjust deferred, not a bug; raw annotate still works. |
| `endpointType?` | `"start" \| "end"` endpoint segments | represented by `ReferenceEndpointKind` on refs only | ⚠️ | C# lacks sentinel segment objects; endpoint sliding is custom. |
| `clone`, `canAppend`, `append`, `splitAt`, `toJSONObject` | virtual segment API | equivalent abstract/virtual methods | ✅ | Method surface is present. |
| `splitAt` metadata copy | copies insert/removes/obliterate refs/local refs/tracking/attribution | copies stamps/properties/local refs only | ⚠️ | Missing tracking/attribution/segmentGroups; local refs use simplified split rules. |
| flattened remove properties | Not TS public shape | `RemovedSeq`, `ObliteratedSeq`, etc. | ➕ | Back-compat adapter; risky if treated as source of truth. |

#### Type: `TextSegment`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | `"TextSegment"` | `"TextSegment"` | ✅ | Equivalent. |
| `text`, length | constructor sets `cachedLength = text.length` | `Text` property updates `CachedLength` | ✅ | Equivalent. |
| `TextSegmentGranularity` | `256` | `256` | ✅ | Equivalent. |
| `fromJSONObject` | string or `{ text, props }` | string or `IJSONTextSegment` | ✅ | Serializer also handles `JsonElement` before constructing. |
| `toJSONObject` | string if `properties` falsy, else `{ text, props }` | string if `Properties == null`, else object | ⚠️ | Empty `PropertySet` serializes as object in C#; TS V1 snapshot elides empty props before `toJSONObject`. Watch empty-props wire drift. |
| `clone` / `split` properties | TS shallow clones property bag | C# clone/split uses `ClonePropertySet` | ✅ | Prior PARITY M10 appears fixed structurally. |
| attribution split/append | `attribution.splitAt/append` | Missing | ❌ | Deferred, not a bug. |

#### Type: `ReferenceType`, `Marker`, marker JSON

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `ReferenceType` values | `Tile=0x1`, `RangeBegin=0x10`, `RangeEnd=0x20`, `SlideOnRemove=0x40`, `StayOnRemove=0x80`, `Transient=0x100` | same enum values | ✅ | Prior PARITY O1/IC2 appears fixed. |
| marker wire shape | `{ marker: { refType }, props? }` | `JSONMarkerSegment` same shape | ✅ | Serializer writes same core fields. |
| `reservedMarkerIdKey` | `markerId` | `markerId` | ✅ | Equivalent. |
| `reservedMarkerSimpleTypeKey` | exported `markerSimpleType` | Missing | ❌ | Usually legacy/low risk unless old marker simple-type consumers appear. |
| `reservedTileLabelsKey` | label key is TS convention via properties | `referenceTileLabels` constant | ✅ | Equivalent key for tile labels. |
| `Marker` implements `ReferencePosition` | `getSegment`, `getOffset`, `getProperties` | Missing those methods | ⚠️ | C# marker is only a segment; APIs resolve marker positions externally. |
| `toString` | `M${id}` | left-to-right mark `\u200E` | ⚠️ | Runtime diagnostic/stringification differs; not wire data. |
| `getId`, tile label helpers | present | present plus `HasTileLabel` | ✅ | Equivalent for marker lookup. |
| marker id mutation guard | TS rejects changing marker id | C# validates in `MergeTree.AnnotateVisibleRange` | ✅ | Prior PARITY M5 appears fixed. |

#### Type: `LocalReferencePosition` / `LocalReferenceCollection`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `refType`, `slidingPreference`, `canSlideToEndpoint`, `properties` | present | present | ✅ | Core reference data exists; enum names differ only by casing. |
| reference validation | mutually exclusive transient/slide/stay | same validation | ✅ | Equivalent. |
| `getSegment()`, `getOffset()`, `getProperties()` | interface methods | internal properties only | ⚠️ | C# callers inside port can inspect, but public TS-like interface is absent. |
| `trackingCollection` | present | Missing | ❌ | Deferred, not a bug. |
| `callbacks.beforeSlide/afterSlide` | present | Missing | ❌ | Silent risk for future interval/reference behavior. |
| `addProperties` | merges reference properties | Missing | ❌ | C# property set is mutable, but TS method shape absent. |
| collection shape | sparse offsets with `before`/`at`/`after` tombstone buckets | flat list per segment | ⚠️ | Major structural drift; affects tombstone ordering and endpoint sliding. |
| endpoint sentinels | refs can link to `startOfTree`/`endOfTree` segments | `ReferenceEndpointKind` enum on ref | ⚠️ | Adapter can represent endpoint, but lacks segment identity and tracking behavior. |
| transient refs | not stored in collection | C# `Detach`/list behavior only | ⚠️ | Some transient semantics exist, but not TS collection membership model. |

### 2. Sequencing, stamps, perspectives

#### Type: `OperationStamp` and stamp utilities

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `seq`, `clientId`, `localSeq?` | readonly interface fields | init-only record fields | ✅ | Prior PARITY O8 appears fixed. |
| insert/remove discriminators | `"insert"`, `"setRemove"`, `"sliceRemove"` | record subclasses with `Type` | ✅ | Equivalent discriminators. |
| stamp comparisons | `lessThan`, `greaterThan`, `lte`, `gte`, `equal` | same helpers | ✅ | Equivalent local-vs-acked ordering. |
| `isLocal`, `isAcked`, `isSquashedOp` | present | present | ✅ | Equivalent. |
| `spliceIntoList`, `hasAnyAckedOperation`, `compare` | present | present | ✅ | Equivalent. |
| `SliceRemoveOperationStamp` boundary sides | side metadata lives in obliterate info/range logic, not base stamp interface | C# adds `StartSide`, `EndSide` on stamp | ➕ | Useful local adapter, but not exact TS data placement. |

#### Type: `Perspective`, `PriorPerspective`, `CollaborationWindow`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `Perspective` interface | `refSeq`, `clientId`, `localSeq?`, `hasOccurred`, `isSegmentPresent` | no named interface; static `HasOccurredAt`/`IsSegmentPresentAt` | ⚠️ | Behavior is embedded, not first-class data structure. |
| `PriorPerspective` | class with same-client and refSeq visibility | emulated with `refSeq`/`clientId` args | ⚠️ | Equivalent for common remote op application, but not reusable/typed. |
| `LocalReconnectingPerspective` | class with `localSeq` cutoff | emulated when `localSeq` provided | ⚠️ | Partial support in reconnect helpers; no named type. |
| `LocalDefaultPerspective` | all known edits visible | current local view via `refSeq < 0` or current seq | ⚠️ | Semantics represented by sentinels, not explicit perspective. |
| `RemoteObliteratePerspective`, `LocalSquashPerspective`, `allAckedChangesPerspective` | present | Missing | ❌ | Obliterate beyond minimum and squash are deferred/not bugs, but shape gap is real. |
| `CollaborationWindow.clientId` | short local client id | split between `ClientId` string, MergeTree `_clientIds`, and `CollabWindow*` seq fields | ⚠️ | Source-of-truth split can drift. |
| `CollaborationWindow.localSeq` | authoritative local op counter | `_clientSeq` on `Client` | ⚠️ | Same role, different owner; affects pending segment semantics. |
| `mintNextLocalOperationStamp()` | creates local stamps and increments localSeq if collaborating | custom `NextClientSeq` plus later stamp mutation | ⚠️ | Equivalent for simple local ops, less TS-parallel for ack/rebase. |

### 3. MergeTree and Client collections of truth

#### Type: `MergeTree`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `root` | `IRootMergeBlock` with `mergeTree` back-pointer | `MergeBlock Root` | ⚠️ | Core tree exists; no root-specific type/back-pointer. |
| `collabWindow` | single object with clientId/currentSeq/minSeq/localSeq/collaborating | `CurrentSeq`, `MinSeq`; client id map separate | ⚠️ | Major source-of-truth split. |
| `pendingSegments` | `DoublyLinkedList<SegmentGroup>` | Missing | ❌ | Replaced by Client `_pendingOps`; ack/reconnect metadata is not TS-parallel. |
| `segmentsToScour` | heap of LRU segments | Missing | ❌ | C# zamboni scans/pack/coalesces differently. |
| `obliterates` | dedicated `Obliterates` index | custom nested `ObliterateStamp`/coverage scans | ⚠️ | Minimum obliterate support exists; index shape differs and beyond-minimum is deferred. |
| `options` | `IMergeTreeOptionsInternal` | no options object on `MergeTree` | ❌ | Feature flags and snapshot policy are not modeled. |
| `attributionPolicy` | optional policy | Missing | ❌ | Deferred, not a bug. |
| `idToMarker` | marker id map inside MergeTree | `MarkerIndex` on Client | ⚠️ | Marker identity is tracked outside tree; must be kept in sync manually. |
| `startOfTree`, `endOfTree` | sentinel segment objects | Missing | ❌ | Endpoint refs use enum, not segment identity. |
| `mergeTreeDeltaCallback`, `mergeTreeMaintenanceCallback` | callback fields | absent in MergeTree; C# returns `MergeTreeDelta` from Client | ⚠️ | Event source shape differs. |
| `localPerspective` | getter from collab window | not named | ⚠️ | Emulated with local current view methods. |
| `InsertSegments`, `MarkRangeRemoved`, `ObliterateRange`, `AnnotateRange` | present | present | ✅ | Core edit APIs exist. |
| `create/removeLocalReferencePosition`, `referencePositionToLocalPosition` | present | present but simplified | ⚠️ | C# APIs exist, but data structure behind them differs. |
| `reloadFromSegments` / rebuild | present | `RebuildFromSegments` | ✅ | Equivalent purpose. |
| block-aware length queries | `nodeLength` + partial lengths | `GetLength` + `PartialLengths.Query` | ✅ | Function exists; algorithm differs. |

#### Type: `Client`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `longClientId` | mutable public string | `ClientId` readonly string | ⚠️ | C# cannot represent TS reconnect client-id replacement exactly. |
| `clientNameToIds`, `shortClientIdMap` | red-black tree + array | `_clientIds` dictionary on MergeTree | ⚠️ | Short-id mapping exists, but owner and reverse lookup differ. |
| `_mergeTree` | readonly MergeTree | `MergeTree` internal property | ✅ | Equivalent. |
| `specToSegment` | constructor-injected segment factory | `SegmentFromSpec` internal switch | ⚠️ | Less extensible; non-text/marker segment specs are not first-class. |
| logger/events | `TypedEventEmitter` callbacks | `MergeTreeDelta` return values/events in SharedString layer | ⚠️ | Event data shape differs; sibling audit may cover API layer. |
| `getMinInFlightRefSeq` | callback used for MSN preservation | Missing | ❌ | High-risk minSeq with in-flight ops mismatch (PARITY C9). |
| `pendingSegments` access | via `_mergeTree.pendingSegments` | `_pendingOps: List<PendingOpEntry>` | ⚠️ | Custom collection of truth; no segment group nodes. |
| `rollback`, `applyStashedOp`, `regeneratePendingOp` | TS reconnect/rollback model | partial custom `RebasePendingOps`; no TS rollback/stash | ⚠️ | Deferred, not a bug per handoff, but structural drift. |
| `applyMsg` sequencing | updates collabWindow, acks pending, applies remote | `ApplyOp` / `AcknowledgeLocalOp` paths | ⚠️ | Core behavior exists; shape and metadata differ. |
| local-reference APIs | public methods route into MergeTree | present in C# Client | ✅ | Method purpose exists with simplified refs. |
| marker search | block-cache search | linear walk + marker index | ⚠️ | Correct for small docs; not TS cache structure. |
| interval collections | SharedString/interval layer uses TS separate envelope | C# keeps interval collections and legacy interval op DTOs here | ➕ | Out of MergeTree audit scope; deferred/not a bug. |

#### Type: `PendingOpEntry` / TS segment groups

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| op-local metadata | `SegmentGroup` and local op metadata | `PendingOpEntry.Op`, `Segments`, flags | ⚠️ | Carries similar information, but not same identity/lifecycle. |
| `localSeq` | stamp on segments and collab window | `LocalSeq` field plus segment aliases | ✅ | Present. |
| `clientSequenceNumberObserved` / batch correlation | runtime message metadata and segment groups | `ClientSeq` plus `AssociatePendingOpsWithClientSequence` | ⚠️ | Similar goal; shape differs and can group incorrectly after transformations. |
| pending interval references | TS interval endpoint local refs/metadata | `IntervalStartReference`, `IntervalEndReference` | ➕ | C# adapter for interval subset; out of MergeTree scope. |
| `DropAfterRebase` | no exact field | C# flag | ➕ | Custom rebase control state. |

### 4. Ops and wire DTOs

#### Type: `MergeTreeDeltaType`, op union, op builders

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| insert/remove/annotate/group/obliterate/sided values | `0..5` | `0..5` | ✅ | Core discriminants match. |
| interval op discriminants | not MergeTreeDeltaType; SharedString interval-map envelope | `10..13` extra enum values | ➕ | Deferred/out of scope; not a MergeTree parity bug if not emitted to TS clients. |
| `IMergeTreeInsertMsg` | `pos1?`, `relativePos1?`, `pos2?`, `relativePos2?`, `seg?` | same fields, relative typed `object?` | ⚠️ | Wire names exist; C# loses `IRelativePosition` shape and validation. |
| `IMergeTreeRemoveMsg` | same positional fields | same fields | ✅ | Equivalent core shape. |
| `IMergeTreeAnnotateMsg` | required `props`, `adjust?: never` | nullable `Props` | ⚠️ | Missing/empty props tolerance differs. |
| `IMergeTreeAnnotateAdjustMsg` | annotate with `adjust` | Missing | ❌ | Deferred, not a bug; wire field still unsupported if received. |
| `IMergeTreeObliterateMsg` | numeric `pos1/pos2`; relative fields `never` | numeric fields plus relative object fields | ⚠️ | Extra relative fields should not appear on TS wire. |
| `IMergeTreeObliterateSidedMsg` | `{ pos, before }` endpoints | `SequencePlace` class serialized as `{ pos, before }` | ✅ | Serializer maps `Side.Before` to `before: true`. |
| `IMergeTreeGroupMsg` | `ops: IMergeTreeDeltaOp[]` | `List<MergeTreeOp>` | ⚠️ | C# allows nested group/interval `MergeTreeOp` values unless validation blocks them. |
| `ClientSeq` on op objects | runtime message has client sequence number, not core op field | `ClientSeq` property on `IMergeTreeOp` | ➕ | Adapter for local queue; should not leak on wire. Serializer appears not to write it. |
| `OpBuilder` helpers | TS `opBuilder.ts` | subset in `Ops.cs` | ✅ | Common insert/remove/annotate/obliterate builders exist. |

#### Type: `SequencePlace`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| union shape | `number \| "start" \| "end" \| { pos, side }` | class with `Position`, `Side`, `IsStart`, `IsEnd` | ⚠️ | C# cannot represent the bare-number/string union directly; serializer normalizes endpoints. |
| `Side` enum | `Before=0`, `After=1` | same values | ✅ | Equivalent. |
| default side for bare number | `Side.Before` | `SequencePlace.At(..., Side.Before)` helper | ✅ | Equivalent when caller uses helper. |
| `"start"`, `"end"` | normalized to `pos=-1` with side conventions | static `Start`/`End` | ✅ | Equivalent after normalization. |
| `defaultSide` export | exported constant | Missing | ❌ | Low risk; C# callers hardcode helper defaults. |

#### Type: op serializer behavior relevant to MergeTree

| Wire shape | TS | C# | Status | Notes |
|---|---|---|---|---|
| `type` | numeric delta type | numeric enum value | ✅ | Equivalent. |
| relative positions | serialized as `relativePos1/2` objects | read/write supported as generic JSON values | ⚠️ | Recent fix likely addresses PARITY O4, but type validation remains weaker. |
| segment payload | string, text object, marker object, unknown segment specs via `specToSegment` | string/text/marker/JsonElement only | ⚠️ | Unknown segment round-trip still limited. |
| handle values in props/relative pos | Fluid serializer handles serialized handles | C# uses `HandleWireFormat.MakeHandlesSerializable/ResolveSerializedHandle` | ✅ | Shape appears intentionally round-trippable, including optional `payloadPending`. |
| empty props | TS snapshots elide empty props; op objects can carry `{}` if caller gave it | C# writes any non-null `PropertySet` | ⚠️ | Empty-props drift still possible outside snapshot extraction. |
| annotate-adjust wire | supported when option enabled | not read/written | ❌ | Deferred, not a bug. |

### 5. Properties, partial lengths, snapshots

#### Type: `PropertySet`, `PropertyMap`, `PropertiesManager`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| `PropertySet` | object map with any JSON-serializable values | `Dictionary<string, object?>` | ✅ | Equivalent broad shape. |
| `createMap` | `Object.create(null)` | `Dictionary<string,T>` | ⚠️ | Prototype behavior irrelevant in C#, but key equality/ordering differs. |
| `matchProperties` | recursive, treats missing/undefined distinctly | recursive map comparison | ⚠️ | C# cannot represent JS `undefined`; JSON value normalization helps but not identical. |
| `extend` | skips `undefined`, deletes on `null` | deletes on `null`; no `undefined` | ✅ | Equivalent after JSON materialization. |
| `clone`, `addProperties`, `extendIfUndefined`, `combine` | present | present | ✅ | Core helpers exist. |
| `PropertiesManager` | tracks local/remote property changes and annotate-adjust consensus | Missing | ❌ | Annotate-adjust/rollback deferred, not a bug; raw annotate is simpler. |
| `copyPropertiesAndManager` | preserves property manager through split | Missing | ❌ | Relevant only when property manager is ported. |

#### Type: `PartialSequenceLengths` / `PartialLengths`

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| min-seq baseline | `minLength`, `minSeq` | no explicit `minLength`; cumulative deltas from zero/current cache | ⚠️ | Structurally non-parallel; old-refSeq lengths may diverge near MSN. |
| overall partial list | `PartialSequenceLengthsSet partialLengths` | `_deltasBySeq` + cumulative list | ✅ | Same high-level purpose. |
| per-client adjustments | `perClientAdjustments[]` | `_clientDeltasBySeq` | ⚠️ | Similar but not TS algorithm/shape. |
| unsequenced local records | `unsequencedRecords.partialLengths/perRefSeqAdjustments/cachedAdjustmentByRefSeq` | `_localUnackedSegmentsByClient` only | ⚠️ | LocalSeq/refSeq reconnect lengths are not TS-parallel. |
| segment count | `segmentCount` | Missing | ❌ | Used by TS verification/zamboni. |
| `combine` / `fromLeaves` | static combine with collabWindow and local partial option | rebuild/update from block | ⚠️ | Correctness depends on custom logic; not branch-for-branch. |
| `update` incremental path | TS update with verifier hooks | `UpdateLengthFor`, `UpdateStructureFor` | ⚠️ | Similar purpose; invariants differ. |
| zamboni support | internal `zamboni(collabWindow)` | outside `PartialLengths`; MergeTree zamboni rebuilds | ⚠️ | Deferred/non-parallel. |

#### Type: snapshot chunks and merge-info DTOs

| Field / Method | TS | C# | Status | Notes |
|---|---|---|---|---|
| V1 chunk fields | `version`, `startIndex`, `segmentCount`, `length`, `segments`, `headerMetadata`, `attribution?` | DTO has all except `attribution` | ⚠️ | Attribution deferred, not a bug; field is omitted. |
| legacy chunk fields | `chunkStartSegmentIndex`, `segmentTexts`, chunk seq/minSeq fields | loader accepts legacy path | ✅ | Loader maps legacy/current chunk forms. |
| header metadata | `totalLength`, `totalSegmentCount`, `orderedChunkMetadata`, `sequenceNumber`, `minSequenceNumber` | same plus catchup blob names | ✅ | Extra catchup names are SharedString loader adapter. |
| segment merge info | `json`, `client`, `seq`, `removedClientIds`, `removedSeq`, `movedClientIds`, `movedSeq`, `movedSeqs` | DTO stores scalar insert plus `RemoveStamps`; loader reads arrays | ✅ | Prior PARITY S1 appears partly fixed for loading. |
| removed seq arrays | TS writes one `removedSeq` for all remove clients | C# also accepts `removedSeqs` | ➕ | Tolerant adapter; should not emit non-TS field. |
| attribution chunk | `SerializedAttributionCollection` | Missing | ❌ | Deferred, not a bug unless loading documents with attribution. |
| unknown segment specs | TS uses injected `specToSegment` | C# marks unknown and skips | ⚠️ | Silent document content loss if non-text/marker segment appears. |

### 6. Deferred structural surfaces

| Surface | TS | C# | Status | Notes |
|---|---|---|---|---|
| Attribution | collection + policy + snapshot chunks | omitted | ❌ | Deferred, not a bug per handoff; type shape gap documented. |
| Tracking groups | segment/ref tracking collections | omitted | ❌ | Deferred, not a bug. |
| Full segment groups | pending op collections for rollback/resubmit | custom pending queue | ⚠️ | Rollback/resubmit/stashed ops deferred; current queue is not TS-parallel. |
| Annotate-adjust | op type + `PropertiesManager` | omitted | ❌ | Deferred, not a bug. |
| Obliterate beyond minimum surface | full `Obliterates` index, reconnect/squash | minimum support plus custom coverage | ⚠️ | Deferred, not a bug; keep tests scoped. |

## Reverse pass — C# types with no exact TS counterpart

| C# type | Classification | Notes |
|---|---|---|
| `PropertySet : Dictionary<string, object?>` | Legitimate adapter | Represents TS object map in .NET. |
| `MarkerIndex` | Legitimate adapter, but source-of-truth risk | TS stores `idToMarker` in `MergeTree`; C# stores marker id map on `Client`. |
| `PendingOpEntry` | Invented type | Replaces TS segment groups/local op metadata; high-risk for ack/reconnect drift. |
| `RebasedSegmentRun` | Invented type | Custom reconnect helper; acceptable only while TS reconnect model is deferred. |
| `MergeTreeDelta` / `MergeTreeDeltaRange` | Legitimate adapter | .NET event/return payload, but not TS `SequenceDeltaEvent`. |
| `SequencePlace` class | Legitimate adapter | Normalizes TS union into a class; cannot preserve bare-number/string authoring shape. |
| interval op classes in `Ops.cs` | Invented/out-of-scope | Intervals are covered by sibling audit; do not treat as MergeTree bugs. |
| snapshot DTO classes | Legitimate loader adapter | Useful for parsing; missing attribution is deferred. |
| nested `ObliterateStamp`, `NearestObliterate`, `ObliterateCoverage`, `SlideTarget` | Invented internal types | Replace TS `Obliterates`/perspective model; high-risk if obliterate scope expands. |

## Silent-bug risk findings

| ID | Risk | Impact | Notes |
|---|---|---|---|
| DS-MT-01 | Segment stamp history still has two sources of truth (`RemoveStamps` and flattened aliases) | Wrong historical visibility for overlapping remove/obliterate | Related to PARITY M1/M2/M13/P4/S1. |
| DS-MT-02 | Collaboration window is split across `Client`, `MergeTree`, `_clientIds`, and `_clientSeq` | Wrong minSeq/localSeq visibility, especially with in-flight ops | Related to PARITY C3/C9/P5. |
| DS-MT-03 | `pendingSegments`/segment groups are replaced by `PendingOpEntry` | Acks/rebase can update wrong segments after split/zamboni/remote edits | Related to PARITY C3/C4/M6. |
| DS-MT-04 | Partial-length algorithm is structurally non-parallel | Incorrect `getLength`/position at older refSeq or localSeq | Related to PARITY P1-P7. |
| DS-MT-05 | Local refs are flat lists without before/at/after tombstone buckets or callbacks | Interval endpoints and slide-on-remove may silently drift | Related to PARITY L1-L8/M4. |
| DS-MT-06 | Snapshot attribution/unknown segment specs are not represented | Data loss or load divergence for attribution/non-text documents | Attribution is deferred; unknown segment skip is still a wire-shape risk. |
| DS-MT-07 | Empty property maps are not consistently elided | Snapshot/op shape can differ from TS string segment normalization | Related to PARITY M14/O6 and earlier empty-props bug. |
| DS-MT-08 | Extra interval op discriminants and generic relative-position objects can leak | TS clients may reject or misinterpret non-TS op shapes | Intervals out of scope; serializer should keep core merge-tree wire strict. |

## Cross-reference with `PARITY-AUDIT.md`

| Data-structure finding | Related PARITY findings | Current note |
|---|---|---|
| DS-MT-01 stamp history/source of truth | M1, M2, M13, P4, S1, O8 | Stamp records and snapshot array loading improved; flattened aliases remain a drift risk. |
| DS-MT-02 collab window/localSeq | C3, C9, P5 | Still structurally non-parallel. |
| DS-MT-03 pending op/segment groups | C3, C4, C10, C11, C12, M6 | Deferred reconnect/rollback surfaces should remain explicitly scoped. |
| DS-MT-04 partial lengths | P1-P7 | Still a custom algorithm; some per-client/local support exists but not TS shape. |
| DS-MT-05 local references/endpoints | L1-L8, M4 | Core fields improved, collection model still differs. |
| DS-MT-06 MergeBlock tile caches | M8, SS5 | C# marker search is functional adapter, not TS cache shape. |
| DS-MT-07 ops/wire union | C5, C6, O2-O7, O9, O10 | ReferenceType and relative serializer improved; annotate-adjust/interval extras remain. |
| DS-MT-08 snapshot chunks | S1-S8, M11 | Remove/move arrays improved; attribution/unknown segment policy still gap. |
| DS-MT-09 marker identity/index | M5, C8, M8 | Marker id immutability appears fixed; index owner differs. |
| DS-MT-10 properties | M14, O6, C6 | Raw helpers exist; no `PropertiesManager`. |
| DS-MT-11 zamboni/scour | M6, M7, P6 | Custom zamboni remains high-risk around pending state. |
| DS-MT-12 deferred attribution/tracking | M11, M12 | Deferred, not bugs for current scope. |

## Recommended actions

### Must fix before broad TS-client interoperability

1. Pick one segment metadata source of truth: either remove flattened remove/obliterate aliases from decision paths or make them read-only projections of `RemoveStamps`.
2. Move toward a TS-shaped `CollaborationWindow` object (`clientId`, `collaborating`, `minSeq`, `currentSeq`, `localSeq`, `localPerspective`) and make `Client` use it for every local stamp.
3. Port or explicitly replace `pendingSegments`/`SegmentGroup` with an equivalent identity-preserving collection before relying on reconnect, zamboni, or pending interval rebasing.
4. Add targeted parity tests for historical length/position queries over overlapping insert/remove/obliterate across block splits.

### Wire-shape safeguards

1. Keep `ClientSeq`, interval op discriminants, and C#-only fields out of serialized MergeTree ops.
2. Normalize empty `PropertySet` to absent props in snapshot/op paths where TS would emit a raw string segment.
3. Reject or explicitly surface unsupported `adjust` and unknown segment specs rather than silently dropping content.
4. If attribution is not supported, detect attribution chunks and fail loudly or document a lossy load contract.

### Deferred, not bugs

1. Attribution collection/policy and tracking groups can remain absent if WN does not consume them.
2. Annotate-adjust, rollback/resubmit/stashed ops, and full obliterate reconnect should stay behind explicit unsupported/deferred notes.
3. Interval op DTOs should be audited by the interval sibling agent, not fixed under this MergeTree-only report.

## Files opened / compared for context

- `packages/dds/map/src/csharp-port/DATA-STRUCTURE-AUDIT.md`
- `packages/dds/sequence/src/csharp-port/PARITY-AUDIT.md`
- `packages/dds/sequence/src/csharp-port/HANDOFF.md`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/*.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedStringOpSerializer.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedStringSnapshotLoader.cs`
- `packages/dds/merge-tree/src/mergeTree.ts`
- `packages/dds/merge-tree/src/client.ts`
- `packages/dds/merge-tree/src/mergeTreeNodes.ts`
- `packages/dds/merge-tree/src/ops.ts`
- `packages/dds/merge-tree/src/localReference.ts`
- `packages/dds/merge-tree/src/segmentInfos.ts`
- `packages/dds/merge-tree/src/stamps.ts`
- `packages/dds/merge-tree/src/perspective.ts`
- `packages/dds/merge-tree/src/partialLengths.ts`
- `packages/dds/merge-tree/src/snapshotChunks.ts`
- `packages/dds/merge-tree/src/snapshotV1.ts`
- `packages/dds/merge-tree/src/attributionCollection.ts`
- `packages/dds/merge-tree/src/segmentPropertiesManager.ts`
- `packages/dds/merge-tree/src/textSegment.ts`
- `packages/dds/merge-tree/src/sequencePlace.ts`
- `packages/dds/merge-tree/src/constants.ts`

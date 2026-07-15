# SharedString TS Parity Audit

**Base:** microsoft/main state (TS source paths listed in request; local checkout audited statically)
**Port:** transpiledir `packages/dds/sequence/src/csharp-port/*`
**Audited:** 2026-07-14
**Goal:** exhaustive comparison against TS mergeTree + sequence sources; flag every deviation.

## Summary

| Category | Count |
|---|---:|
| BUG (silent) | 53 |
| BUG (crash) | 5 |
| Deviation (correct, non-parallel) | 17 |
| Deferred (POC choice, documented) | 10 |
| N/A (WN doesn't use / doesn't apply) | 0 |
| **Total findings** | **85** |

Static audit only. No source/test files were modified and no build/test commands were run. Findings intentionally over-report structural drift because the review priority is TS-parallel logic and high reliability.

## Findings — merge-tree core

### Finding M1: Segment removal history is flattened instead of storing TS remove stamp lists
- **Impact:** BUG (silent)
- **TS:** `mergeTree.ts:2083-2383`, `segmentInfos.ts`, `stamps.ts:29-172` — segments carry insertion stamps plus ordered remove/obliterate stamp history.
- **C#:** `MergeTree.cs:500-694, 1670-1833`, `ISegment.cs`, `Stamps.cs:19-244` — segments expose single `RemovedSeq`/`ObliteratedSeq` style fields plus simplified stamps.
- **What TS does:** TS can keep multiple `RemoveOperationStamp`s on one segment, compare them in order, and answer perspective-specific visibility for overlapping removes/obliterates.
- **What our port does:** The port keeps only one effective remove/obliterate metadata path for several decisions, then infers visibility from flattened fields.
- **Why it matters:** Overlapping remote/local remove and obliterate scenarios can choose the wrong winning operation, silently changing text length and visibility at older refSeqs.
- **Suggested fix:** ~150-250 LOC / ~2-4 h: port the stamp-list model and use it in visibility, zamboni, partial-length, snapshot, and serializer paths.
- **Regression-test suggestion:** Port `obliterate.concurrent.spec.ts` cases for multiple obliterates on the same segment, overlapping remove+obliterate, and “takes the correct remove clientId/stamp”.

### Finding M2: Obliterate-on-insert metadata is not TS-parallel
- **Impact:** BUG (silent)
- **TS:** `mergeTree.ts:2083-2290` and `segmentInfos.ts` — TS tracks `obliteratePrecedingInsertion` / insertion-ref-seq stamping when an insert loses to an obliterate.
- **C#:** `MergeTree.cs:1670-1833`, `ISegment.cs` — C# has custom obliterated-on-insert fields and local tests rather than the same metadata graph.
- **What TS does:** TS records that a segment was removed by a slice remove at insertion time and later uses that stamp for partial lengths and ack traversal.
- **What our port does:** The port has a custom “covered by obliterate” path and does not preserve all of TS’s insertion-time provenance fields.
- **Why it matters:** A segment inserted in an obliterated slice can be visible to one client but not another, especially during reconnect or old-refSeq length queries.
- **Suggested fix:** ~120-180 LOC / ~2-3 h: align the C# metadata names and state transitions with TS `SegmentInfo` and use identical tests.
- **Regression-test suggestion:** Port `obliterate.concurrent.spec.ts` “segment obliterated on insert overlaps with local obliterate” and “wasRemovedOnInsert computation remains accurate after leaf node is split”.

### Finding M3: Default numeric obliterate endpoint sides are not encoded like TS
- **Impact:** BUG (silent)
- **TS:** `mergeTree.ts:2083-2290`, `sequencePlace.ts` — TS converts a numeric range to sided places with a before start and after end-side semantics for the last covered position.
- **C#:** `SharedString.cs:304-340`, `Client.cs:448-538`, `MergeTree.cs:611-694`, `SequencePlace.cs` — public `ObliterateRange` sends `Side.Before`/`Side.Before` for the numeric path.
- **What TS does:** TS distinguishes numeric obliterate, sided obliterate, and boundary insert eating through `Side.Before`/`Side.After` on `InteriorSequencePlace`.
- **What our port does:** C# uses a different default side pair and then applies custom cover checks.
- **Why it matters:** Boundary concurrent inserts can be eaten or spared differently from TS without obvious failures in existing happy-path tests.
- **Suggested fix:** ~60-100 LOC / ~1-2 h: make numeric `ObliterateRange` produce the same two `SequencePlace` values as TS and re-run sided tests.
- **Regression-test suggestion:** Port `obliterate.spec.ts` endpoint behavior and `obliterate.rangeExpansion.spec.ts` sided-obliterate boundary cases.

### Finding M4: Interval-boundary splitting hook is absent from core editing paths
- **Impact:** BUG (silent)
- **TS:** `sequenceInterval.ts:686-1002`, `mergeTree.ts:1484-1517, 2292-2383` — TS creates endpoint references and preserves endpoint segment boundaries during edits.
- **C#:** `MergeTree.cs:500-694`, `Intervals/SequenceInterval.cs:154-181` — C# split logic is range-boundary oriented and intervals are a separate layer.
- **What TS does:** TS interval endpoints are local references that participate in split/slide behavior so edits around endpoints keep intervals stable.
- **What our port does:** The port updates interval objects around merge-tree edits but does not make merge-tree core maintain the same endpoint boundary invariants.
- **Why it matters:** Edits that split, merge, or zamboni endpoint segments can produce interval endpoints that still look valid but point to different logical ranges.
- **Suggested fix:** ~150 LOC / ~2-3 h: route interval endpoint references through the local-reference collection and enforce split boundaries like TS.
- **Regression-test suggestion:** Port `intervalRebasing.spec.ts` cases for endpoint changes while disconnected and zamboni avoiding pending interval changes.

### Finding M5: Annotating marker IDs does not enforce TS marker-ID immutability
- **Impact:** BUG (silent)
- **TS:** `sharedString.ts:255-339`, `client.annotateMarker.spec.ts:30-31` — TS rejects changing `reservedMarkerIdKey` except to the existing value.
- **C#:** `SharedString.cs:342-375`, `MergeTree.cs:569-609` — C# `AnnotateRange` can apply arbitrary properties to marker segments.
- **What TS does:** TS treats marker IDs as stable identity and explicitly tests null/new/undefined updates as failures.
- **What our port does:** The port lets marker-id property changes flow through normal range annotation unless callers avoid it.
- **Why it matters:** Marker lookup, interval endpoint identity, and remote marker correlation can silently drift from the TS document.
- **Suggested fix:** ~30-60 LOC / ~30-60 min: add marker-specific annotation validation matching TS before applying local and remote annotate ops.
- **Regression-test suggestion:** Port `sharedString.spec.ts` marker ID update failure tests and `client.annotateMarker.spec.ts` valid marker annotation.

### Finding M6: Zamboni does not have TS pending-interval/pending-segment protections
- **Impact:** BUG (silent)
- **TS:** `mergeTree.ts:830-1271`, `intervalRebasing.spec.ts:248-273` — TS avoids zamboni changes that would break pending interval changes or segment groups.
- **C#:** `MergeTree.cs:830-1271` — C# zamboni uses simplified dead/merge checks and has no TS segment-group equivalent.
- **What TS does:** TS keeps segments alive when pending operations still need their identity for resubmit, rollback, and interval rebasing.
- **What our port does:** The port can pack/merge based on structural visibility only.
- **Why it matters:** A reconnect after zamboni can regenerate interval or range ops against segments that no longer exist, yielding wrong ranges without crashing.
- **Suggested fix:** ~120-200 LOC / ~2-4 h: port segment-group blockers or disable zamboni for segments with pending interval/reference state.
- **Regression-test suggestion:** Port `intervalRebasing.spec.ts` “zamboni avoids modifying segments with pending interval changes” and `client.rollback.spec.ts` zamboni rollback cases.

### Finding M7: Zamboni minimum-sequence comparison appears off by one
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `mergeTree.ts:830-1271` — TS scours segments once the collaboration window makes their remove metadata stable.
- **C#:** `MergeTree.cs:830-1271` — C# dead checks use a simplified `MinSeq` threshold.
- **What TS does:** TS has exact MSN semantics tied to segment metadata and partial lengths.
- **What our port does:** The port uses a simpler comparison and may retain just-eligible tombstones longer.
- **Why it matters:** Usually this is memory/performance only, but it is structurally non-parallel and can mask pending-state bugs.
- **Suggested fix:** ~20 LOC / ~30 min: align the condition with TS after metadata parity lands.
- **Regression-test suggestion:** Port `mergeTree.zamboni.spec.ts` and a small tombstone-at-MSN boundary test.

### Finding M8: MergeBlock tile caches and ordinals are not ported
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `mergeTreeNodes.ts`, `mergeTree.ts` marker search paths — TS maintains block-level ordinals and tile caches.
- **C#:** `MergeBlock.cs`, `MergeTree.cs:221-240`, `Client.cs:641-665` — C# walks segments/marker indexes directly.
- **What TS does:** TS marker search relies on block caches for correctness and performance over large trees.
- **What our port does:** The port has a simpler tree shape and no rightmost/leftmost tile cache.
- **Why it matters:** Small tests pass, but large documents may differ in marker search around split/pack boundaries and will not expose TS cache invariants.
- **Suggested fix:** ~100-180 LOC / ~2-3 h: port cache fields and update paths, or explicitly document/scaffold equivalent invariants.
- **Regression-test suggestion:** Port `client.searchForMarker.spec.ts` multi-block and forward/backward excursion equivalence tests.

### Finding M9: `getPosition` absent/detached semantics differ from TS `-1` contract
- **Impact:** BUG (silent)
- **TS:** `client.getPosition.spec.ts:33-65`, `mergeTree.ts` position helpers — TS returns `-1` for removed/detached segments.
- **C#:** `MergeTree.cs:1352-1458`, `LocalReference.cs:23-74` — C# APIs commonly return nullable positions (`null`) for detached/removed references.
- **What TS does:** TS consumers can distinguish valid position `0` from absent `-1` consistently.
- **What our port does:** The port uses nullability in some paths and no public TS-equivalent `getPosition` surface.
- **Why it matters:** Interval and local-reference code can silently skip absent endpoints instead of reporting TS’s sentinel value.
- **Suggested fix:** ~40-80 LOC / ~1 h: expose TS-parallel position helpers and normalize all absent returns to `-1` where the TS API does.
- **Regression-test suggestion:** Port all four `client.getPosition.spec.ts` scenarios and interval tests checking reference position `-1` after obliterate.

### Finding M10: TextSegment split shares property dictionaries instead of TS clone semantics
- **Impact:** BUG (silent)
- **TS:** `textSegment.ts` — TS `TextSegment` split/clone creates independent property bags via merge-tree property utilities.
- **C#:** `TextSegment.cs`, `ISegment.cs` — `CreateSplitSegmentAt` plus metadata copying can share the same `PropertySet` object across split siblings.
- **What TS does:** TS lets later annotation mutate only the intended split range.
- **What our port does:** The port risks both halves observing later dictionary mutations unless every annotation path replaces rather than mutates.
- **Why it matters:** Range-specific annotations can bleed into adjacent text without a crash.
- **Suggested fix:** ~25-50 LOC / ~30-60 min: deep/structural clone property dictionaries when splitting and snapshotting segments.
- **Regression-test suggestion:** Port `sharedString.spec.ts` “can annotate single multi-character segment” and `mergeTree.annotate.spec.ts` split annotation cases.

### Finding M11: Attribution fields are omitted from segment metadata
- **Impact:** Deferred (POC choice, documented)
- **TS:** `segmentInfos.ts`, `snapshotV1.ts:192-365`, attribution tests — TS preserves attribution collections through edits and snapshots.
- **C#:** `ISegment.cs`, `TextSegment.cs`, `SharedStringSnapshotLoader.cs:607-622` — C# has no attribution data model.
- **What TS does:** TS can answer who inserted text and serializes attribution when enabled.
- **What our port does:** The port ignores attribution entirely.
- **Why it matters:** WN may not need attribution today, but parity audit should mark it as an explicit deferred surface.
- **Suggested fix:** ~300+ LOC / ~1-2 days: port attribution collection model, policies, and snapshot format if product requires it.
- **Regression-test suggestion:** If required, port `attributionCollection.spec.ts`, `attributionPolicy.spec.ts`, and `snapshot.spec.ts` attribution cases.

### Finding M12: Tracking groups are absent
- **Impact:** Deferred (POC choice, documented)
- **TS:** `tracking.spec.ts:22-181`, `mergeTree.ts` — TS segments and local references maintain tracking groups.
- **C#:** `ISegment.cs`, `LocalReference.cs`, `MergeTree.cs` — no equivalent tracking-group collection appears in the port.
- **What TS does:** TS uses tracking groups to follow segments/references through split, append, unlink, rollback, and zamboni.
- **What our port does:** The port has no API or data structure for this.
- **Why it matters:** If any C# caller or future interval work depends on tracking groups, behavior cannot be replicated.
- **Suggested fix:** ~150-250 LOC / ~3-5 h: port tracking collections or document as permanently unsupported.
- **Regression-test suggestion:** Port `tracking.spec.ts` only if C# product code will use tracking groups.

### Finding M13: Removed/obliterated segment traversal through hierarchy is custom
- **Impact:** BUG (silent)
- **TS:** `mergeTree.ts:2083-2383`, `obliterate.concurrent.spec.ts:652-674, 1465-1488` — TS deliberately traverses hierarchy when partial length at refSeq is non-zero or ack traversal must pass tombstones.
- **C#:** `MergeTree.cs:1670-1833`, `PartialLengths.cs:169-283` — C# traversal is reimplemented around visible length and custom obliterate coverage.
- **What TS does:** TS traversal rules decide whether to descend into blocks with zero current length but non-zero historical length.
- **What our port does:** The port may prune or traverse using local visible length instead of the same historical criteria.
- **Why it matters:** Concurrent insert/remove/obliterate across block splits can produce wrong text or length without failing simple tests.
- **Suggested fix:** ~100-180 LOC / ~2-3 h: mirror TS traversal predicates and add asserts for parent/child length consistency.
- **Regression-test suggestion:** Port `obliterate.concurrent.spec.ts` “traverses hier block…” and “obliterate ack traversal is not stopped by obliterated segment”.

### Finding M14: Property matching/null semantics are not proven equivalent
- **Impact:** BUG (silent)
- **TS:** `properties.spec.ts:10-51`, `propertyManager.spec.ts` — TS differentiates missing, `undefined`, empty, null, and complex property matching.
- **C#:** `SharedStringOpSerializer.cs:624-1209`, `MergeTree.cs:569-609` — C# maps JSON values into CLR dictionaries/`JsonElement` and annotations into `PropertySet`.
- **What TS does:** TS property helpers have well-tested equality, match, combine, and rollback semantics.
- **What our port does:** The port uses normal .NET dictionary/value equality in several places and may preserve `JsonElement` wrappers.
- **Why it matters:** Annotation filters and marker lookup by labels can silently mismatch TS for null/empty/complex values.
- **Suggested fix:** ~80-140 LOC / ~2 h: port property helper functions exactly and normalize JSON values on deserialize.
- **Regression-test suggestion:** Port `properties.spec.ts`, `propertyManager.spec.ts`, and SharedString null/empty annotation cases.

## Findings — client + op dispatch

### Finding C1: Remote sequence delta events use original op coordinates instead of transformed deltas
- **Impact:** BUG (silent)
- **TS:** `sequence.ts:880-950`, `mergeTree.*.deltaCallback.spec.ts`, `sequenceDeltaEvent.spec.ts:273-3210` — TS emits events from merge-tree delta callbacks after operation transformation.
- **C#:** `SharedString.cs:499-585`, `Client.cs:1680-1908` — C# constructs remote event args from op `Pos1`/`Pos2` and operation type around apply.
- **What TS does:** TS event ranges reflect the actual local view after concurrency, splits, and group transformation.
- **What our port does:** The port can report stale positions/lengths from the submitted op rather than the transformed local effect.
- **Why it matters:** Consumers can update UI/model state at the wrong offsets while the document text itself appears correct.
- **Suggested fix:** ~120-200 LOC / ~2-4 h: drive events from merge-tree delta callbacks and include segment ranges exactly as TS.
- **Regression-test suggestion:** Port the collab matrix from `sequenceDeltaEvent.spec.ts`, especially overlapping insert/delete/annotate and combination cases.

### Finding C2: Group op event aggregation is not TS-parallel
- **Impact:** BUG (silent)
- **TS:** `client.ts:558-666`, `ops.ts:231-235`, `sequenceDeltaEvent.spec.ts` — TS applies each grouped op through the same delta callback pipeline.
- **C#:** `SharedString.cs:792-858`, `Client.cs:548-633` — C# batches and emits simplified aggregate local/remote event paths.
- **What TS does:** TS preserves per-range operation order and transformed positions within a group.
- **What our port does:** The port risks one event with grouped original coordinates or missing child operation details.
- **Why it matters:** Batch consumers can miss sub-deltas or apply them in the wrong order.
- **Suggested fix:** ~80-140 LOC / ~1-2 h: represent group op event ranges exactly as TS `SequenceDeltaEvent` does.
- **Regression-test suggestion:** Port group/batch portions of `client.applyMsg.spec.ts` and `sequenceDeltaEvent.spec.ts` combination cases.

### Finding C3: Ack correlation is pending-op-queue based instead of TS local-op metadata/segment groups
- **Impact:** BUG (silent)
- **TS:** `client.ts:1358-1542`, `segmentGroupCollection.spec.ts` — TS correlates acks with local op metadata and segment groups.
- **C#:** `Client.cs:740-1294, 1680-1908` — C# uses `_pendingOps`, local sequence matching, and custom rebase paths.
- **What TS does:** TS keeps exact segment ranges affected by the submitted op so ack, rollback, and regenerate update the same objects.
- **What our port does:** The port replays/rebases from simplified op objects and current tree positions.
- **Why it matters:** Acking after splits, zamboni, or remote deletes can update the wrong segment or leave stale pending state.
- **Suggested fix:** ~200-350 LOC / ~4-8 h: port local op metadata and segment group collection before adding more reconnect features.
- **Regression-test suggestion:** Port `segmentGroupCollection.spec.ts`, `resetPendingSegmentsToOp.spec.ts`, and `client.rollback.spec.ts` ack/rollback cases.

### Finding C4: Reconnect regeneration is custom and omits TS rollback/squash/stash model
- **Impact:** BUG (silent)
- **TS:** `client.ts:1358-1542`, `client.reconnectFarm.spec.ts`, `client.rollback.spec.ts`, `client.applyStashedOpFarm.spec.ts` — TS regenerates pending ops with local metadata and stashed-op support.
- **C#:** `Client.cs:740-1294` — C# `RebasePendingOps` rewrites pending ops with custom interval and merge-tree handling.
- **What TS does:** TS has tested behavior for rollback, resubmit, squash, stashed ops, and concurrent remote edits.
- **What our port does:** The port implements a subset without the same metadata contract.
- **Why it matters:** Disconnected editing can converge for simple strings but diverge for intervals, markers, obliterate, and annotate over removed ranges.
- **Suggested fix:** ~300+ LOC / ~1-2 days: port TS regenerate/rollback/stash concepts or narrow the supported reconnect contract explicitly.
- **Regression-test suggestion:** Port `client.rebasePosition.spec.ts`, `client.rollback.spec.ts`, `client.reconnectFarm.spec.ts` non-fuzz smoke cases, and interval rebasing tests.

### Finding C5: Relative positions are present in models but unsupported in apply/serialize paths
- **Impact:** BUG (silent)
- **TS:** `client.ts:674-815`, `ops.ts:112-188` — TS resolves `relativePos1`/`relativePos2` in `getValidOpRange`.
- **C#:** `Ops.cs`, `SharedStringOpSerializer.cs:77-505`, `Client.cs:548-633` — C# op classes mention relative positions, but serializers/apply expect numeric `Pos1`/`Pos2`.
- **What TS does:** TS can insert/remove/annotate relative to tombstoned segments loaded from snapshots or concurrent edits.
- **What our port does:** The port drops or rejects those fields and cannot resolve them on receive.
- **Why it matters:** Ops produced by TS clients using relative positions can be no-ops, crashes, or wrong-position edits in C#.
- **Suggested fix:** ~100-180 LOC / ~2-3 h: port relative-position serialization and `getValidOpRange` resolution exactly.
- **Regression-test suggestion:** Port `snapshot.spec.ts` “insert segments relative to removed/obliterated segment” and `client.applyMsg.spec.ts` `getContainingSegment with op`.

### Finding C6: Annotate-adjust is missing
- **Impact:** BUG (crash)
- **TS:** `client.applyMsg.spec.ts:730-924`, `ops.ts:190-223` — TS supports `IMergeTreeAnnotateAdjustMsg` with combine/min/max validation.
- **C#:** `Ops.cs:494-495`, `SharedStringOpSerializer.cs:77-505`, `Client.cs:548-633` — C# has TODO/unsupported handling for annotate-adjust.
- **What TS does:** TS accepts adjust ops locally/remotely and validates min/max constraints.
- **What our port does:** The port cannot process the op shape.
- **Why it matters:** A TS peer can send a valid annotate-adjust op that crashes or is ignored by the C# port.
- **Suggested fix:** ~120-200 LOC / ~2-4 h: port op model, serializer, validation, and property adjust combine rules.
- **Regression-test suggestion:** Port all `client.applyMsg.spec.ts` `annotateRangeAdjust` tests.

### Finding C7: Empty text insert throws instead of TS no-op
- **Impact:** BUG (crash)
- **TS:** `mergeTree.ts:1484-1517`, `sharedString.ts:161-205` — TS zero-length text insert yields no segment/no-op behavior.
- **C#:** `Client.cs:192-247`, `SharedString.cs:225-248` — C# rejects `text.Length == 0`.
- **What TS does:** TS callers can safely issue empty inserts, often through replace or normalized user input.
- **What our port does:** The port throws an exception.
- **Why it matters:** Valid TS-level calls can crash C# even though no document change is required.
- **Suggested fix:** ~10 LOC / ~15 min: treat empty text as no-op and suppress op emission.
- **Regression-test suggestion:** Add a C# test matching TS zero-length insert/replace-zero-range behavior.

### Finding C8: Insert marker requires `markerId`, unlike TS
- **Impact:** BUG (crash)
- **TS:** `sharedString.ts:220-255`, `Marker.make` — TS marker insertion requires `ReferenceType` and optional properties, not a marker id.
- **C#:** `SharedString.cs:250-278`, `Client.cs` marker insert path — C# throws unless props contain non-empty `markerId`.
- **What TS does:** TS supports anonymous markers and marker search by label/properties.
- **What our port does:** The port enforces a stricter C# identity requirement.
- **Why it matters:** Valid TS operations or TS-authored snapshots with marker segments lacking marker IDs can fail in C#.
- **Suggested fix:** ~30-50 LOC / ~30-60 min: allow marker insertion without ID and maintain optional ID index separately.
- **Regression-test suggestion:** Port `sharedString.spec.ts` “can insert marker” with only label/refType properties.

### Finding C9: Minimum-sequence tracking with in-flight ops differs from TS
- **Impact:** BUG (silent)
- **TS:** `client.applyMsg.spec.ts:1086-1130`, `client.ts` collab-window logic — TS updates MSN to min(message.minSeq, pending local constraints).
- **C#:** `Client.cs:1680-1908` — C# updates minSeq during ack/remote apply with simplified pending queue constraints.
- **What TS does:** TS prevents zamboni/partial-length cleanup while in-flight local messages still reference old segments.
- **What our port does:** The port may advance or delay cleanup with different criteria.
- **Why it matters:** Length caches and tombstone cleanup can diverge under concurrent local ops.
- **Suggested fix:** ~50-100 LOC / ~1-2 h: port TS minSeq update formula and add assertions around pending local messages.
- **Regression-test suggestion:** Port `client.applyMsg.spec.ts` “updates minSeq” cases.

### Finding C10: Rollback/revertible APIs are absent
- **Impact:** Deferred (POC choice, documented)
- **TS:** `client.rollback.spec.ts:41-658`, `revertibles.spec.ts`, `sharedString.spec.ts:846-887` — TS supports rollback and revertible operations.
- **C#:** `SharedString.cs`, `Client.cs` — no public rollback/revertible API parity.
- **What TS does:** TS can undo local inserts/removes/annotates and restore local references.
- **What our port does:** The port does not expose or implement this surface.
- **Why it matters:** If WN never uses rollback this is N/A operationally, but it is a major TS surface omission.
- **Suggested fix:** ~300+ LOC / ~1-2 days if required; otherwise explicitly document unsupported.
- **Regression-test suggestion:** Do not require farm/fuzz tests, but port the three SharedString revertible smoke tests if rollback is added.

### Finding C11: Remote/local processing path does not model TS stashed ops
- **Impact:** Deferred (POC choice, documented)
- **TS:** `client.applyStashedOpFarm.spec.ts`, `intervalStashedOps.spec.ts:78-315`, `sequence.ts` load/process APIs — TS supports applying stashed ops.
- **C#:** `SharedString.cs`, `Client.cs`, `IntervalCollection.cs` — no stashed-op API or validation gate is implemented.
- **What TS does:** TS can replay serialized local operations from detached/offline state.
- **What our port does:** The port only handles normal local pending ops and remote messages.
- **Why it matters:** Offline integration scenarios cannot be TS-parallel.
- **Suggested fix:** ~200+ LOC / ~1 day if required; otherwise document as unsupported.
- **Regression-test suggestion:** Port `intervalStashedOps.spec.ts` only if stashed/offline ops are in scope.

### Finding C12: Obliterate reconnect normalization is incomplete
- **Impact:** BUG (silent)
- **TS:** `obliterate.reconnect.spec.ts:35-242`, `client.ts` regenerate path — TS normalizes obliterate groups and preserves segment groups across reconnects.
- **C#:** `Client.cs:740-1294` — C# has a custom rebase path for pending obliterates and intervals.
- **What TS does:** TS prevents obliterate from expanding during rebase while still deleting intended concurrently inserted content.
- **What our port does:** The port mixes numeric/sided range fields with current tree positions and simplified pending metadata.
- **Why it matters:** Reconnect after local obliterate can spare text TS deletes or delete boundary text TS spares.
- **Suggested fix:** ~150-250 LOC / ~4-6 h: port TS obliterate regenerate logic and side normalization.
- **Regression-test suggestion:** Port all non-fuzz `obliterate.reconnect.spec.ts` cases, especially separated group ops and sided obliterate reconnect.

## Findings — partialLengths

### Finding P1: Partial-length algorithm is structurally non-parallel
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `partialLengths.ts:179-229, 341-504` — TS stores `minLength`, ordered `partialLengths`, per-client adjustments, and ref-seq adjustments.
- **C#:** `PartialLengths.cs:51-75, 169-283` — C# uses cumulative dictionaries/rebuild/update helpers.
- **What TS does:** TS length math is designed around multiple perspectives and incremental structural updates.
- **What our port does:** The port implements a different algorithm that attempts equivalent outputs.
- **Why it matters:** Even if tests pass now, reviewer concern is valid: logic is not “as-parallel-to-TS as possible”.
- **Suggested fix:** ~250-400 LOC / ~1 day: port TS structures directly or add proof-style tests covering all TS cases.
- **Regression-test suggestion:** Port `partialLength.spec.ts` and `obliterate.partialLength.spec.ts` before relying on this path.

### Finding P2: Per-client adjustment model is missing
- **Impact:** BUG (silent)
- **TS:** `partialLengths.ts:511-830` — TS accounts local vs remote visibility through per-client adjustments.
- **C#:** `PartialLengths.cs:447-633` — C# tracks local unacked sets and cumulative deltas differently.
- **What TS does:** TS length at refSeq depends on the querying client’s own ops and remote ops up to refSeq.
- **What our port does:** The port infers this from current pending/local metadata.
- **Why it matters:** Lengths can be wrong for remote views of local unacked inserts/removes.
- **Suggested fix:** ~120-200 LOC / ~2-3 h: port per-client adjustment maps and update calls.
- **Regression-test suggestion:** Port `partialLength.spec.ts` single insert/remove local-vs-remote view cases.

### Finding P3: Per-refSeq adjustment map is absent
- **Impact:** BUG (silent)
- **TS:** `partialLengths.ts:845-1005` — TS records adjustments keyed by reference sequence for historical queries.
- **C#:** `PartialLengths.cs:51-75, 169-283` — C# query computes from cumulative state and current sequence bounds.
- **What TS does:** TS supports length queries from older perspectives after later operations have been applied.
- **What our port does:** The port does not maintain the same refSeq-indexed adjustment history.
- **Why it matters:** Snapshot, remote op validation, and relative-position resolution can use wrong historical lengths.
- **Suggested fix:** ~100-160 LOC / ~2 h: port per-refSeq adjustment storage and tests.
- **Regression-test suggestion:** Port `partialLength.spec.ts` aggregation/permutation and `client.rebasePosition.spec.ts` seqNumberFrom cases.

### Finding P4: Overlapping obliterate/remove lengths rely on flattened metadata
- **Impact:** BUG (silent)
- **TS:** `obliterate.partialLength.spec.ts:132-444`, `obliterate.concurrent.spec.ts:392-506` — TS length cache distinguishes overlapping remove and slice-remove stamps.
- **C#:** `PartialLengths.cs:169-283, 447-633`, `ISegment.cs` — C# consumes simplified removed/obliterated state.
- **What TS does:** TS can count a segment removed by one operation from one perspective and by another from a different perspective.
- **What our port does:** The port can only derive one effective historical status in several paths.
- **Why it matters:** Client length can become negative, too high, or inconsistent across peers.
- **Suggested fix:** ~150-250 LOC / ~4 h after M1: update length calculation to use ordered remove stamps.
- **Regression-test suggestion:** Port overlapping remove+obliterate and overlapping obliterate+obliterate tests from `obliterate.partialLength.spec.ts`.

### Finding P5: Local unacked length accounting is not tied to TS segment groups
- **Impact:** BUG (silent)
- **TS:** `partialLengths.ts:511-830`, `client.ts` local metadata — TS accounts local segments through segment groups and localSeq stamps.
- **C#:** `PartialLengths.cs:447-633`, `Client.cs:740-1294` — C# uses local unacked segment sets/pending ops.
- **What TS does:** TS can update a partially split local operation through ack/rebase without losing the original operation identity.
- **What our port does:** The port reconstructs relationships from current segments.
- **Why it matters:** Acking after split/merge can leave stale local length adjustments.
- **Suggested fix:** ~100-180 LOC / ~2-3 h: hook partial lengths to TS-style segment groups.
- **Regression-test suggestion:** Port `client.rollback.spec.ts` “rollback insert and validate partial lengths” and `obliterate.concurrent.spec.ts` “partial lengths updated when local insert is acked”.

### Finding P6: Incremental update path needs TS parent/child invariant checks
- **Impact:** BUG (silent)
- **TS:** `partialLengths.ts:341-504`, `obliterate.concurrent.spec.ts:1591-1883` — TS has extensive incremental partial-length update coverage.
- **C#:** `PartialLengths.cs:169-283` — C# incremental rebuild/update path is custom.
- **What TS does:** TS combines child partial lengths and remote-obliterated lengths with strict parent/child invariants.
- **What our port does:** The port updates compact aggregates without identical invariants/asserts.
- **Why it matters:** Subtree splits can silently make parent length differ from children under rare edit orders.
- **Suggested fix:** ~80-140 LOC / ~2 h: add TS-equivalent combine/update invariants and port tests.
- **Regression-test suggestion:** Port non-fuzz incremental cases under `obliterate.concurrent.spec.ts:1591-1883`, excluding fuzz-labeled cases as required.

### Finding P7: Negative/zero historical length edge cases are under-tested
- **Impact:** BUG (silent)
- **TS:** `obliterate.concurrent.spec.ts:506-530, 1850-1959` — TS has regressions for negative partial lengths and post-insertion wins.
- **C#:** `PartialLengths.cs:51-75, 169-283` — no equivalent static evidence of guards for negative aggregate deltas.
- **What TS does:** TS explicitly guards old regressions where overlapping operations produced negative lengths.
- **What our port does:** The port lacks those regression scenarios and uses different math.
- **Why it matters:** The bug would only surface as wrong lengths in complex collaboration, not simple text mismatch.
- **Suggested fix:** ~40-80 LOC tests plus guard asserts / ~1-2 h.
- **Regression-test suggestion:** Port non-fuzz negative-length regression scenarios; note fuzz-labeled cases should be flagged but not required.

## Findings — localReference

### Finding L1: LocalReferenceCollection structure is collapsed to flat segment lists
- **Impact:** BUG (silent)
- **TS:** `localReference.ts:232-629` — TS maintains before/at/after/tombstone buckets, endpoints, and ordered reference collections.
- **C#:** `LocalReference.cs:23-74`, `MergeTree.cs:2212-2286` — C# has simplified reference objects/lists attached to segments.
- **What TS does:** TS can slide references differently based on side, offset, and tombstone state.
- **What our port does:** The port stores less state and applies simpler slide/detach loops.
- **Why it matters:** References around removes/obliterates can land on the wrong side of content.
- **Suggested fix:** ~200-300 LOC / ~1 day: port `LocalReferenceCollection` substantially as-is.
- **Regression-test suggestion:** Port `client.localReference.spec.ts` removal, offset, split/append, and multi-reference sliding cases.

### Finding L2: StayOnRemove is detached/slid instead of retained on removed segments
- **Impact:** BUG (silent)
- **TS:** `localReference.ts:25-59, 232-629`, `client.localReference.spec.ts:648-703` — TS keeps StayOnRemove references on removed segments in concurrent delete cases.
- **C#:** `MergeTree.cs:2212-2286`, `LocalReference.cs:13-18` — C# remove sliding path detaches/stops references based on simplified flags.
- **What TS does:** TS uses StayOnRemove to preserve references through removal and later resolution.
- **What our port does:** The port can remove or slide them away.
- **Why it matters:** Intervals/markers that should remain anchored to removed content can disappear or move.
- **Suggested fix:** ~80-140 LOC / ~2 h: implement exact StayOnRemove behavior and validation.
- **Regression-test suggestion:** Port `client.localReference.spec.ts` “avoids removing StayOnRemove references on local + remote concurrent delete”.

### Finding L3: Transient reference detach behavior is incomplete
- **Impact:** BUG (silent)
- **TS:** `localReference.ts:25-59, 507-537` — TS allows transient references on removed segments and detaches them in specific unlink paths.
- **C#:** `LocalReference.cs:13-18`, `MergeTree.cs:2212-2286` — C# has `Transient` flags but simplified detach logic.
- **What TS does:** TS differentiates transient references from slide/stay references and validates mutually exclusive flags.
- **What our port does:** The port does not expose the same validation or tombstone handling.
- **Why it matters:** Transient interval endpoints can leak or detach too early.
- **Suggested fix:** ~60-120 LOC / ~1-2 h: port flag validation and detach/link behavior.
- **Regression-test suggestion:** Port `client.localReference.spec.ts` “Transient references can be created on removed segments”.

### Finding L4: SlidingPreference enum is not TS-parallel
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `localReference.ts:25-59` — TS has `SlidingPreference.BACKWARD`/`FORWARD` and flag-based behavior.
- **C#:** `LocalReference.cs:13-18` — C# adds `None` and maps preferences through custom code.
- **What TS does:** TS behavior is expressed through reference type flags and a two-value preference.
- **What our port does:** The port introduces an extra state not present in TS.
- **Why it matters:** This increases the chance callers create references whose behavior has no TS analog.
- **Suggested fix:** ~20 LOC / ~30 min: remove or fully isolate `None` and use TS names/semantics.
- **Regression-test suggestion:** Port backward-sliding tests from `client.localReference.spec.ts` and interval stickiness tests.

### Finding L5: Local reference callbacks and tracking groups are absent
- **Impact:** Deferred (POC choice, documented)
- **TS:** `localReference.ts:87-170, 232-629` — TS references have callbacks, tracking groups, and event hooks.
- **C#:** `LocalReference.cs:23-74` — C# reference is a simple data object.
- **What TS does:** TS consumers can observe reference movement/detach and track references through edits.
- **What our port does:** The port cannot notify or maintain those memberships.
- **Why it matters:** Unsupported if WN does not expose callbacks; otherwise a parity gap.
- **Suggested fix:** ~120-200 LOC / ~3 h if required.
- **Regression-test suggestion:** Port tracking/reference callback tests only if the API is required.

### Finding L6: Endpoint/special segment handling for reconnect is missing
- **Impact:** BUG (silent)
- **TS:** `intervalRebasing.spec.ts:432-462`, `localReference.ts` — TS uses endpoint/special segments to preserve sliding preference across reconnect.
- **C#:** `LocalReference.cs`, `Client.cs:740-1294` — C# has no endpoint segment mechanism.
- **What TS does:** TS can reattach interval endpoints after remote deletes and reconnect while preserving slide direction.
- **What our port does:** The port reuses current segment references and custom rebase rules.
- **Why it matters:** Intervals can land one character off after reconnect.
- **Suggested fix:** ~150 LOC / ~3 h: port endpoint segment strategy or equivalent metadata.
- **Regression-test suggestion:** Port `intervalRebasing.spec.ts` sliding preference after ack/reconnect cases.

### Finding L7: Reference position absent value uses `null` instead of TS `-1`
- **Impact:** BUG (silent)
- **TS:** `referencePositions.ts`, `intervalRebasing.spec.ts:125-143` — TS reports `-1` for obliterated/detached reference positions.
- **C#:** `SequenceInterval.cs:154-181`, `LocalReference.cs:23-74` — C# reference-to-position APIs return nullable positions.
- **What TS does:** TS interval and reference consumers depend on `-1` as a numeric sentinel.
- **What our port does:** The port can serialize/compare `null` differently or omit endpoints.
- **Why it matters:** Compatibility with TS snapshots/events is broken for obliterated endpoints.
- **Suggested fix:** ~30-60 LOC / ~1 h: normalize public/serialized positions to `-1` where TS does.
- **Regression-test suggestion:** Port `intervalRebasing.spec.ts` “reference is -1 for obliterated segment”.

### Finding L8: `canSlideToEndpoint` rules are not ported
- **Impact:** BUG (silent)
- **TS:** `localReference.ts:232-629` — TS decides if sliding may move to start/end sentinels.
- **C#:** `MergeTree.cs:2212-2286` — C# sliding scans visible segments and detaches when no target is found.
- **What TS does:** TS sometimes allows endpoint sliding and sometimes detaches based on flags and range shape.
- **What our port does:** The port has simpler no-target behavior.
- **Why it matters:** References removed at document boundaries can incorrectly vanish or stay.
- **Suggested fix:** ~60-100 LOC / ~1-2 h: implement `canSlideToEndpoint` parity.
- **Regression-test suggestion:** Port local-reference tests “Remove segments to end”, “Remove all segments”, and “slides off end of string”.

## Findings — snapshot

### Finding S1: Snapshot segment metadata arrays are collapsed to scalar fields
- **Impact:** BUG (silent)
- **TS:** `snapshotV1.ts:192-365` — TS serializes `removedSeqs`, `removedClientIds`, `movedSeqs`, `movedClientIds`, and related arrays.
- **C#:** `SharedStringSnapshotLoader.cs:21-115, 607-622` — C# DTOs keep scalar removed/obliterated client/seq fields.
- **What TS does:** TS can preserve multiple remove/obliterate stamps on one segment across snapshot/load.
- **What our port does:** The port can only reload one effective removal/move per segment.
- **Why it matters:** A snapshot taken after overlapping operations can load with different visibility/length in C#.
- **Suggested fix:** ~120-200 LOC / ~3 h after M1: align DTOs and loader with TS arrays.
- **Regression-test suggestion:** Port `snapshot.spec.ts` removals/obliterates above MSN plus overlapping obliterate snapshot regressions.

### Finding S2: Attribution chunks are omitted
- **Impact:** Deferred (POC choice, documented)
- **TS:** `snapshotLoader.ts:105-181`, `snapshotV1.ts:192-365`, `snapshot.spec.ts:167-207` — TS loads/stores attribution chunks.
- **C#:** `SharedStringSnapshotLoader.cs:21-115, 306-548` — no attribution fields are parsed or emitted.
- **What TS does:** TS preserves attribution policy and data across summaries.
- **What our port does:** The port ignores the field entirely.
- **Why it matters:** Not needed if WN does not consume attribution, but it is a documented parity omission.
- **Suggested fix:** ~200+ LOC / ~1 day if required.
- **Regression-test suggestion:** Port `snapshot.spec.ts` with/without attribution only if attribution is in scope.

### Finding S3: Catchup ops are applied during populate instead of returned like TS loader
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `snapshotLoader.ts:183-348` — TS loader returns loaded tree plus catchup messages for normal processing.
- **C#:** `SharedStringSnapshotLoader.cs:306-548` — C# appears to apply catchup ops as part of load/populate.
- **What TS does:** TS routes catchup through the same client apply path, producing normal metadata/event handling.
- **What our port does:** The port has a loader-specific apply path.
- **Why it matters:** This may be functionally OK for simple text, but bypasses TS-parallel op dispatch semantics.
- **Suggested fix:** ~80-140 LOC / ~2 h: return and process catchup ops through `Client.ApplyOp` equivalent.
- **Regression-test suggestion:** Port snapshot catchup tests once a catchup fixture is available.

### Finding S4: Snapshot wire/chunk protocol is simplified
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `snapshotLoader.ts`, `snapshotV1.ts` — TS supports header/body/chunks, legacy V1 extraction, and serializer abstraction.
- **C#:** `SharedStringSnapshotLoader.cs:185-248, 306-548` — C# loader uses bespoke JSON DTOs and validation.
- **What TS does:** TS supports old and current summary layouts.
- **What our port does:** The port supports only the shapes implemented in the POC loader.
- **Why it matters:** TS-authored summaries outside the tested subset can fail to load.
- **Suggested fix:** ~150-250 LOC / ~4 h: port the loader state machine and legacy extraction cases.
- **Regression-test suggestion:** Port `snapshotlegacy.spec.ts`, `snapshotVersion.spec.ts`, and `marshalling.spec.ts` segment spec tests.

### Finding S5: Snapshot text-length validation can ignore marker segment length
- **Impact:** BUG (silent)
- **TS:** `snapshotLoader.ts:105-181`, `Marker`/`TextSegment` specs — TS segment length includes markers but `getText` omits marker text unless marker-aware APIs are used.
- **C#:** `SharedStringSnapshotLoader.cs:1009-1067`, `MergeTree.cs:191-208` — C# validates using text output while marker segments are TODO placeholders.
- **What TS does:** TS distinguishes sequence length from rendered text length.
- **What our port does:** The port’s validation can compare the wrong measure when snapshots contain markers.
- **Why it matters:** A marker-containing snapshot may be rejected or accepted with an incorrect length invariant.
- **Suggested fix:** ~40-80 LOC / ~1 h: validate merge-tree length, not rendered text, and add marker snapshot tests.
- **Regression-test suggestion:** Port `sharedString.spec.ts` marker summary/load and `snapshot.spec.ts` annotated segment load cases.

### Finding S6: Relative inserts after loading tombstones are not supported
- **Impact:** BUG (silent)
- **TS:** `snapshot.spec.ts:68-121` — TS can insert after/relative to removed or obliterated segments loaded from snapshot.
- **C#:** `SharedStringSnapshotLoader.cs:306-548`, `Client.cs:548-633` — C# reloads tombstones but relative op resolution is not wired.
- **What TS does:** TS preserves tombstone metadata so future ops can target positions around them.
- **What our port does:** The port may load the tombstone but cannot later resolve relative positions against it.
- **Why it matters:** Interop with TS clients after snapshot load can diverge.
- **Suggested fix:** ~100-180 LOC / ~2-3 h with C5.
- **Regression-test suggestion:** Port `snapshot.spec.ts` “can insert segments relative to removed/obliterated segment loaded from snapshot”.

### Finding S7: Summary/load policy flags are not TS-parallel
- **Impact:** Deferred (POC choice, documented)
- **TS:** `sequence.ts`, `snapshotLoader.ts`, `snapshot.spec.ts:191-207` — TS initialization and snapshot presence can override attribution/policy flags.
- **C#:** `SharedStringSnapshotLoader.cs`, `README.md:57-84` — C# exposes a feasibility-loader scope, not full policy handling.
- **What TS does:** TS loader can alter runtime behavior based on snapshot contents.
- **What our port does:** The port omits those policy decisions.
- **Why it matters:** Future feature flags may behave differently after load.
- **Suggested fix:** ~80-120 LOC / ~2 h if policies enter scope.
- **Regression-test suggestion:** Port `snapshot.spec.ts` policy override tests if attribution/policies are required.

### Finding S8: Snapshot serializer does not round-trip unknown/unsupported segment specs like TS
- **Impact:** BUG (crash)
- **TS:** `marshalling.spec.ts:35-60`, `snapshotLoader.ts:105-181` — TS returns `undefined` for unrecognized segment JSON specs where appropriate.
- **C#:** `SharedStringOpSerializer.cs:624-1209`, `SharedStringSnapshotLoader.cs` — C# deserialization validates strongly and throws on unsupported shapes.
- **What TS does:** TS can skip/handle unknown specs in marshalling tests.
- **What our port does:** The port can crash when encountering a valid TS fallback/unknown case.
- **Why it matters:** Forward compatibility with TS summaries is weaker.
- **Suggested fix:** ~40-80 LOC / ~1 h: mirror TS unknown-spec handling and add serializer tests.
- **Regression-test suggestion:** Port `marshalling.spec.ts` “returns undefined for unrecognized JSON spec”.

## Findings — ops / wire format

### Finding O1: `ReferenceType` numeric values do not match TS wire values
- **Impact:** BUG (silent)
- **TS:** `ops.ts:10-48` — TS values: `Tile=0x1`, `RangeBegin=0x10`, `RangeEnd=0x20`, `SlideOnRemove=0x40`, `StayOnRemove=0x80`, `Transient=0x100`.
- **C#:** `Marker.cs` — C# values include `NestBegin=0x2`, `NestEnd=0x4`, `RemoveOnInsert=0x100`, `SlideOnRemove=0x200`, `StayOnRemove=0x400`, `Transient=0x800`.
- **What TS does:** TS serializes these bit flags directly in marker/reference/interval operations.
- **What our port does:** The port emits and reads different bits for the same semantic names.
- **Why it matters:** This is a top-priority wire corruption bug: TS peers interpret C# markers/interval endpoints incorrectly and vice versa.
- **Suggested fix:** ~20 LOC / ~30 min plus migration check: change enum constants to exact TS values and remove non-TS bits or gate them off-wire.
- **Regression-test suggestion:** Add serializer round-trip tests for every `ReferenceType` flag and port marker/interval endpoint tests.

### Finding O2: Interval operations are modeled as merge-tree delta types
- **Impact:** BUG (silent)
- **TS:** `ops.ts:61-71` — TS `MergeTreeDeltaType` covers merge-tree ops only; interval ops live in sequence interval collection/map operation layers.
- **C#:** `Ops.cs:16-67, 406-481` — C# adds `IntervalAdd`, `IntervalDelete`, `IntervalChange`, and `IntervalPropertyChanged` delta enum values.
- **What TS does:** TS wire discriminators keep interval collection ops separate from merge-tree op codes.
- **What our port does:** The port overloads the merge-tree delta enum with C#-specific values.
- **Why it matters:** A TS peer will not understand these operation types as merge-tree ops; C# may also mis-route TS interval ops.
- **Suggested fix:** ~120-220 LOC / ~4 h: move interval op serialization to TS interval collection op shape.
- **Regression-test suggestion:** Port `intervalCollection.spec.ts` remote collaboration and serializer tests.

### Finding O3: Annotate-adjust wire fields are not serialized/deserialized
- **Impact:** BUG (crash)
- **TS:** `ops.ts:190-223` — TS op shape includes adjust/combine/min/max fields.
- **C#:** `SharedStringOpSerializer.cs:77-505`, `Ops.cs:494-495` — C# does not write/read/apply these fields.
- **What TS does:** TS peers can send valid annotate-adjust messages.
- **What our port does:** The port treats them as unsupported or malformed.
- **Why it matters:** Interop crashes or data loss occur for numeric property adjustments.
- **Suggested fix:** ~120-200 LOC / ~2-4 h with C6.
- **Regression-test suggestion:** Port `client.applyMsg.spec.ts` annotateRangeAdjust tests.

### Finding O4: Relative position fields are omitted on the wire
- **Impact:** BUG (silent)
- **TS:** `ops.ts:112-188`, `client.ts:674-815` — TS supports `relativePos1`/`relativePos2` fields in op messages.
- **C#:** `SharedStringOpSerializer.cs:77-505` — C# writes numeric `pos1`/`pos2` and does not preserve relative fields.
- **What TS does:** TS uses relative fields for ops around removed/obliterated segments and snapshot tombstones.
- **What our port does:** The port drops those fields during serialization.
- **Why it matters:** A C# relay/round-trip can turn a precise relative op into an invalid or different absolute op.
- **Suggested fix:** ~80-140 LOC / ~2 h with C5.
- **Regression-test suggestion:** Add JSON golden tests using TS relative op fixtures.

### Finding O5: Group op nesting/validation differs from TS
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `opBuilder.ts`, `ops.ts:231-235`, `client.ts:558-666` — TS constructs and applies group ops through standard op handling.
- **C#:** `SharedStringOpSerializer.cs:77-251`, `Client.cs:548-633` — C# rejects/simplifies nested or unsupported group shapes.
- **What TS does:** TS’s exact accepted group shape is part of the wire protocol.
- **What our port does:** The port enforces a narrower shape.
- **Why it matters:** This may be acceptable if C# never receives nested groups, but it is not TS-parallel.
- **Suggested fix:** ~30-60 LOC: match TS validation exactly or document unsupported shapes.
- **Regression-test suggestion:** Add group op golden serializer tests and port group-related `client.applyMsg.spec.ts` cases.

### Finding O6: Property JSON normalization can leak .NET value wrappers
- **Impact:** BUG (silent)
- **TS:** `properties.spec.ts`, `sharedString.ts` annotation paths — TS properties are plain JS values with JS equality/null semantics.
- **C#:** `SharedStringOpSerializer.cs:624-1209` — C# deserializes arbitrary JSON into CLR/`JsonElement`-style values in property dictionaries.
- **What TS does:** TS compares and serializes property values without `JsonElement` wrappers.
- **What our port does:** The port can compare wrapper identity or serialize a wrapper shape unless normalized.
- **Why it matters:** Property matching, marker labels, and interval properties can silently fail.
- **Suggested fix:** ~80-120 LOC / ~2 h: normalize JSON recursively into stable CLR primitives/arrays/maps and implement TS equality helpers.
- **Regression-test suggestion:** Port `properties.spec.ts` and SharedString null/empty/complex property tests.

### Finding O7: Sided obliterate wire defaults are not TS-parallel
- **Impact:** BUG (silent)
- **TS:** `sequencePlace.ts`, `ops.ts`, `obliterate.rangeExpansion.spec.ts:323+` — TS has explicit sided place wire shape and numeric-range normalization.
- **C#:** `SequencePlace.cs`, `SharedStringOpSerializer.cs:519-590` — C# serializes its own sided place defaults and nullable side fields.
- **What TS does:** TS peers expect exact side field names/default behavior.
- **What our port does:** The port may emit a semantically similar but not identical shape.
- **Why it matters:** Boundary behavior changes when TS and C# clients collaborate.
- **Suggested fix:** ~60-100 LOC / ~1-2 h: align field names, omitted-field defaults, and numeric normalization.
- **Regression-test suggestion:** Add wire golden tests for before/after start/end combinations.

### Finding O8: Operation stamps are mutable classes rather than immutable TS records
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `stamps.ts:24-47` — TS stamps are readonly and treated as immutable.
- **C#:** `Stamps.cs:19-35` — C# stamp properties are settable.
- **What TS does:** TS avoids mutating historical operation provenance.
- **What our port does:** The port can accidentally mutate a stamp shared by segment metadata.
- **Why it matters:** No current evidence of mutation bugs, but mutability is a structural deviation with silent-bug risk.
- **Suggested fix:** ~20-40 LOC / ~30-60 min: make stamps immutable records or init-only properties.
- **Regression-test suggestion:** Port `stamps.spec.ts` equality/comparison plus add mutation-safety regression around shared stamps.

### Finding O9: OpBuilder helpers are not ported
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `opBuilder.ts` — TS centralizes op construction including ranges, side conversion, and group creation.
- **C#:** `SharedString.cs`, `Client.cs`, `Ops.cs` — C# constructs op objects manually at call sites.
- **What TS does:** TS reduces drift by using one builder path for local and regenerated ops.
- **What our port does:** The port duplicates construction logic.
- **Why it matters:** Future fixes can patch one path and miss another; this already shows up in sided/relative defaults.
- **Suggested fix:** ~80-120 LOC / ~2 h: add a C# OpBuilder mirroring TS and use it everywhere.
- **Regression-test suggestion:** Add unit tests for every local public API producing expected JSON.

### Finding O10: Reserved marker/reference flags include non-TS bits
- **Impact:** BUG (silent)
- **TS:** `referencePositions.ts`, `ops.ts:10-48` — TS currently defines the reference flag set consumed on the wire.
- **C#:** `Marker.cs`, `IntervalUtils.cs` — C# includes `NestBegin`, `NestEnd`, `RemoveOnInsert`, and interval types not present in current TS wire flags.
- **What TS does:** TS readers may treat unknown bits as meaningful old flags or invalid values.
- **What our port does:** The port can emit flags that no current TS code expects.
- **Why it matters:** Wire compatibility and future upgrades become brittle.
- **Suggested fix:** ~20-40 LOC: remove unsupported flags from wire enum or isolate legacy aliases behind conversion.
- **Regression-test suggestion:** Add a test that C# never emits unknown `ReferenceType` bits for markers or interval endpoints.

## Findings — SharedString public API

### Finding SS1: `insertText` lacks optional properties
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `sharedString.ts:161-205` — `insertText(pos, text, props?)` can create annotated text.
- **C#:** `SharedString.cs:225-248` — `InsertText(int position, string text)` has no properties argument.
- **What TS does:** TS callers can insert text and properties atomically.
- **What our port does:** The port requires separate annotation or cannot express the op.
- **Why it matters:** This changes event/op shapes and can expose transient unannotated text.
- **Suggested fix:** ~40-80 LOC / ~1 h: add optional `PropertySet` to insert text and serialize segment properties.
- **Regression-test suggestion:** Port `sharedString.spec.ts` local/remote insert-with-properties and annotation-on-insert scenarios.

### Finding SS2: `removeText` is renamed to `DeleteText`
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `sharedString.ts:161-205` — public API is `removeText(start, end)`.
- **C#:** `SharedString.cs:280-302` — public C# API is `DeleteText(start, end)`.
- **What TS does:** TS API naming is the parity target.
- **What our port does:** The port uses a .NET-style alternate name only.
- **Why it matters:** This may be acceptable wrapper design, but it is not parallel and complicates test-porting.
- **Suggested fix:** ~10 LOC: add `RemoveText` alias and keep `DeleteText` if desired.
- **Regression-test suggestion:** Port `sharedString.spec.ts` remove text tests using TS names in C# wrapper tests.

### Finding SS3: `replaceText` is missing as a TS-parallel public API
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `sharedString.ts:105-115, 529-546` — TS exposes and tests `replaceText` including zero/negative ranges.
- **C#:** `SharedString.cs` — no direct TS-named `ReplaceText` wrapper was found.
- **What TS does:** TS replace normalizes remove+insert behavior behind one API.
- **What our port does:** The port requires callers/tests to compose delete+insert manually.
- **Why it matters:** Composed ops may not match TS group/event behavior.
- **Suggested fix:** ~30-60 LOC / ~1 h: add `ReplaceText` emitting the same op/group semantics as TS.
- **Regression-test suggestion:** Port `sharedString.spec.ts` replace local/remote/zero/negative range cases.

### Finding SS4: Relative insertion APIs are missing
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `sharedString.ts`, `sequence.ts:140-372` — TS exposes insertion at/reference-position helpers through the SharedSegmentSequence layer.
- **C#:** `SharedString.cs` — no TS-parallel relative insert API is exposed.
- **What TS does:** TS can insert relative to markers/references/tombstones.
- **What our port does:** The port exposes only numeric positions.
- **Why it matters:** Interop with APIs that intentionally avoid absolute positions is missing.
- **Suggested fix:** ~80-120 LOC / ~2 h after C5: expose and wire relative APIs.
- **Regression-test suggestion:** Port snapshot-relative insertion tests and `client.rebasePosition.spec.ts`.

### Finding SS5: Marker search signature and label semantics differ
- **Impact:** BUG (silent)
- **TS:** `sharedString.ts` / `client.searchForMarker.spec.ts:35-661` — TS `searchForMarker(startPos, markerLabel, forwards?)` requires a marker label and searches by that label.
- **C#:** `SharedString.cs:378-392` — C# signature is `SearchForMarker(int startPos, bool forwards = true, string? tileLabel = null)`.
- **What TS does:** TS cannot ask for “any tile” by omitting the label and parameter order is label before direction.
- **What our port does:** The port changes both call shape and semantics.
- **Why it matters:** Callers/tests can pass a string as the second argument in TS-equivalent code and get a different C# overload/behavior; unlabeled matches can mask missing labels.
- **Suggested fix:** ~30-60 LOC: add TS-shaped overload and require label matching by default.
- **Regression-test suggestion:** Port the full `client.searchForMarker.spec.ts` non-farm suite.

### Finding SS6: Marker-aware text APIs are missing
- **Impact:** Deferred (POC choice, documented)
- **TS:** `sharedString.ts` — TS exposes `getTextWithPlaceholders`, `getTextRangeWithMarkers`, and related marker-aware extraction helpers.
- **C#:** `SharedString.cs:124-139`, `MergeTree.cs:191-208` — C# `GetText` skips marker placeholders with a TODO.
- **What TS does:** TS callers can include marker placeholders or marker objects in text ranges.
- **What our port does:** The port only returns concatenated text segments.
- **Why it matters:** Any consumer relying on marker positions in extracted ranges cannot be ported directly.
- **Suggested fix:** ~80-150 LOC / ~2-3 h: implement marker-aware range extraction and placeholder option.
- **Regression-test suggestion:** Port marker range/text extraction tests from sequence/shared-string tests when in scope.

### Finding SS7: `SequenceDeltaEvent` shape is simplified
- **Impact:** BUG (silent)
- **TS:** `sequenceDeltaEvent.spec.ts:3140-3210`, `sequence.ts` — TS event ranges expose segment arrays, operation kind, and continuous/noncontinuous ranges.
- **C#:** `SharedString.cs:17-60, 499-585` — C# event args expose simplified text/position/range fields.
- **What TS does:** TS consumers can inspect the exact segments affected by an edit.
- **What our port does:** The port exposes a smaller, non-TS event contract.
- **Why it matters:** Tests that only compare final text will miss UI/model event regressions.
- **Suggested fix:** ~150-250 LOC / ~4 h: port TS event classes or provide a parallel compatibility layer.
- **Regression-test suggestion:** Port `SequenceDeltaEventClass.ranges` tests and collab delta tests.

### Finding SS8: SharedString reentrancy guard is absent
- **Impact:** BUG (silent)
- **TS:** `reentrancy.spec.ts:25-195` — TS can prevent or log reentrant SharedString mutations during event handling.
- **C#:** `SharedString.cs` — no equivalent reentrancy option/guard was found.
- **What TS does:** TS protects internal consistency when callbacks mutate the same SharedString.
- **What our port does:** The port allows reentrant changes through normal public methods.
- **Why it matters:** Reentrant event handlers can corrupt event ordering or local pending state without final-text tests covering it.
- **Suggested fix:** ~80-120 LOC / ~2 h: add a mutation guard and option matching TS behavior.
- **Regression-test suggestion:** Port `reentrancy.spec.ts` local reentrancy, consistency, and logging-count tests.

### Finding SS9: Load/process/summarize APIs are not TS-parallel
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `sequence.ts:880-950`, `sharedString.ts` — TS SharedString inherits SharedObject/channel lifecycle APIs.
- **C#:** `SharedString.cs`, `SharedStringSnapshotLoader.cs` — C# exposes bespoke load/save/send interfaces for POC usage.
- **What TS does:** TS lifecycle integrates op processing, summarization, GC/handles, and interval collections.
- **What our port does:** The port uses standalone helpers.
- **Why it matters:** This may be the intended WN integration layer, but it is a large structural deviation.
- **Suggested fix:** ~N/A unless targeting full Fluid runtime parity; otherwise document the boundary clearly.
- **Regression-test suggestion:** Add integration tests around whatever C# runtime adapter is expected to provide.

### Finding SS10: README describes stale POC scope
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `sharedString.ts`, interval and obliterate sources — current TS surface includes markers, intervals, local references, obliterate, snapshots.
- **C#:** `README.md:57-84` — documentation still describes several items as deferred/POC despite code implementing subsets.
- **What TS does:** TS parity review relies on clear documented intent.
- **What our port does:** The port documentation does not reflect actual implementation boundaries.
- **Why it matters:** Reviewers can misclassify bugs as intentional or miss unsupported areas.
- **Suggested fix:** ~20-40 LOC docs / ~30 min: update README after deciding which findings are deferred vs bugs.
- **Regression-test suggestion:** No code test; add docs review checklist after fixes.

## Findings — IntervalCollection

### Finding IC1: Interval collection is explicitly a subset, not TS-parallel
- **Impact:** Deferred (POC choice, documented)
- **TS:** `intervalCollection.ts:379-1852` — TS collection has add/change/delete, indexes, events, pending state, rollback, resubmit, labels, and serialization.
- **C#:** `Intervals/IntervalCollection.cs:1-8, 126-238` — C# file declares a POC/subset-style implementation.
- **What TS does:** TS interval collections are a full SharedString feature.
- **What our port does:** The port implements a selected subset.
- **Why it matters:** This is acceptable only if product scope explicitly excludes the rest; otherwise many following findings are bugs.
- **Suggested fix:** ~Large: port feature-by-feature or document supported subset.
- **Regression-test suggestion:** Start with non-fuzz `intervalCollection.spec.ts` happy-path add/change/delete/query tests.

### Finding IC2: Interval endpoint reference flags use the wrong `ReferenceType` bits
- **Impact:** BUG (silent)
- **TS:** `sequenceInterval.ts:686-1002` — TS interval endpoints use `RangeBegin`/`RangeEnd`, `StayOnRemove` while pending, then `SlideOnRemove` after ack.
- **C#:** `Intervals/SequenceInterval.cs:154-181`, `Marker.cs` — C# endpoint flags map through the non-TS enum and simplified preferences.
- **What TS does:** TS serializes endpoint marker/reference flags directly.
- **What our port does:** The port emits/consumes different flag numbers and lifecycle transitions.
- **Why it matters:** Intervals exchanged with TS peers can have endpoints interpreted as unrelated flags.
- **Suggested fix:** ~40-80 LOC after O1: fix enum values and pending/acked flag transitions.
- **Regression-test suggestion:** Port `intervalCollection.spec.ts` endpoint reference-type tests and `intervalRebasing.spec.ts` sliding-preference cases.

### Finding IC3: Stickiness/startSide/endSide are missing
- **Impact:** BUG (silent)
- **TS:** `intervalUtils.ts:193-271`, `sequenceInterval.ts:133-186, 327-684` — TS supports interval stickiness and serialized endpoint sides.
- **C#:** `Intervals/IntervalUtils.cs`, `Intervals/SequenceInterval.cs:34-93` — C# lacks equivalent stickiness fields and side serialization.
- **What TS does:** TS determines how interval endpoints move when text is inserted at boundaries.
- **What our port does:** The port uses simplified sliding preferences.
- **Why it matters:** Boundary insertions can put intervals on the wrong side of new text.
- **Suggested fix:** ~120-200 LOC / ~3-5 h: port `IntervalStickiness`, side fields, serialization, and movement rules.
- **Regression-test suggestion:** Port `intervalStashedOps.spec.ts` stickiness round-trip and `intervalRebasing.spec.ts` full-stickiness cases.

### Finding IC4: Pending interval consensus/resubmit metadata is missing
- **Impact:** BUG (silent)
- **TS:** `intervalCollection.ts:680-1697` — TS tracks pending add/change/delete ops, local metadata, ACK, resubmit, and remote rebasing.
- **C#:** `Intervals/IntervalCollection.cs:250-333, 356-527` — C# has simpler remote/ack handling.
- **What TS does:** TS can reconcile interval changes that race with text edits and remote interval ops.
- **What our port does:** The port updates interval state directly with limited pending metadata.
- **Why it matters:** Concurrent interval edits can converge incorrectly while text converges.
- **Suggested fix:** ~250-400 LOC / ~1 day: port pending-event metadata and resubmit hooks.
- **Regression-test suggestion:** Port `intervalRebasing.spec.ts`, `intervalCollection.rollback.spec.ts`, and non-fuzz interval collaboration tests.

### Finding IC5: Interval rollback/stashed ops are unsupported
- **Impact:** Deferred (POC choice, documented)
- **TS:** `intervalCollection.rollback.spec.ts`, `intervalStashedOps.spec.ts:78-315` — TS handles rollback and stashed interval ops.
- **C#:** `Intervals/IntervalCollection.cs` — no equivalent public support found.
- **What TS does:** TS can undo or replay interval local operations.
- **What our port does:** The port does not implement those features.
- **Why it matters:** Offline/revertible interval workflows are not portable.
- **Suggested fix:** ~200+ LOC if required; otherwise document unsupported.
- **Regression-test suggestion:** Port rollback/stashed interval tests only if those workflows enter scope.

### Finding IC6: Interval events do not match TS event semantics
- **Impact:** BUG (silent)
- **TS:** `intervalCollection.events.spec.ts`, `intervalCollection.ts:379-678` — TS emits add/change/delete/property events with specific local/remote metadata.
- **C#:** `Intervals/IntervalCollection.cs:126-238, 250-333` — C# exposes simplified events and arguments.
- **What TS does:** TS consumers can distinguish local vs remote, changed endpoints, property deltas, and deleted interval states.
- **What our port does:** The port sends fewer/different event details.
- **Why it matters:** Event-driven code can become inconsistent even when intervals end at the same positions.
- **Suggested fix:** ~120-200 LOC / ~3 h: port event types and metadata payloads.
- **Regression-test suggestion:** Port `intervalCollection.events.spec.ts` complete non-fuzz suite.

### Finding IC7: Interval serialized format is not TS V1/V2 compatible
- **Impact:** BUG (silent)
- **TS:** `intervalCollection.snapshot.spec.ts`, `intervalCollection.ts` serialization paths — TS supports labels, compressed intervals, and versioned summary format.
- **C#:** `Intervals/IntervalCollection.cs`, `SharedStringSnapshotLoader.cs` — C# uses custom DTO/op serialization for intervals.
- **What TS does:** TS summaries can round-trip interval collections across versions.
- **What our port does:** The port can only understand its own subset shape.
- **Why it matters:** A TS-authored document with intervals may load without intervals or with wrong endpoints/properties.
- **Suggested fix:** ~200-350 LOC / ~1 day: port TS interval collection summary format.
- **Regression-test suggestion:** Port `intervalCollection.snapshot.spec.ts` and detached interval collection tests.

### Finding IC8: Interval API signatures differ from TS
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `intervalCollection.ts:1722-1852` — TS exposes iterators and query APIs such as `findOverlappingIntervals`, `findIntervalsWithEndpointInRange`, labels, and property helpers.
- **C#:** `Intervals/IntervalCollection.cs:126-238` — C# exposes a smaller .NET-shaped API.
- **What TS does:** TS tests and callers use a broad query surface.
- **What our port does:** The port requires alternative calls or lacks some queries.
- **Why it matters:** This is a structural deviation and hides index parity gaps.
- **Suggested fix:** ~100-180 LOC: add TS-shaped wrappers and map to C# types.
- **Regression-test suggestion:** Port query API tests from `intervalCollection.spec.ts`, `endpointInRangeIndex.spec.ts`, and `startpointInRangeIndex.spec.ts`.

### Finding IC9: Interval property delta semantics are simplified
- **Impact:** BUG (silent)
- **TS:** `intervalCollection.ts`, `sequenceInterval.ts:327-684` — TS property managers track previous values, null removals, property changed ops, and events.
- **C#:** `Intervals/IntervalCollection.cs:356-527`, `SequenceInterval.cs:34-93` — C# stores property dictionaries with simpler update behavior.
- **What TS does:** TS distinguishes property overwrite, delete/null, and no-op changes.
- **What our port does:** The port may treat null or missing values differently.
- **Why it matters:** Remote property updates can silently fail or emit wrong deltas.
- **Suggested fix:** ~80-140 LOC / ~2 h: port property manager semantics used by intervals.
- **Regression-test suggestion:** Port interval property change tests from `intervalCollection.spec.ts`, `intervalCollection.events.spec.ts`, and `intervalStashedOps.spec.ts`.

### Finding IC10: Interval deletion on slide-off differs from TS
- **Impact:** BUG (silent)
- **TS:** `intervalRebasing.spec.ts:86-162, 596-762` — TS deletes intervals that slide off and suppresses undefined-interval events in specific cases.
- **C#:** `Intervals/IntervalCollection.cs:250-333, 356-527` — C# has simplified slide/delete handling.
- **What TS does:** TS has subtle event and state rules when endpoints detach during remove/obliterate/rebase.
- **What our port does:** The port can keep an invalid interval, delete it too early, or emit an event with missing interval data.
- **Why it matters:** Consumers will see stale intervals or wrong events after collaboration.
- **Suggested fix:** ~120-180 LOC / ~3 h: port TS endpoint detach/delete rules and event filtering.
- **Regression-test suggestion:** Port all non-fuzz `intervalRebasing.spec.ts` slide-off/event tests.

## Findings — interval indexes

### Finding II1: Indexes are list-based rather than TS red-black/interval indexes
- **Impact:** Deviation (correct, non-parallel)
- **TS:** `intervalIndex/*.ts` — TS uses specialized indexes with comparator utilities for overlap/endpoint queries.
- **C#:** `Intervals/*Index.cs` — C# implementations are list-based scans.
- **What TS does:** TS index structure affects deterministic ordering and performance.
- **What our port does:** The port likely returns correct small results but does not preserve algorithmic behavior.
- **Why it matters:** Large interval collections can degrade and order ties differently.
- **Suggested fix:** ~150-250 LOC / ~4 h: port comparator/index utilities or add order-normalizing wrappers.
- **Regression-test suggestion:** Port index massive/random tests as deterministic non-fuzz equivalents.

### Finding II2: Endpoint-in-range boundary inclusivity differs
- **Impact:** BUG (silent)
- **TS:** `endpointInRangeIndex.spec.ts`, `intervalIndex/endpointInRangeIndex.ts` — TS treats query endpoints inclusively as specified by tests.
- **C#:** `Intervals/EndpointInRangeIndex.cs` — C# method naming/logic uses range comparisons that appear half-open in places.
- **What TS does:** TS returns intervals whose endpoints exactly equal query boundaries.
- **What our port does:** The port can exclude end-boundary matches.
- **Why it matters:** Boundary interval queries silently miss results.
- **Suggested fix:** ~30-60 LOC / ~1 h: align comparisons and add boundary tests.
- **Regression-test suggestion:** Port `endpointInRangeIndex.spec.ts` exact-boundary, negative, empty, and duplicate endpoint cases.

### Finding II3: Startpoint-in-range boundary/invalid-query semantics differ
- **Impact:** BUG (silent)
- **TS:** `startpointInRangeIndex.spec.ts:91-230` — TS defines empty results for start>end, negative endpoints, exact boundaries, duplicates, and massive random inputs.
- **C#:** `Intervals/StartpointInRangeIndex.cs` — C# simplified scan lacks evidence of all TS invalid/boundary cases.
- **What TS does:** TS behavior is precisely tested for boundary inclusivity and invalid ranges.
- **What our port does:** The port may rely on caller validation or half-open rules.
- **Why it matters:** Queries can silently omit or include intervals at range edges.
- **Suggested fix:** ~30-60 LOC / ~1 h: port exact predicates from TS index.
- **Regression-test suggestion:** Port all non-random `startpointInRangeIndex.spec.ts` cases and one deterministic large case.

### Finding II4: Overlap comparator/tie-breaker is not TS-parallel
- **Impact:** BUG (silent)
- **TS:** `overlappingIntervalsIndex.ts`, `intervalIndexUtils.ts`, `sequenceInterval.ts:192-325` — TS uses interval comparison utilities for ordering and overlap queries.
- **C#:** `Intervals/OverlappingIntervalsIndex.cs`, `SequenceInterval.cs:100-129` — C# uses simplified comparisons.
- **What TS does:** TS returns intervals in deterministic order when endpoints tie and handles 0-length intervals.
- **What our port does:** The port can return a different order or treat collapsed intervals differently.
- **Why it matters:** Consumers relying on stable interval ordering can diverge without text differences.
- **Suggested fix:** ~50-100 LOC / ~1-2 h: port `compare`, `overlaps`, and index sort functions exactly.
- **Regression-test suggestion:** Port overlapping interval query/order tests from `intervalCollection.spec.ts` and `overlappingIntervalsIndex` coverage if present.

### Finding II5: ID index duplicate/delete behavior is under-specified
- **Impact:** BUG (silent)
- **TS:** `idIntervalIndex.ts`, interval collection tests — TS maps interval IDs and updates the index on add/change/delete/dispose.
- **C#:** `Intervals/IdIntervalIndex.cs` — C# has a simple ID map.
- **What TS does:** TS enforces identity semantics across pending and remote operations.
- **What our port does:** The port may overwrite duplicates or fail to remove stale IDs after remote delete/change.
- **Why it matters:** Lookup by interval id can return deleted or wrong intervals.
- **Suggested fix:** ~40-80 LOC / ~1 h: port TS ID index update/duplicate semantics.
- **Regression-test suggestion:** Port interval add/delete/change by id tests, including duplicate add and remote delete.

### Finding II6: Indexes are not updated from local-reference endpoint movement like TS
- **Impact:** BUG (silent)
- **TS:** `intervalCollection.ts:147-341`, `sequenceInterval.ts` — TS local interval indexes observe endpoint movement and re-index intervals.
- **C#:** `Intervals/IntervalCollection.cs:250-527`, `*Index.cs` — C# updates indexes during explicit interval operations, not full local-reference movement parity.
- **What TS does:** TS query results change when text edits slide endpoints.
- **What our port does:** The port can leave indexes stale after remove/obliterate/rebase slides endpoints.
- **Why it matters:** Subsequent queries return pre-edit interval positions.
- **Suggested fix:** ~100-180 LOC / ~3 h: wire endpoint movement notifications into every index.
- **Regression-test suggestion:** Port `intervalRebasing.spec.ts` plus index query checks after text edits.

## Areas confirmed matching TS

- `Stamps.cs` comparison helpers (`lessThan`, `greaterThan`, `compare`, `spliceIntoList`) are very close to `stamps.ts` behavior; main deviation is C# mutability, not ordering math.
- Basic visible text insertion/removal/annotation paths cover the normal SharedString happy path; current 239/239 C# tests passing is consistent with this subset.
- Text segment length and append basics are structurally aligned for simple text-only segments.
- The C# port has explicit models for sided `SequencePlace`, group ops, obliterate, local references, intervals, snapshots, and partial lengths; the audit issues are mostly semantic completeness rather than total absence.
- Local/remote op application is split into client/core layers similar to TS, which should make future parity refactors feasible.
- The serializer preserves standard insert/remove/annotate/obliterate top-level intent for simple numeric-position ops.
- The interval subsystem has the same broad nouns as TS (collection, sequence interval, endpoint/overlap/id indexes), making targeted replacement with TS-parallel logic possible.

## Recommended test coverage additions

Recommended non-fuzz scenarios to port: **108**. Fuzz/farm tests should be noted as future stress coverage, but they are not required in this recommended-port set per request.

### From `merge-tree/src/test/*.spec.ts`

- `client.applyMsg.spec.ts` interleaved inserts/annotates/deletes — catches C1/C2
- `client.applyMsg.spec.ts` overlapping deletes — catches M1/P4/C1
- `client.applyMsg.spec.ts` overlapping insert and delete — catches C1/P2
- `client.applyMsg.spec.ts` intersecting insert after local delete — catches C3/P5
- `client.applyMsg.spec.ts` conflicting insert over local delete — catches C3/M13
- `client.applyMsg.spec.ts` concurrent insert into removed segment across block boundary — catches M13/P6
- `client.applyMsg.spec.ts` annotateRangeAdjust combine/min/max — catches C6/O3
- `client.applyMsg.spec.ts` minSeq with no ops and in-flight ops — catches C9/M7
- `client.localReference.spec.ts` non-sliding reference removed with segment — catches L1
- `client.localReference.spec.ts` sliding reference removed in middle — catches L1/L8
- `client.localReference.spec.ts` remove to end/from end with sliding refs — catches L8
- `client.localReference.spec.ts` remove+obliterate first-ack sliding — catches L1/L2
- `client.localReference.spec.ts` offsets on removed segments — catches L1/L3
- `client.localReference.spec.ts` transient refs on removed segments — catches L3
- `client.localReference.spec.ts` split/append segment with references — catches L1/M10
- `client.localReference.spec.ts` StayOnRemove concurrent delete matrix — catches L2
- `client.localReference.spec.ts` backward sliding preference — catches L4/L8
- `client.searchForMarker.spec.ts` marker at search position both directions — catches SS5/M8
- `client.searchForMarker.spec.ts` label-specific forward/backward search — catches SS5/O1
- `client.searchForMarker.spec.ts` multi-block marker search — catches M8
- `client.searchForMarker.spec.ts` distant forward/backward marker — catches M8
- `client.searchForMarker.spec.ts` rolled-back marker search cases — catches C10/M8
- `client.searchForMarker.spec.ts` removed/obliterated marker getMarkerById — catches SS5/O1
- `client.getPosition.spec.ts` existing/deleted/detached/removed segment — catches M9/L7
- `client.rebasePosition.spec.ts` rebase past remote insert/delete — catches C4/P3
- `client.rebasePosition.spec.ts` rebase mid local delete — catches C4/L1
- `client.rollback.spec.ts` rollback insert marker and annotate marker — catches C10/M5
- `client.rollback.spec.ts` rollback annotate over split string — catches M10/C10
- `client.rollback.spec.ts` rollback delete and restore local references — catches L1/C10
- `client.rollback.spec.ts` group op after re-submit split — catches C2/C4
- `mergeTree.annotate.spec.ts` local/remote interleaved annotation — catches M10/O6
- `mergeTree.annotate.spec.ts` local rewrite before remote — catches M14
- `mergeTree.markRangeRemoved.spec.ts` local/remote race to insert at removed segment — catches M13/P4
- `mergeTree.markRangeRemoved.spec.ts` local remove followed by remote overlapping remove — catches M1/P4
- `mergeTree.zamboni.spec.ts` one and many segments to scour — catches M6/M7
- `partialLength.spec.ts` single inserted element local/remote view — catches P2
- `partialLength.spec.ts` single removed segment local/remote view — catches P2
- `partialLength.spec.ts` aggregation permutations — catches P3
- `partialLength.spec.ts` concurrent overlapping deletes — catches P4
- `obliterate.spec.ts` concurrent obliterate and insert — catches M2/M3
- `obliterate.spec.ts` endpoint behavior at start/end — catches M3/O7
- `obliterate.partialLength.spec.ts` local remove after local obliterate — catches P4
- `obliterate.partialLength.spec.ts` overlapping remove+obliterate matrix — catches P4
- `obliterate.partialLength.spec.ts` overlapping obliterate+obliterate — catches P4
- `obliterate.partialLength.spec.ts` concurrent insert middle/start/end — catches P4/M3
- `obliterate.reconnect.spec.ts` obliterate does not expand during rebase — catches C12/M3
- `obliterate.reconnect.spec.ts` reconnected insert into obliterate range — catches C12/M2
- `obliterate.reconnect.spec.ts` separated group ops delete concurrent insert — catches C12/C2
- `obliterate.reconnect.spec.ts` sided obliterate reconnect — catches C12/O7
- `obliterate.concurrent.spec.ts` deletes concurrent insert before/after obliterate — catches M2/P4
- `obliterate.concurrent.spec.ts` partial lens consider overlapping obliterates — catches P4
- `obliterate.concurrent.spec.ts` clones removes array during insert — catches M1
- `obliterate.concurrent.spec.ts` keeps track of all obliterates on a segment — catches M1
- `obliterate.concurrent.spec.ts` many overlapping obliterates — catches M1/P4
- `obliterate.concurrent.spec.ts` multiple obliterates choose correct clientId/stamp — catches M1/P4
- `obliterate.concurrent.spec.ts` traversal past obliterated/non-obliterated tombstones — catches M13/P6
- `obliterate.rangeExpansion.spec.ts` removes prior insert from same client — catches M3/C12
- `obliterate.rangeExpansion.spec.ts` does not remove subsequent insert same client — catches M3/C12
- `obliterate.rangeExpansion.spec.ts` zero-length and sided obliterates — catches M3/O7
- `snapshot.spec.ts` excludes unacked/includes acked above MSN — catches S1/C9
- `snapshot.spec.ts` removals/obliterates above MSN — catches S1/P3
- `snapshot.spec.ts` insert relative to removed/obliterated loaded segment — catches S6/C5
- `snapshotlegacy.spec.ts` legacy summary load — catches S4
- `stamps.spec.ts` equality/comparison/spliceIntoList — catches O8/M1
- `properties.spec.ts` match undefined/empty/null/complex props — catches M14/O6
- `tracking.spec.ts` split/zamboni tracking groups (optional/deferred) — catches M12/L5

### From `sequence/src/test/*.spec.ts`

- `sharedString.spec.ts` local insert text with properties — catches SS1/O6
- `sharedString.spec.ts` local replace text and replace zero range — catches SS3/C7
- `sharedString.spec.ts` remove text local/remote — catches SS2/C1
- `sharedString.spec.ts` null and empty annotations — catches M14/O6
- `sharedString.spec.ts` annotate single multi-character segment — catches M10
- `sharedString.spec.ts` insert marker without markerId requirement — catches C8/O1
- `sharedString.spec.ts` annotate marker valid path — catches M5
- `sharedString.spec.ts` markerId update rejected for new/null/undefined — catches M5
- `sharedString.spec.ts` connected remote insert/replace/remove/annotate — catches C1
- `sharedString.spec.ts` reconnect resends unacked ops — catches C4
- `sharedString.spec.ts` multibyte/surrogate-pair insertion — catches SS3/C4
- `sharedString.obliterate.spec.ts` / Shared String Obliterate zero-length middle — catches M3
- `sequenceDeltaEvent.spec.ts` non-collab insert/remove/annotate ranges — catches C1/SS7
- `sequenceDeltaEvent.spec.ts` collab insert matrix same/overlapping positions — catches C1
- `sequenceDeltaEvent.spec.ts` collab delete matrix overlapping/shadowing — catches C1/P4
- `sequenceDeltaEvent.spec.ts` collab annotate same/different property matrix — catches C1/M14
- `sequenceDeltaEvent.spec.ts` insert/delete combination matrix — catches C1/C2
- `sequenceDeltaEvent.spec.ts` SequenceDeltaEventClass continuous/noncontinuous ranges — catches SS7
- `reentrancy.spec.ts` local reentrancy throws when enabled — catches SS8
- `reentrancy.spec.ts` consistency after reentrant mutation when guard disabled — catches SS8
- `marshalling.spec.ts` text segment to/from spec unannotated/annotated — catches S4/M10
- `marshalling.spec.ts` unrecognized JSON spec returns undefined — catches S8
- `intervalCollection.spec.ts` add/change/delete intervals locally and remotely — catches IC1/IC4
- `intervalCollection.spec.ts` interval property change/null/delete semantics — catches IC9/O6
- `intervalCollection.events.spec.ts` add/change/delete/property event payloads — catches IC6
- `intervalCollection.snapshot.spec.ts` summarize/load intervals with labels/properties — catches IC7
- `intervalCollection.detached.spec.ts` detached interval operations — catches IC4/IC7
- `intervalCollection.rollback.spec.ts` rollback add/change/delete interval — catches IC5
- `intervalRebasing.spec.ts` interval on concurrently removed text does not crash — catches IC4/L6
- `intervalRebasing.spec.ts` basic interval sliding for obliterate — catches IC10/L8
- `intervalRebasing.spec.ts` reference is -1 for obliterated segment — catches L7/IC10
- `intervalRebasing.spec.ts` slides two refs on same segment to different segments — catches L1/IC10
- `intervalRebasing.spec.ts` sliding preference after ack/reconnect — catches L6/IC3
- `intervalRebasing.spec.ts` changing endpoint to concurrently deleted segment detaches interval — catches IC10
- `intervalRebasing.spec.ts` delete events when interval slides off — catches IC10/IC6
- `intervalStashedOps.spec.ts` apply stashed add/delete/change/property ops (optional) — catches C11/IC5
- `intervalStashedOps.spec.ts` stickiness round-trip/gate tests — catches IC3
- `endpointInRangeIndex.spec.ts` exact boundary and duplicate endpoint cases — catches II2/II4
- `startpointInRangeIndex.spec.ts` invalid/negative/exact-boundary cases — catches II3
- `startpointInRangeIndex.spec.ts` deterministic large input vs brute force — catches II1/II3
- `endpointInRangeIndex.spec.ts` empty index and remove-missing interval cases — catches II2/II5
- `snapshotVersion.spec.ts` summary version compatibility — catches S4/IC7

### Fuzz/farm tests to flag but not require

- `client.conflictFarm.spec.ts`, `client.localReferenceFarm.spec.ts`, `client.reconnectFarm.spec.ts`, `client.rollbackFarm.spec.ts`, `client.obliterateFarm.spec.ts` — excellent later stress coverage, but excluded from the required port set.
- `obliterate.concurrent.spec.ts` fuzz-labeled regressions — keep their seeds/issues as notes, but first port deterministic neighboring cases.
- `interval` fuzz/rebasing stress tests — defer until deterministic interval parity is fixed.

## Recommended fix priority

1. [BUG-silent] Finding O1 — `ReferenceType` wire values corrupt marker/interval/reference semantics.
2. [BUG-silent] Finding C1 — Remote delta events report original positions instead of transformed local effects.
3. [BUG-silent] Finding M1 — Flattened remove/obliterate metadata loses multi-stamp visibility history.
4. [BUG-silent] Finding P4 — Partial lengths over overlapping remove/obliterate can be wrong.
5. [BUG-silent] Finding C4/C12 — Reconnect/regenerate paths are custom and high-risk for pending ops/obliterate.
6. [BUG-silent] Finding IC2/IC3 — Interval endpoints use wrong flags and omit stickiness/side semantics.
7. [BUG-crash] Finding C6/O3 — Annotate-adjust messages are unsupported.
8. [BUG-crash] Finding C8 — Valid TS marker inserts without markerId crash in C#.
9. [BUG-silent] Finding M10 — Text segment split may share property dictionaries across siblings.
10. [BUG-silent] Finding S1/S6 — Snapshots collapse metadata and cannot support relative ops to loaded tombstones.
11. [Deviation] Findings P1/O9/SS10 — after bugs, reduce structural drift by adding TS-parallel builders/docs and replacing bespoke algorithms.

## Files opened / compared for context

- `.claude/CLAUDE.md`
- `packages/dds/map/src/csharp-port/PARITY-AUDIT.md`
- `packages/dds/sequence/src/csharp-port/README.md`
- `packages/dds/sequence/src/sharedString.ts`
- `packages/dds/sequence/src/sequence.ts`
- `packages/dds/sequence/src/intervalCollection.ts`
- `packages/dds/sequence/src/intervals/sequenceInterval.ts`
- `packages/dds/sequence/src/intervals/intervalUtils.ts`
- `packages/dds/sequence/src/intervalIndex/*.ts`
- `packages/dds/merge-tree/src/mergeTree.ts`
- `packages/dds/merge-tree/src/client.ts`
- `packages/dds/merge-tree/src/partialLengths.ts`
- `packages/dds/merge-tree/src/localReference.ts`
- `packages/dds/merge-tree/src/snapshotLoader.ts`
- `packages/dds/merge-tree/src/snapshotV1.ts`
- `packages/dds/merge-tree/src/ops.ts`
- `packages/dds/merge-tree/src/opBuilder.ts`
- `packages/dds/merge-tree/src/stamps.ts`
- `packages/dds/merge-tree/src/sequencePlace.ts`
- `packages/dds/merge-tree/src/mergeTreeNodes.ts`
- `packages/dds/merge-tree/src/referencePositions.ts`
- `packages/dds/merge-tree/src/textSegment.ts`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedString.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedStringOpSerializer.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/SharedStringSnapshotLoader.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/MergeTree/*.cs`
- `packages/dds/sequence/src/csharp-port/SharedString/cs-out/Intervals/*.cs`
- `packages/dds/merge-tree/src/test/*.spec.ts` inventory
- `packages/dds/sequence/src/test/*.spec.ts` inventory


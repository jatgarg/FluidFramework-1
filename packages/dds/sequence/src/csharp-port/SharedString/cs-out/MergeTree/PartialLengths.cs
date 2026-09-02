// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/partialLengths.ts.
//
// Scope: per-block seq-indexed length cache for O(log N) position queries.
// Skipped: attribution, zamboni GC integration,
// legacy V0 snapshot format, verbose asserts / telemetry.
//
// Style chosen: B — compact rebuild-from-tree cache keeps the port correct and
// debuggable while preserving the important seq-indexed query shape.
// Cache invalidation: path-local incremental updates for edits, with full
// rebuild retained for snapshot load / structural recompute.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    /// <summary>
    /// Sequence-indexed cached length information for one <see cref="MergeBlock" /> subtree.
    /// </summary>
    /// <remarks>
    /// The cache is rebuilt bottom-up after structural or sequencing metadata changes. Queries use
    /// cumulative per-sequence deltas plus per-client deltas to reproduce merge-tree reference-sequence
    /// visibility without walking every descendant segment.
    /// </remarks>
    public sealed class PartialLengths
    {
        private readonly SortedDictionary<long, int> _deltasBySeq = new();
        private readonly Dictionary<int, SortedDictionary<long, int>> _clientDeltasBySeq = new();
        private readonly Dictionary<int, HashSet<ISegment>> _localUnackedSegmentsByClient = new();
        private readonly Dictionary<int, List<CumulativeLengthDelta>> _cumulativeClientDeltas = new();
        private List<CumulativeLengthDelta> _cumulativeDeltas = new();
        private bool _cumulativeLengthsDirty = true;

        /// <summary>
        /// Gets the current local-view length for the block represented by this cache.
        /// </summary>
        public int TotalCurrentLength { get; private set; }

        /// <summary>
        /// Gets the block length visible from a reference-sequence perspective.
        /// </summary>
        /// <param name="refSeq">The reference sequence number to view, or <see cref="Constants.UnassignedSequenceNumber" /> for the current local view.</param>
        /// <param name="clientId">Optional short client id for same-client and local-unacked visibility.</param>
        /// <param name="perspectiveSeq">Optional sequence currently being applied for same-client prior-operation visibility.</param>
        /// <returns>The visible length of this block from the supplied perspective.</returns>
        public int Query(long refSeq, int? clientId = null, long? perspectiveSeq = null)
        {
            if (refSeq < 0)
            {
                return clientId is int currentClientId
                    ? GetCurrentLengthForClient(currentClientId)
                    : TotalCurrentLength;
            }

            EnsureCumulativeLengths();
            int length = GetCumulativeLengthAtOrBefore(_cumulativeDeltas, refSeq);
            if (clientId is int perspectiveClientId)
            {
                if (perspectiveSeq is long seq
                    && _cumulativeClientDeltas.TryGetValue(perspectiveClientId, out List<CumulativeLengthDelta>? clientDeltas))
                {
                    length += GetCumulativeLengthBefore(clientDeltas, seq)
                        - GetCumulativeLengthAtOrBefore(clientDeltas, refSeq);
                }

                length += GetLocalUnackedAdjustment(perspectiveClientId, refSeq, perspectiveSeq);
            }

            return length;
        }

        /// <summary>
        /// Gets the visible length of one segment from a reference-sequence perspective.
        /// </summary>
        /// <param name="segment">The segment to test.</param>
        /// <param name="refSeq">The reference sequence number to view, or <see cref="Constants.UnassignedSequenceNumber" /> for the current local view.</param>
        /// <param name="clientId">Optional short client id for same-client and local-unacked visibility.</param>
        /// <param name="perspectiveSeq">Optional sequence currently being applied for same-client prior-operation visibility.</param>
        /// <returns>The segment length when visible; otherwise, 0.</returns>
        public static int GetSegmentVisibleLength(
            ISegment segment,
            long refSeq,
            int? clientId = null,
            long? perspectiveSeq = null)
        {
            if (segment is null)
            {
                throw new ArgumentNullException(nameof(segment));
            }

            return IsVisibleAt(segment, refSeq, clientId, perspectiveSeq, includeLocalUnacked: true)
                ? segment.CachedLength
                : 0;
        }

        internal void AddSegment(ISegment segment)
        {
            if (segment is null)
            {
                throw new ArgumentNullException(nameof(segment));
            }

            if (!IsHiddenInCurrentView(segment))
            {
                TotalCurrentLength += segment.CachedLength;
            }

            SegmentOperationStamp? firstAckedHidingStamp = GetFirstAckedHidingStamp(segment);
            if (firstAckedHidingStamp is SegmentOperationStamp hiddenOnInsertStamp
                && IsHiddenOnInsert(segment, hiddenOnInsertStamp))
            {
                AddHiddenOnInsertAdjustment(segment, hiddenOnInsertStamp);
                return;
            }

            if (segment.Seq == Constants.UnassignedSequenceNumber)
            {
                AddLocalUnackedSegment(segment.ClientId, segment);
            }
            else
            {
                AddSequencedDelta(segment.Seq, segment.ClientId, segment.CachedLength);
            }

            AddHidingDeltas(segment);
        }

        internal void AddChild(PartialLengths child)
        {
            if (child is null)
            {
                throw new ArgumentNullException(nameof(child));
            }

            TotalCurrentLength += child.TotalCurrentLength;
            foreach (KeyValuePair<long, int> delta in child._deltasBySeq)
            {
                AddDelta(_deltasBySeq, delta.Key, delta.Value);
            }

            foreach (KeyValuePair<int, SortedDictionary<long, int>> clientEntry in child._clientDeltasBySeq)
            {
                SortedDictionary<long, int> clientDeltas = GetClientDeltas(clientEntry.Key);
                foreach (KeyValuePair<long, int> delta in clientEntry.Value)
                {
                    AddDelta(clientDeltas, delta.Key, delta.Value);
                }
            }

            foreach (KeyValuePair<int, HashSet<ISegment>> localEntry in child._localUnackedSegmentsByClient)
            {
                foreach (ISegment segment in localEntry.Value)
                {
                    AddLocalUnackedSegment(localEntry.Key, segment);
                }
            }
        }

        internal void Seal()
        {
            _cumulativeLengthsDirty = true;
        }

        internal static int RebuildFromBlock(MergeBlock block, bool recurse, bool updateOrdinals)
        {
            if (block is null)
            {
                throw new ArgumentNullException(nameof(block));
            }

            int length = 0;
            PartialLengths partialLengths = new();
            for (int i = 0; i < block.ChildCount; i++)
            {
                IMergeNode? child = block.GetChild(i);
                if (child is null)
                {
                    continue;
                }

                child.Parent = block;
                child.Index = i;
                if (updateOrdinals)
                {
                    block.SetOrdinal(child, i);
                }

                if (child is MergeBlock childBlock)
                {
                    if (recurse)
                    {
                        RebuildFromBlock(childBlock, recurse: true, updateOrdinals: updateOrdinals);
                    }

                    length += childBlock.CachedLength;
                    partialLengths.AddChild(childBlock.PartialLengths);
                }
                else if (child is ISegment segment)
                {
                    length += GetSegmentVisibleLength(segment, Constants.UnassignedSequenceNumber);
                    partialLengths.AddSegment(segment);
                }
            }

            block.CachedLength = length;
            partialLengths.Seal();
            block.PartialLengths = partialLengths;
            return length;
        }

        internal static int UpdateStructureFor(MergeBlock block)
        {
            return RebuildFromBlock(block, recurse: false, updateOrdinals: false);
        }

        internal static int UpdateLengthFor(
            MergeBlock block,
            OperationStamp stamp,
            IReadOnlyCollection<ISegment>? changedSegments = null)
        {
            if (block is null)
            {
                throw new ArgumentNullException(nameof(block));
            }

            if (stamp is null)
            {
                throw new ArgumentNullException(nameof(stamp));
            }

            PartialLengths partialLengths = block.PartialLengths;
            int currentLength = 0;
            int seqDelta = 0;
            Dictionary<int, int> clientDeltas = new();
            bool updateSeqDeltas = stamp.Seq != Constants.UnassignedSequenceNumber;

            for (int i = 0; i < block.ChildCount; i++)
            {
                IMergeNode? child = block.GetChild(i);
                if (child is null)
                {
                    continue;
                }

                if (child is MergeBlock childBlock)
                {
                    currentLength += childBlock.CachedLength;
                    if (updateSeqDeltas)
                    {
                        seqDelta += childBlock.PartialLengths.GetDelta(stamp.Seq);
                        foreach (KeyValuePair<int, int> childClientDelta in childBlock.PartialLengths.GetClientDeltas(stamp.Seq))
                        {
                            AddDelta(clientDeltas, childClientDelta.Key, childClientDelta.Value);
                        }
                    }
                }
                else if (child is ISegment segment)
                {
                    currentLength += GetSegmentVisibleLength(segment, Constants.UnassignedSequenceNumber);
                    if (updateSeqDeltas)
                    {
                        AddSegmentContributionsForSeq(segment, stamp.Seq, ref seqDelta, clientDeltas);
                    }
                }
            }

            block.CachedLength = currentLength;
            partialLengths.TotalCurrentLength = currentLength;
            if (updateSeqDeltas)
            {
                partialLengths.SetDelta(stamp.Seq, seqDelta);
                partialLengths.SetClientDeltas(stamp.Seq, clientDeltas);
            }

            partialLengths.SyncLocalUnackedSegments(block, changedSegments);
            partialLengths.Seal();
            return currentLength;
        }

        internal static void ValidateBlockPartialLengthInvariants(MergeBlock block)
        {
            if (block is null)
            {
                throw new ArgumentNullException(nameof(block));
            }

            PartialLengths expected = new();
            int expectedCurrentLength = 0;
            for (int i = 0; i < block.ChildCount; i++)
            {
                IMergeNode? child = block.GetChild(i);
                if (child is null)
                {
                    continue;
                }

                if (child is MergeBlock childBlock)
                {
                    if (childBlock.CachedLength != childBlock.PartialLengths.TotalCurrentLength)
                    {
                        throw new LoggingError("Child block cached length does not match partial lengths.");
                    }

                    expectedCurrentLength += childBlock.CachedLength;
                    expected.AddChild(childBlock.PartialLengths);
                }
                else if (child is ISegment segment)
                {
                    int segmentLength = GetSegmentVisibleLength(segment, Constants.UnassignedSequenceNumber);
                    expectedCurrentLength += segmentLength;
                    expected.AddSegment(segment);
                }
            }

            expected.Seal();
            if (block.CachedLength != expectedCurrentLength
                || block.PartialLengths.TotalCurrentLength != expectedCurrentLength)
            {
                throw new LoggingError("MergeBlock cached length does not match child partial lengths.");
            }

            if (!AreDeltasEqual(block.PartialLengths._deltasBySeq, expected._deltasBySeq)
                || !AreClientDeltasEqual(block.PartialLengths._clientDeltasBySeq, expected._clientDeltasBySeq)
                || !AreLocalUnackedSegmentsEqual(
                    block.PartialLengths._localUnackedSegmentsByClient,
                    expected._localUnackedSegmentsByClient))
            {
                throw new LoggingError("Partial-length aggregates do not match child contributions.");
            }
        }

        internal int GetDelta(long seq)
        {
            return _deltasBySeq.TryGetValue(seq, out int delta) ? delta : 0;
        }

        internal IEnumerable<KeyValuePair<int, int>> GetClientDeltas(long seq)
        {
            foreach (KeyValuePair<int, SortedDictionary<long, int>> clientEntry in _clientDeltasBySeq)
            {
                if (clientEntry.Value.TryGetValue(seq, out int delta))
                {
                    yield return new KeyValuePair<int, int>(clientEntry.Key, delta);
                }
            }
        }

        private static void AddSegmentContributionsForSeq(
            ISegment segment,
            long seq,
            ref int seqDelta,
            Dictionary<int, int> clientDeltas)
        {
            SegmentOperationStamp? firstAckedHidingStamp = GetFirstAckedHidingStamp(segment);
            if (firstAckedHidingStamp is SegmentOperationStamp hiddenOnInsertStamp
                && IsHiddenOnInsert(segment, hiddenOnInsertStamp))
            {
                if (hiddenOnInsertStamp.Seq == seq && !WasRemovedByClient(segment, segment.ClientId))
                {
                    AddDelta(clientDeltas, segment.ClientId, segment.CachedLength);
                }

                return;
            }

            if (segment.Seq == seq)
            {
                seqDelta += segment.CachedLength;
                AddDelta(clientDeltas, segment.ClientId, segment.CachedLength);
            }

            if (firstAckedHidingStamp is SegmentOperationStamp hidingStamp
                && segment.Seq != Constants.UnassignedSequenceNumber
                && hidingStamp.Seq == seq)
            {
                seqDelta -= segment.CachedLength;
                foreach (int clientId in GetRemoveClientIds(segment))
                {
                    AddDelta(clientDeltas, clientId, -segment.CachedLength);
                }
            }
        }

        private void SetDelta(long seq, int delta)
        {
            SetDelta(_deltasBySeq, seq, delta);
        }

        private static bool AreDeltasEqual(
            IReadOnlyDictionary<long, int> actual,
            IReadOnlyDictionary<long, int> expected)
        {
            if (actual.Count != expected.Count)
            {
                return false;
            }

            foreach (KeyValuePair<long, int> expectedDelta in expected)
            {
                if (!actual.TryGetValue(expectedDelta.Key, out int actualDelta)
                    || actualDelta != expectedDelta.Value)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool AreClientDeltasEqual(
            IReadOnlyDictionary<int, SortedDictionary<long, int>> actual,
            IReadOnlyDictionary<int, SortedDictionary<long, int>> expected)
        {
            if (actual.Count != expected.Count)
            {
                return false;
            }

            foreach (KeyValuePair<int, SortedDictionary<long, int>> expectedEntry in expected)
            {
                if (!actual.TryGetValue(expectedEntry.Key, out SortedDictionary<long, int>? actualDeltas)
                    || !AreDeltasEqual(actualDeltas, expectedEntry.Value))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool AreLocalUnackedSegmentsEqual(
            IReadOnlyDictionary<int, HashSet<ISegment>> actual,
            IReadOnlyDictionary<int, HashSet<ISegment>> expected)
        {
            if (actual.Count != expected.Count)
            {
                return false;
            }

            foreach (KeyValuePair<int, HashSet<ISegment>> expectedEntry in expected)
            {
                if (!actual.TryGetValue(expectedEntry.Key, out HashSet<ISegment>? actualSegments)
                    || actualSegments.Count != expectedEntry.Value.Count)
                {
                    return false;
                }

                foreach (ISegment expectedSegment in expectedEntry.Value)
                {
                    if (!actualSegments.Contains(expectedSegment))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private void SetClientDeltas(long seq, IReadOnlyDictionary<int, int> clientDeltas)
        {
            List<int> emptyClients = new();
            foreach (KeyValuePair<int, SortedDictionary<long, int>> clientEntry in _clientDeltasBySeq)
            {
                clientEntry.Value.Remove(seq);
                if (clientEntry.Value.Count == 0)
                {
                    emptyClients.Add(clientEntry.Key);
                }
            }

            foreach (int clientId in emptyClients)
            {
                _clientDeltasBySeq.Remove(clientId);
            }

            foreach (KeyValuePair<int, int> clientDelta in clientDeltas)
            {
                if (clientDelta.Value != 0)
                {
                    GetClientDeltas(clientDelta.Key)[seq] = clientDelta.Value;
                }
            }
        }

        private void SyncLocalUnackedSegments(MergeBlock block, IReadOnlyCollection<ISegment>? changedSegments)
        {
            if (changedSegments is null || changedSegments.Count == 0)
            {
                return;
            }

            foreach (ISegment segment in changedSegments)
            {
                if (IsDescendantOf(segment, block))
                {
                    RemoveLocalUnackedSegment(segment);
                    AddCurrentLocalUnackedSegment(segment.Seq, segment.ClientId, segment);
                    foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
                    {
                        AddCurrentLocalUnackedSegment(stamp.Seq, stamp.ClientId, segment);
                    }
                }
            }
        }

        private static bool IsDescendantOf(IMergeNode node, MergeBlock block)
        {
            MergeBlock? current = node.Parent;
            while (current is not null)
            {
                if (ReferenceEquals(current, block))
                {
                    return true;
                }

                current = current.Parent;
            }

            return false;
        }

        private void AddCurrentLocalUnackedSegment(long? seq, int clientId, ISegment segment)
        {
            if (seq == Constants.UnassignedSequenceNumber)
            {
                AddLocalUnackedSegment(clientId, segment);
            }
        }

        private void RemoveLocalUnackedSegment(ISegment segment)
        {
            List<int> emptyClients = new();
            foreach (KeyValuePair<int, HashSet<ISegment>> localEntry in _localUnackedSegmentsByClient)
            {
                localEntry.Value.Remove(segment);
                if (localEntry.Value.Count == 0)
                {
                    emptyClients.Add(localEntry.Key);
                }
            }

            foreach (int clientId in emptyClients)
            {
                _localUnackedSegmentsByClient.Remove(clientId);
            }
        }

        private void EnsureCumulativeLengths()
        {
            if (!_cumulativeLengthsDirty)
            {
                return;
            }

            _cumulativeDeltas = BuildCumulativeList(_deltasBySeq);
            _cumulativeClientDeltas.Clear();
            foreach (KeyValuePair<int, SortedDictionary<long, int>> clientEntry in _clientDeltasBySeq)
            {
                List<CumulativeLengthDelta> cumulative = BuildCumulativeList(clientEntry.Value);
                if (cumulative.Count > 0)
                {
                    _cumulativeClientDeltas[clientEntry.Key] = cumulative;
                }
            }

            _cumulativeLengthsDirty = false;
        }

        private static bool IsVisibleAt(
            ISegment segment,
            long refSeq,
            int? perspectiveClientId,
            long? perspectiveSeq,
            bool includeLocalUnacked)
        {
            if (refSeq < 0)
            {
                return perspectiveClientId is int currentClientId
                    ? IsVisibleInCurrentView(segment, currentClientId)
                    : !IsHiddenInCurrentView(segment);
            }

            return IsInsertionVisibleAt(segment, refSeq, perspectiveClientId, perspectiveSeq, includeLocalUnacked)
                && !IsRemovalVisibleAt(segment, refSeq, perspectiveClientId, perspectiveSeq, includeLocalUnacked);
        }

        private static bool IsInsertionVisibleAt(
            ISegment segment,
            long refSeq,
            int? perspectiveClientId,
            long? perspectiveSeq,
            bool includeLocalUnacked)
        {
            if (segment.Seq == Constants.UnassignedSequenceNumber)
            {
                return includeLocalUnacked
                    && perspectiveClientId is int clientId
                    && segment.ClientId == clientId;
            }

            return segment.Seq <= refSeq
                || IsSameClientPriorOperation(segment.Seq, segment.ClientId, perspectiveClientId, perspectiveSeq);
        }

        private static bool IsRemovalVisibleAt(
            ISegment segment,
            long refSeq,
            int? perspectiveClientId,
            long? perspectiveSeq,
            bool includeLocalUnacked)
        {
            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (stamp.Seq == Constants.UnassignedSequenceNumber)
                {
                    if (includeLocalUnacked
                        && perspectiveClientId is int clientId
                        && stamp.ClientId == clientId)
                    {
                        return true;
                    }
                }
                else if (stamp.Seq <= refSeq
                    || IsSameClientPriorOperation(stamp.Seq, stamp.ClientId, perspectiveClientId, perspectiveSeq))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsSameClientPriorOperation(
            long operationSeq,
            int operationClientId,
            int? perspectiveClientId,
            long? perspectiveSeq)
        {
            return perspectiveClientId is int clientId
                && perspectiveSeq is long seq
                && operationClientId == clientId
                && operationSeq < seq;
        }

        private int GetLocalUnackedAdjustment(int clientId, long refSeq, long? perspectiveSeq)
        {
            if (!_localUnackedSegmentsByClient.TryGetValue(clientId, out HashSet<ISegment>? segments))
            {
                return 0;
            }

            int adjustment = 0;
            foreach (ISegment segment in segments)
            {
                int exactLength = IsVisibleAt(segment, refSeq, clientId, perspectiveSeq, includeLocalUnacked: true)
                    ? segment.CachedLength
                    : 0;
                int ackedOnlyLength = IsVisibleAt(segment, refSeq, clientId, perspectiveSeq, includeLocalUnacked: false)
                    ? segment.CachedLength
                    : 0;
                adjustment += exactLength - ackedOnlyLength;
            }

            return adjustment;
        }

        private int GetCurrentLengthForClient(int clientId)
        {
            int length = TotalCurrentLength;
            HashSet<ISegment> adjustedSegments = new();
            foreach (HashSet<ISegment> segments in _localUnackedSegmentsByClient.Values)
            {
                foreach (ISegment segment in segments)
                {
                    if (!adjustedSegments.Add(segment))
                    {
                        continue;
                    }

                    int currentContribution = IsHiddenInCurrentView(segment) ? 0 : segment.CachedLength;
                    int clientContribution = IsVisibleInCurrentView(segment, clientId) ? segment.CachedLength : 0;
                    length += clientContribution - currentContribution;
                }
            }

            return length;
        }

        private static bool IsVisibleInCurrentView(ISegment segment, int clientId)
        {
            bool inserted = segment.Seq != Constants.UnassignedSequenceNumber
                || segment.ClientId == clientId;
            if (!inserted)
            {
                return false;
            }

            return !IsHiddenInCurrentView(segment, clientId);
        }

        private static bool IsHiddenInCurrentView(ISegment segment, int? clientId = null)
        {
            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (IsCurrentViewHiddenByStamp(stamp, clientId))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsCurrentViewHiddenByStamp(RemoveOperationStamp stamp, int? perspectiveClientId)
        {
            return stamp.Seq != Constants.UnassignedSequenceNumber
                    || perspectiveClientId is null
                    || stamp.ClientId == perspectiveClientId;
        }

        private void AddHidingDeltas(ISegment segment)
        {
            SegmentOperationStamp? firstAckedHidingStamp = GetFirstAckedHidingStamp(segment);
            foreach (RemoveOperationStamp removeStamp in segment.RemoveStamps)
            {
                AddLocalHidingStamp(removeStamp);
            }

            if (firstAckedHidingStamp is SegmentOperationStamp stamp
                && segment.Seq != Constants.UnassignedSequenceNumber)
            {
                AddSequencedDelta(stamp.Seq, stamp.ClientId, -segment.CachedLength);
                foreach (int clientId in GetRemoveClientIds(segment))
                {
                    if (clientId == stamp.ClientId)
                    {
                        continue;
                    }

                    AddDelta(GetClientDeltas(clientId), stamp.Seq, -segment.CachedLength);
                    if (segment.Seq != Constants.UnassignedSequenceNumber && clientId != segment.ClientId)
                    {
                        AddDelta(GetClientDeltas(clientId), segment.Seq, segment.CachedLength);
                    }
                }
            }

            void AddLocalHidingStamp(RemoveOperationStamp removeStamp)
            {
                if (removeStamp.Seq == Constants.UnassignedSequenceNumber)
                {
                    AddLocalUnackedSegment(removeStamp.ClientId, segment);
                }
            }
        }

        private static SegmentOperationStamp? GetFirstAckedHidingStamp(ISegment segment)
        {
            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (stamp.Seq != Constants.UnassignedSequenceNumber)
                {
                    return new SegmentOperationStamp(stamp.Seq, stamp.ClientId);
                }
            }

            return null;
        }

        private static bool IsHiddenOnInsert(ISegment segment, SegmentOperationStamp firstAckedHidingStamp)
        {
            return segment.Seq != Constants.UnassignedSequenceNumber
                && CompareSegmentStamp(firstAckedHidingStamp, new SegmentOperationStamp(segment.Seq, segment.ClientId)) < 0;
        }

        private void AddHiddenOnInsertAdjustment(ISegment segment, SegmentOperationStamp firstAckedHidingStamp)
        {
            if (!WasRemovedByClient(segment, segment.ClientId))
            {
                AddDelta(GetClientDeltas(segment.ClientId), firstAckedHidingStamp.Seq, segment.CachedLength);
            }
        }

        private static IEnumerable<int> GetRemoveClientIds(ISegment segment)
        {
            HashSet<int> clientIds = new();
            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (clientIds.Add(stamp.ClientId))
                {
                    yield return stamp.ClientId;
                }
            }
        }

        private static bool WasRemovedByClient(ISegment segment, int clientId)
        {
            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (stamp.ClientId == clientId)
                {
                    return true;
                }
            }

            return false;
        }

        private static int CompareSegmentStamp(SegmentOperationStamp a, SegmentOperationStamp b)
        {
            int seqComparison = a.Seq.CompareTo(b.Seq);
            return seqComparison != 0 ? seqComparison : a.ClientId.CompareTo(b.ClientId);
        }

        private void AddSequencedDelta(long seq, int clientId, int delta)
        {
            AddDelta(_deltasBySeq, seq, delta);
            AddDelta(GetClientDeltas(clientId), seq, delta);
        }

        private SortedDictionary<long, int> GetClientDeltas(int clientId)
        {
            if (!_clientDeltasBySeq.TryGetValue(clientId, out SortedDictionary<long, int>? deltas))
            {
                deltas = new SortedDictionary<long, int>();
                _clientDeltasBySeq[clientId] = deltas;
            }

            return deltas;
        }

        private void AddLocalUnackedSegment(int clientId, ISegment segment)
        {
            if (!_localUnackedSegmentsByClient.TryGetValue(clientId, out HashSet<ISegment>? segments))
            {
                segments = new HashSet<ISegment>();
                _localUnackedSegmentsByClient[clientId] = segments;
            }

            segments.Add(segment);
        }

        private static void AddDelta(SortedDictionary<long, int> deltas, long seq, int delta)
        {
            if (delta == 0)
            {
                return;
            }

            int newDelta = deltas.TryGetValue(seq, out int existingDelta)
                ? existingDelta + delta
                : delta;
            if (newDelta == 0)
            {
                deltas.Remove(seq);
            }
            else
            {
                deltas[seq] = newDelta;
            }
        }

        private static void SetDelta(SortedDictionary<long, int> deltas, long seq, int delta)
        {
            if (delta == 0)
            {
                deltas.Remove(seq);
            }
            else
            {
                deltas[seq] = delta;
            }
        }

        private static void AddDelta(Dictionary<int, int> deltas, int clientId, int delta)
        {
            if (delta == 0)
            {
                return;
            }

            int newDelta = deltas.TryGetValue(clientId, out int existingDelta)
                ? existingDelta + delta
                : delta;
            if (newDelta == 0)
            {
                deltas.Remove(clientId);
            }
            else
            {
                deltas[clientId] = newDelta;
            }
        }

        private static List<CumulativeLengthDelta> BuildCumulativeList(SortedDictionary<long, int> deltas)
        {
            List<CumulativeLengthDelta> cumulative = new();
            int length = 0;
            foreach (KeyValuePair<long, int> delta in deltas)
            {
                length += delta.Value;
                cumulative.Add(new CumulativeLengthDelta(delta.Key, length));
            }

            return cumulative;
        }

        private static int GetCumulativeLengthAtOrBefore(IReadOnlyList<CumulativeLengthDelta> deltas, long seq)
        {
            int index = LatestIndex(deltas, seq, inclusive: true);
            return index >= 0 ? deltas[index].Length : 0;
        }

        private static int GetCumulativeLengthBefore(IReadOnlyList<CumulativeLengthDelta> deltas, long seq)
        {
            int index = LatestIndex(deltas, seq, inclusive: false);
            return index >= 0 ? deltas[index].Length : 0;
        }

        private static int LatestIndex(IReadOnlyList<CumulativeLengthDelta> deltas, long seq, bool inclusive)
        {
            int low = 0;
            int high = deltas.Count - 1;
            int result = -1;
            while (low <= high)
            {
                int mid = low + ((high - low) / 2);
                bool isMatch = inclusive ? deltas[mid].Seq <= seq : deltas[mid].Seq < seq;
                if (isMatch)
                {
                    result = mid;
                    low = mid + 1;
                }
                else
                {
                    high = mid - 1;
                }
            }

            return result;
        }

        private readonly struct CumulativeLengthDelta
        {
            public CumulativeLengthDelta(long seq, int length)
            {
                Seq = seq;
                Length = length;
            }

            public long Seq { get; }

            public int Length { get; }
        }

        private readonly struct SegmentOperationStamp
        {
            public SegmentOperationStamp(long seq, int clientId)
            {
                Seq = seq;
                ClientId = clientId;
            }

            public long Seq { get; }

            public int ClientId { get; }
        }
    }
}

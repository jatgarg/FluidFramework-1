// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/mergeTree.ts.
//
// Scope: insert / mark-removed / basic obliterate / walk / get-text / get-length
// / position-query. Position queries use the partialLengths cache for
// block-aware descent.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    /// <summary>
    /// Merge-tree storage for SharedString text segments.
    /// </summary>
    public sealed class MergeTree
    {
        /// <summary>
        /// The sequence number of an op before it is acked.
        /// </summary>
        public const long UnassignedSequenceNumber = Constants.UnassignedSequenceNumber;

        /// <summary>
        /// The sequence number which can be seen by all ops.
        /// </summary>
        public const long UniversalSequenceNumber = Constants.UniversalSequenceNumber;

        /// <summary>
        /// Sentinel returned for detached reference positions.
        /// </summary>
        public const int DetachedReferencePosition = -1;

        /// <summary>
        /// The lower bound for the collaboration window's minimum sequence number.
        /// </summary>
        public const long CollaborationWindowMinSequenceNumber = 0;

        private readonly Dictionary<string, int> _clientIds = new(StringComparer.Ordinal);
        private int _nextClientId = 1;
        private const int MaxChildrenBeforeSplit = MergeBlock.MaxChildren - 1;
        private const int DefaultZamboniSegmentsMaxCount = 250;

        /// <summary>
        /// Initializes a new instance of the <see cref="MergeTree" /> class.
        /// </summary>
        public MergeTree()
        {
            Root = MakeBlock();
            CurrentSeq = UniversalSequenceNumber;
            MinSeq = CollaborationWindowMinSequenceNumber;
        }

        /// <summary>
        /// Gets the root block of the merge tree.
        /// </summary>
        public MergeBlock Root { get; private set; }

        /// <summary>
        /// Gets or sets the highest sequence number seen from the server.
        /// </summary>
        public long CurrentSeq { get; set; }

        /// <summary>
        /// Gets or sets the minimum sequence number.
        /// </summary>
        /// <remarks>
        /// Segments removed or obliterated before this value are zamboni candidates.
        /// </remarks>
        public long MinSeq { get; set; }

        /// <summary>
        /// Inserts one or more segments at a position.
        /// </summary>
        /// <param name="position">The position at which to insert, computed in the <paramref name="refSeq" /> view.</param>
        /// <param name="segments">The segments to insert.</param>
        /// <param name="refSeq">The sequence number whose view was used to compute <paramref name="position" />.</param>
        /// <param name="seq">The sequence number assigned to this insert, or <see cref="UnassignedSequenceNumber" /> for local unacked inserts.</param>
        /// <param name="clientId">The client that submitted the insert.</param>
        public List<ISegment> InsertSegments(int position, ISegment[] segments, long refSeq, long seq, string clientId, long? perspectiveSeq = null)
        {
            if (segments is null)
            {
                throw new ArgumentNullException(nameof(segments));
            }

            return InsertSegments(position, segments, refSeq, seq, GetClientId(clientId), perspectiveSeq);
        }

        /// <summary>
        /// Marks a range as removed.
        /// </summary>
        /// <param name="start">The inclusive start position, computed in the <paramref name="refSeq" /> view.</param>
        /// <param name="end">The exclusive end position, computed in the <paramref name="refSeq" /> view.</param>
        /// <param name="refSeq">The sequence number whose view was used to compute the range.</param>
        /// <param name="seq">The sequence number assigned to this remove, or <see cref="UnassignedSequenceNumber" /> for local unacked removes.</param>
        /// <param name="clientId">The client that submitted the remove.</param>
        public List<ISegment> MarkRangeRemoved(int start, int end, long refSeq, long seq, string clientId, long? perspectiveSeq = null)
        {
            return MarkRangeRemoved(start, end, refSeq, seq, GetClientId(clientId), perspectiveSeq);
        }

        /// <summary>
        /// Marks a range as obliterated.
        /// </summary>
        /// <param name="start">The inclusive start position, computed in the <paramref name="refSeq" /> view.</param>
        /// <param name="end">The exclusive end position, computed in the <paramref name="refSeq" /> view.</param>
        /// <param name="refSeq">The sequence number whose view was used to compute the range.</param>
        /// <param name="seq">The sequence number assigned to this obliterate, or <see cref="UnassignedSequenceNumber" /> for local unacked obliterates.</param>
        /// <param name="clientId">The client that submitted the obliterate.</param>
        public List<ISegment> ObliterateRange(int start, int end, long refSeq, long seq, string clientId, long? perspectiveSeq = null)
        {
            return ObliterateRange(start, end, refSeq, seq, GetClientId(clientId), perspectiveSeq);
        }

        /// <summary>
        /// Marks a sided range as obliterated.
        /// </summary>
        /// <param name="start">The start place, computed in the <paramref name="refSeq" /> view.</param>
        /// <param name="end">The end place, computed in the <paramref name="refSeq" /> view.</param>
        /// <param name="refSeq">The sequence number whose view was used to compute the range.</param>
        /// <param name="seq">The sequence number assigned to this obliterate, or <see cref="UnassignedSequenceNumber" /> for local unacked obliterates.</param>
        /// <param name="clientId">The client that submitted the obliterate.</param>
        public List<ISegment> ObliterateRangeSided(SequencePlace start, SequencePlace end, long refSeq, long seq, string clientId, long? perspectiveSeq = null)
        {
            return ObliterateRangeSided(start, end, refSeq, seq, GetClientId(clientId), perspectiveSeq);
        }

        /// <summary>
        /// Applies properties to all segments in the range [start, end).
        /// Splits segments at boundaries as needed so that the annotation applies to
        /// exactly the requested range.
        /// </summary>
        /// <param name="start">Inclusive start position.</param>
        /// <param name="end">Exclusive end position.</param>
        /// <param name="props">Properties to merge. A null value on a key deletes that property.</param>
        /// <param name="refSeq">Caller's view of sequence numbers at annotation time.</param>
        /// <param name="seq">Server-assigned sequence number, or UnassignedSequenceNumber for unacked local ops.</param>
        /// <param name="clientId">Client applying the annotation.</param>
        public List<ISegment> AnnotateRange(int start, int end, PropertySet props, long refSeq, long seq, string clientId, long? perspectiveSeq = null)
        {
            if (props is null)
            {
                throw new ArgumentNullException(nameof(props));
            }

            return AnnotateRange(start, end, props, refSeq, seq, GetClientId(clientId), perspectiveSeq);
        }

        /// <summary>
        /// Gets the total length of the merge tree in the current local view.
        /// </summary>
        /// <returns>The sum of visible segment lengths.</returns>
        public int GetLength()
        {
            return GetLength(UnassignedSequenceNumber);
        }

        /// <summary>
        /// Gets the total length of the merge tree as visible at a reference sequence number.
        /// </summary>
        /// <param name="refSeq">The reference sequence number to view.</param>
        /// <param name="clientId">The client perspective for same-client and local-unacked visibility.</param>
        /// <returns>The sum of segment lengths visible at <paramref name="refSeq" />.</returns>
        public int GetLength(long refSeq, string? clientId = null)
        {
            int? perspectiveClientId = clientId is null ? null : GetClientId(clientId);
            return Root.PartialLengths.Query(refSeq, perspectiveClientId, CurrentSeq + 1);
        }

        internal int GetLengthAt(long refSeq, string? clientId, long? localSeq)
        {
            int? perspectiveClientId = clientId is null ? null : GetClientId(clientId);
            return GetNodeLengthAt(Root, refSeq, perspectiveClientId, localSeq);
        }

        /// <summary>
        /// Extracts text from the current local view.
        /// </summary>
        /// <returns>The concatenated text of all visible <see cref="TextSegment" /> leaves.</returns>
        public string GetText()
        {
            return GetText(0, GetLength());
        }

        /// <summary>
        /// Extracts text from a range in the current local view.
        /// </summary>
        /// <param name="start">The inclusive start position.</param>
        /// <param name="end">The exclusive end position.</param>
        /// <returns>The concatenated text in the requested range.</returns>
        public string GetText(int start, int end)
        {
            ValidateRange(start, end, GetLength());

            StringBuilder builder = new();
            foreach ((ISegment segment, int startOffset, int endOffset) in GetSegments(start, end))
            {
                if (segment is TextSegment textSegment)
                {
                    builder.Append(textSegment.Text, startOffset, endOffset - startOffset);
                }
                else
                {
                    // TODO: represent Marker/non-TextSegment leaves with placeholders.
                }
            }

            return builder.ToString();
        }

        /// <summary>
        /// Enumerates visible segment slices in a range in the current local view.
        /// </summary>
        /// <param name="start">The inclusive start position.</param>
        /// <param name="end">The exclusive end position.</param>
        /// <returns>Live segments and offsets covering the requested range.</returns>
        public IEnumerable<(ISegment segment, int startOffset, int endOffset)> GetSegments(int start, int end)
        {
            ValidateRange(start, end, GetLength());

            int position = 0;
            foreach (ISegment segment in WalkAllSegments())
            {
                int length = VisibleLength(segment, UnassignedSequenceNumber);
                if (length == 0)
                {
                    continue;
                }

                int nextPosition = position + length;
                if (nextPosition <= start)
                {
                    position = nextPosition;
                    continue;
                }

                if (position >= end)
                {
                    yield break;
                }

                int startOffset = Math.Max(0, start - position);
                int endOffset = Math.Min(length, end - position);
                yield return (segment, startOffset, endOffset);
                position = nextPosition;
            }
        }

        /// <summary>
        /// Creates a <see cref="LocalReferencePosition" /> at the given current local-view position.
        /// </summary>
        /// <param name="position">The absolute position to attach the reference to.</param>
        /// <param name="refType">The reference type flags.</param>
        /// <param name="slidingPreference">The preferred slide direction for this reference.</param>
        /// <param name="properties">Optional reference properties.</param>
        /// <returns>The created reference.</returns>
        public LocalReferencePosition CreateReferencePosition(
            int position,
            ReferenceType refType,
            SlidingPreference slidingPreference = SlidingPreference.Forward,
            PropertySet? properties = null,
            bool canSlideToEndpoint = false)
        {
            (ISegment segment, int offset) = GetSegmentForReferencePosition(position);
            return CreateReferencePosition(segment, offset, refType, slidingPreference, properties, canSlideToEndpoint);
        }

        /// <summary>
        /// Creates a <see cref="LocalReferencePosition" /> on the supplied segment and offset.
        /// </summary>
        /// <param name="segment">The segment to attach the reference to.</param>
        /// <param name="offset">The offset within the segment.</param>
        /// <param name="refType">The reference type flags.</param>
        /// <param name="slidingPreference">The preferred slide direction for this reference.</param>
        /// <param name="properties">Optional reference properties.</param>
        /// <param name="canSlideToEndpoint">Whether the reference may slide to a start/end endpoint sentinel.</param>
        /// <returns>The created reference.</returns>
        public LocalReferencePosition CreateReferencePosition(
            ISegment segment,
            int offset,
            ReferenceType refType,
            SlidingPreference slidingPreference = SlidingPreference.Forward,
            PropertySet? properties = null,
            bool canSlideToEndpoint = false)
        {
            ArgumentNullException.ThrowIfNull(segment);
            if (offset < 0 || offset > segment.CachedLength)
            {
                throw new ArgumentOutOfRangeException(nameof(offset), "Offset must be within the segment.");
            }

            if (VisibleLength(segment, UnassignedSequenceNumber) == 0
                && !RefTypeIncludesFlag(
                    refType,
                    ReferenceType.SlideOnRemove | ReferenceType.Transient | ReferenceType.StayOnRemove))
            {
                throw new ArgumentException(
                    "Can only create SlideOnRemove, Transient, or StayOnRemove local reference positions on removed segments.",
                    nameof(refType));
            }

            LocalReferencePosition reference = new(segment, offset, refType, slidingPreference, properties, canSlideToEndpoint);
            segment.LocalRefs.Add(reference);
            return reference;
        }

        /// <summary>
        /// <summary>
        /// Creates a <see cref="LocalReferencePosition" /> anchored to the synthetic
        /// start-of-tree or end-of-tree endpoint. Matches TS's `startOfTree`/`endOfTree`
        /// sentinel segments used when the wire encodes an interval endpoint as
        /// `"start"` or `"end"`.
        /// </summary>
        internal LocalReferencePosition CreateReferencePositionAtEndpoint(
            ReferenceEndpointKind endpoint,
            ReferenceType refType,
            SlidingPreference slidingPreference = SlidingPreference.Forward,
            PropertySet? properties = null)
        {
            if (endpoint == ReferenceEndpointKind.None)
            {
                throw new ArgumentException(
                    "Endpoint must be Start or End.",
                    nameof(endpoint));
            }

            return new LocalReferencePosition(endpoint, refType, slidingPreference, properties, canSlideToEndpoint: true);
        }

        /// <summary>
        /// Removes a <see cref="LocalReferencePosition" /> from its segment.
        /// </summary>
        /// <param name="reference">The reference to remove.</param>
        /// <returns><see langword="true" /> if the reference was attached and present on its segment.</returns>
        public bool RemoveReferencePosition(LocalReferencePosition reference)
        {
            ArgumentNullException.ThrowIfNull(reference);

            if (reference.Segment is null)
            {
                return reference.Detach();
            }

            return reference.Detach();
        }

        /// <summary>
        /// Resolves a local reference to its current local-view absolute position.
        /// </summary>
        /// <param name="reference">The reference to resolve.</param>
        /// <returns>The current position, or <see cref="DetachedReferencePosition" /> when detached.</returns>
        public int? GetPositionOfReference(LocalReferencePosition reference)
        {
            ArgumentNullException.ThrowIfNull(reference);

            if (reference.Endpoint == ReferenceEndpointKind.Start)
            {
                return 0;
            }

            if (reference.Endpoint == ReferenceEndpointKind.End)
            {
                return GetLength();
            }

            if (reference.Segment is null)
            {
                return DetachedReferencePosition;
            }

            int position = 0;
            List<ISegment> segments = FlattenSegments();
            for (int i = 0; i < segments.Count; i++)
            {
                ISegment segment = segments[i];
                int length = VisibleLength(segment, UnassignedSequenceNumber);
                if (ReferenceEquals(segment, reference.Segment))
                {
                    return length == 0
                        ? GetPositionOfRemovedReference(segments, i, reference)
                        : position + Math.Min(reference.Offset, length);
                }

                position += length;
            }

            return DetachedReferencePosition;
        }

        /// <summary>
        /// Returns the current absolute position of a segment by walking the tree.
        /// </summary>
        /// <param name="segment">The segment to locate.</param>
        /// <returns>The current local-view position, or <see cref="DetachedReferencePosition" /> when the segment is not in the tree or is removed.</returns>
        public int? GetPositionOfSegment(ISegment segment)
        {
            ArgumentNullException.ThrowIfNull(segment);

            int position = 0;
            foreach (ISegment currentSegment in WalkAllSegments())
            {
                int length = VisibleLength(currentSegment, UnassignedSequenceNumber);
                if (ReferenceEquals(currentSegment, segment))
                {
                    if (length == 0)
                    {
                        return HasUnackedRemoveStamp(segment) ? null : DetachedReferencePosition;
                    }

                    return position;
                }

                position += length;
            }

            return DetachedReferencePosition;
        }

        private static bool HasUnackedRemoveStamp(ISegment segment)
        {
            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (stamp.Seq == UnassignedSequenceNumber)
                {
                    return true;
                }
            }

            return false;
        }

        internal int? GetPositionOfSegmentInCurrentView(ISegment segment)
        {
            ArgumentNullException.ThrowIfNull(segment);

            MergeBlock? parent = segment.Parent;
            if (parent is null)
            {
                return null;
            }

            int position = 0;
            IMergeNode childOnPath = segment;
            while (parent is not null)
            {
                for (int i = 0; i < parent.ChildCount; i++)
                {
                    IMergeNode? child = parent.GetChild(i);
                    if (ReferenceEquals(child, childOnPath))
                    {
                        break;
                    }

                    if (child is not null)
                    {
                        position += GetNodeLength(child, UnassignedSequenceNumber, null, null);
                    }
                }

                childOnPath = parent;
                parent = parent.Parent;
            }

            return position;
        }

        internal int? GetPositionOfSegmentAt(
            ISegment segment,
            long refSeq,
            string? clientId,
            long? localSeq)
        {
            ArgumentNullException.ThrowIfNull(segment);

            int? perspectiveClientId = clientId is null ? null : GetClientId(clientId);
            MergeBlock? parent = segment.Parent;
            if (parent is null)
            {
                return DetachedReferencePosition;
            }

            int position = 0;
            IMergeNode childOnPath = segment;
            while (parent is not null)
            {
                for (int i = 0; i < parent.ChildCount; i++)
                {
                    IMergeNode? child = parent.GetChild(i);
                    if (ReferenceEquals(child, childOnPath))
                    {
                        break;
                    }

                    if (child is not null)
                    {
                        position += GetNodeLengthAt(child, refSeq, perspectiveClientId, localSeq);
                    }
                }

                childOnPath = parent;
                parent = parent.Parent;
            }

            return ReferenceEquals(childOnPath, Root) ? position : DetachedReferencePosition;
        }

        internal int? GetPositionOfSegmentForReconnect(ISegment segment, string clientId, long localSeq)
        {
            ArgumentNullException.ThrowIfNull(segment);
            if (string.IsNullOrEmpty(clientId))
            {
                throw new ArgumentException("Client id must be provided.", nameof(clientId));
            }

            int reconnectClientId = GetClientId(clientId);
            int? position = GetPositionOfSegmentAt(segment, CurrentSeq, clientId, localSeq);
            if (position is not int segmentPosition || position == DetachedReferencePosition)
            {
                return null;
            }

            return GetSegmentLengthAt(segment, CurrentSeq, reconnectClientId, localSeq) == 0 ? null : segmentPosition;
        }

        internal int GetVisibleLengthForReconnect(ISegment segment, string clientId, long localSeq)
        {
            ArgumentNullException.ThrowIfNull(segment);
            if (string.IsNullOrEmpty(clientId))
            {
                throw new ArgumentException("Client id must be provided.", nameof(clientId));
            }

            return GetSegmentLengthAt(segment, CurrentSeq, GetClientId(clientId), localSeq);
        }

        internal int RebasePosition(int stalePosition, long staleRefSeq, string clientId)
        {
            if (string.IsNullOrEmpty(clientId))
            {
                throw new ArgumentException("Client id must be provided.", nameof(clientId));
            }

            int staleLength = GetLength(staleRefSeq, clientId);
            ValidatePosition(stalePosition, staleLength, allowEnd: true);
            if (stalePosition == staleLength)
            {
                return GetLength();
            }

            (ISegment segment, int offsetInSegment)? segmentInfo = TryGetContainingSegment(stalePosition, staleRefSeq, clientId);
            if (segmentInfo is null)
            {
                return Math.Min(stalePosition, GetLength());
            }

            int? currentPosition = GetPositionOfSegment(segmentInfo.Value.segment);
            if (currentPosition is null || currentPosition == DetachedReferencePosition)
            {
                return Math.Min(stalePosition, GetLength());
            }

            return Math.Min(currentPosition.Value + segmentInfo.Value.offsetInSegment, GetLength());
        }

        /// <summary>
        /// Finds the segment containing a position in a reference sequence view.
        /// </summary>
        /// <param name="position">The absolute position to resolve.</param>
        /// <param name="refSeq">The reference sequence number to view.</param>
        /// <param name="clientId">The client perspective for same-client and local-unacked visibility.</param>
        /// <returns>The containing segment and offset within that segment.</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when <paramref name="position" /> is outside the visible length.</exception>
        public (ISegment segment, int offsetInSegment) GetContainingSegment(
            int position,
            long refSeq,
            string? clientId = null)
        {
            (ISegment segment, int offsetInSegment)? result = TryGetContainingSegment(position, refSeq, clientId);
            if (result is null)
            {
                throw new ArgumentOutOfRangeException(nameof(position), "Position must identify an existing visible segment.");
            }

            return result.Value;
        }

        /// <summary>
        /// Attempts to find the segment containing a position in a reference sequence view.
        /// </summary>
        /// <param name="position">The absolute position to resolve.</param>
        /// <param name="refSeq">The reference sequence number to view.</param>
        /// <param name="clientId">The client perspective for same-client and local-unacked visibility.</param>
        /// <returns>The containing segment and offset, or <see langword="null" /> when the position is out of range.</returns>
        public (ISegment segment, int offsetInSegment)? TryGetContainingSegment(
            int position,
            long refSeq,
            string? clientId = null)
        {
            if (position < 0)
            {
                return null;
            }

            int? perspectiveClientId = clientId is null ? null : GetClientId(clientId);
            long perspectiveSeq = CurrentSeq + 1;
            if (position >= GetNodeLength(Root, refSeq, perspectiveClientId, perspectiveSeq))
            {
                return null;
            }

            return FindContainingSegment(Root, position, refSeq, perspectiveClientId, perspectiveSeq);
        }

        /// <summary>
        /// Walks all leaf segments in tree order.
        /// </summary>
        /// <returns>All leaf segments, including removed segments.</returns>
        public IEnumerable<ISegment> WalkAllSegments()
        {
            foreach (ISegment segment in WalkSegments(Root))
            {
                yield return segment;
            }
        }

        internal bool SetMinSeq(long minSeq)
        {
            if (minSeq < MinSeq)
            {
                throw new ArgumentOutOfRangeException(nameof(minSeq), "Minimum sequence number cannot move backwards.");
            }

            if (minSeq == MinSeq)
            {
                return false;
            }

            MinSeq = minSeq;
            return true;
        }

        internal void ZamboniSegments(int maxCount = DefaultZamboniSegmentsMaxCount)
        {
            if (maxCount <= 0 || Root.ChildCount == 0)
            {
                return;
            }

            int processedCount = 0;
            bool changed = ZamboniBlock(Root, maxCount, ref processedCount);
            if (processedCount < maxCount)
            {
                changed = CoalesceAcrossBlockBoundaries(maxCount, ref processedCount) || changed;
            }

            if (changed)
            {
                PackSparseBlocks(Root);
            }
        }

        internal List<ISegment> InsertSegments(int position, ISegment[] segments, long refSeq, long seq, int clientId, long? perspectiveSeq = null)
        {
            if (segments is null)
            {
                throw new ArgumentNullException(nameof(segments));
            }

            List<ISegment> newSegments = new();
            foreach (ISegment segment in segments)
            {
                if (segment is null)
                {
                    throw new ArgumentException("Inserted segments cannot contain null entries.", nameof(segments));
                }

                if (segment.CachedLength == 0)
                {
                    continue;
                }

                PrepareInsertedSegment(segment, seq, clientId);
                newSegments.Add(segment);
            }

            long effectivePerspectiveSeq = perspectiveSeq ?? seq;
            List<ISegment> flatSegments = FlattenSegments();
            ValidatePosition(
                position,
                GetVisibleLength(flatSegments, refSeq, clientId, effectivePerspectiveSeq),
                allowEnd: true);
            if (newSegments.Count == 0)
            {
                return new List<ISegment>();
            }

            OperationStamp stamp = new()
            {
                Seq = seq,
                ClientId = clientId,
            };
            int insertionIndex = FindInsertionIndex(flatSegments, position, refSeq, effectivePerspectiveSeq, seq, clientId, stamp);
            for (int i = 0; i < newSegments.Count; i++)
            {
                InsertSegmentAtFlatIndex(flatSegments, insertionIndex + i, newSegments[i], stamp);
            }

            SlideReferencesForInsertedSegments(flatSegments, insertionIndex, newSegments);
            List<ISegment> obliteratedSegments = MarkInsertedSegmentsObliteratedOnInsert(
                flatSegments,
                insertionIndex,
                newSegments.Count,
                refSeq,
                effectivePerspectiveSeq,
                seq,
                clientId);
            SlideReferencesOffRemovedSegments(flatSegments, obliteratedSegments);
            UpdateAffectedBlockPathsForObliterationOnInsert(CollectParentBlocks(obliteratedSegments), stamp, obliteratedSegments);
            RecordSequence(seq);
            List<ISegment> deltaSegments = new();
            foreach (ISegment segment in newSegments)
            {
                if (segment.RemoveStamps.Count == 0)
                {
                    deltaSegments.Add(segment);
                }
            }

            return deltaSegments;
        }

        private static void SlideReferencesForInsertedSegments(
            IReadOnlyList<ISegment> segments,
            int insertionIndex,
            IReadOnlyList<ISegment> insertedSegments)
        {
            if (insertedSegments.Count == 0)
            {
                return;
            }

            ISegment firstInsertedSegment = insertedSegments[0];
            ISegment lastInsertedSegment = insertedSegments[insertedSegments.Count - 1];
            ISegment? leftSegment = insertionIndex > 0 ? segments[insertionIndex - 1] : null;
            ISegment? rightSegment = insertionIndex + insertedSegments.Count < segments.Count
                ? segments[insertionIndex + insertedSegments.Count]
                : null;

            if (leftSegment is not null)
            {
                List<LocalReferencePosition> leftReferences = new(leftSegment.LocalRefs);
                foreach (LocalReferencePosition reference in leftReferences)
                {
                    if (reference.Offset == leftSegment.CachedLength
                        && reference.SlidingPreference != SlidingPreference.Backward)
                    {
                        MoveReferenceToSegment(reference, lastInsertedSegment, lastInsertedSegment.CachedLength);
                    }
                }
            }

            if (rightSegment is not null)
            {
                List<LocalReferencePosition> rightReferences = new(rightSegment.LocalRefs);
                foreach (LocalReferencePosition reference in rightReferences)
                {
                    if (reference.Offset == 0 && reference.SlidingPreference == SlidingPreference.Backward)
                    {
                        MoveReferenceToSegment(reference, firstInsertedSegment, 0);
                    }
                }
            }
        }

        internal List<ISegment> MarkRangeRemoved(int start, int end, long refSeq, long seq, int clientId, long? perspectiveSeq = null)
        {
            long effectivePerspectiveSeq = perspectiveSeq ?? seq;
            List<ISegment> flatSegments = FlattenSegments();
            ValidateRange(start, end, GetVisibleLength(flatSegments, refSeq, clientId, effectivePerspectiveSeq));
            if (start == end)
            {
                return new List<ISegment>();
            }

            OperationStamp stamp = new()
            {
                Seq = seq,
                ClientId = clientId,
            };
            HashSet<MergeBlock> affectedBlocks = new();
            EnsureBoundary(flatSegments, start, refSeq, effectivePerspectiveSeq, clientId, stamp, affectedBlocks);
            EnsureBoundary(flatSegments, end, refSeq, effectivePerspectiveSeq, clientId, stamp, affectedBlocks);
            List<ISegment> removedSegments = MarkVisibleRangeRemoved(
                flatSegments,
                start,
                end,
                refSeq,
                effectivePerspectiveSeq,
                seq,
                clientId,
                out List<ISegment> changedSegments,
                out List<ISegment> deltaSegments,
                out bool requiresStructureRecompute);
            SlideReferencesOffRemovedSegments(flatSegments, removedSegments);
            AddParentBlocks(affectedBlocks, changedSegments);
            if (requiresStructureRecompute || seq == UnassignedSequenceNumber || perspectiveSeq.HasValue)
            {
                RebuildFromRoot();
            }
            else
            {
                UpdateAffectedBlockPaths(affectedBlocks, stamp, changedSegments);
            }

            RecordSequence(seq);
            return deltaSegments;
        }

        internal List<ISegment> ObliterateRange(int start, int end, long refSeq, long seq, int clientId, long? perspectiveSeq = null)
        {
            return ObliterateRangeCore(start, end, Side.Before, Side.After, refSeq, seq, clientId, perspectiveSeq);
        }

        internal List<ISegment> ObliterateRangeSided(SequencePlace start, SequencePlace end, long refSeq, long seq, int clientId, long? perspectiveSeq = null)
        {
            if (start is null)
            {
                throw new ArgumentNullException(nameof(start));
            }

            if (end is null)
            {
                throw new ArgumentNullException(nameof(end));
            }

            long effectivePerspectiveSeq = perspectiveSeq ?? seq;
            List<ISegment> flatSegments = FlattenSegments();
            int length = GetVisibleLength(flatSegments, refSeq, clientId, effectivePerspectiveSeq);
            int startPosition = NormalizeSequencePlacePosition(start, length);
            int endPosition = NormalizeSequencePlacePosition(end, length);
            ValidateRange(startPosition, endPosition, length);
            return ObliterateRangeCore(startPosition, endPosition, start.Side, end.Side, refSeq, seq, clientId, perspectiveSeq, flatSegments);
        }

        private List<ISegment> ObliterateRangeCore(
            int start,
            int end,
            Side startSide,
            Side endSide,
            long refSeq,
            long seq,
            int clientId,
            long? perspectiveSeq = null,
            List<ISegment>? flatSegmentsArg = null)
        {
            long effectivePerspectiveSeq = perspectiveSeq ?? seq;
            List<ISegment> flatSegments = flatSegmentsArg ?? FlattenSegments();
            ValidateRange(start, end, GetVisibleLength(flatSegments, refSeq, clientId, effectivePerspectiveSeq));
            if (start == end)
            {
                return new List<ISegment>();
            }

            OperationStamp stamp = new()
            {
                Seq = seq,
                ClientId = clientId,
            };
            HashSet<MergeBlock> affectedBlocks = new();
            EnsureBoundary(flatSegments, start, refSeq, effectivePerspectiveSeq, clientId, stamp, affectedBlocks);
            EnsureBoundary(flatSegments, end, refSeq, effectivePerspectiveSeq, clientId, stamp, affectedBlocks);
            List<ISegment> obliteratedSegments = MarkVisibleRangeObliterated(
                flatSegments,
                start,
                end,
                refSeq,
                effectivePerspectiveSeq,
                seq,
                clientId,
                startSide,
                endSide,
                out List<ISegment> changedSegments,
                out List<ISegment> deltaSegments,
                out bool requiresStructureRecompute);
            SlideReferencesOffRemovedSegments(flatSegments, obliteratedSegments);
            AddParentBlocks(affectedBlocks, changedSegments);
            if (requiresStructureRecompute || seq == UnassignedSequenceNumber || perspectiveSeq.HasValue)
            {
                RebuildFromRoot();
            }
            else
            {
                UpdateAffectedBlockPaths(affectedBlocks, stamp, changedSegments);
            }

            RecordSequence(seq);
            return deltaSegments;
        }

        internal List<ISegment> AnnotateRange(int start, int end, PropertySet props, long refSeq, long seq, int clientId, long? perspectiveSeq = null)
        {
            if (props is null)
            {
                throw new ArgumentNullException(nameof(props));
            }

            long effectivePerspectiveSeq = perspectiveSeq ?? seq;
            List<ISegment> flatSegments = FlattenSegments();
            ValidateRange(start, end, GetVisibleLength(flatSegments, refSeq, clientId, effectivePerspectiveSeq));
            if (start == end || props.Count == 0)
            {
                return new List<ISegment>();
            }

            OperationStamp stamp = new()
            {
                Seq = seq,
                ClientId = clientId,
            };
            HashSet<MergeBlock> affectedBlocks = new();
            EnsureBoundary(flatSegments, start, refSeq, effectivePerspectiveSeq, clientId, stamp, affectedBlocks);
            EnsureBoundary(flatSegments, end, refSeq, effectivePerspectiveSeq, clientId, stamp, affectedBlocks);
            List<ISegment> deltaSegments = AnnotateVisibleRange(flatSegments, start, end, refSeq, effectivePerspectiveSeq, clientId, props);
            if (affectedBlocks.Count > 0)
            {
                UpdateAffectedBlockPaths(affectedBlocks, stamp);
            }

            RecordSequence(seq);
            return deltaSegments;
        }

        internal ISegment? FirstSegment()
        {
            foreach (ISegment segment in WalkAllSegments())
            {
                return segment;
            }

            return null;
        }

        internal ISegment? LastSegment()
        {
            ISegment? last = null;
            foreach (ISegment segment in WalkAllSegments())
            {
                last = segment;
            }

            return last;
        }

        internal void RebuildFromRoot()
        {
            RebuildFromRoot(updateOrdinals: false);
        }

        private void RebuildFromRoot(bool updateOrdinals)
        {
            PartialLengths.RebuildFromBlock(Root, recurse: true, updateOrdinals: updateOrdinals);
        }

        internal int GetOrAddShortClientId(string? clientId)
        {
            return GetClientId(clientId);
        }

        private static MergeBlock MakeBlock()
        {
            return new MergeBlock(0)
            {
                Ordinal = string.Empty,
            };
        }

        private MergeBlock Split(MergeBlock block)
        {
            int halfCount = MergeBlock.MaxChildren / 2;
            MergeBlock splitBlock = MakeBlock();
            block.ChildCount = halfCount;
            NodeUpdateOrdinals(block);
            for (int i = 0; i < halfCount; i++)
            {
                int sourceIndex = halfCount + i;
                IMergeNode child = block.Children[sourceIndex]
                    ?? throw new InvalidOperationException("Cannot split a block with missing children.");
                block.Children[sourceIndex] = null;
                splitBlock.SetChild(i, child, updateOrdinal: false);
            }

            NodeUpdateLengthNewStructure(block);
            NodeUpdateLengthNewStructure(splitBlock);
            return splitBlock;
        }

        private void UpdateRoot(MergeBlock splitBlock)
        {
            MergeBlock oldRoot = Root;
            MergeBlock newRoot = MakeBlock();
            newRoot.SetChild(0, oldRoot, updateOrdinal: false);
            newRoot.SetChild(1, splitBlock, updateOrdinal: false);
            Root = newRoot;
            Root.Parent = null;
            Root.Index = 0;
            Root.Ordinal = string.Empty;
            NodeUpdateOrdinals(Root);
            NodeUpdateLengthNewStructure(Root);
        }

        private void BlockUpdatePathLengths(
            MergeBlock startNode,
            OperationStamp stamp,
            bool newStructure = false,
            IReadOnlyCollection<ISegment>? changedSegments = null,
            bool validatePartialLengths = true)
        {
            MergeBlock? block = startNode;
            List<MergeBlock>? updatedBlocks = validatePartialLengths ? new List<MergeBlock>() : null;
            while (block is not null)
            {
                if (newStructure)
                {
                    NodeUpdateLengthNewStructure(block);
                }
                else
                {
                    BlockUpdateLength(block, stamp, changedSegments);
                }

                updatedBlocks?.Add(block);
                block = block.Parent;
            }

            if (updatedBlocks is not null)
            {
                foreach (MergeBlock updatedBlock in updatedBlocks)
                {
                    PartialLengths.ValidateBlockPartialLengthInvariants(updatedBlock);
                }
            }
        }

        private void BlockUpdateLength(
            MergeBlock block,
            OperationStamp stamp,
            IReadOnlyCollection<ISegment>? changedSegments)
        {
            PartialLengths.UpdateLengthFor(block, stamp, changedSegments);
        }

        private void NodeUpdateLengthNewStructure(IMergeNode node)
        {
            if (node is MergeBlock block)
            {
                PartialLengths.UpdateStructureFor(block);
            }
        }

        private void NodeUpdateOrdinals(MergeBlock block)
        {
            for (int i = 0; i < block.ChildCount; i++)
            {
                IMergeNode? child = block.GetChild(i);
                if (child is null)
                {
                    continue;
                }

                block.SetOrdinal(child, i);
                if (child is MergeBlock childBlock)
                {
                    NodeUpdateOrdinals(childBlock);
                }
            }
        }

        private bool ZamboniBlock(
            MergeBlock block,
            int maxCount,
            ref int processedCount)
        {
            bool blockStructureChanged = false;
            bool descendantStructureChanged = false;
            ISegment? previousSegment = null;
            for (int i = 0; i < block.ChildCount;)
            {
                if (processedCount >= maxCount)
                {
                    break;
                }

                IMergeNode? child = block.GetChild(i);
                if (child is null)
                {
                    i++;
                    previousSegment = null;
                    continue;
                }

                if (child is MergeBlock childBlock)
                {
                    descendantStructureChanged = ZamboniBlock(childBlock, maxCount, ref processedCount)
                        || descendantStructureChanged;
                    if (childBlock.ChildCount == 0)
                    {
                        block.RemoveChildAt(i);
                        blockStructureChanged = true;
                        previousSegment = null;
                        continue;
                    }

                    i++;
                    previousSegment = null;
                    continue;
                }

                ISegment segment = (ISegment)child;
                if (IsZamboniDead(segment))
                {
                    SlideReferencesOffZamboniDroppedSegment(segment);
                    block.RemoveChildAt(i);
                    blockStructureChanged = true;
                    previousSegment = PreviousCoalesceCandidateInBlock(block, i - 1);
                    processedCount++;
                    continue;
                }

                if (previousSegment is not null && CanZamboniCoalesce(previousSegment, segment))
                {
                    previousSegment.Append(segment);
                    block.RemoveChildAt(i);
                    blockStructureChanged = true;
                    processedCount++;
                    continue;
                }

                previousSegment = IsZamboniCoalesceCandidate(segment) ? segment : null;
                i++;
            }

            if (blockStructureChanged)
            {
                if (block.ChildCount > 0)
                {
                    NodeUpdateOrdinals(block);
                }

                BlockUpdatePathLengths(block, CreateTreeMaintenanceStamp(), newStructure: true);
            }

            return blockStructureChanged || descendantStructureChanged;
        }

        private bool CoalesceAcrossBlockBoundaries(
            int maxCount,
            ref int processedCount)
        {
            List<MergeBlock> leafBlocks = new();
            CollectLeafBlocks(Root, leafBlocks);
            if (leafBlocks.Count < 2)
            {
                return false;
            }

            bool changed = false;
            HashSet<MergeBlock> affectedBlocks = new();
            MergeBlock? leftBlock = null;
            foreach (MergeBlock rightBlock in leafBlocks)
            {
                if (processedCount >= maxCount)
                {
                    break;
                }

                if (!IsBlockAttached(rightBlock) || rightBlock.ChildCount == 0)
                {
                    if (rightBlock.ChildCount == 0)
                    {
                        changed = DropEmptyBlockAndAncestors(rightBlock, affectedBlocks) || changed;
                    }

                    continue;
                }

                if (leftBlock is null || !IsBlockAttached(leftBlock) || leftBlock.ChildCount == 0)
                {
                    leftBlock = rightBlock;
                    continue;
                }

                while (processedCount < maxCount
                    && IsBlockAttached(rightBlock)
                    && rightBlock.ChildCount > 0
                    && TryCoalesceLeafBlockBoundary(leftBlock, rightBlock, affectedBlocks))
                {
                    processedCount++;
                    changed = true;
                    if (rightBlock.ChildCount == 0)
                    {
                        changed = DropEmptyBlockAndAncestors(rightBlock, affectedBlocks) || changed;
                        break;
                    }
                }

                if (IsBlockAttached(rightBlock) && rightBlock.ChildCount > 0)
                {
                    leftBlock = rightBlock;
                }
            }

            if (changed)
            {
                UpdateZamboniChangedBlocks(affectedBlocks);
            }

            return changed;
        }

        private static void CollectLeafBlocks(MergeBlock block, List<MergeBlock> leafBlocks)
        {
            bool hasChildBlock = false;
            for (int i = 0; i < block.ChildCount; i++)
            {
                if (block.GetChild(i) is MergeBlock childBlock)
                {
                    hasChildBlock = true;
                    CollectLeafBlocks(childBlock, leafBlocks);
                }
            }

            if (!hasChildBlock)
            {
                leafBlocks.Add(block);
            }
        }

        private bool TryCoalesceLeafBlockBoundary(
            MergeBlock leftBlock,
            MergeBlock rightBlock,
            HashSet<MergeBlock> affectedBlocks)
        {
            IMergeNode? leftChild = leftBlock.ChildCount == 0 ? null : leftBlock.GetChild(leftBlock.ChildCount - 1);
            IMergeNode? rightChild = rightBlock.ChildCount == 0 ? null : rightBlock.GetChild(0);
            if (leftChild is not ISegment leftSegment
                || rightChild is not ISegment rightSegment
                || !CanZamboniCoalesce(leftSegment, rightSegment))
            {
                return false;
            }

            leftSegment.Append(rightSegment);
            rightBlock.RemoveChildAt(0);
            affectedBlocks.Add(leftBlock);
            affectedBlocks.Add(rightBlock);
            return true;
        }

        private bool DropEmptyBlockAndAncestors(
            MergeBlock block,
            HashSet<MergeBlock> affectedBlocks)
        {
            bool changed = false;
            MergeBlock current = block;
            while (current.ChildCount == 0 && current.Parent is MergeBlock parent)
            {
                if (!parent.DropEmptyBlockAt(current.Index))
                {
                    break;
                }

                affectedBlocks.Add(parent);
                current = parent;
                changed = true;
            }

            return changed;
        }

        private void UpdateZamboniChangedBlocks(HashSet<MergeBlock> affectedBlocks)
        {
            NodeUpdateOrdinals(Root);
            OperationStamp stamp = CreateTreeMaintenanceStamp();
            foreach (MergeBlock block in affectedBlocks)
            {
                if (IsBlockAttached(block))
                {
                    BlockUpdatePathLengths(block, stamp, newStructure: true);
                }
            }
        }

        private bool PackSparseBlocks(MergeBlock block)
        {
            bool changed = false;
            for (int i = 0; i < block.ChildCount; i++)
            {
                if (block.GetChild(i) is MergeBlock childBlock)
                {
                    changed = PackSparseBlocks(childBlock) || changed;
                }
            }

            if (HasOnlyBlockChildren(block) && HasUnderflowChildBlock(block))
            {
                PackParent(block);
                changed = true;
            }

            if (ReferenceEquals(block, Root))
            {
                changed = CollapseRoot() || changed;
            }

            return changed;
        }

        private void PackParent(MergeBlock parent)
        {
            List<MergeBlock> sourceBlocks = new();
            int totalNodeCount = 0;
            for (int i = 0; i < parent.ChildCount; i++)
            {
                MergeBlock childBlock = (MergeBlock)(parent.GetChild(i)
                    ?? throw new InvalidOperationException("Cannot pack a missing child block."));
                sourceBlocks.Add(childBlock);
                totalNodeCount += childBlock.ChildCount;
            }

            for (int i = 0; i < parent.ChildCount; i++)
            {
                if (parent.Children[i] is MergeBlock childBlock)
                {
                    childBlock.Parent = null;
                    childBlock.Index = 0;
                    childBlock.Ordinal = string.Empty;
                }

                parent.Children[i] = null;
            }

            parent.ChildCount = 0;
            if (totalNodeCount == 0)
            {
                NodeUpdateOrdinals(parent);
                BlockUpdatePathLengths(parent, CreateTreeMaintenanceStamp(), newStructure: true);
                return;
            }

            int halfOfMaxNodeCount = MergeBlock.MaxChildren / 2;
            int childCount = Math.Min(
                MaxChildrenBeforeSplit,
                totalNodeCount / halfOfMaxNodeCount);
            if (childCount < 1)
            {
                childCount = 1;
            }

            int baseNodesInBlockCount = totalNodeCount / childCount;
            int remainderCount = totalNodeCount % childCount;
            int sourceBlockIndex = 0;
            for (int nodeIndex = 0; nodeIndex < childCount; nodeIndex++)
            {
                int nodeCount = baseNodesInBlockCount;
                if (remainderCount > 0)
                {
                    nodeCount++;
                    remainderCount--;
                }

                MergeBlock packedBlock = MakeBlock();
                while (packedBlock.ChildCount < nodeCount)
                {
                    while (sourceBlockIndex < sourceBlocks.Count
                        && sourceBlocks[sourceBlockIndex].ChildCount == 0)
                    {
                        sourceBlockIndex++;
                    }

                    if (sourceBlockIndex >= sourceBlocks.Count)
                    {
                        throw new InvalidOperationException("Packed block source nodes were exhausted unexpectedly.");
                    }

                    MergeBlock sourceBlock = sourceBlocks[sourceBlockIndex];
                    int moveCount = Math.Min(nodeCount - packedBlock.ChildCount, sourceBlock.ChildCount);
                    sourceBlock.MoveChildrenTo(packedBlock, moveCount, updateOrdinal: false);
                }

                NodeUpdateLengthNewStructure(packedBlock);
                parent.AppendChild(packedBlock, updateOrdinal: false);
            }

            NodeUpdateOrdinals(parent);
            BlockUpdatePathLengths(parent, CreateTreeMaintenanceStamp(), newStructure: true);
        }

        private bool CollapseRoot()
        {
            bool changed = false;
            while (Root.ChildCount == 1 && Root.GetChild(0) is MergeBlock onlyChild)
            {
                Root = onlyChild;
                Root.Parent = null;
                Root.Index = 0;
                Root.Ordinal = string.Empty;
                NodeUpdateOrdinals(Root);
                BlockUpdatePathLengths(Root, CreateTreeMaintenanceStamp(), newStructure: true);
                changed = true;
            }

            return changed;
        }

        private static bool HasOnlyBlockChildren(MergeBlock block)
        {
            if (block.ChildCount == 0)
            {
                return false;
            }

            for (int i = 0; i < block.ChildCount; i++)
            {
                if (block.GetChild(i) is not MergeBlock)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool HasUnderflowChildBlock(MergeBlock block)
        {
            for (int i = 0; i < block.ChildCount; i++)
            {
                if (block.GetChild(i) is MergeBlock childBlock
                    && childBlock.ChildCount < MergeBlock.MinChildren)
                {
                    return true;
                }
            }

            return false;
        }

        private bool IsBlockAttached(MergeBlock block)
        {
            MergeBlock current = block;
            while (current.Parent is MergeBlock parent)
            {
                current = parent;
            }

            return ReferenceEquals(current, Root);
        }

        private ISegment? PreviousCoalesceCandidateInBlock(MergeBlock block, int index)
        {
            if (index < 0)
            {
                return null;
            }

            IMergeNode? child = block.GetChild(index);
            if (child is not ISegment segment)
            {
                return null;
            }

            return IsZamboniCoalesceCandidate(segment) ? segment : null;
        }

        private void SlideReferencesOffZamboniDroppedSegment(ISegment segment)
        {
            List<ISegment> attachedSegments = FlattenSegments();
            int removedIndex = IndexOfSegment(attachedSegments, segment);
            if (removedIndex < 0 || segment.LocalRefs.Count == 0)
            {
                return;
            }

            List<LocalReferencePosition> references = new(segment.LocalRefs);
            foreach (LocalReferencePosition reference in references)
            {
                SlideReferenceOffRemovedSegment(attachedSegments, removedIndex, reference);
            }
        }

        private bool IsZamboniDead(ISegment segment)
        {
            if (segment.RemoveStamps.Count == 0)
            {
                return false;
            }

            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (!IsAssignedSequenceNumber(stamp.Seq) || stamp.Seq >= MinSeq)
                {
                    return false;
                }
            }

            return true;
        }

        private bool CanZamboniCoalesce(ISegment first, ISegment second)
        {
            return IsZamboniCoalesceCandidate(first)
                && IsZamboniCoalesceCandidate(second)
                && first is TextSegment
                && second is TextSegment
                && string.Equals(first.Type, second.Type, StringComparison.Ordinal)
                && MatchSegmentProperties(first.Properties, second.Properties);
        }

        private static bool MatchSegmentProperties(
            IReadOnlyDictionary<string, object?>? first,
            IReadOnlyDictionary<string, object?>? second)
        {
            int firstCount = first?.Count ?? 0;
            int secondCount = second?.Count ?? 0;
            if (firstCount != secondCount)
            {
                return false;
            }

            if (first is null)
            {
                return true;
            }

            foreach (KeyValuePair<string, object?> property in first)
            {
                if (second is null || !second.TryGetValue(property.Key, out object? secondValue))
                {
                    return false;
                }

                if (!MatchSegmentPropertyValue(property.Value, secondValue))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool MatchSegmentPropertyValue(object? first, object? second)
        {
            if (ReferenceEquals(first, second))
            {
                return true;
            }

            first = NormalizeJsonElement(first);
            second = NormalizeJsonElement(second);
            if (IsNumeric(first) && IsNumeric(second))
            {
                return Convert.ToDecimal(first) == Convert.ToDecimal(second);
            }

            bool firstIsDictionary = TryAsPropertyDictionary(first, out IReadOnlyDictionary<string, object?>? firstDictionary);
            bool secondIsDictionary = TryAsPropertyDictionary(second, out IReadOnlyDictionary<string, object?>? secondDictionary);
            if (firstIsDictionary || secondIsDictionary)
            {
                return firstIsDictionary && secondIsDictionary && MatchSegmentProperties(firstDictionary, secondDictionary);
            }

            bool firstIsList = TryAsPropertyList(first, out List<object?>? firstList);
            bool secondIsList = TryAsPropertyList(second, out List<object?>? secondList);
            if (firstIsList || secondIsList)
            {
                if (!firstIsList || !secondIsList || firstList!.Count != secondList!.Count)
                {
                    return false;
                }

                for (int i = 0; i < firstList.Count; i++)
                {
                    if (!MatchSegmentPropertyValue(firstList[i], secondList[i]))
                    {
                        return false;
                    }
                }

                return true;
            }

            return Equals(first, second);
        }

        private static object? NormalizeJsonElement(object? value)
        {
            return value is JsonElement jsonElement ? JsonElementToObject(jsonElement) : value;
        }

        private static bool IsNumeric(object? value)
        {
            return value is byte
                || value is sbyte
                || value is short
                || value is ushort
                || value is int
                || value is uint
                || value is long
                || value is ulong
                || value is float
                || value is double
                || value is decimal;
        }

        private static bool TryAsPropertyDictionary(
            object? value,
            out IReadOnlyDictionary<string, object?>? dictionary)
        {
            if (value is IReadOnlyDictionary<string, object?> propertyDictionary)
            {
                dictionary = propertyDictionary;
                return true;
            }

            if (value is IDictionary nonGenericDictionary)
            {
                Dictionary<string, object?> converted = new();
                foreach (DictionaryEntry entry in nonGenericDictionary)
                {
                    if (entry.Key is not string key)
                    {
                        dictionary = null;
                        return false;
                    }

                    converted[key] = entry.Value;
                }

                dictionary = converted;
                return true;
            }

            dictionary = null;
            return false;
        }

        private static bool TryAsPropertyList(object? value, out List<object?>? list)
        {
            if (value is null || value is string || value is IReadOnlyDictionary<string, object?> || value is IDictionary)
            {
                list = null;
                return false;
            }

            if (value is IEnumerable enumerable)
            {
                list = new List<object?>();
                foreach (object? item in enumerable)
                {
                    list.Add(item);
                }

                return true;
            }

            list = null;
            return false;
        }

        private static object? JsonElementToObject(JsonElement jsonElement)
        {
            switch (jsonElement.ValueKind)
            {
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;

                case JsonValueKind.String:
                    return jsonElement.GetString();

                case JsonValueKind.Number:
                    if (jsonElement.TryGetInt32(out int intValue))
                    {
                        return intValue;
                    }

                    if (jsonElement.TryGetInt64(out long longValue))
                    {
                        return longValue;
                    }

                    return jsonElement.GetDouble();

                case JsonValueKind.True:
                    return true;

                case JsonValueKind.False:
                    return false;

                case JsonValueKind.Object:
                    PropertySet properties = new();
                    foreach (JsonProperty property in jsonElement.EnumerateObject())
                    {
                        properties[property.Name] = JsonElementToObject(property.Value);
                    }

                    return properties;

                case JsonValueKind.Array:
                    List<object?> values = new();
                    foreach (JsonElement item in jsonElement.EnumerateArray())
                    {
                        values.Add(JsonElementToObject(item));
                    }

                    return values;

                default:
                    return null;
            }
        }

        private bool IsZamboniCoalesceCandidate(ISegment segment)
        {
            return segment.RemoveStamps.Count == 0
                && IsAssignedSequenceNumber(segment.Seq)
                && segment.Seq < MinSeq;
        }

        private static bool IsAssignedSequenceNumber(long? seq)
        {
            return seq is long value && value != UnassignedSequenceNumber;
        }

        private static long? GetAssignedSequenceNumber(long? seq)
        {
            return IsAssignedSequenceNumber(seq) ? seq : null;
        }

        private static OperationStamp CreateTreeMaintenanceStamp()
        {
            return new OperationStamp()
            {
                Seq = Constants.TreeMaintenanceSequenceNumber,
                ClientId = Constants.NonCollabClient,
            };
        }

        internal List<ISegment> ObliterateInsertedSegmentsIfCovered(
            IReadOnlyList<ISegment> insertedSegments,
            long insertRefSeq,
            long insertSeq,
            int insertClientId)
        {
            if (insertedSegments is null)
            {
                throw new ArgumentNullException(nameof(insertedSegments));
            }

            if (insertedSegments.Count == 0)
            {
                return new List<ISegment>();
            }

            List<ISegment> flatSegments = FlattenSegments();
            List<ISegment> obliteratedSegments = new();
            foreach ((int startIndex, int count) in GetContiguousSegmentRuns(flatSegments, insertedSegments))
            {
                obliteratedSegments.AddRange(MarkInsertedSegmentsObliteratedOnInsert(
                    flatSegments,
                    startIndex,
                    count,
                    insertRefSeq,
                    insertSeq,
                    insertSeq,
                    insertClientId));
            }

            if (obliteratedSegments.Count > 0)
            {
                SlideReferencesOffRemovedSegments(flatSegments, obliteratedSegments);
                OperationStamp stamp = new()
                {
                    Seq = insertSeq,
                    ClientId = insertClientId,
                };
                UpdateAffectedBlockPathsForObliterationOnInsert(CollectParentBlocks(obliteratedSegments), stamp, obliteratedSegments);
            }

            return obliteratedSegments;
        }

        private static void PrepareInsertedSegment(ISegment segment, long seq, int clientId)
        {
            segment.InsertionStamp = new InsertOperationStamp()
            {
                Seq = seq,
                ClientId = clientId,
            };
            segment.RemoveStamps.Clear();
            segment.PendingAnnotates = null;
        }

        private static int VisibleLength(
            ISegment segment,
            long refSeq,
            int? perspectiveClientId = null,
            long? perspectiveSeq = null)
        {
            return IsVisibleAt(segment, refSeq, perspectiveClientId, perspectiveSeq) ? segment.CachedLength : 0;
        }

        private static bool IsVisibleAt(
            ISegment segment,
            long refSeq,
            int? perspectiveClientId = null,
            long? perspectiveSeq = null)
        {
            if (refSeq < 0)
            {
                return segment.RemoveStamps.Count == 0;
            }

            bool insertedAtRefSeq = segment.Seq != UnassignedSequenceNumber
                && (segment.Seq <= refSeq
                    || IsSameClientPriorOperation(segment.Seq, segment.ClientId, perspectiveClientId, perspectiveSeq));
            if (!insertedAtRefSeq)
            {
                return false;
            }

            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (stamp.Seq != UnassignedSequenceNumber
                    && (stamp.Seq <= refSeq
                        || IsSameClientPriorOperation(
                            stamp.Seq,
                            stamp.ClientId,
                            perspectiveClientId,
                            perspectiveSeq)))
                {
                    return false;
                }
            }

            return true;
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

        private static int GetVisibleLength(
            IEnumerable<ISegment> segments,
            long refSeq,
            int? perspectiveClientId = null,
            long? perspectiveSeq = null)
        {
            int length = 0;
            foreach (ISegment segment in segments)
            {
                length += VisibleLength(segment, refSeq, perspectiveClientId, perspectiveSeq);
            }

            return length;
        }

        private (ISegment segment, int offset) GetSegmentForReferencePosition(int position)
        {
            int length = GetLength();
            ValidatePosition(position, length, allowEnd: true);
            if (length == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(position), "Cannot create a reference in an empty merge tree.");
            }

            if (position == length)
            {
                ISegment? lastVisibleSegment = null;
                foreach (ISegment segment in WalkAllSegments())
                {
                    if (VisibleLength(segment, UnassignedSequenceNumber) > 0)
                    {
                        lastVisibleSegment = segment;
                    }
                }

                if (lastVisibleSegment is null)
                {
                    throw new ArgumentOutOfRangeException(nameof(position), "Position must identify an existing visible segment.");
                }

                return (lastVisibleSegment, lastVisibleSegment.CachedLength);
            }

            return GetContainingSegment(position, UnassignedSequenceNumber);
        }

        private static int GetNodeLength(
            IMergeNode node,
            long refSeq,
            int? perspectiveClientId,
            long? perspectiveSeq)
        {
            if (node is MergeBlock block)
            {
                return block.PartialLengths.Query(refSeq, perspectiveClientId, perspectiveSeq);
            }

            return PartialLengths.GetSegmentVisibleLength((ISegment)node, refSeq, perspectiveClientId, perspectiveSeq);
        }

        private static (ISegment segment, int offsetInSegment)? FindContainingSegment(
            IMergeNode node,
            int position,
            long refSeq,
            int? perspectiveClientId,
            long? perspectiveSeq)
        {
            if (node is ISegment segment)
            {
                int length = PartialLengths.GetSegmentVisibleLength(segment, refSeq, perspectiveClientId, perspectiveSeq);
                return position < length ? (segment, position) : null;
            }

            MergeBlock block = (MergeBlock)node;
            int remaining = position;
            for (int i = 0; i < block.ChildCount; i++)
            {
                IMergeNode? child = block.GetChild(i);
                if (child is null)
                {
                    continue;
                }

                int childLength = GetNodeLength(child, refSeq, perspectiveClientId, perspectiveSeq);
                if (childLength == 0)
                {
                    continue;
                }

                if (remaining < childLength)
                {
                    return FindContainingSegment(child, remaining, refSeq, perspectiveClientId, perspectiveSeq);
                }

                remaining -= childLength;
            }

            return null;
        }

        private static int GetNodeLengthAt(
            IMergeNode node,
            long refSeq,
            int? perspectiveClientId,
            long? localSeq)
        {
            if (node is ISegment segment)
            {
                return GetSegmentLengthAt(segment, refSeq, perspectiveClientId, localSeq);
            }

            MergeBlock block = (MergeBlock)node;
            int length = 0;
            for (int i = 0; i < block.ChildCount; i++)
            {
                IMergeNode? child = block.GetChild(i);
                if (child is not null)
                {
                    length += GetNodeLengthAt(child, refSeq, perspectiveClientId, localSeq);
                }
            }

            return length;
        }

        private static int GetSegmentLengthAt(
            ISegment segment,
            long refSeq,
            int? perspectiveClientId,
            long? localSeq)
        {
            return IsSegmentPresentAt(segment, refSeq, perspectiveClientId, localSeq)
                ? segment.CachedLength
                : 0;
        }

        private static bool IsSegmentPresentAt(
            ISegment segment,
            long refSeq,
            int? perspectiveClientId,
            long? localSeq)
        {
            if (!HasOccurredAt(segment.InsertionStamp, refSeq, perspectiveClientId, localSeq))
            {
                return false;
            }

            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (HasOccurredAt(stamp, refSeq, perspectiveClientId, localSeq))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool HasOccurredAt(
            OperationStamp? stamp,
            long refSeq,
            int? perspectiveClientId,
            long? localSeq)
        {
            if (stamp is null)
            {
                return true;
            }

            if (stamp.Seq != UnassignedSequenceNumber && stamp.Seq <= refSeq)
            {
                return true;
            }

            if (perspectiveClientId is not int clientId || stamp.ClientId != clientId)
            {
                return false;
            }

            if (localSeq is long localPerspectiveSeq)
            {
                return stamp.LocalSeq is long stampLocalSeq && stampLocalSeq <= localPerspectiveSeq;
            }

            return true;
        }

        private static void ValidatePosition(int position, int length, bool allowEnd)
        {
            int upperBound = allowEnd ? length : length - 1;
            if (position < 0 || position > upperBound)
            {
                throw new ArgumentOutOfRangeException(nameof(position), $"Position must be between 0 and {upperBound}.");
            }
        }

        private static void ValidateRange(int start, int end, int length)
        {
            if (start < 0 || start > length)
            {
                throw new ArgumentOutOfRangeException(nameof(start), $"Start must be between 0 and {length}.");
            }

            if (end < start || end > length)
            {
                throw new ArgumentOutOfRangeException(nameof(end), $"End must be between start and {length}.");
            }
        }

        private static int NormalizeSequencePlacePosition(SequencePlace place, int length)
        {
            if (place.IsStart)
            {
                return 0;
            }

            if (place.IsEnd)
            {
                return length;
            }

            return place.Position;
        }

        private static int CompareStamp(long seqA, int clientIdA, long seqB, int clientIdB)
        {
            if (seqA == UnassignedSequenceNumber)
            {
                if (seqB == UnassignedSequenceNumber)
                {
                    return clientIdA.CompareTo(clientIdB);
                }

                return 1;
            }

            if (seqB == UnassignedSequenceNumber)
            {
                return -1;
            }

            int seqComparison = seqA.CompareTo(seqB);
            return seqComparison != 0 ? seqComparison : clientIdA.CompareTo(clientIdB);
        }

        private static bool ShouldInsertBeforeZeroLengthSegment(
            ISegment segment,
            long insertSeq,
            int insertClientId)
        {
            if (CompareStamp(insertSeq, insertClientId, segment.Seq, segment.ClientId) > 0)
            {
                return true;
            }

            foreach (RemoveOperationStamp stamp in segment.RemoveStamps)
            {
                if (stamp.Seq != UnassignedSequenceNumber
                    && CompareStamp(stamp.Seq, stamp.ClientId, insertSeq, insertClientId) > 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static List<ISegment> MarkInsertedSegmentsObliteratedOnInsert(
            IReadOnlyList<ISegment> segments,
            int insertionIndex,
            int insertedCount,
            long insertRefSeq,
            long perspectiveSeq,
            long insertSeq,
            int insertClientId)
        {
            List<ISegment> obliteratedSegments = new();
            if (insertRefSeq < 0 || insertedCount == 0)
            {
                return obliteratedSegments;
            }

            List<ObliterateCoverage> coverages = FindCoveringObliterates(
                segments,
                insertionIndex,
                insertedCount,
                insertRefSeq,
                perspectiveSeq,
                insertClientId);
            if (coverages.Count == 0)
            {
                return obliteratedSegments;
            }

            for (int i = insertionIndex; i < insertionIndex + insertedCount; i++)
            {
                ISegment segment = segments[i];
                bool wasVisibleInCurrentView = segment.RemoveStamps.Count == 0;
                foreach (ObliterateCoverage obliterateCoverage in coverages)
                {
                    SliceRemoveOperationStamp stamp = new()
                    {
                        Seq = obliterateCoverage.Stamp.Seq,
                        ClientId = obliterateCoverage.Stamp.ClientId,
                        LocalSeq = obliterateCoverage.Stamp.LocalSeq,
                        StartSide = obliterateCoverage.AtStartBoundary && i == insertionIndex ? Side.After : null,
                        EndSide = obliterateCoverage.AtEndBoundary && i == insertionIndex + insertedCount - 1 ? Side.After : null,
                    };

                    AddRemoveStamp(segment, stamp);
                }
                if (wasVisibleInCurrentView)
                {
                    obliteratedSegments.Add(segment);
                }
            }

            return obliteratedSegments;
        }

        private static List<ObliterateCoverage> FindCoveringObliterates(
            IReadOnlyList<ISegment> segments,
            int insertionIndex,
            int insertedCount,
            long insertRefSeq,
            long perspectiveSeq,
            int insertClientId)
        {
            Dictionary<ObliterateStamp, NearestObliterate> leftObliterates = FindNearestEligibleObliterates(
                segments,
                insertionIndex - 1,
                step: -1,
                insertRefSeq,
                perspectiveSeq,
                insertClientId);
            Dictionary<ObliterateStamp, NearestObliterate> rightObliterates = FindNearestEligibleObliterates(
                segments,
                insertionIndex + insertedCount,
                step: 1,
                insertRefSeq,
                perspectiveSeq,
                insertClientId);

            List<ObliterateCoverage> coverages = new();
            foreach (KeyValuePair<ObliterateStamp, NearestObliterate> entry in leftObliterates)
            {
                if (rightObliterates.ContainsKey(entry.Key))
                {
                    coverages.Add(new ObliterateCoverage(entry.Key, atStartBoundary: false, atEndBoundary: false));
                }
                else if (entry.Value.EndSide == Side.After)
                {
                    coverages.Add(new ObliterateCoverage(entry.Key, atStartBoundary: false, atEndBoundary: true));
                }
            }

            foreach (KeyValuePair<ObliterateStamp, NearestObliterate> entry in rightObliterates)
            {
                if (!leftObliterates.ContainsKey(entry.Key) && entry.Value.StartSide == Side.After)
                {
                    coverages.Add(new ObliterateCoverage(entry.Key, atStartBoundary: true, atEndBoundary: false));
                }
            }

            if (coverages.Count == 0)
            {
                return coverages;
            }

            ObliterateCoverage newest = coverages[0];
            foreach (ObliterateCoverage coverage in coverages)
            {
                if (CompareObliterateStamp(coverage.Stamp, newest.Stamp) > 0)
                {
                    newest = coverage;
                }
            }

            if (newest.Stamp.ClientId == insertClientId)
            {
                return new List<ObliterateCoverage>();
            }

            List<ObliterateCoverage> nonLocalCoverages = new();
            foreach (ObliterateCoverage coverage in coverages)
            {
                if (coverage.Stamp.ClientId != insertClientId)
                {
                    nonLocalCoverages.Add(coverage);
                }
            }

            nonLocalCoverages.Sort((a, b) => CompareObliterateStamp(a.Stamp, b.Stamp));
            return nonLocalCoverages;
        }

        private static Dictionary<ObliterateStamp, NearestObliterate> FindNearestEligibleObliterates(
            IReadOnlyList<ISegment> segments,
            int startIndex,
            int step,
            long insertRefSeq,
            long perspectiveSeq,
            int insertClientId)
        {
            Dictionary<ObliterateStamp, NearestObliterate> obliterates = new();
            for (int i = startIndex; i >= 0 && i < segments.Count; i += step)
            {
                ISegment segment = segments[i];
                foreach (NearestObliterate obliterate in GetEligibleObliterates(segment, insertRefSeq))
                {
                    obliterates.TryAdd(obliterate.Stamp, obliterate);
                }

                if (VisibleLength(segment, insertRefSeq, insertClientId, perspectiveSeq) > 0)
                {
                    break;
                }
            }

            return obliterates;
        }

        private static IEnumerable<NearestObliterate> GetEligibleObliterates(
            ISegment segment,
            long insertRefSeq)
        {
            foreach (RemoveOperationStamp removeStamp in segment.RemoveStamps)
            {
                if (removeStamp is not SliceRemoveOperationStamp sliceRemoveStamp)
                {
                    continue;
                }

                if (sliceRemoveStamp.Seq != UnassignedSequenceNumber && sliceRemoveStamp.Seq <= insertRefSeq)
                {
                    continue;
                }

                ObliterateStamp stamp = new(
                    sliceRemoveStamp.Seq,
                    sliceRemoveStamp.ClientId,
                    sliceRemoveStamp.LocalSeq);
                yield return new NearestObliterate(
                    stamp,
                    sliceRemoveStamp.StartSide ?? Side.Before,
                    sliceRemoveStamp.EndSide ?? Side.Before);
            }
        }

        private static IEnumerable<(int startIndex, int count)> GetContiguousSegmentRuns(
            IReadOnlyList<ISegment> allSegments,
            IReadOnlyList<ISegment> targetSegments)
        {
            List<int> indexes = new();
            foreach (ISegment targetSegment in targetSegments)
            {
                int index = IndexOfSegment(allSegments, targetSegment);
                if (index >= 0)
                {
                    indexes.Add(index);
                }
            }

            indexes.Sort();
            for (int i = 0; i < indexes.Count;)
            {
                int startIndex = indexes[i];
                int count = 1;
                i++;
                while (i < indexes.Count && indexes[i] == startIndex + count)
                {
                    count++;
                    i++;
                }

                yield return (startIndex, count);
            }
        }

        private int SplitSegmentAt(
            List<ISegment> segments,
            int index,
            int offset,
            OperationStamp stamp,
            HashSet<MergeBlock>? affectedBlocks = null)
        {
            ISegment segment = segments[index];
            if (offset <= 0)
            {
                return index;
            }

            if (offset >= segment.CachedLength)
            {
                return index + 1;
            }

            ISegment? rightSegment = segment.SplitAt(offset);
            if (rightSegment is null)
            {
                throw new InvalidOperationException("Segment split failed.");
            }

            MergeBlock parent = segment.Parent
                ?? throw new InvalidOperationException("Cannot split a detached segment.");
            InsertChildIntoBlock(parent, segment.Index + 1, rightSegment);
            UpdateAfterChildInsertion(parent, rightSegment, stamp, rebuildAfterInsertion: true);
            if (rightSegment.Parent is MergeBlock rightParent)
            {
                affectedBlocks?.Add(rightParent);
            }

            segments.Insert(index + 1, rightSegment);
            return index + 1;
        }

        private int FindInsertionIndex(
            List<ISegment> segments,
            int position,
            long refSeq,
            long perspectiveSeq,
            long insertSeq,
            int insertClientId,
            OperationStamp stamp)
        {
            int remaining = position;
            for (int i = 0; i < segments.Count; i++)
            {
                ISegment segment = segments[i];
                int length = VisibleLength(segment, refSeq, insertClientId, perspectiveSeq);
                if (remaining < length)
                {
                    return SplitSegmentAt(segments, i, remaining, stamp);
                }

                if (remaining == length)
                {
                    if (length == 0)
                    {
                        if (ShouldInsertBeforeZeroLengthSegment(segment, insertSeq, insertClientId))
                        {
                            return i;
                        }
                    }
                    else
                    {
                        remaining = 0;
                    }

                    continue;
                }

                remaining -= length;
            }

            if (remaining == 0)
            {
                return segments.Count;
            }

            throw new ArgumentOutOfRangeException(nameof(position), "Position is outside the reference sequence view.");
        }

        private int EnsureBoundary(
            List<ISegment> segments,
            int position,
            long refSeq,
            long perspectiveSeq,
            int perspectiveClientId,
            OperationStamp stamp,
            HashSet<MergeBlock> affectedBlocks)
        {
            int remaining = position;
            for (int i = 0; i < segments.Count; i++)
            {
                ISegment segment = segments[i];
                int length = VisibleLength(segment, refSeq, perspectiveClientId, perspectiveSeq);
                if (remaining < length)
                {
                    return SplitSegmentAt(segments, i, remaining, stamp, affectedBlocks);
                }

                if (remaining == length)
                {
                    if (length > 0)
                    {
                        remaining = 0;
                    }

                    continue;
                }

                remaining -= length;
            }

            if (remaining == 0)
            {
                return segments.Count;
            }

            throw new ArgumentOutOfRangeException(nameof(position), "Boundary is outside the reference sequence view.");
        }

        private void InsertSegmentAtFlatIndex(
            List<ISegment> segments,
            int insertionIndex,
            ISegment segment,
            OperationStamp stamp)
        {
            MergeBlock parent;
            int childIndex;
            if (segments.Count == 0)
            {
                parent = Root;
                childIndex = 0;
            }
            else if (insertionIndex < segments.Count)
            {
                ISegment nextSegment = segments[insertionIndex];
                parent = nextSegment.Parent
                    ?? throw new InvalidOperationException("Cannot insert before a detached segment.");
                childIndex = nextSegment.Index;
            }
            else
            {
                ISegment previousSegment = segments[segments.Count - 1];
                parent = previousSegment.Parent
                    ?? throw new InvalidOperationException("Cannot insert after a detached segment.");
                childIndex = previousSegment.Index + 1;
            }

            InsertChildIntoBlock(parent, childIndex, segment);
            UpdateAfterChildInsertion(parent, segment, stamp);
            segments.Insert(insertionIndex, segment);
        }

        private void InsertChildIntoBlock(MergeBlock block, int childIndex, IMergeNode child)
        {
            if (childIndex < 0 || childIndex > block.ChildCount)
            {
                throw new ArgumentOutOfRangeException(nameof(childIndex), "Child index must be within the populated child range.");
            }

            if (block.ChildCount >= MergeBlock.MaxChildren)
            {
                throw new InvalidOperationException("Cannot insert into an already-full merge block.");
            }

            for (int i = block.ChildCount; i > childIndex; i--)
            {
                IMergeNode? movedChild = block.Children[i - 1];
                block.Children[i] = movedChild;
                if (movedChild is not null)
                {
                    movedChild.Parent = block;
                    movedChild.Index = i;
                }
            }

            block.Children[childIndex] = child;
            block.ChildCount++;
            child.Parent = block;
            child.Index = childIndex;
            block.SetOrdinal(child, childIndex);
        }

        private void UpdateAfterChildInsertion(
            MergeBlock block,
            IMergeNode insertedChild,
            OperationStamp stamp,
            bool rebuildAfterInsertion = false)
        {
            MergeBlock current = block;
            IReadOnlyCollection<ISegment>? changedSegments = insertedChild is ISegment insertedSegment
                ? new ISegment[] { insertedSegment }
                : null;
            while (true)
            {
                if (current.ChildCount < MergeBlock.MaxChildren)
                {
                    bool newStructure = insertedChild is MergeBlock;
                    if (insertedChild is MergeBlock insertedBlock)
                    {
                        NodeUpdateOrdinals(insertedBlock);
                    }

                    BlockUpdatePathLengths(current, stamp, newStructure, changedSegments, validatePartialLengths: !rebuildAfterInsertion);
                    if (rebuildAfterInsertion)
                    {
                        RebuildFromRoot();
                    }

                    return;
                }

                MergeBlock splitBlock = Split(current);
                if (current.Parent is null)
                {
                    UpdateRoot(splitBlock);
                    if (rebuildAfterInsertion)
                    {
                        RebuildFromRoot();
                    }

                    return;
                }

                MergeBlock parent = current.Parent;
                InsertChildIntoBlock(parent, current.Index + 1, splitBlock);
                current = parent;
                insertedChild = splitBlock;
            }
        }

        private static List<ISegment> MarkVisibleRangeRemoved(
            IEnumerable<ISegment> segments,
            int start,
            int end,
            long refSeq,
            long perspectiveSeq,
            long seq,
            int clientId,
            out List<ISegment> changedSegments,
            out List<ISegment> deltaSegments,
            out bool requiresStructureRecompute)
        {
            List<ISegment> removedSegments = new();
            changedSegments = new List<ISegment>();
            deltaSegments = new List<ISegment>();
            requiresStructureRecompute = false;
            int position = 0;
            foreach (ISegment segment in segments)
            {
                int length = VisibleLength(segment, refSeq, clientId, perspectiveSeq);
                if (length == 0)
                {
                    continue;
                }

                int nextPosition = position + length;
                if (nextPosition <= start)
                {
                    position = nextPosition;
                    continue;
                }

                if (position >= end)
                {
                    return removedSegments;
                }

                SetRemoveOperationStamp stamp = new()
                {
                    Seq = seq,
                    ClientId = clientId,
                };
                requiresStructureRecompute = requiresStructureRecompute || segment.RemoveStamps.Count > 0;
                bool wasVisibleInCurrentView = segment.RemoveStamps.Count == 0;
                bool shouldSlideLocalOverlap = ShouldSlideLocalOverlap(segment, stamp);
                AddRemoveStamp(segment, stamp);
                changedSegments.Add(segment);
                if (wasVisibleInCurrentView)
                {
                    removedSegments.Add(segment);
                    deltaSegments.Add(segment);
                }
                else if (shouldSlideLocalOverlap)
                {
                    removedSegments.Add(segment);
                }

                position = nextPosition;
            }

            return removedSegments;
        }

        private static List<ISegment> MarkVisibleRangeObliterated(
            IEnumerable<ISegment> segments,
            int start,
            int end,
            long refSeq,
            long perspectiveSeq,
            long seq,
            int clientId,
            Side startSide,
            Side endSide,
            out List<ISegment> changedSegments,
            out List<ISegment> deltaSegments,
            out bool requiresStructureRecompute)
        {
            List<ISegment> obliteratedSegments = new();
            changedSegments = new List<ISegment>();
            deltaSegments = new List<ISegment>();
            requiresStructureRecompute = false;
            int position = 0;
            ISegment? firstStampedSegment = null;
            ISegment? lastStampedSegment = null;
            foreach (ISegment segment in segments)
            {
                int length = VisibleLength(segment, refSeq, clientId, perspectiveSeq);
                if (length == 0)
                {
                    continue;
                }

                int nextPosition = position + length;
                if (nextPosition <= start)
                {
                    position = nextPosition;
                    continue;
                }

                if (position >= end)
                {
                    break;
                }

                SliceRemoveOperationStamp stamp = new()
                {
                    Seq = seq,
                    ClientId = clientId,
                };
                requiresStructureRecompute = requiresStructureRecompute || segment.RemoveStamps.Count > 0;
                bool wasVisibleInCurrentView = segment.RemoveStamps.Count == 0;
                bool shouldSlideLocalOverlap = ShouldSlideLocalOverlap(segment, stamp);
                AddRemoveStamp(segment, stamp);
                firstStampedSegment ??= segment;
                lastStampedSegment = segment;
                changedSegments.Add(segment);
                if (wasVisibleInCurrentView)
                {
                    obliteratedSegments.Add(segment);
                    deltaSegments.Add(segment);
                }
                else if (shouldSlideLocalOverlap)
                {
                    obliteratedSegments.Add(segment);
                }

                position = nextPosition;
            }

            if (firstStampedSegment is not null)
            {
                UpdateLastSliceRemoveStamp(
                    firstStampedSegment,
                    seq,
                    clientId,
                    stamp => stamp with { StartSide = startSide });
                UpdateLastSliceRemoveStamp(
                    lastStampedSegment!,
                    seq,
                    clientId,
                    stamp => stamp with { EndSide = endSide });
            }

            return obliteratedSegments;
        }

        private static void AddRemoveStamp(ISegment segment, RemoveOperationStamp stamp)
        {
            foreach (RemoveOperationStamp existingStamp in segment.RemoveStamps)
            {
                if (Stamps.Equal(existingStamp, stamp) && string.Equals(existingStamp.Type, stamp.Type, StringComparison.Ordinal))
                {
                    return;
                }
            }

            Stamps.SpliceIntoList(segment.RemoveStamps, stamp);
        }

        private static bool ShouldSlideLocalOverlap(ISegment segment, RemoveOperationStamp newStamp)
        {
            return newStamp.Seq != UnassignedSequenceNumber
                && segment.RemoveStamps.Count > 0
                && !Stamps.HasAnyAckedOperation(segment.RemoveStamps);
        }

        private static void UpdateLastSliceRemoveStamp(
            ISegment segment,
            long seq,
            int clientId,
            Func<SliceRemoveOperationStamp, SliceRemoveOperationStamp> update)
        {
            for (int i = segment.RemoveStamps.Count - 1; i >= 0; i--)
            {
                if (segment.RemoveStamps[i] is SliceRemoveOperationStamp stamp
                    && stamp.Seq == seq
                    && stamp.ClientId == clientId)
                {
                    segment.RemoveStamps[i] = update(stamp);
                    return;
                }
            }
        }

        private static void SlideReferencesOffRemovedSegments(
            IReadOnlyList<ISegment> segments,
            IEnumerable<ISegment> removedSegments)
        {
            foreach (ISegment removedSegment in removedSegments)
            {
                int removedIndex = IndexOfSegment(segments, removedSegment);
                if (removedIndex < 0 || removedSegment.LocalRefs.Count == 0)
                {
                    continue;
                }

                List<LocalReferencePosition> references = new(removedSegment.LocalRefs);
                foreach (LocalReferencePosition reference in references)
                {
                    SlideReferenceOffRemovedSegment(segments, removedIndex, reference);
                }
            }
        }

        private static void SlideReferenceOffRemovedSegment(
            IReadOnlyList<ISegment> segments,
            int removedIndex,
            LocalReferencePosition reference)
        {
            if (RefTypeIncludesFlag(reference.RefType, ReferenceType.StayOnRemove))
            {
                return;
            }

            if (!RefTypeIncludesFlag(reference.RefType, ReferenceType.SlideOnRemove))
            {
                DetachReference(reference);
                return;
            }

            if (TryGetSlideTarget(segments, removedIndex, reference, out SlideTarget slideTarget))
            {
                if (slideTarget.Segment is ISegment targetSegment)
                {
                    MoveReferenceToSegment(reference, targetSegment, slideTarget.Offset);
                }
                else
                {
                    MoveReferenceToEndpoint(reference, slideTarget.Endpoint);
                }

                return;
            }

            DetachReference(reference);
        }

        private static int GetPositionOfRemovedReference(
            IReadOnlyList<ISegment> segments,
            int removedIndex,
            LocalReferencePosition reference)
        {
            if (!RefTypeIncludesFlag(reference.RefType, ReferenceType.StayOnRemove)
                && !RefTypeIncludesFlag(reference.RefType, ReferenceType.Transient))
            {
                return DetachedReferencePosition;
            }

            if (!TryGetSlideTarget(segments, removedIndex, reference, out SlideTarget slideTarget))
            {
                return DetachedReferencePosition;
            }

            if (slideTarget.Endpoint == ReferenceEndpointKind.Start)
            {
                return 0;
            }

            if (slideTarget.Endpoint == ReferenceEndpointKind.End)
            {
                return GetVisibleLength(segments, UnassignedSequenceNumber);
            }

            if (slideTarget.Segment is null)
            {
                return DetachedReferencePosition;
            }

            int segmentPosition = GetPositionOfVisibleSegment(segments, slideTarget.Segment);
            return segmentPosition == DetachedReferencePosition
                ? DetachedReferencePosition
                : segmentPosition + slideTarget.Offset;
        }

        private static bool TryGetSlideTarget(
            IReadOnlyList<ISegment> segments,
            int removedIndex,
            LocalReferencePosition reference,
            out SlideTarget target)
        {
            bool preferBackward = reference.SlidingPreference == SlidingPreference.Backward;
            if (TryGetSlideTargetInDirection(segments, removedIndex, preferBackward, out target))
            {
                return true;
            }

            if (reference.CanSlideToEndpoint)
            {
                target = new SlideTarget(
                    null,
                    0,
                    preferBackward ? ReferenceEndpointKind.Start : ReferenceEndpointKind.End);
                return true;
            }

            return TryGetSlideTargetInDirection(segments, removedIndex, !preferBackward, out target);
        }

        private static bool TryGetSlideTargetInDirection(
            IReadOnlyList<ISegment> segments,
            int removedIndex,
            bool backward,
            out SlideTarget target)
        {
            ISegment? segment = backward
                ? FindPreviousVisibleSegment(segments, removedIndex)
                : FindNextVisibleSegment(segments, removedIndex);
            if (segment is null)
            {
                target = default;
                return false;
            }

            target = new SlideTarget(
                segment,
                backward ? Math.Max(0, segment.CachedLength - 1) : 0,
                ReferenceEndpointKind.None);
            return true;
        }

        private static void DetachReference(LocalReferencePosition reference)
        {
            reference.Detach();
        }

        private static void MoveReferenceToSegment(
            LocalReferencePosition reference,
            ISegment targetSegment,
            int targetOffset)
        {
            if (reference.Segment is not null)
            {
                reference.Segment.LocalRefs.Remove(reference);
            }

            reference.Segment = targetSegment;
            reference.Endpoint = ReferenceEndpointKind.None;
            reference.Offset = targetOffset;
            targetSegment.LocalRefs.Add(reference);
        }

        private static void MoveReferenceToEndpoint(
            LocalReferencePosition reference,
            ReferenceEndpointKind endpoint)
        {
            if (endpoint == ReferenceEndpointKind.None)
            {
                DetachReference(reference);
                return;
            }

            if (reference.Segment is not null)
            {
                reference.Segment.LocalRefs.Remove(reference);
            }

            reference.Segment = null;
            reference.Endpoint = endpoint;
            reference.Offset = 0;
        }

        private static int GetPositionOfVisibleSegment(IReadOnlyList<ISegment> segments, ISegment targetSegment)
        {
            int position = 0;
            foreach (ISegment segment in segments)
            {
                int length = VisibleLength(segment, UnassignedSequenceNumber);
                if (ReferenceEquals(segment, targetSegment))
                {
                    return length == 0 ? DetachedReferencePosition : position;
                }

                position += length;
            }

            return DetachedReferencePosition;
        }

        private static ISegment? FindNextVisibleSegment(IReadOnlyList<ISegment> segments, int index)
        {
            for (int i = index + 1; i < segments.Count; i++)
            {
                if (VisibleLength(segments[i], UnassignedSequenceNumber) > 0)
                {
                    return segments[i];
                }
            }

            return null;
        }

        private static ISegment? FindPreviousVisibleSegment(IReadOnlyList<ISegment> segments, int index)
        {
            for (int i = index - 1; i >= 0; i--)
            {
                if (VisibleLength(segments[i], UnassignedSequenceNumber) > 0)
                {
                    return segments[i];
                }
            }

            return null;
        }

        private static int IndexOfSegment(IReadOnlyList<ISegment> segments, ISegment segment)
        {
            for (int i = 0; i < segments.Count; i++)
            {
                if (ReferenceEquals(segments[i], segment))
                {
                    return i;
                }
            }

            return -1;
        }

        private static bool RefTypeIncludesFlag(ReferenceType refType, ReferenceType flags)
        {
            return (refType & flags) != 0;
        }

        private static List<ISegment> AnnotateVisibleRange(
            IEnumerable<ISegment> segments,
            int start,
            int end,
            long refSeq,
            long perspectiveSeq,
            int clientId,
            IReadOnlyDictionary<string, object?> props)
        {
            List<ISegment> deltaSegments = new();
            int position = 0;
            foreach (ISegment segment in segments)
            {
                int length = VisibleLength(segment, refSeq, clientId, perspectiveSeq);
                if (length == 0)
                {
                    continue;
                }

                int nextPosition = position + length;
                if (nextPosition <= start)
                {
                    position = nextPosition;
                    continue;
                }

                if (position >= end)
                {
                    return deltaSegments;
                }

                ValidateMarkerIdAnnotation(segment, props);
                PropertySet properties = PropertyMap.ClonePropertySet(segment.Properties) ?? new PropertySet();
                PropertyMap.Extend(properties, props);
                segment.Properties = properties.Count == 0 ? null : properties;
                if (segment.RemoveStamps.Count == 0)
                {
                    deltaSegments.Add(segment);
                }

                position = nextPosition;
            }

            return deltaSegments;
        }

        private static void ValidateMarkerIdAnnotation(
            ISegment segment,
            IReadOnlyDictionary<string, object?> props)
        {
            if (segment is not Marker marker || !props.ContainsKey(Marker.ReservedMarkerIdKey))
            {
                return;
            }

            string? existingId = marker.GetId();
            if (existingId is null
                || props[Marker.ReservedMarkerIdKey] is not string proposedId
                || !string.Equals(existingId, proposedId, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Cannot change the markerId of an existing marker.");
            }
        }

        private static int CompareObliterateStamp(ObliterateStamp a, ObliterateStamp b)
        {
            if (a.Seq == UnassignedSequenceNumber && b.Seq == UnassignedSequenceNumber)
            {
                int localSeqComparison = Nullable.Compare(a.LocalSeq, b.LocalSeq);
                return localSeqComparison != 0 ? localSeqComparison : a.ClientId.CompareTo(b.ClientId);
            }

            if (a.Seq == UnassignedSequenceNumber)
            {
                return 1;
            }

            if (b.Seq == UnassignedSequenceNumber)
            {
                return -1;
            }

            int seqComparison = a.Seq.CompareTo(b.Seq);
            return seqComparison != 0 ? seqComparison : a.ClientId.CompareTo(b.ClientId);
        }

        private static IEnumerable<ISegment> WalkSegments(IMergeNode node)
        {
            if (node is ISegment segment)
            {
                yield return segment;
                yield break;
            }

            MergeBlock block = (MergeBlock)node;
            for (int i = 0; i < block.ChildCount; i++)
            {
                IMergeNode? child = block.GetChild(i);
                if (child is null)
                {
                    continue;
                }

                foreach (ISegment childSegment in WalkSegments(child))
                {
                    yield return childSegment;
                }
            }
        }

        private int GetClientId(string? clientId)
        {
            if (string.IsNullOrEmpty(clientId))
            {
                return Constants.LocalClientId;
            }

            if (int.TryParse(clientId, out int parsedClientId))
            {
                // Reserve parsed numeric ids in the allocator so a later
                // non-numeric client cannot be assigned the same integer.
                if (parsedClientId >= _nextClientId)
                {
                    _nextClientId = parsedClientId + 1;
                }

                return parsedClientId;
            }

            if (_clientIds.TryGetValue(clientId, out int existingClientId))
            {
                return existingClientId;
            }

            int allocatedClientId = _nextClientId++;
            _clientIds.Add(clientId, allocatedClientId);
            return allocatedClientId;
        }

        private static HashSet<MergeBlock> CollectParentBlocks(IEnumerable<ISegment> segments)
        {
            HashSet<MergeBlock> blocks = new();
            AddParentBlocks(blocks, segments);
            return blocks;
        }

        private static void AddParentBlocks(HashSet<MergeBlock> blocks, IEnumerable<ISegment> segments)
        {
            foreach (ISegment segment in segments)
            {
                if (segment.Parent is MergeBlock parent)
                {
                    blocks.Add(parent);
                }
            }
        }

        internal void UpdatePartialLengthsForSegments(
            IReadOnlyCollection<ISegment> segments,
            OperationStamp stamp,
            bool newStructure = false)
        {
            if (newStructure)
            {
                RebuildFromRoot();
                return;
            }

            UpdateAffectedBlockPaths(CollectParentBlocks(segments), stamp, segments, newStructure);
        }

        private void UpdateAffectedBlockPaths(
            IEnumerable<MergeBlock> blocks,
            OperationStamp stamp,
            IReadOnlyCollection<ISegment>? changedSegments = null,
            bool newStructure = false,
            bool validatePartialLengths = true)
        {
            foreach (MergeBlock block in blocks)
            {
                BlockUpdatePathLengths(block, stamp, newStructure, changedSegments, validatePartialLengths);
            }
        }

        private void UpdateAffectedBlockPathsForObliterationOnInsert(
            IEnumerable<MergeBlock> blocks,
            OperationStamp insertStamp,
            IReadOnlyCollection<ISegment> obliteratedSegments)
        {
            if (obliteratedSegments.Count == 0)
            {
                return;
            }

            RebuildFromRoot();
        }

        private static void ValidateAffectedBlockPaths(IEnumerable<MergeBlock> blocks)
        {
            HashSet<MergeBlock> visited = new();
            foreach (MergeBlock block in blocks)
            {
                MergeBlock? current = block;
                while (current is not null)
                {
                    if (visited.Add(current))
                    {
                        PartialLengths.ValidateBlockPartialLengthInvariants(current);
                    }

                    current = current.Parent;
                }
            }
        }

        private static IEnumerable<OperationStamp> GetObliterationStamps(IEnumerable<ISegment> segments)
        {
            HashSet<ObliterateStamp> seen = new();
            foreach (ISegment segment in segments)
            {
                foreach (RemoveOperationStamp removeStamp in segment.RemoveStamps)
                {
                    if (removeStamp is not SliceRemoveOperationStamp)
                    {
                        continue;
                    }

                    ObliterateStamp key = new(
                        removeStamp.Seq,
                        removeStamp.ClientId,
                        removeStamp.LocalSeq);
                    if (!seen.Add(key))
                    {
                        continue;
                    }

                    yield return new OperationStamp()
                    {
                        Seq = key.Seq,
                        ClientId = key.ClientId,
                        LocalSeq = key.LocalSeq,
                    };
                }
            }
        }

        private readonly struct SlideTarget
        {
            public SlideTarget(ISegment? segment, int offset, ReferenceEndpointKind endpoint)
            {
                Segment = segment;
                Offset = offset;
                Endpoint = endpoint;
            }

            public ISegment? Segment { get; }

            public int Offset { get; }

            public ReferenceEndpointKind Endpoint { get; }
        }

        private readonly struct NearestObliterate
        {
            public NearestObliterate(ObliterateStamp stamp, Side startSide, Side endSide)
            {
                Stamp = stamp;
                StartSide = startSide;
                EndSide = endSide;
            }

            public ObliterateStamp Stamp { get; }

            public Side StartSide { get; }

            public Side EndSide { get; }
        }

        private readonly struct ObliterateCoverage
        {
            public ObliterateCoverage(ObliterateStamp stamp, bool atStartBoundary, bool atEndBoundary)
            {
                Stamp = stamp;
                AtStartBoundary = atStartBoundary;
                AtEndBoundary = atEndBoundary;
            }

            public ObliterateStamp Stamp { get; }

            public bool AtStartBoundary { get; }

            public bool AtEndBoundary { get; }
        }

        private readonly struct ObliterateStamp : IEquatable<ObliterateStamp>
        {
            public ObliterateStamp(long seq, int clientId, long? localSeq)
            {
                Seq = seq;
                ClientId = clientId;
                LocalSeq = localSeq;
            }

            public long Seq { get; }

            public int ClientId { get; }

            public long? LocalSeq { get; }

            public bool Equals(ObliterateStamp other)
            {
                return Seq == other.Seq && ClientId == other.ClientId && LocalSeq == other.LocalSeq;
            }

            public override bool Equals(object? obj)
            {
                return obj is ObliterateStamp other && Equals(other);
            }

            public override int GetHashCode()
            {
                return HashCode.Combine(Seq, ClientId, LocalSeq);
            }
        }

        private List<ISegment> FlattenSegments()
        {
            List<ISegment> segments = new();
            foreach (ISegment segment in WalkAllSegments())
            {
                segments.Add(segment);
            }

            return segments;
        }

        // Snapshot load rebuilds from a flat segment list in one bottom-up pass.
        // Mid-mutation paths use path-local split/length updates instead.
        internal void RebuildFromSegments(IReadOnlyList<ISegment> segments)
        {
            Root = BuildBalancedTree(segments);
            Root.Parent = null;
            Root.Index = 0;
            Root.Ordinal = string.Empty;
            UpdateOrdinalsAndLengths(Root);
        }

        private MergeBlock BuildBalancedTree(IReadOnlyList<ISegment> segments)
        {
            if (segments.Count == 0)
            {
                return MakeBlock();
            }

            List<IMergeNode> level = new();
            foreach (ISegment segment in segments)
            {
                level.Add(segment);
            }

            while (level.Count > MaxChildrenBeforeSplit)
            {
                level = BuildParentLevel(level);
            }

            MergeBlock root = MakeBlock();
            for (int i = 0; i < level.Count; i++)
            {
                root.SetChild(i, level[i], updateOrdinal: false);
            }

            return root;
        }

        private List<IMergeNode> BuildParentLevel(IReadOnlyList<IMergeNode> level)
        {
            List<IMergeNode> parentLevel = new();
            for (int i = 0; i < level.Count;)
            {
                MergeBlock block = MakeBlock();
                int childCount = Math.Min(MaxChildrenBeforeSplit, level.Count - i);
                for (int childIndex = 0; childIndex < childCount; childIndex++, i++)
                {
                    block.SetChild(childIndex, level[i], updateOrdinal: false);
                }

                parentLevel.Add(block);
            }

            return parentLevel;
        }

        private void UpdateOrdinalsAndLengths(MergeBlock block)
        {
            RebuildFromRoot(updateOrdinals: true);
        }

        private void RecordSequence(long seq)
        {
            if (seq > CurrentSeq)
            {
                CurrentSeq = seq;
            }
        }
    }
}

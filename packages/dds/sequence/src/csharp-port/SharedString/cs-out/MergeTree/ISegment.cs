// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/mergeTreeNodes.ts (subset for POC)
// Part of the SharedString C# feasibility port — Wave 1.
// POC scope: base segment/block types, TextSegment-only, basic obliterate.
// Attribution, local refs, and Marker deferred to later waves.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    /// <summary>
    /// Common information for a node in the merge tree.
    /// </summary>
    public interface IMergeNode
    {
        /// <summary>
        /// Gets or sets the parent merge block, if this node is attached to a tree.
        /// </summary>
        MergeBlock? Parent { get; set; }

        /// <summary>
        /// Gets or sets the index of this node in its parent's child list.
        /// </summary>
        int Index { get; set; }

        /// <summary>
        /// Gets or sets the ordinal used to compare this node's location to other nodes in the same tree.
        /// </summary>
        string Ordinal { get; set; }

        /// <summary>
        /// Gets or sets the cached length of this node's contents or descendants.
        /// </summary>
        int CachedLength { get; set; }

        /// <summary>
        /// Returns whether this node is a segment leaf.
        /// </summary>
        /// <returns><see langword="true" /> for segment leaves; otherwise, <see langword="false" />.</returns>
        bool IsLeaf();
    }

    /// <summary>
    /// Represents an interior merge-tree block.
    /// </summary>
    public interface IHierBlock : IMergeNode
    {
        /// <summary>
        /// Gets or sets the number of populated child slots.
        /// </summary>
        int ChildCount { get; set; }

        /// <summary>
        /// Gets the maximum number of child slots in this block.
        /// </summary>
        int Capacity { get; }

        /// <summary>
        /// Gets the fixed-size child storage for this block.
        /// </summary>
        IMergeNode?[] Children { get; }
    }

    /// <summary>
    /// A segment representing a leaf portion of the merge tree.
    /// </summary>
    public interface ISegment : IMergeNode
    {
        /// <summary>
        /// Gets the segment type discriminator.
        /// </summary>
        string Type { get; }

        /// <summary>
        /// Gets or sets properties that have been added to this segment via annotation.
        /// </summary>
        PropertySet? Properties { get; set; }

        /// <summary>
        /// Gets the local references attached to this segment.
        /// </summary>
        internal List<LocalReferencePosition> LocalRefs { get; }

        /// <summary>
        /// Gets or sets the stamp for the operation that inserted this segment.
        /// </summary>
        internal OperationStamp? InsertionStamp { get; set; }

        /// <summary>
        /// Gets the ordered list of remove and obliterate stamps that have affected this segment.
        /// </summary>
        internal List<RemoveOperationStamp> RemoveStamps { get; }

        /// <summary>
        /// Gets or sets the sequence number that created this segment.
        /// </summary>
        long Seq { get; set; }

        /// <summary>
        /// Gets or sets the local client sequence number that created this segment before ack.
        /// </summary>
        long? LocalSeq { get; set; }

        /// <summary>
        /// Gets or sets the client id that created this segment.
        /// </summary>
        int ClientId { get; set; }

        /// <summary>
        /// Gets or sets the sequence number that removed this segment, if any.
        /// </summary>
        long? RemovedSeq { get; set; }

        /// <summary>
        /// Gets or sets the local client sequence number that removed this segment before ack.
        /// </summary>
        long? RemovedLocalSeq { get; set; }

        /// <summary>
        /// Gets or sets the client id that removed this segment, if any.
        /// </summary>
        int? RemovedClientId { get; set; }

        /// <summary>
        /// Gets or sets the sequence number that obliterated this segment, if any.
        /// </summary>
        long? ObliteratedSeq { get; set; }

        /// <summary>
        /// Gets or sets the local client sequence number that obliterated this segment before ack.
        /// </summary>
        long? ObliteratedLocalSeq { get; set; }

        /// <summary>
        /// Gets or sets the client id that obliterated this segment, if any.
        /// </summary>
        int? ObliteratedClientId { get; set; }

        /// <summary>
        /// Gets or sets the sidedness of the start boundary for the obliterate that stamped this segment.
        /// </summary>
        Side? ObliterateStartSide { get; set; }

        /// <summary>
        /// Gets or sets the sidedness of the end boundary for the obliterate that stamped this segment.
        /// </summary>
        Side? ObliterateEndSide { get; set; }

        /// <summary>
        /// Gets or sets pending local annotate metadata keyed by local sequence number.
        /// </summary>
        Dictionary<long, PropertySet>? PendingAnnotates { get; set; }

        /// <summary>
        /// Creates a copy of this segment.
        /// </summary>
        /// <returns>The cloned segment.</returns>
        ISegment Clone();

        /// <summary>
        /// Returns whether another segment can be appended to this segment.
        /// </summary>
        /// <param name="segment">The candidate segment to append.</param>
        /// <returns><see langword="true" /> if the segment can be appended; otherwise, <see langword="false" />.</returns>
        bool CanAppend(ISegment segment);

        /// <summary>
        /// Appends another segment to this segment.
        /// </summary>
        /// <param name="segment">The segment to append.</param>
        void Append(ISegment segment);

        /// <summary>
        /// Splits this segment at the supplied position.
        /// </summary>
        /// <param name="pos">The position at which to split.</param>
        /// <returns>The right-hand segment, or <see langword="null" /> when no split occurred.</returns>
        ISegment? SplitAt(int pos);

        /// <summary>
        /// Serializes this segment to its JSON-compatible representation.
        /// </summary>
        /// <returns>A JSON-compatible representation of this segment.</returns>
        object ToJSONObject();
    }

    /// <summary>
    /// Base implementation for merge-tree segments.
    /// </summary>
    public abstract class BaseSegment : ISegment
    {
        private readonly List<LocalReferencePosition> _localRefs = new();

        // TODO(post-POC): attribution, marker.

        /// <inheritdoc />
        public MergeBlock? Parent { get; set; }

        /// <inheritdoc />
        public int Index { get; set; }

        /// <inheritdoc />
        public string Ordinal { get; set; } = string.Empty;

        /// <inheritdoc />
        public int CachedLength { get; set; }

        /// <inheritdoc />
        public PropertySet? Properties { get; set; }

        /// <inheritdoc />
        public long Seq
        {
            get => InsertionStamp?.Seq ?? Constants.UniversalSequenceNumber;
            set => InsertionStamp = EnsureInsertionStamp() with { Seq = value };
        }

        /// <inheritdoc />
        public long? LocalSeq
        {
            get => InsertionStamp?.LocalSeq;
            set => InsertionStamp = EnsureInsertionStamp() with { LocalSeq = value };
        }

        /// <inheritdoc />
        public int ClientId
        {
            get => InsertionStamp?.ClientId ?? Constants.NonCollabClient;
            set => InsertionStamp = EnsureInsertionStamp() with { ClientId = value };
        }

        /// <inheritdoc />
        public long? RemovedSeq
        {
            get => PrimarySetRemoveStamp?.Seq;
            set
            {
                if (value is long seq)
                {
                    int index = GetOrCreateSetRemoveStampIndex();
                    RemoveStamps[index] = ((SetRemoveOperationStamp)RemoveStamps[index]) with { Seq = seq };
                    Stamps.SortList(RemoveStamps);
                }
                else
                {
                    RemoveStamps.RemoveAll(stamp => stamp is SetRemoveOperationStamp);
                }
            }
        }

        /// <inheritdoc />
        public long? RemovedLocalSeq
        {
            get => PrimarySetRemoveStamp?.LocalSeq;
            set => UpdatePrimarySetRemoveStampIfPresent(stamp => stamp with { LocalSeq = value });
        }

        /// <inheritdoc />
        public int? RemovedClientId
        {
            get => PrimarySetRemoveStamp?.ClientId;
            set
            {
                if (value is int clientId)
                {
                    UpdatePrimarySetRemoveStampIfPresent(stamp => stamp with { ClientId = clientId });
                }
            }
        }

        /// <inheritdoc />
        public long? ObliteratedSeq
        {
            get => PrimarySliceRemoveStamp?.Seq;
            set
            {
                if (value is long seq)
                {
                    int index = GetOrCreateSliceRemoveStampIndex();
                    RemoveStamps[index] = ((SliceRemoveOperationStamp)RemoveStamps[index]) with { Seq = seq };
                    Stamps.SortList(RemoveStamps);
                }
                else
                {
                    RemoveStamps.RemoveAll(stamp => stamp is SliceRemoveOperationStamp);
                }
            }
        }

        /// <inheritdoc />
        public long? ObliteratedLocalSeq
        {
            get => PrimarySliceRemoveStamp?.LocalSeq;
            set => UpdatePrimarySliceRemoveStampIfPresent(stamp => stamp with { LocalSeq = value });
        }

        /// <inheritdoc />
        public int? ObliteratedClientId
        {
            get => PrimarySliceRemoveStamp?.ClientId;
            set
            {
                if (value is int clientId)
                {
                    UpdatePrimarySliceRemoveStampIfPresent(stamp => stamp with { ClientId = clientId });
                }
            }
        }

        /// <inheritdoc />
        public Side? ObliterateStartSide
        {
            get => PrimarySliceRemoveStamp?.StartSide;
            set
            {
                UpdatePrimarySliceRemoveStampIfPresent(stamp => stamp with { StartSide = value }, sort: false);
            }
        }

        /// <inheritdoc />
        public Side? ObliterateEndSide
        {
            get => PrimarySliceRemoveStamp?.EndSide;
            set
            {
                UpdatePrimarySliceRemoveStampIfPresent(stamp => stamp with { EndSide = value }, sort: false);
            }
        }

        /// <inheritdoc />
        public Dictionary<long, PropertySet>? PendingAnnotates { get; set; }

        /// <inheritdoc />
        public abstract string Type { get; }

        /// <inheritdoc />
        public OperationStamp? InsertionStamp { get; set; }

        /// <inheritdoc />
        public List<RemoveOperationStamp> RemoveStamps { get; } = new();

        List<LocalReferencePosition> ISegment.LocalRefs => _localRefs;

        /// <summary>
        /// Initializes a new instance of the <see cref="BaseSegment" /> class.
        /// </summary>
        /// <param name="properties">The optional properties attached to this segment.</param>
        protected BaseSegment(PropertySet? properties = null)
        {
            Properties = PropertyMap.ClonePropertySet(properties);
        }

        /// <summary>
        /// Returns whether this segment has a property with the supplied key.
        /// </summary>
        /// <param name="key">The property key to test.</param>
        /// <returns><see langword="true" /> if the property is present; otherwise, <see langword="false" />.</returns>
        public bool HasProperty(string key)
        {
            return Properties is System.Collections.Generic.IDictionary<string, object?> properties
                && properties.ContainsKey(key);
        }

        /// <inheritdoc />
        public virtual bool IsLeaf()
        {
            return true;
        }

        /// <inheritdoc />
        public virtual bool CanAppend(ISegment segment)
        {
            return false;
        }

        /// <inheritdoc />
        public virtual void Append(ISegment segment)
        {
            int offset = CachedLength;
            CachedLength += segment.CachedLength;
            MoveAppendedLocalReferences(segment, offset);
        }

        /// <inheritdoc />
        public virtual ISegment? SplitAt(int pos)
        {
            if (pos <= 0)
            {
                return null;
            }

            BaseSegment? splitSegment = CreateSplitSegmentAt(pos);
            if (splitSegment is null)
            {
                return null;
            }

            CopyMetadataTo(splitSegment);
            splitSegment.Parent = Parent;
            splitSegment.Index = Index + 1;
            splitSegment.Ordinal = string.Concat(Ordinal, '\0');
            MoveLocalReferencesToSplitSegment(splitSegment, pos);

            return splitSegment;
        }

        /// <inheritdoc />
        public abstract ISegment Clone();

        /// <inheritdoc />
        public abstract object ToJSONObject();

        /// <summary>
        /// Creates the right-hand segment for a split at the supplied position.
        /// </summary>
        /// <param name="pos">The position at which to split.</param>
        /// <returns>The new right-hand segment, or <see langword="null" /> if no split occurred.</returns>
        protected abstract BaseSegment? CreateSplitSegmentAt(int pos);

        /// <summary>
        /// Copies tree and sequencing metadata to another segment.
        /// </summary>
        /// <param name="segment">The target segment.</param>
        protected void CopyMetadataTo(ISegment segment)
        {
            segment.Parent = Parent;
            segment.Index = Index;
            segment.Ordinal = Ordinal;
            segment.Properties = PropertyMap.ClonePropertySet(Properties);
            segment.InsertionStamp = CloneOperationStamp(InsertionStamp);
            segment.RemoveStamps.Clear();
            foreach (RemoveOperationStamp removeStamp in RemoveStamps)
            {
                segment.RemoveStamps.Add(CloneRemoveStamp(removeStamp));
            }

            segment.PendingAnnotates = ClonePendingAnnotates(PendingAnnotates);
        }

        private OperationStamp EnsureInsertionStamp()
        {
            InsertionStamp ??= new InsertOperationStamp()
            {
                Seq = Constants.UniversalSequenceNumber,
                ClientId = Constants.NonCollabClient,
            };
            return InsertionStamp;
        }

        private SetRemoveOperationStamp? PrimarySetRemoveStamp
        {
            get
            {
                int index = PrimarySetRemoveStampIndex;
                return index >= 0 ? (SetRemoveOperationStamp)RemoveStamps[index] : null;
            }
        }

        private int PrimarySetRemoveStampIndex
        {
            get
            {
                for (int i = 0; i < RemoveStamps.Count; i++)
                {
                    if (RemoveStamps[i] is SetRemoveOperationStamp)
                    {
                        return i;
                    }
                }

                return -1;
            }
        }

        private SliceRemoveOperationStamp? PrimarySliceRemoveStamp
        {
            get
            {
                int index = PrimarySliceRemoveStampIndex;
                return index >= 0 ? (SliceRemoveOperationStamp)RemoveStamps[index] : null;
            }
        }

        private int PrimarySliceRemoveStampIndex
        {
            get
            {
                for (int i = 0; i < RemoveStamps.Count; i++)
                {
                    if (RemoveStamps[i] is SliceRemoveOperationStamp)
                    {
                        return i;
                    }
                }

                return -1;
            }
        }

        private void UpdatePrimarySetRemoveStampIfPresent(
            Func<SetRemoveOperationStamp, SetRemoveOperationStamp> update,
            bool sort = true)
        {
            int index = PrimarySetRemoveStampIndex;
            if (index >= 0)
            {
                RemoveStamps[index] = update((SetRemoveOperationStamp)RemoveStamps[index]);
                if (sort)
                {
                    Stamps.SortList(RemoveStamps);
                }
            }
        }

        private void UpdatePrimarySliceRemoveStampIfPresent(
            Func<SliceRemoveOperationStamp, SliceRemoveOperationStamp> update,
            bool sort = true)
        {
            int index = PrimarySliceRemoveStampIndex;
            if (index >= 0)
            {
                RemoveStamps[index] = update((SliceRemoveOperationStamp)RemoveStamps[index]);
                if (sort)
                {
                    Stamps.SortList(RemoveStamps);
                }
            }
        }

        private int GetOrCreateSetRemoveStampIndex()
        {
            int index = PrimarySetRemoveStampIndex;
            if (index >= 0)
            {
                return index;
            }

            SetRemoveOperationStamp stamp = CreateSetRemoveStamp();
            Stamps.SpliceIntoList(RemoveStamps, stamp);
            return IndexOfReference(RemoveStamps, stamp);
        }

        private int GetOrCreateSliceRemoveStampIndex()
        {
            int index = PrimarySliceRemoveStampIndex;
            if (index >= 0)
            {
                return index;
            }

            SliceRemoveOperationStamp stamp = CreateSliceRemoveStamp();
            Stamps.SpliceIntoList(RemoveStamps, stamp);
            return IndexOfReference(RemoveStamps, stamp);
        }

        private static int IndexOfReference<T>(IReadOnlyList<T> list, T value)
            where T : class
        {
            for (int i = 0; i < list.Count; i++)
            {
                if (ReferenceEquals(list[i], value))
                {
                    return i;
                }
            }

            throw new InvalidOperationException("Inserted stamp was not found in the segment stamp list.");
        }

        private SetRemoveOperationStamp CreateSetRemoveStamp()
        {
            return new SetRemoveOperationStamp()
            {
                Seq = Constants.UnassignedSequenceNumber,
                ClientId = ClientId,
            };
        }

        private SliceRemoveOperationStamp CreateSliceRemoveStamp()
        {
            return new SliceRemoveOperationStamp()
            {
                Seq = Constants.UnassignedSequenceNumber,
                ClientId = ClientId,
            };
        }

        private static OperationStamp? CloneOperationStamp(OperationStamp? stamp)
        {
            if (stamp is null)
            {
                return null;
            }

            return stamp is InsertOperationStamp
                ? new InsertOperationStamp()
                {
                    Seq = stamp.Seq,
                    ClientId = stamp.ClientId,
                    LocalSeq = stamp.LocalSeq,
                }
                : new OperationStamp()
                {
                    Seq = stamp.Seq,
                    ClientId = stamp.ClientId,
                    LocalSeq = stamp.LocalSeq,
                };
        }

        private static RemoveOperationStamp CloneRemoveStamp(RemoveOperationStamp stamp)
        {
            if (stamp is SliceRemoveOperationStamp sliceRemoveStamp)
            {
                return new SliceRemoveOperationStamp()
                {
                    Seq = sliceRemoveStamp.Seq,
                    ClientId = sliceRemoveStamp.ClientId,
                    LocalSeq = sliceRemoveStamp.LocalSeq,
                    StartSide = sliceRemoveStamp.StartSide,
                    EndSide = sliceRemoveStamp.EndSide,
                };
            }

            return new SetRemoveOperationStamp()
            {
                Seq = stamp.Seq,
                ClientId = stamp.ClientId,
                LocalSeq = stamp.LocalSeq,
            };
        }

        private static Dictionary<long, PropertySet>? ClonePendingAnnotates(Dictionary<long, PropertySet>? pendingAnnotates)
        {
            if (pendingAnnotates is null)
            {
                return null;
            }

            Dictionary<long, PropertySet> clone = new();
            foreach (KeyValuePair<long, PropertySet> pendingAnnotate in pendingAnnotates)
            {
                clone[pendingAnnotate.Key] = PropertyMap.ClonePropertySet(pendingAnnotate.Value) ?? new PropertySet();
            }

            return clone;
        }

        private void MoveLocalReferencesToSplitSegment(ISegment splitSegment, int splitOffset)
        {
            List<LocalReferencePosition> rightRefs = splitSegment.LocalRefs;
            for (int i = _localRefs.Count - 1; i >= 0; i--)
            {
                LocalReferencePosition reference = _localRefs[i];
                bool moveRight = reference.Offset > splitOffset
                    || (reference.Offset == splitOffset
                        && reference.SlidingPreference != SlidingPreference.Backward);
                if (!moveRight)
                {
                    continue;
                }

                _localRefs.RemoveAt(i);
                reference.Segment = splitSegment;
                reference.Offset -= splitOffset;
                rightRefs.Add(reference);
            }
        }

        private void MoveAppendedLocalReferences(ISegment appendedSegment, int offset)
        {
            List<LocalReferencePosition> appendedRefs = appendedSegment.LocalRefs;
            if (appendedRefs.Count == 0)
            {
                return;
            }

            foreach (LocalReferencePosition reference in appendedRefs)
            {
                reference.Segment = this;
                reference.Offset += offset;
                _localRefs.Add(reference);
            }

            appendedRefs.Clear();
        }
    }
}

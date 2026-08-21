// -----------------------------------------------------------------------------
// Ported from packages/dds/merge-tree/src/ops.ts
// Scope: Insert + Remove + Annotate + basic Obliterate + Group + interval envelopes.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid.MergeTree
{
    /// <summary>
    /// Numeric discriminator for merge-tree delta operations.
    /// </summary>
    public enum MergeTreeDeltaType
    {
        /// <summary>
        /// Insert operation. Wire value: 0.
        /// </summary>
        Insert = 0,

        /// <summary>
        /// Remove operation. Wire value: 1.
        /// </summary>
        Remove = 1,

        /// <summary>
        /// Annotate operation. Wire value: 2.
        /// </summary>
        Annotate = 2,

        /// <summary>
        /// Group operation. Wire value: 3.
        /// </summary>
        Group = 3,

        /// <summary>
        /// Obliterate operation. Wire value: 4.
        /// </summary>
        Obliterate = 4,

        /// <summary>
        /// Sided obliterate operation. Wire value: 5.
        /// </summary>
        ObliterateSided = 5,

        // Interval discriminators (10-13) are port-internal only. TS
        // routes interval ops through the intervalCollectionMap "act"
        // envelope. Kept for existing port dispatch code; new code should
        // prefer IntervalOpKind or the "act" wire shape.
        /// <summary>Port-internal: interval add. Not a TS wire discriminator.</summary>
        IntervalAdd = 10,

        /// <summary>Port-internal: interval delete. Not a TS wire discriminator.</summary>
        IntervalDelete = 11,

        /// <summary>Port-internal: interval change. Not a TS wire discriminator.</summary>
        IntervalChange = 12,

        /// <summary>Port-internal: interval property-changed. Not a TS wire discriminator.</summary>
        IntervalPropertyChanged = 13,
    }

    /// <summary>
    /// Placeholder marker for segment JSON forms that can appear in merge-tree ops.
    /// </summary>
    public interface IJSONSegment
    {
    }

    /// <summary>
    /// Base interface for all merge-tree ops.
    /// </summary>
    public interface IMergeTreeOp
    {
        /// <summary>
        /// Numeric discriminator of the operation type.
        /// </summary>
        MergeTreeDeltaType Type { get; }

        /// <summary>
        /// Local client sequence number stamped on an outbound op.
        /// </summary>
        long? ClientSeq { get; set; }
    }

    /// <summary>
    /// Insert operation contract.
    /// </summary>
    public interface IMergeTreeInsertMsg : IMergeTreeOp
    {
        /// <summary>
        /// Start position for the insert operation.
        /// </summary>
        int? Pos1 { get; set; }

        /// <summary>
        /// Relative start position for the insert operation.
        /// </summary>
        object? RelativePos1 { get; set; }

        /// <summary>
        /// End position for the insert operation.
        /// </summary>
        int? Pos2 { get; set; }

        /// <summary>
        /// Relative end position for the insert operation.
        /// </summary>
        object? RelativePos2 { get; set; }

        /// <summary>
        /// Segment payload to insert.
        /// </summary>
        object? Seg { get; set; }
    }

    /// <summary>
    /// Remove operation contract.
    /// </summary>
    public interface IMergeTreeRemoveMsg : IMergeTreeOp
    {
        /// <summary>
        /// Start position for the remove operation.
        /// </summary>
        int? Pos1 { get; set; }

        /// <summary>
        /// Relative start position for the remove operation.
        /// </summary>
        object? RelativePos1 { get; set; }

        /// <summary>
        /// End position for the remove operation.
        /// </summary>
        int? Pos2 { get; set; }

        /// <summary>
        /// Relative end position for the remove operation.
        /// </summary>
        object? RelativePos2 { get; set; }
    }

    /// <summary>
    /// Obliterate operation contract.
    /// </summary>
    public interface IMergeTreeObliterateMsg : IMergeTreeOp
    {
        /// <summary>
        /// Start position for the obliterate operation.
        /// </summary>
        int? Pos1 { get; set; }

        /// <summary>
        /// Relative start position for the obliterate operation. Basic C# parity does not support this yet.
        /// </summary>
        object? RelativePos1 { get; set; }

        /// <summary>
        /// End position for the obliterate operation.
        /// </summary>
        int? Pos2 { get; set; }

        /// <summary>
        /// Relative end position for the obliterate operation. Basic C# parity does not support this yet.
        /// </summary>
        object? RelativePos2 { get; set; }
    }

    /// <summary>
    /// Sided obliterate operation contract.
    /// </summary>
    public interface IMergeTreeObliterateSidedMsg : IMergeTreeOp
    {
        /// <summary>
        /// Start place for the sided obliterate operation.
        /// </summary>
        SequencePlace Pos1 { get; set; }

        /// <summary>
        /// Relative start position for the obliterate operation. Basic C# parity does not support this yet.
        /// </summary>
        object? RelativePos1 { get; set; }

        /// <summary>
        /// End place for the sided obliterate operation.
        /// </summary>
        SequencePlace Pos2 { get; set; }

        /// <summary>
        /// Relative end position for the obliterate operation. Basic C# parity does not support this yet.
        /// </summary>
        object? RelativePos2 { get; set; }
    }

    /// <summary>
    /// Group operation contract.
    /// </summary>
    public interface IMergeTreeGroupMsg : IMergeTreeOp
    {
        /// <summary>
        /// Ordered merge-tree delta operations in this group.
        /// </summary>
        List<MergeTreeOp> Ops { get; }
    }

    /// <summary>
    /// Base type for all merge-tree ops.
    /// </summary>
    public abstract class MergeTreeOp : IMergeTreeOp
    {
        /// <summary>
        /// Numeric discriminator of the operation type.
        /// </summary>
        public abstract MergeTreeDeltaType Type { get; }

        /// <inheritdoc />
        public long? ClientSeq { get; set; }
    }

    /// <summary>
    /// Base type for merge-tree delta messages.
    /// </summary>
    public abstract class MergeTreeDeltaMsg : MergeTreeOp
    {
    }

    /// <summary>
    /// Insert operation.
    /// </summary>
    public sealed class MergeTreeInsertMsg : MergeTreeDeltaMsg, IMergeTreeInsertMsg
    {
        /// <summary>
        /// Numeric discriminator of the operation type.
        /// </summary>
        public override MergeTreeDeltaType Type => MergeTreeDeltaType.Insert;

        /// <summary>
        /// Start position for the insert operation.
        /// </summary>
        public int? Pos1 { get; set; }

        /// <summary>
        /// Relative start position for the insert operation.
        /// </summary>
        public object? RelativePos1 { get; set; }

        /// <summary>
        /// End position for the insert operation.
        /// </summary>
        public int? Pos2 { get; set; }

        /// <summary>
        /// Relative end position for the insert operation.
        /// </summary>
        public object? RelativePos2 { get; set; }

        /// <summary>
        /// Segment payload to insert.
        /// </summary>
        public object? Seg { get; set; }
    }

    /// <summary>
    /// Remove operation.
    /// </summary>
    public sealed class MergeTreeRemoveMsg : MergeTreeDeltaMsg, IMergeTreeRemoveMsg
    {
        /// <summary>
        /// Numeric discriminator of the operation type.
        /// </summary>
        public override MergeTreeDeltaType Type => MergeTreeDeltaType.Remove;

        /// <summary>
        /// Start position for the remove operation.
        /// </summary>
        public int? Pos1 { get; set; }

        /// <summary>
        /// Relative start position for the remove operation.
        /// </summary>
        public object? RelativePos1 { get; set; }

        /// <summary>
        /// End position for the remove operation.
        /// </summary>
        public int? Pos2 { get; set; }

        /// <summary>
        /// Relative end position for the remove operation.
        /// </summary>
        public object? RelativePos2 { get; set; }
    }

    /// <summary>
    /// Basic obliterate operation.
    /// </summary>
    public sealed class MergeTreeObliterateMsg : MergeTreeDeltaMsg, IMergeTreeObliterateMsg
    {
        /// <summary>
        /// Numeric discriminator of the operation type.
        /// </summary>
        public override MergeTreeDeltaType Type => MergeTreeDeltaType.Obliterate;

        /// <summary>
        /// Start position for the obliterate operation.
        /// </summary>
        public int? Pos1 { get; set; }

        /// <summary>
        /// Relative start position for the obliterate operation. Basic C# parity does not support this yet.
        /// </summary>
        public object? RelativePos1 { get; set; }

        /// <summary>
        /// End position for the obliterate operation.
        /// </summary>
        public int? Pos2 { get; set; }

        /// <summary>
        /// Relative end position for the obliterate operation. Basic C# parity does not support this yet.
        /// </summary>
        public object? RelativePos2 { get; set; }
    }

    /// <summary>
    /// Sided obliterate operation.
    /// </summary>
    public sealed class MergeTreeObliterateSidedMsg : MergeTreeDeltaMsg, IMergeTreeObliterateSidedMsg
    {
        /// <summary>
        /// Numeric discriminator of the operation type.
        /// </summary>
        public override MergeTreeDeltaType Type => MergeTreeDeltaType.ObliterateSided;

        /// <summary>
        /// Start place for the sided obliterate operation.
        /// </summary>
        public SequencePlace Pos1 { get; set; } = new();

        /// <summary>
        /// Relative start position for the obliterate operation. Basic C# parity does not support this yet.
        /// </summary>
        public object? RelativePos1 { get; set; }

        /// <summary>
        /// End place for the sided obliterate operation.
        /// </summary>
        public SequencePlace Pos2 { get; set; } = new();

        /// <summary>
        /// Relative end position for the obliterate operation. Basic C# parity does not support this yet.
        /// </summary>
        public object? RelativePos2 { get; set; }
    }

    /// <summary>
    /// Op: apply properties to a range of the sequence.
    /// </summary>
    /// <remarks>
    /// Ported from packages/dds/merge-tree/src/ops.ts IMergeTreeAnnotateMsg.
    /// </remarks>
    public sealed class MergeTreeAnnotateMsg : MergeTreeDeltaMsg
    {
        /// <summary>
        /// Numeric discriminator of the operation type.
        /// </summary>
        public override MergeTreeDeltaType Type => MergeTreeDeltaType.Annotate;

        /// <summary>
        /// Inclusive start position of the range.
        /// </summary>
        public int? Pos1 { get; set; }

        /// <summary>
        /// Relative start position of the range.
        /// </summary>
        public object? RelativePos1 { get; set; }

        /// <summary>
        /// Exclusive end position of the range.
        /// </summary>
        public int? Pos2 { get; set; }

        /// <summary>
        /// Relative end position of the range.
        /// </summary>
        public object? RelativePos2 { get; set; }

        /// <summary>
        /// Properties to merge into segments in [Pos1, Pos2). A null value on a key
        /// means "delete this property from the segment".
        /// </summary>
        public PropertySet? Props { get; set; }
    }

    /// <summary>
    /// Op: applies multiple merge-tree delta operations under one envelope.
    /// </summary>
    public sealed class MergeTreeGroupMsg : MergeTreeOp, IMergeTreeGroupMsg
    {
        /// <summary>
        /// Numeric discriminator of the operation type.
        /// </summary>
        public override MergeTreeDeltaType Type => MergeTreeDeltaType.Group;

        /// <inheritdoc />
        public List<MergeTreeOp> Ops { get; } = new();
    }

    // -----------------------------------------------------------------------------
    // Interval-collection op types.
    // -----------------------------------------------------------------------------

    public enum IntervalOpKind
    {
        Add = 0,
        Delete = 1,
        Change = 2,
        PropertyChanged = 3,
    }

    /// <summary>
    /// Envelope for interval-collection operations.
    /// </summary>
    public abstract class IntervalOpMsg : MergeTreeOp
    {
        public abstract IntervalOpKind IntervalOpKind { get; }

        /// <inheritdoc />
        public override MergeTreeDeltaType Type
        {
            get
            {
                return IntervalOpKind switch
                {
                    IntervalOpKind.Add => MergeTreeDeltaType.IntervalAdd,
                    IntervalOpKind.Delete => MergeTreeDeltaType.IntervalDelete,
                    IntervalOpKind.Change => MergeTreeDeltaType.IntervalChange,
                    IntervalOpKind.PropertyChanged => MergeTreeDeltaType.IntervalPropertyChanged,
                    _ => throw new OcsException(OcsGateErrorCode.UnknownOp, $"Unknown interval op kind: {(int)IntervalOpKind}"),
                };
            }
        }

        public string CollectionName { get; set; } = string.Empty;

        public string IntervalId { get; set; } = string.Empty;
    }

    /// <summary>
    /// Wire-endpoint sentinel kind on interval add/change ops. TS encodes
    /// `"start"`/`"end"` string sentinels for endpoints that anchor to the
    /// synthetic start-of-tree / end-of-tree segment. See TS
    /// `packages/dds/sequence/src/intervals/sequenceInterval.ts` `createPositionReference`.
    /// </summary>
    public enum EndpointSentinel
    {
        None = 0,
        Start = 1,
        End = 2,
    }

    public sealed class IntervalAddOpMsg : IntervalOpMsg
    {
        public override IntervalOpKind IntervalOpKind => IntervalOpKind.Add;

        public int Start { get; set; }

        public int End { get; set; }

        // Set when the wire encoded the endpoint as the string sentinel
        // `"start"` or `"end"`. When non-None, the numeric Start/End field is
        // meaningless and the endpoint must be anchored via
        // MergeTree.CreateReferencePositionAtEndpoint.
        public EndpointSentinel StartSentinel { get; set; } = EndpointSentinel.None;

        public EndpointSentinel EndSentinel { get; set; } = EndpointSentinel.None;

        public Microsoft.Office.Web.Fluid.Intervals.IntervalType IntervalType { get; set; } =
            Microsoft.Office.Web.Fluid.Intervals.IntervalType.SlideOnRemove;

        public Microsoft.Office.Web.Fluid.Intervals.IntervalStickiness Stickiness { get; set; } =
            Microsoft.Office.Web.Fluid.Intervals.IntervalStickiness.End;

        public Microsoft.Office.Web.Fluid.Intervals.Side StartSide { get; set; } =
            Microsoft.Office.Web.Fluid.Intervals.Side.Before;

        public Microsoft.Office.Web.Fluid.Intervals.Side EndSide { get; set; } =
            Microsoft.Office.Web.Fluid.Intervals.Side.Before;

        public PropertySet? Props { get; set; }
    }

    public sealed class IntervalDeleteOpMsg : IntervalOpMsg
    {
        public override IntervalOpKind IntervalOpKind => IntervalOpKind.Delete;
    }

    public sealed class IntervalChangeOpMsg : IntervalOpMsg
    {
        public override IntervalOpKind IntervalOpKind => IntervalOpKind.Change;

        public int? Start { get; set; }

        public int? End { get; set; }

        // See IntervalAddOpMsg.StartSentinel/EndSentinel.
        public EndpointSentinel StartSentinel { get; set; } = EndpointSentinel.None;

        public EndpointSentinel EndSentinel { get; set; } = EndpointSentinel.None;

        public Microsoft.Office.Web.Fluid.Intervals.IntervalStickiness? Stickiness { get; set; }

        public Microsoft.Office.Web.Fluid.Intervals.Side? StartSide { get; set; }

        public Microsoft.Office.Web.Fluid.Intervals.Side? EndSide { get; set; }

        // TS ref: packages/dds/sequence/src/intervalCollection.ts changeInterval —
        // TS emits a single op carrying both endpoint delta and property delta.
        // Nullable so property-only or endpoint-only wire shapes remain valid.
        public PropertySet? Props { get; set; }
    }

    public sealed class IntervalPropertyChangedOpMsg : IntervalOpMsg
    {
        public override IntervalOpKind IntervalOpKind => IntervalOpKind.PropertyChanged;

        public PropertySet Props { get; set; } = new();
    }

    /// <summary>
    /// Helpers for constructing merge-tree operations using the same centralized shapes as TS opBuilder.ts.
    /// </summary>
    public static class OpBuilder
    {
        public static MergeTreeAnnotateMsg? CreateAnnotateMarkerOp(Marker marker, PropertySet props)
        {
            string? id = marker.GetId();
            if (string.IsNullOrEmpty(id))
            {
                return null;
            }

            return new MergeTreeAnnotateMsg()
            {
                Props = PropertyMap.ClonePropertySet(props),
                RelativePos1 = new PropertySet()
                {
                    ["id"] = id,
                    ["before"] = true,
                },
                RelativePos2 = new PropertySet()
                {
                    ["id"] = id,
                },
            };
        }

        public static MergeTreeAnnotateMsg CreateAnnotateRangeOp(int start, int end, PropertySet props)
        {
            return new MergeTreeAnnotateMsg()
            {
                Pos1 = start,
                Pos2 = end,
                Props = PropertyMap.ClonePropertySet(props),
            };
        }

        public static MergeTreeRemoveMsg CreateRemoveRangeOp(int start, int end)
        {
            return new MergeTreeRemoveMsg()
            {
                Pos1 = start,
                Pos2 = end,
            };
        }

        public static MergeTreeObliterateMsg CreateObliterateRangeOp(int start, int end)
        {
            return new MergeTreeObliterateMsg()
            {
                Pos1 = start,
                Pos2 = end,
            };
        }

        public static MergeTreeObliterateSidedMsg CreateObliterateRangeOpSided(SequencePlace start, SequencePlace end)
        {
            return new MergeTreeObliterateSidedMsg()
            {
                Pos1 = start.Clone(),
                Pos2 = end.Clone(),
            };
        }

        public static MergeTreeObliterateSidedMsg CreateObliterateRangeOpSided(int start, int end)
        {
            return CreateObliterateRangeOpSided(
                SequencePlace.At(start, Side.Before),
                SequencePlace.At(end - 1, Side.After));
        }

        public static MergeTreeInsertMsg CreateInsertSegmentOp(int position, ISegment segment)
        {
            return CreateInsertOp(position, segment.ToJSONObject());
        }

        public static MergeTreeInsertMsg CreateInsertOp(int position, object? segmentSpec)
        {
            return new MergeTreeInsertMsg()
            {
                Pos1 = position,
                Seg = segmentSpec,
            };
        }

        public static MergeTreeGroupMsg CreateGroupOp(params MergeTreeOp[] ops)
        {
            MergeTreeGroupMsg group = new();
            group.Ops.AddRange(ops);
            return group;
        }
    }

    /// <summary>
    /// Envelope carrying the merge-tree op corresponding to a delta.
    /// </summary>
    public sealed class MergeTreeDeltaOpArgs
    {
        /// <summary>
        /// Merge-tree operation corresponding to the delta.
        /// </summary>
        public MergeTreeOp? Op { get; set; }
    }

    // TODO: Annotate-adjust,
    // ReferenceType, and typed relative/reference position helpers are deferred for the port.
}

// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervals/{sequenceInterval,intervalUtils}.ts
// (subset for POC — Simple + SlideOnRemove + Transient intervals).
// Part of the SharedString C# feasibility port — Wave 12b.
// -----------------------------------------------------------------------------

#nullable enable

using System;

namespace Microsoft.Office.Web.Fluid.Intervals
{
    /// <summary>
    /// Which side of a segment an endpoint attaches to.
    /// Matches TS packages/dds/merge-tree/src/sequencePlace.ts Side.
    /// </summary>
    public enum Side
    {
        /// <summary>
        /// The endpoint attaches before the segment position.
        /// </summary>
        Before,

        /// <summary>
        /// The endpoint attaches after the segment position.
        /// </summary>
        After,
    }

    /// <summary>
    /// Discriminator for interval types.
    /// Matches TS packages/dds/sequence/src/intervals/intervalUtils.ts IntervalType flags.
    /// </summary>
    [Flags]
    public enum IntervalType
    {
        /// <summary>
        /// A simple interval.
        /// </summary>
        Simple = 0x0,

        /// <summary>
        /// A nested interval flag retained for wire compatibility.
        /// </summary>
        Nest = 0x1,

        /// <summary>
        /// Endpoints slide when referenced segments are removed.
        /// </summary>
        SlideOnRemove = 0x2,

        /// <summary>
        /// Endpoints detach when referenced segments are removed.
        /// </summary>
        Transient = 0x4,

        /// <summary>
        /// A cursor interval flag retained for wire compatibility.
        /// </summary>
        Cursor = 0x8,
    }

    /// <summary>
    /// Determines how an interval expands when segments are inserted adjacent to its endpoints.
    /// Matches TS packages/dds/sequence/src/intervals/intervalUtils.ts IntervalStickiness values.
    /// </summary>
    [Flags]
    public enum IntervalStickiness
    {
        /// <summary>
        /// Interval does not expand to include adjacent segments.
        /// </summary>
        None = 0b00,

        /// <summary>
        /// Interval expands to include segments inserted adjacent to the start.
        /// </summary>
        Start = 0b01,

        /// <summary>
        /// Interval expands to include segments inserted adjacent to the end.
        /// </summary>
        End = 0b10,

        /// <summary>
        /// Interval expands to include all adjacent inserted segments.
        /// </summary>
        Full = 0b11,
    }

    /// <summary>
    /// Utility helpers for SharedString intervals.
    /// </summary>
    public static class IntervalUtils
    {
        /// <summary>
        /// The side used by bare numeric sequence positions.
        /// </summary>
        public const Side DefaultSide = Side.Before;

        /// <summary>
        /// Normalizes a possibly absent endpoint position and side.
        /// </summary>
        /// <param name="position">The endpoint position, or <see langword="null" /> when absent.</param>
        /// <param name="side">The explicitly requested side, or <see langword="null" /> to use the default for numeric positions.</param>
        /// <returns>The normalized endpoint position and side.</returns>
        public static (int? Position, Side? Side) EndpointPosAndSide(int? position, Side? side = null)
        {
            if (position is null)
            {
                return (null, null);
            }

            return (position, side ?? DefaultSide);
        }

        /// <summary>
        /// Normalizes a possibly absent endpoint position and side using the TypeScript helper name.
        /// </summary>
        /// <param name="position">The endpoint position, or <see langword="null" /> when absent.</param>
        /// <param name="side">The explicitly requested side, or <see langword="null" /> to use the default for numeric positions.</param>
        /// <returns>The normalized endpoint position and side.</returns>
        public static (int? Position, Side? Side) endpointPosAndSide(int? position, Side? side = null)
        {
            return EndpointPosAndSide(position, side);
        }

        /// <summary>
        /// Computes the start and end sides that represent the supplied interval stickiness.
        /// </summary>
        /// <param name="stickiness">The interval stickiness.</param>
        /// <returns>The start and end endpoint sides.</returns>
        public static (Side StartSide, Side EndSide) SidesFromStickiness(IntervalStickiness stickiness)
        {
            Side startSide = (stickiness & IntervalStickiness.Start) == 0 ? Side.Before : Side.After;
            Side endSide = (stickiness & IntervalStickiness.End) == 0 ? Side.After : Side.Before;
            return (startSide, endSide);
        }

        /// <summary>
        /// Computes interval stickiness from the endpoint sides.
        /// </summary>
        /// <param name="startSide">The side for the start endpoint.</param>
        /// <param name="endSide">The side for the end endpoint.</param>
        /// <returns>The interval stickiness represented by the endpoint sides.</returns>
        public static IntervalStickiness ComputeStickinessFromSide(Side startSide, Side endSide)
        {
            IntervalStickiness stickiness = IntervalStickiness.None;
            if (startSide == Side.After)
            {
                stickiness |= IntervalStickiness.Start;
            }

            if (endSide == Side.Before)
            {
                stickiness |= IntervalStickiness.End;
            }

            return stickiness;
        }

        /// <summary>
        /// Gets the sliding preference for an interval start endpoint.
        /// </summary>
        /// <param name="startSide">The side for the start endpoint.</param>
        /// <param name="endSide">The side for the end endpoint.</param>
        /// <returns>The start reference sliding preference.</returns>
        public static Microsoft.Office.Web.Fluid.MergeTree.SlidingPreference StartReferenceSlidingPreference(Side startSide, Side endSide)
        {
            IntervalStickiness stickiness = ComputeStickinessFromSide(startSide, endSide);
            return (stickiness & IntervalStickiness.Start) == 0
                ? Microsoft.Office.Web.Fluid.MergeTree.SlidingPreference.Forward
                : Microsoft.Office.Web.Fluid.MergeTree.SlidingPreference.Backward;
        }

        /// <summary>
        /// Gets the sliding preference for an interval end endpoint.
        /// </summary>
        /// <param name="startSide">The side for the start endpoint.</param>
        /// <param name="endSide">The side for the end endpoint.</param>
        /// <returns>The end reference sliding preference.</returns>
        public static Microsoft.Office.Web.Fluid.MergeTree.SlidingPreference EndReferenceSlidingPreference(Side startSide, Side endSide)
        {
            IntervalStickiness stickiness = ComputeStickinessFromSide(startSide, endSide);
            return (stickiness & IntervalStickiness.End) == 0
                ? Microsoft.Office.Web.Fluid.MergeTree.SlidingPreference.Backward
                : Microsoft.Office.Web.Fluid.MergeTree.SlidingPreference.Forward;
        }

        /// <summary>
        /// Compares endpoint positions, treating detached endpoints as sorting before attached endpoints.
        /// </summary>
        /// <param name="left">The first endpoint position.</param>
        /// <param name="right">The second endpoint position.</param>
        /// <returns>A standard comparison result.</returns>
        public static int ComparePositions(int? left, int? right)
        {
            if (left.HasValue && right.HasValue)
            {
                return left.Value.CompareTo(right.Value);
            }

            if (left.HasValue)
            {
                return 1;
            }

            if (right.HasValue)
            {
                return -1;
            }

            return 0;
        }

        /// <summary>
        /// Compares endpoint sides using the TypeScript SequenceInterval ordering.
        /// </summary>
        /// <param name="left">The first side.</param>
        /// <param name="right">The second side.</param>
        /// <returns>A standard comparison result.</returns>
        public static int CompareSides(Side left, Side right)
        {
            if (left == right)
            {
                return 0;
            }

            return left == Side.Before ? 1 : -1;
        }

        /// <summary>
        /// Determines whether two half-open ranges intersect.
        /// </summary>
        /// <param name="leftStart">The inclusive start of the first range.</param>
        /// <param name="leftEnd">The exclusive end of the first range.</param>
        /// <param name="rightStart">The inclusive start of the second range.</param>
        /// <param name="rightEnd">The exclusive end of the second range.</param>
        /// <returns><see langword="true" /> when the ranges have a non-empty intersection.</returns>
        public static bool RangesOverlap(int leftStart, int leftEnd, int rightStart, int rightEnd)
        {
            return leftStart < leftEnd && rightStart < rightEnd && leftStart < rightEnd && rightStart < leftEnd;
        }
    }
}

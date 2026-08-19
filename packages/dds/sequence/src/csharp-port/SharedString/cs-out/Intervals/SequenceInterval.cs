// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervals/{sequenceInterval,intervalUtils}.ts
// (subset for POC — Simple + SlideOnRemove + Transient intervals).
// Part of the SharedString C# feasibility port — Wave 12b.
// -----------------------------------------------------------------------------

#nullable enable

using System;

using MergeTree = Microsoft.Office.Web.Fluid.MergeTree;

namespace Microsoft.Office.Web.Fluid.Intervals
{
    /// <summary>
    /// Common interface. Both SequenceInterval and future custom interval types implement this.
    /// </summary>
    public interface IInterval
    {
        /// <summary>
        /// Gets the interval identifier, if any.
        /// </summary>
        string? Id { get; }

        /// <summary>
        /// Gets or sets the interval properties.
        /// </summary>
        MergeTree.PropertySet? Properties { get; set; }
    }

    /// <summary>
    /// An interval over positions in a SharedString. Endpoints slide with text edits via LocalReferencePosition.
    /// </summary>
    public sealed class SequenceInterval : IInterval
    {
        internal SequenceInterval(
            MergeTree.MergeTree tree,
            MergeTree.LocalReferencePosition start,
            MergeTree.LocalReferencePosition end,
            IntervalType intervalType,
            Side startSide,
            Side endSide,
            string? id = null,
            MergeTree.PropertySet? properties = null)
        {
            ArgumentNullException.ThrowIfNull(tree);
            ArgumentNullException.ThrowIfNull(start);
            ArgumentNullException.ThrowIfNull(end);

            MergeTree = tree;
            Start = start;
            End = end;
            IntervalType = intervalType;
            StartSide = startSide;
            EndSide = endSide;
            Id = id;
            Properties = Microsoft.Office.Web.Fluid.MergeTree.PropertyMap.ClonePropertySet(properties);
        }

        /// <summary>
        /// Gets the start endpoint reference.
        /// </summary>
        public MergeTree.LocalReferencePosition Start { get; internal set; }

        /// <summary>
        /// Gets the end endpoint reference.
        /// </summary>
        public MergeTree.LocalReferencePosition End { get; internal set; }

        /// <summary>
        /// Gets the interval identifier, if any.
        /// </summary>
        public string? Id { get; internal set; }

        /// <summary>
        /// Gets the interval type flags.
        /// </summary>
        public IntervalType IntervalType { get; init; }

        /// <summary>
        /// Gets the side for the start endpoint.
        /// </summary>
        public Side StartSide { get; internal set; }

        /// <summary>
        /// Gets the side for the end endpoint.
        /// </summary>
        public Side EndSide { get; internal set; }

        /// <summary>
        /// Gets the stickiness represented by the endpoint sides.
        /// </summary>
        public IntervalStickiness Stickiness => IntervalUtils.ComputeStickinessFromSide(StartSide, EndSide);

        /// <summary>
        /// Gets or sets the interval properties.
        /// </summary>
        public MergeTree.PropertySet? Properties { get; set; }

        internal MergeTree.MergeTree MergeTree { get; }

        /// <summary>
        /// Gets the current start position, or <see langword="null" /> when the start reference is detached.
        /// </summary>
        public int? StartPosition => MergeTree.GetPositionOfReference(Start);

        /// <summary>
        /// Gets the current end position, or <see langword="null" /> when the end reference is detached.
        /// </summary>
        public int? EndPosition => MergeTree.GetPositionOfReference(End);

        internal bool HasDetachedEndpoint => Start.IsDetached || End.IsDetached;

        /// <summary>
        /// Compares this interval to another interval by start position and then start side.
        /// </summary>
        /// <param name="other">The interval to compare with this interval.</param>
        /// <returns>A standard comparison result.</returns>
        public int CompareStart(SequenceInterval other)
        {
            ArgumentNullException.ThrowIfNull(other);

            // TS ref: packages/dds/sequence/src/intervals/sequenceInterval.ts compareStart —
            // TS compares only start position, then start side. The port
            // previously fell through to end-position comparison, producing
            // ordering where TS returns equality.
            int startComparison = IntervalUtils.ComparePositions(StartPosition, other.StartPosition);
            if (startComparison != 0)
            {
                return startComparison;
            }

            return IntervalUtils.CompareSides(StartSide, other.StartSide);
        }

        /// <summary>
        /// Compares this interval to another interval by end position and then end side.
        /// </summary>
        /// <param name="other">The interval to compare with this interval.</param>
        /// <returns>A standard comparison result.</returns>
        public int CompareEnd(SequenceInterval other)
        {
            ArgumentNullException.ThrowIfNull(other);

            // TS ref: packages/dds/sequence/src/intervals/sequenceInterval.ts compareEnd —
            // TS compares only end position, then end side. The port previously
            // fell through to start-position comparison, producing ordering
            // where TS returns equality.
            int endComparison = IntervalUtils.ComparePositions(EndPosition, other.EndPosition);
            if (endComparison != 0)
            {
                return endComparison;
            }

            return IntervalUtils.CompareSides(other.EndSide, EndSide);
        }

        internal void SetStart(MergeTree.LocalReferencePosition newStart, Side newStartSide)
        {
            ArgumentNullException.ThrowIfNull(newStart);
            Start = newStart;
            StartSide = newStartSide;
        }

        internal void SetStart(MergeTree.LocalReferencePosition newStart)
        {
            SetStart(newStart, StartSide);
        }

        internal void SetEnd(MergeTree.LocalReferencePosition newEnd, Side newEndSide)
        {
            ArgumentNullException.ThrowIfNull(newEnd);
            End = newEnd;
            EndSide = newEndSide;
        }

        internal void SetEnd(MergeTree.LocalReferencePosition newEnd)
        {
            SetEnd(newEnd, EndSide);
        }

        internal void SetProperties(MergeTree.PropertySet? props)
        {
            if (props is null)
            {
                return;
            }

            MergeTree.PropertySet merged = Microsoft.Office.Web.Fluid.MergeTree.PropertyMap.AddProperties(Properties, props);
            Properties = merged.Count == 0 ? null : merged;
        }

        /// <summary>
        /// Determines whether this interval intersects the half-open range [<paramref name="start" />, <paramref name="end" />).
        /// </summary>
        /// <param name="start">The inclusive start of the range to test.</param>
        /// <param name="end">The exclusive end of the range to test.</param>
        /// <returns><see langword="true" /> when the interval and range have a non-empty intersection.</returns>
        public bool Overlaps(int start, int end)
        {
            int? intervalStartPosition = StartPosition;
            int? intervalEndPosition = EndPosition;
            if (!intervalStartPosition.HasValue || !intervalEndPosition.HasValue)
            {
                return false;
            }

            int intervalStart = Math.Min(intervalStartPosition.Value, intervalEndPosition.Value);
            int intervalEnd = Math.Max(intervalStartPosition.Value, intervalEndPosition.Value);
            return IntervalUtils.RangesOverlap(intervalStart, intervalEnd, start, end);
        }

        /// <summary>
        /// Gets the current start and end positions for serialization.
        /// </summary>
        /// <returns>The current reference positions, with <see langword="null" /> for detached endpoints.</returns>
        public (int? start, int? end) GetReferencePositions()
        {
            return (StartPosition, EndPosition);
        }
    }
}

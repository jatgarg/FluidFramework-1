// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalIndex/overlappingIntervalsIndex.ts.
// Uses List/SortedDictionary-based implementations instead of the TS
// RedBlackTree-backed indexes.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;

namespace Microsoft.Office.Web.Fluid.Intervals
{
    public sealed class OverlappingIntervalsIndex : IIntervalIndex
    {
        private readonly List<SequenceInterval> _intervals = new();

        public void Add(SequenceInterval interval) =>
            IntervalIndexComparers.Put(_intervals, interval, IntervalIndexComparers.CompareByIntervalThenId);

        public void Remove(SequenceInterval interval) =>
            IntervalIndexComparers.Remove(_intervals, interval, IntervalIndexComparers.CompareByIntervalThenId);

        /// <summary>Returns intervals overlapping [start, end].</summary>
        public IEnumerable<SequenceInterval> FindOverlapping(int start, int end)
        {
            if (end < start || _intervals.Count == 0)
            {
                return Enumerable.Empty<SequenceInterval>();
            }

            // Materialize at call time to match TS's snapshot semantics
            // (intervalIndex/overlappingIntervalsIndex.ts).
            return _intervals.Where(interval => !interval.HasDetachedEndpoint && OverlapsInclusive(interval, start, end))
                .OrderBy(interval => interval, IntervalIndexComparers.IntervalComparer)
                .ToArray();
        }

        // SS-A04: TS's compareReferencePositions places references on
        // different segment/ordinal positions when the side differs, so the
        // boundary comparisons in `SequenceInterval.overlaps` implicitly
        // depend on start/end side. The port collapses each endpoint to a
        // single numeric position via StartPosition/EndPosition — so we
        // apply the side comparison explicitly at the boundaries. A numeric
        // query has both endpoints implicitly Side.Before (defaultSide),
        // matching TS's normalizePlace of a plain number.
        //
        // TS overlap:
        //   compareReferencePositions(iv.start, qEnd) <= 0
        //     && compareReferencePositions(iv.end, qStart) >= 0
        //
        // Same-position case with qEnd.side=Before / qStart.side=Before:
        //   iv.start <= qEnd iff iv.start.pos < qEnd.pos
        //                       || (equal AND iv.startSide == Before)
        //   iv.end >= qStart iff iv.end.pos >= qStart.pos
        //                       (any iv.endSide is >= Before at equal pos)
        private static bool OverlapsInclusive(SequenceInterval interval, int start, int end)
        {
            if (interval.StartPosition is not int intervalStart || interval.EndPosition is not int intervalEnd)
            {
                return false;
            }

            bool startNotAfterQueryEnd = intervalStart < end
                || (intervalStart == end && interval.StartSide == Side.Before);
            bool endNotBeforeQueryStart = intervalEnd >= start;
            return startNotAfterQueryEnd && endNotBeforeQueryStart;
        }
    }
}

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
        /// <summary>Returns intervals overlapping [start, end], inclusive on both bounds with side awareness.</summary>
        /// <remarks>
        /// Materializes at call time so the returned collection is a snapshot
        /// of the index at query time. Deferred LINQ over the mutable
        /// backing list would surface subsequent Add/Remove as
        /// InvalidOperationException on enumeration; a snapshot avoids that
        /// and matches TS's return-array contract.
        /// </remarks>
        public IEnumerable<SequenceInterval> FindOverlapping(int start, int end)
        {
            if (end < start || _intervals.Count == 0)
            {
                return Enumerable.Empty<SequenceInterval>();
            }

            return _intervals.Where(interval => !interval.HasDetachedEndpoint && OverlapsInclusive(interval, start, end))
                .OrderBy(interval => interval, IntervalIndexComparers.IntervalComparer)
                .ToArray();
        }

        // Side-aware boundary comparisons match TS's compareReferencePositions
        // semantics: at equal position, a start with Side.Before is <= the
        // query end, and any end is >= a Before-sided query start. The
        // port collapses each endpoint to a numeric position, so we apply
        // the side comparison explicitly at the boundaries.
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

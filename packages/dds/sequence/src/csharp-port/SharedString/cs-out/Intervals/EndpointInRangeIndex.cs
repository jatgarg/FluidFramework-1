// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalIndex/endpointInRangeIndex.ts.
// Uses List/SortedDictionary-based implementations instead of the TS
// RedBlackTree-backed indexes.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;

namespace Microsoft.Office.Web.Fluid.Intervals
{
    public sealed class EndpointInRangeIndex : IIntervalIndex
    {
        private readonly List<SequenceInterval> _intervals = new();

        public void Add(SequenceInterval interval) =>
            IntervalIndexComparers.Put(_intervals, interval, IntervalIndexComparers.CompareByEndThenId);

        public void Remove(SequenceInterval interval) =>
            IntervalIndexComparers.Remove(_intervals, interval, IntervalIndexComparers.CompareByEndThenId);

        /// <summary>Returns intervals whose end position is in [start, end].</summary>
        /// <remarks>
        /// Materializes at call time so the returned collection is a snapshot
        /// of the index at query time. Deferred LINQ over the mutable
        /// backing list would surface subsequent Add/Remove as
        /// InvalidOperationException on enumeration; a snapshot avoids that
        /// and matches TS's return-array contract.
        /// </remarks>
        public IEnumerable<SequenceInterval> FindEndpointsInRange(int start, int end)
        {
            if (start <= 0 || start > end || _intervals.Count == 0)
            {
                return Enumerable.Empty<SequenceInterval>();
            }

            // Side-aware boundary at the upper edge. Query bounds are numeric
            // so both use defaultSide (Before). An interval's end is "in
            // range" iff pos >= start (any side qualifies at pos==start; see
            // OverlappingIntervalsIndex.OverlapsInclusive for reasoning) AND
            // (pos < end OR pos == end with EndSide == Before).
            return _intervals.Where(
                interval => !interval.HasDetachedEndpoint
                    && interval.EndPosition is int position
                    && position >= start
                    && (position < end
                        || (position == end && interval.EndSide == Side.Before)))
                .OrderBy(interval => interval, IntervalIndexComparers.EndpointComparer)
                .ToArray();
        }
    }
}

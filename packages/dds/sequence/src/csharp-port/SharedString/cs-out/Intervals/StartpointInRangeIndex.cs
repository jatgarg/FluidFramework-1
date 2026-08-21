// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalIndex/startpointInRangeIndex.ts.
// Uses List/SortedDictionary-based implementations instead of the TS
// RedBlackTree-backed indexes.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;

namespace Microsoft.Office.Web.Fluid.Intervals
{
    public sealed class StartpointInRangeIndex : IIntervalIndex
    {
        private readonly List<SequenceInterval> _intervals = new();

        public void Add(SequenceInterval interval) =>
            IntervalIndexComparers.Put(_intervals, interval, IntervalIndexComparers.CompareByStartThenId);

        public void Remove(SequenceInterval interval) =>
            IntervalIndexComparers.Remove(_intervals, interval, IntervalIndexComparers.CompareByStartThenId);

        /// <summary>Returns intervals whose start position is in [start, end].</summary>
        public IEnumerable<SequenceInterval> FindStartpointsInRange(int start, int end)
        {
            if (start <= 0 || start > end || _intervals.Count == 0)
            {
                return Enumerable.Empty<SequenceInterval>();
            }

            // side-aware boundary. Query bounds are numeric so both
            // use defaultSide (Before). An interval's start is "in range" iff
            //   start >= query.start (pos >= start; any side >= Before at ==)
            //   AND start <= query.end (pos < end OR pos == end AND
            //                           interval.StartSide == Before)
            return _intervals.Where(
                interval => !interval.HasDetachedEndpoint
                    && interval.StartPosition is int position
                    && position >= start
                    && (position < end
                        || (position == end && interval.StartSide == Side.Before)))
                .OrderBy(interval => interval, IntervalIndexComparers.StartpointComparer)
                .ToArray();
        }
    }
}

// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalIndex/startpointInRangeIndex.ts (subset for POC).
// Part of the SharedString C# feasibility port — Wave 12b.
// POC uses List/SortedDictionary-based implementations instead of the TS
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

            // Materialize at call time to match TS's snapshot semantics
            // (intervalIndex/startpointInRangeIndex.ts).
            return _intervals.Where(
                interval => !interval.HasDetachedEndpoint
                    && interval.StartPosition is int position
                    && position >= start
                    && position <= end)
                .OrderBy(interval => interval, IntervalIndexComparers.StartpointComparer)
                .ToArray();
        }
    }
}

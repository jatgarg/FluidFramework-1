// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalIndex/endpointInRangeIndex.ts (subset for POC).
// Part of the SharedString C# feasibility port — Wave 12b.
// POC uses List/SortedDictionary-based implementations instead of the TS
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
        public IEnumerable<SequenceInterval> FindEndpointsInRange(int start, int end)
        {
            if (start <= 0 || start > end || _intervals.Count == 0)
            {
                return Enumerable.Empty<SequenceInterval>();
            }

            // TS ref: packages/dds/sequence/src/intervalIndex/endpointInRangeIndex.ts —
            // TS builds and returns a snapshot array at call time. Materialize
            // the LINQ query so callers get a stable snapshot and are safe
            // against subsequent index mutations while enumerating.
            return _intervals.Where(
                interval => !interval.HasDetachedEndpoint
                    && interval.EndPosition is int position
                    && position >= start
                    && position <= end)
                .OrderBy(interval => interval, IntervalIndexComparers.EndpointComparer)
                .ToArray();
        }
    }
}

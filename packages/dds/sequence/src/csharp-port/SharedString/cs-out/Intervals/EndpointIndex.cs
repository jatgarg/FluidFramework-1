// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalIndex/endpointIndex.ts.
// Uses List/SortedDictionary-based implementations instead of the TS
// RedBlackTree-backed indexes.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;

namespace Microsoft.Office.Web.Fluid.Intervals
{
    public sealed class EndpointIndex : IIntervalIndex
    {
        private readonly List<SequenceInterval> _intervals = new();

        public void Add(SequenceInterval interval) =>
            IntervalIndexComparers.Put(_intervals, interval, IntervalIndexComparers.CompareByEndThenId);

        public void Remove(SequenceInterval interval) =>
            IntervalIndexComparers.Remove(_intervals, interval, IntervalIndexComparers.CompareByEndThenId);

        // Side-aware boundary comparisons at same-position endpoints.
        // Numeric query positions are treated as Side.Before (defaultSide),
        // which orders before Side.After at the same numeric position.
        // See OverlappingIntervalsIndex.OverlapsInclusive for the detailed
        // reasoning.

        /// <summary>Returns the previous interval based on the given position number.</summary>
        public SequenceInterval? PreviousInterval(int position)
        {
            // Match intervals whose end is strictly before the query, plus
            // intervals whose end is at the same numeric position with
            // EndSide == Before (which sorts equal to the query's Before).
            return _intervals
                .Where(interval => !interval.HasDetachedEndpoint
                    && interval.EndPosition is int endPosition
                    && (endPosition < position
                        || (endPosition == position && interval.EndSide == Side.Before)))
                .OrderBy(interval => interval, IntervalIndexComparers.EndpointComparer)
                .LastOrDefault();
        }

        /// <summary>Returns the next interval based on the given position number.</summary>
        public SequenceInterval? NextInterval(int position)
        {
            // Match intervals whose end is at-or-after the query. At the
            // same numeric position, either EndSide qualifies: Before sorts
            // equal to the query, After sorts strictly after it — both
            // satisfy "at or after," so a plain numeric >= is sufficient.
            return _intervals
                .Where(interval => !interval.HasDetachedEndpoint
                    && interval.EndPosition is int endPosition
                    && endPosition >= position)
                .OrderBy(interval => interval, IntervalIndexComparers.EndpointComparer)
                .FirstOrDefault();
        }

        /// <summary>Returns intervals with either endpoint in [start, end].</summary>
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

            // Either endpoint must be in [start, end] with side awareness at
            // the upper bound (pos == end → side must be Before).
            return _intervals.Where(
                interval =>
                    !interval.HasDetachedEndpoint &&
                    ((interval.StartPosition is int startPosition
                        && startPosition >= start
                        && (startPosition < end
                            || (startPosition == end && interval.StartSide == Side.Before))) ||
                        (interval.EndPosition is int endPosition
                            && endPosition >= start
                            && (endPosition < end
                                || (endPosition == end && interval.EndSide == Side.Before)))))
                .OrderBy(interval => interval, IntervalIndexComparers.IntervalComparer)
                .ToArray();
        }
    }
}

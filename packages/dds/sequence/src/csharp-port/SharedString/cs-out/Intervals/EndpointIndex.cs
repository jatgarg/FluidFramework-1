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

        // SS-A04: side-aware boundary comparisons. Query positions are
        // numeric so implicitly use defaultSide (Before). See detailed
        // reasoning in OverlappingIntervalsIndex.OverlapsInclusive.

        /// <summary>Returns the previous interval based on the given position number.</summary>
        public SequenceInterval? PreviousInterval(int position)
        {
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
            // end >= query.pos (any side >= Before at ==) → numeric >= suffices.
            return _intervals
                .Where(interval => !interval.HasDetachedEndpoint
                    && interval.EndPosition is int endPosition
                    && endPosition >= position)
                .OrderBy(interval => interval, IntervalIndexComparers.EndpointComparer)
                .FirstOrDefault();
        }

        /// <summary>Returns intervals with either endpoint in [start, end].</summary>
        public IEnumerable<SequenceInterval> FindEndpointsInRange(int start, int end)
        {
            if (start <= 0 || start > end || _intervals.Count == 0)
            {
                return Enumerable.Empty<SequenceInterval>();
            }

            // Materialize at call time to match TS's snapshot semantics.
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

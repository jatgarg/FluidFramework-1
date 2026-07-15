// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalIndex/overlappingIntervalsIndex.ts (subset for POC).
// Part of the SharedString C# feasibility port — Wave 12b.
// POC uses List/SortedDictionary-based implementations instead of the TS
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

			return _intervals.Where(interval => !interval.HasDetachedEndpoint && OverlapsInclusive(interval, start, end))
				.OrderBy(interval => interval, IntervalIndexComparers.IntervalComparer);
		}

		private static bool OverlapsInclusive(SequenceInterval interval, int start, int end)
		{
			if (interval.StartPosition is not int intervalStart || interval.EndPosition is not int intervalEnd)
			{
				return false;
			}

			return intervalStart <= end && intervalEnd >= start;
		}
	}
}

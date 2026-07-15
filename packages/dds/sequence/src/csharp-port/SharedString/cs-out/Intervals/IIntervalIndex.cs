// -----------------------------------------------------------------------------
// Ported from packages/dds/sequence/src/intervalIndex/intervalIndex.ts (subset for POC).
// Part of the SharedString C# feasibility port — Wave 12b.
// POC uses List/SortedDictionary-based implementations instead of the TS
// RedBlackTree-backed indexes.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;

namespace Microsoft.Office.Web.Fluid.Intervals
{
	public interface IIntervalIndex
	{
		/// <summary>Add an interval to the index.</summary>
		void Add(SequenceInterval interval);

		/// <summary>Remove an interval from the index.</summary>
		void Remove(SequenceInterval interval);
	}

	internal static class IntervalIndexComparers
	{
		public static readonly IComparer<SequenceInterval> StartpointComparer =
			Comparer<SequenceInterval>.Create(CompareByStartThenId);

		public static readonly IComparer<SequenceInterval> EndpointComparer =
			Comparer<SequenceInterval>.Create(CompareByEndThenId);

		public static readonly IComparer<SequenceInterval> IntervalComparer =
			Comparer<SequenceInterval>.Create(CompareByIntervalThenId);

		public static int CompareByStartThenId(SequenceInterval left, SequenceInterval right)
		{
			int startComparison = CompareStarts(left, right);
			return startComparison != 0 ? startComparison : CompareIds(left, right);
		}

		public static int CompareByEndThenId(SequenceInterval left, SequenceInterval right)
		{
			int endComparison = CompareEnds(left, right);
			return endComparison != 0 ? endComparison : CompareIds(left, right);
		}

		public static int CompareByIntervalThenId(SequenceInterval left, SequenceInterval right)
		{
			int startComparison = CompareStarts(left, right);
			if (startComparison != 0)
			{
				return startComparison;
			}

			int endComparison = CompareEnds(left, right);
			return endComparison != 0 ? endComparison : CompareIds(left, right);
		}

		public static void Put(
			IList<SequenceInterval> intervals,
			SequenceInterval interval,
			Comparison<SequenceInterval> comparison)
		{
			ArgumentNullException.ThrowIfNull(intervals);
			ArgumentNullException.ThrowIfNull(interval);
			ArgumentNullException.ThrowIfNull(comparison);

			for (int i = 0; i < intervals.Count; i++)
			{
				if (comparison(interval, intervals[i]) == 0)
				{
					intervals[i] = interval;
					return;
				}
			}

			intervals.Add(interval);
		}

		public static void Remove(
			IList<SequenceInterval> intervals,
			SequenceInterval interval,
			Comparison<SequenceInterval> comparison)
		{
			ArgumentNullException.ThrowIfNull(intervals);
			ArgumentNullException.ThrowIfNull(interval);
			ArgumentNullException.ThrowIfNull(comparison);

			for (int i = 0; i < intervals.Count; i++)
			{
				if (object.ReferenceEquals(intervals[i], interval))
				{
					intervals.RemoveAt(i);
					return;
				}
			}

			for (int i = 0; i < intervals.Count; i++)
			{
				if (comparison(interval, intervals[i]) == 0)
				{
					intervals.RemoveAt(i);
					return;
				}
			}
		}

		private static int CompareStarts(SequenceInterval left, SequenceInterval right)
		{
			int startComparison = IntervalUtils.ComparePositions(left.StartPosition, right.StartPosition);
			return startComparison != 0
				? startComparison
				: IntervalUtils.CompareSides(left.StartSide, right.StartSide);
		}

		private static int CompareEnds(SequenceInterval left, SequenceInterval right)
		{
			int endComparison = IntervalUtils.ComparePositions(left.EndPosition, right.EndPosition);
			return endComparison != 0
				? endComparison
				: IntervalUtils.CompareSides(right.EndSide, left.EndSide);
		}

		private static int CompareIds(SequenceInterval left, SequenceInterval right)
		{
			if (left.Id is not null && right.Id is not null)
			{
				return StringComparer.Ordinal.Compare(left.Id, right.Id);
			}

			return 0;
		}
	}
}

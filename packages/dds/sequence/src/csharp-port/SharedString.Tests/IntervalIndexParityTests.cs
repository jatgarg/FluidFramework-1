// -----------------------------------------------------------------------------
// Interval index parity regressions for II-series audit findings.
// -----------------------------------------------------------------------------

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Web.Fluid.Intervals;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class IntervalIndexParityTests
	{
		[Fact]
		public void StartpointInRange_RejectsInvalidRangesAndIncludesBothBoundaries()
		{
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval first = collection.Add(1, 3, intervalId: "a");
			SequenceInterval second = collection.Add(2, 3, intervalId: "b");
			StartpointInRangeIndex index = new();
			index.Add(first);
			index.Add(second);

			Assert.Empty(index.FindStartpointsInRange(2, 1));
			Assert.Empty(index.FindStartpointsInRange(0, 2));
			Assert.Empty(index.FindStartpointsInRange(-1, 1));
			AssertIds(index.FindStartpointsInRange(1, 1), "a");
			AssertIds(index.FindStartpointsInRange(1, 2), "a", "b");
		}

		[Fact]
		public void StartpointInRange_OrdersSameStartByIdLikeTsComparator()
		{
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval b = collection.Add(2, 4, intervalId: "b");
			SequenceInterval a = collection.Add(2, 3, intervalId: "a");
			SequenceInterval c = collection.Add(3, 4, intervalId: "c");
			StartpointInRangeIndex index = new();
			index.Add(b);
			index.Add(a);
			index.Add(c);

			AssertIds(index.FindStartpointsInRange(2, 3), "a", "b", "c");
		}

		[Fact]
		public void EndpointInRange_RejectsInvalidRangesAndIncludesBothBoundaries()
		{
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval first = collection.Add(1, 1, intervalId: "a");
			SequenceInterval second = collection.Add(1, 3, intervalId: "b");
			EndpointInRangeIndex index = new();
			index.Add(first);
			index.Add(second);

			Assert.Empty(index.FindEndpointsInRange(2, 1));
			Assert.Empty(index.FindEndpointsInRange(0, 1));
			Assert.Empty(index.FindEndpointsInRange(-1, 1));
			AssertIds(index.FindEndpointsInRange(1, 1), "a");
			AssertIds(index.FindEndpointsInRange(1, 3), "a", "b");
		}

		[Fact]
		public void EndpointInRange_OrdersSameEndpointByIdLikeTsComparator()
		{
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval b = collection.Add(2, 4, intervalId: "b");
			SequenceInterval a = collection.Add(3, 4, intervalId: "a");
			SequenceInterval c = collection.Add(3, 5, intervalId: "c");
			EndpointInRangeIndex index = new();
			index.Add(b);
			index.Add(a);
			index.Add(c);

			AssertIds(index.FindEndpointsInRange(4, 5), "a", "b", "c");
		}

		[Fact]
		public void OverlappingIntervals_IncludesTouchingAndCollapsedIntervals()
		{
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval point = collection.Add(1, 1, intervalId: "point");
			SequenceInterval span = collection.Add(1, 3, intervalId: "span");
			SequenceInterval after = collection.Add(4, 6, intervalId: "after");
			OverlappingIntervalsIndex index = new();
			index.Add(point);
			index.Add(span);
			index.Add(after);

			AssertIds(index.FindOverlapping(1, 1), "point", "span");
			AssertIds(index.FindOverlapping(3, 4), "span", "after");
			Assert.Empty(index.FindOverlapping(6, 5));
		}

		[Fact]
		public void OverlappingIntervals_OrdersMatchesByStartEndThenId()
		{
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval b = collection.Add(1, 4, intervalId: "b");
			SequenceInterval a = collection.Add(1, 3, intervalId: "a");
			SequenceInterval c = collection.Add(2, 3, intervalId: "c");
			OverlappingIntervalsIndex index = new();
			index.Add(b);
			index.Add(a);
			index.Add(c);

			AssertIds(index.FindOverlapping(0, 10), "a", "b", "c");
		}

		[Fact]
		public void EndpointIndex_PreviousAndNextUseInclusiveEndpointOrdering()
		{
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval a = collection.Add(1, 1, intervalId: "a");
			SequenceInterval b = collection.Add(1, 3, intervalId: "b");
			EndpointIndex index = new();
			index.Add(a);
			index.Add(b);

			Assert.Same(a, index.PreviousInterval(1));
			Assert.Same(b, index.NextInterval(3));
			Assert.Null(index.PreviousInterval(0));
			Assert.Null(index.NextInterval(4));
			AssertIds(index.FindEndpointsInRange(1, 1), "a", "b");
		}

		[Fact]
		public void IdIntervalIndex_DuplicateAddAndRemoveFollowTsMapSemantics()
		{
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval first = collection.Add(1, 2, intervalId: "dup");
			SequenceInterval second = CloneWithId(collection.Add(3, 4, intervalId: "other"), "dup");
			IdIntervalIndex index = new();

			index.Add(first);
			index.Add(second);

			Assert.Same(second, index.GetIntervalById("dup"));
			index.Remove(first);
			Assert.Null(index.GetIntervalById("dup"));
		}

		[Fact]
		public void IdIntervalIndex_RejectsIntervalsWithoutIds()
		{
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval withoutId = CloneWithId(collection.Add(1, 2, intervalId: "source"), null);
			IdIntervalIndex index = new();

			Assert.Throws<ArgumentException>(() => index.Add(withoutId));
			Assert.Throws<ArgumentException>(() => index.Remove(withoutId));
		}

		private static IntervalCollection CreateCollectionWithText(string text)
		{
			SharedString sharedString = new();
			sharedString.InsertText(0, text);
			return sharedString.GetIntervalCollection("comments");
		}

		private static SequenceInterval CloneWithId(SequenceInterval source, string? id)
		{
			return new SequenceInterval(
				source.MergeTree,
				source.Start,
				source.End,
				source.IntervalType,
				source.StartSide,
				source.EndSide,
				id,
				source.Properties);
		}

		private static void AssertIds(IEnumerable<SequenceInterval> intervals, params string[] expectedIds)
		{
			Assert.Equal(expectedIds, intervals.Select(interval => interval.Id).ToArray());
		}
	}
}

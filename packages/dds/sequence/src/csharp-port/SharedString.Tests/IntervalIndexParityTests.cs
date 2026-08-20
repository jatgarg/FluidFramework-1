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

		[Fact]
		public void IdIntervalIndex_DetachedNormalInterval_IsStillAddressableById()
		{
			// TS ref: packages/dds/sequence/src/intervalIndex/idIntervalIndex.ts —
			// TS keeps detached intervals addressable via getIntervalById; the
			// interval reports -1 endpoints but stays reachable. The port
			// previously filtered detached intervals out of lookup and iteration.
			// Transient intervals still auto-remove on detach (a port-specific
			// interval type not present in TS).
			SharedString sharedString = new();
			sharedString.InsertText(0, "abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(2, 4, intervalId: "i1");

			sharedString.DeleteText(2, 3);

			// Regular sliding intervals detach (endpoints slide to -1) but remain
			// addressable — Transient behavior is not the SlideOnRemove default.
			Assert.NotNull(collection.GetIntervalById("i1"));
			Assert.Same(interval, collection.GetIntervalById("i1"));
		}

		[Fact]
		public void SequenceInterval_CompareStart_ReturnsEqualityWhenNamedEndpointMatches()
		{
			// TS ref: packages/dds/sequence/src/intervals/sequenceInterval.ts
			// compareStart / compareEnd — returns 0 as soon as the named
			// endpoint's position and side match.
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval same = collection.Add(2, 4, intervalId: "same");
			SequenceInterval sameStartDifferentEnd = collection.Add(2, 6, intervalId: "sameStart");

			Assert.Equal(0, same.CompareStart(sameStartDifferentEnd));
			Assert.Equal(0, sameStartDifferentEnd.CompareStart(same));

			SequenceInterval sameEndDifferentStart = collection.Add(3, 6, intervalId: "sameEnd");
			SequenceInterval endTwin = collection.Add(1, 6, intervalId: "sameEnd2");

			Assert.Equal(0, sameEndDifferentStart.CompareEnd(endTwin));
			Assert.Equal(0, endTwin.CompareEnd(sameEndDifferentStart));
		}

		[Fact]
		public void PropertyMap_MatchProperties_DeepComparesArrayValues()
		{
			// TS ref: packages/dds/merge-tree/src/properties.ts matchProperties
			// deep-compares array property values, so equal-content arrays
			// with different references compare equal.
			MergeTree.PropertySet a = new()
			{
				["tags"] = new List<object?> { "one", "two", "three" },
			};
			MergeTree.PropertySet b = new()
			{
				["tags"] = new List<object?> { "one", "two", "three" },
			};

			Assert.True(MergeTree.PropertyMap.MatchProperties(a, b));

			MergeTree.PropertySet c = new()
			{
				["tags"] = new List<object?> { "one", "two", "four" },
			};

			Assert.False(MergeTree.PropertyMap.MatchProperties(a, c));
		}

		[Fact]
		public void OverlappingIntervalsIndex_ResultsSurviveIndexMutationDuringEnumeration()
		{
			// TS ref: packages/dds/sequence/src/intervalIndex/overlappingIntervalsIndex.ts —
			// TS returns a snapshot array. Deferred LINQ views over a mutable
			// backing structure would surface subsequent adds/removes as
			// iteration mutations; materializing at call time makes the result
			// snapshot-stable.
			IntervalCollection collection = CreateCollectionWithText("abcdefghij");
			SequenceInterval a = collection.Add(1, 3, intervalId: "a");
			SequenceInterval b = collection.Add(4, 6, intervalId: "b");
			OverlappingIntervalsIndex index = new();
			index.Add(a);
			index.Add(b);

			IEnumerable<SequenceInterval> results = index.FindOverlapping(0, 10);
			// Mutate the index AFTER capturing the query but BEFORE enumerating.
			index.Add(collection.Add(7, 9, intervalId: "c"));

			SequenceInterval[] snapshot = results.ToArray();

			Assert.Equal(new[] { "a", "b" }, snapshot.Select(interval => interval.Id).ToArray());
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

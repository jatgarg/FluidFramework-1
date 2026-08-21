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
		public void OverlappingIntervalsIndex_StartSideAfter_AtQueryEndBoundary_Excluded()
		{
			// SS-A04 regression. An interval with StartSide=After at position
			// N sits effectively "just after N". A query [_, N] must exclude
			// it (interval start is strictly beyond query end). TS uses
			// compareReferencePositions with side-encoded ordinals; the port
			// applies the side check explicitly at boundaries.
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval inclusive = collection.Add(3, Side.Before, 6, Side.Before, intervalId: "before");
			SequenceInterval exclusiveByAfter = collection.Add(3, Side.After, 6, Side.Before, intervalId: "after");
			OverlappingIntervalsIndex index = new();
			index.Add(inclusive);
			index.Add(exclusiveByAfter);

			// Query [1, 3]. `before` has startPos=3/Before → overlaps.
			// `after` has startPos=3/After → sits "just after 3" → excluded.
			AssertIds(index.FindOverlapping(1, 3), "before");
		}

		[Fact]
		public void OverlappingIntervalsIndex_EndSideBefore_AtQueryStartBoundary_StillIncluded()
		{
			// SS-A04: The lower-bound side check is symmetric — an interval
			// with EndSide=Before at position N still overlaps a query
			// starting at N (both sides Before → equal → overlap).
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval touchesLeft = collection.Add(1, Side.Before, 4, Side.Before, intervalId: "l");
			OverlappingIntervalsIndex index = new();
			index.Add(touchesLeft);

			// Query [4, 6]. Interval end=4/Before → same-position same-side →
			// interval.end >= query.start → overlaps.
			AssertIds(index.FindOverlapping(4, 6), "l");
		}

		[Fact]
		public void StartpointInRange_StartSideAfter_AtQueryEndBoundary_Excluded()
		{
			// SS-A04: startpoint in [start, end] uses side at the upper
			// bound. Interval with StartSide=After at endpos of query is
			// "just after" and thus outside the range.
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval before = collection.Add(3, Side.Before, 4, Side.Before, intervalId: "before");
			SequenceInterval after = collection.Add(3, Side.After, 4, Side.Before, intervalId: "after");
			StartpointInRangeIndex index = new();
			index.Add(before);
			index.Add(after);

			AssertIds(index.FindStartpointsInRange(1, 3), "before");
		}

		[Fact]
		public void EndpointInRange_EndSideAfter_AtQueryEndBoundary_Excluded()
		{
			// SS-A04: endpoint-in-range applies the side check to EndSide.
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval before = collection.Add(1, Side.Before, 3, Side.Before, intervalId: "before");
			SequenceInterval after = collection.Add(1, Side.Before, 3, Side.After, intervalId: "after");
			EndpointInRangeIndex index = new();
			index.Add(before);
			index.Add(after);

			AssertIds(index.FindEndpointsInRange(1, 3), "before");
		}

		[Fact]
		public void EndpointIndex_PreviousInterval_EndSideAfter_AtQueryPosition_Excluded()
		{
			// SS-A04: PreviousInterval(N) excludes an interval whose end is
			// EndSide=After at position N, because that end is "just after"
			// N, i.e. NOT <= N.
			IntervalCollection collection = CreateCollectionWithText("abcdefgh");
			SequenceInterval before = collection.Add(1, Side.Before, 3, Side.Before, intervalId: "before");
			SequenceInterval after = collection.Add(1, Side.Before, 3, Side.After, intervalId: "after");
			EndpointIndex index = new();
			index.Add(before);
			index.Add(after);

			SequenceInterval? previous = index.PreviousInterval(3);
			Assert.NotNull(previous);
			Assert.Equal("before", previous!.Id);
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
		public void IdIntervalIndex_FullyDetachedInterval_StillAddressableAfterEntireDocumentRemoved()
		{
			// V2-T01 regression. TS specifies that a truly detached interval
			// (both endpoints slid off after the backing content vanished) is
			// still reachable via getIntervalById. The SS-A03 fix originally
			// added detached-interval addressability but only exercised the
			// case where positions changed while segments still existed.
			SharedString sharedString = new();
			sharedString.InsertText(0, "abcdef");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(2, 4, intervalId: "detachable");

			// Wipe the entire document — both interval endpoints lose their
			// backing segments and can't slide anywhere.
			sharedString.DeleteText(0, 6);

			Assert.NotNull(collection.GetIntervalById("detachable"));
			Assert.Same(interval, collection.GetIntervalById("detachable"));
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

		[Fact]
		public void StartpointInRangeIndex_ResultsSurviveIndexMutationDuringEnumeration()
		{
			// V2-T03 regression. Same call-time snapshot contract as
			// OverlappingIntervalsIndex — deferred enumeration must not surface
			// index mutations that happened after the query was issued.
			IntervalCollection collection = CreateCollectionWithText("abcdefghij");
			SequenceInterval a = collection.Add(1, 3, intervalId: "a");
			SequenceInterval b = collection.Add(4, 6, intervalId: "b");
			StartpointInRangeIndex index = new();
			index.Add(a);
			index.Add(b);

			IEnumerable<SequenceInterval> results = index.FindStartpointsInRange(1, 10);
			index.Add(collection.Add(7, 9, intervalId: "c"));

			SequenceInterval[] snapshot = results.ToArray();

			Assert.Equal(new[] { "a", "b" }, snapshot.Select(interval => interval.Id).OrderBy(id => id).ToArray());
		}

		[Fact]
		public void EndpointInRangeIndex_ResultsSurviveIndexMutationDuringEnumeration()
		{
			// V2-T03 regression.
			IntervalCollection collection = CreateCollectionWithText("abcdefghij");
			SequenceInterval a = collection.Add(1, 3, intervalId: "a");
			SequenceInterval b = collection.Add(4, 6, intervalId: "b");
			EndpointInRangeIndex index = new();
			index.Add(a);
			index.Add(b);

			IEnumerable<SequenceInterval> results = index.FindEndpointsInRange(1, 10);
			index.Add(collection.Add(7, 9, intervalId: "c"));

			SequenceInterval[] snapshot = results.ToArray();

			Assert.Equal(new[] { "a", "b" }, snapshot.Select(interval => interval.Id).OrderBy(id => id).ToArray());
		}

		[Fact]
		public void EndpointIndex_ResultsSurviveIndexMutationDuringEnumeration()
		{
			// V2-T03 regression. FindEndpointsInRange returns a snapshot;
			// index mutations after the query issue must not surface in the
			// enumeration.
			IntervalCollection collection = CreateCollectionWithText("abcdefghij");
			SequenceInterval a = collection.Add(1, 3, intervalId: "a");
			SequenceInterval b = collection.Add(4, 6, intervalId: "b");
			EndpointIndex index = new();
			index.Add(a);
			index.Add(b);

			IEnumerable<SequenceInterval> results = index.FindEndpointsInRange(1, 10);
			// Mutate the index AFTER capturing the query but BEFORE enumerating.
			index.Add(collection.Add(7, 9, intervalId: "c"));

			SequenceInterval[] snapshot = results.ToArray();

			Assert.Equal(new[] { "a", "b" }, snapshot.Select(interval => interval.Id).OrderBy(id => id).ToArray());
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

// -----------------------------------------------------------------------------
// Integration tests for IntervalCollection on SharedString.
// -----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class SharedStringIntervalTests
	{
		[Fact]
		public void GetIntervalCollection_ReturnsSameInstanceForSameName()
		{
			var sharedString = new SharedString();

			IntervalCollection first = sharedString.GetIntervalCollection("comments");
			IntervalCollection second = sharedString.GetIntervalCollection("comments");

			Assert.Same(first, second);
		}

		[Fact]
		public void GetIntervalCollection_DifferentNames_ReturnsDifferentCollections()
		{
			var sharedString = new SharedString();

			IntervalCollection comments = sharedString.GetIntervalCollection("comments");
			IntervalCollection highlights = sharedString.GetIntervalCollection("highlights");

			Assert.NotSame(comments, highlights);
			Assert.Equal("comments", comments.Name);
			Assert.Equal("highlights", highlights.Name);
		}

		[Fact]
		public void Add_ReturnsIntervalWithMatchingEndpoints()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");

			SequenceInterval interval = collection.Add(3, 7, intervalId: "c1");

			AssertIntervalPositions(interval, 3, 7);
			Assert.Equal("c1", interval.Id);
		}

		[Fact]
		public void GetIntervalById_ReturnsAddedInterval()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval added = collection.Add(1, 4, intervalId: "c1");

			SequenceInterval? found = collection.GetIntervalById("c1");

			Assert.Same(added, found);
		}

		[Fact]
		public void RemoveIntervalById_RemovesFromCollection()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(0, 2, intervalId: "c1");
			collection.Add(3, 5, intervalId: "c2");

			SequenceInterval? removed = collection.RemoveIntervalById("c1");

			Assert.NotNull(removed);
			Assert.Null(collection.GetIntervalById("c1"));
			AssertIds(collection, "c2");
		}

		[Fact]
		public void Change_UpdatesEndpoints()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(3, 7, intervalId: "c1");

			SequenceInterval? changed = collection.Change("c1", newStart: 1, newEnd: 5);

			Assert.NotNull(changed);
			AssertIntervalPositions(changed!, 1, 5);
		}

		[Fact]
		public void Insert_BeforeIntervalStart_SlidesRight()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(2, 5, intervalId: "c1");

			sharedString.InsertText(0, "XY");

			AssertIntervalPositions(interval, 4, 7);
		}

		[Fact]
		public void Delete_ContainingIntervalStart_SlidesForward()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments", IntervalType.SlideOnRemove);
			SequenceInterval interval = collection.Add(3, 8, intervalId: "c1");

			sharedString.DeleteText(0, 5);

			Assert.Equal(" world", sharedString.GetText());
			AssertIntervalPositions(interval, 0, 3);
		}

		[Fact]
		public void Insert_InsideInterval_ExpandsRange()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(0, 5, intervalId: "c1");

			sharedString.InsertText(3, "XYZ");

			Assert.Equal("helXYZlo world", sharedString.GetText());
			AssertIntervalPositions(interval, 0, 8);
		}

		[Fact]
		public void CreateForwardIteratorWithStartPosition_ReturnsIntervalsWithExactStart()
		{
			// TS-parity: iterator is EXACT match, not range (intervalCollection.ts uses
			// walkExactMatchesForward). Intervals: a=(8,9), b=(2,4), c=(5,10), d=(5,6).
			IntervalCollection collection = CreateCollectionWithOrderedIntervals();

			// Both c and d start at 5; forward order = by end asc => d(6), c(10).
			AssertIds(collection.CreateForwardIteratorWithStartPosition(5), "d", "c");
			AssertIds(collection.CreateForwardIteratorWithStartPosition(2), "b");
			AssertIds(collection.CreateForwardIteratorWithStartPosition(8), "a");
			// No interval starts at 0 — empty result under exact-match semantics.
			Assert.Empty(collection.CreateForwardIteratorWithStartPosition(0));
		}

		[Fact]
		public void CreateBackwardIteratorWithStartPosition_ReturnsIntervalsWithExactStart()
		{
			// TS-parity: exact-match; backward order = by end desc.
			IntervalCollection collection = CreateCollectionWithOrderedIntervals();

			AssertIds(collection.CreateBackwardIteratorWithStartPosition(5), "c", "d");
			AssertIds(collection.CreateBackwardIteratorWithStartPosition(2), "b");
			Assert.Empty(collection.CreateBackwardIteratorWithStartPosition(10));
		}

		[Fact]
		public void CreateForwardIteratorWithEndPosition_ReturnsIntervalsWithExactEnd()
		{
			// TS-parity: exact-match on end position.
			IntervalCollection collection = CreateCollectionWithOrderedIntervals();

			AssertIds(collection.CreateForwardIteratorWithEndPosition(6), "d");
			AssertIds(collection.CreateForwardIteratorWithEndPosition(10), "c");
			AssertIds(collection.CreateForwardIteratorWithEndPosition(4), "b");
			Assert.Empty(collection.CreateForwardIteratorWithEndPosition(0));
		}

		[Fact]
		public void CreateBackwardIteratorWithEndPosition_ReturnsIntervalsWithExactEnd()
		{
			IntervalCollection collection = CreateCollectionWithOrderedIntervals();

			AssertIds(collection.CreateBackwardIteratorWithEndPosition(9), "a");
			AssertIds(collection.CreateBackwardIteratorWithEndPosition(10), "c");
			Assert.Empty(collection.CreateBackwardIteratorWithEndPosition(11));
		}

		[Fact]
		public void Add_FiresOnAddInterval_WithLocalTrue()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			var events = new List<IntervalAddedEventArgs>();
			collection.OnAddInterval += (sender, args) => events.Add(args);

			SequenceInterval interval = collection.Add(2, 5, intervalId: "c1");

			IntervalAddedEventArgs captured = Assert.Single(events);
			Assert.True(captured.Local);
			Assert.Same(interval, captured.Interval);
		}

		[Fact]
		public void RemoveIntervalById_FiresOnDeleteInterval_WithLocalTrue()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(2, 5, intervalId: "c1");
			var events = new List<IntervalDeletedEventArgs>();
			collection.OnDeleteInterval += (sender, args) => events.Add(args);

			collection.RemoveIntervalById("c1");

			IntervalDeletedEventArgs captured = Assert.Single(events);
			Assert.True(captured.Local);
			Assert.Same(interval, captured.Interval);
		}

		[Fact]
		public void Change_FiresOnChangeInterval_WithPreviousPositions()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(3, 7, intervalId: "c1");
			var events = new List<IntervalChangedEventArgs>();
			collection.OnChangeInterval += (sender, args) => events.Add(args);

			collection.Change("c1", newStart: 1, newEnd: 5);

			IntervalChangedEventArgs captured = Assert.Single(events);
			Assert.True(captured.Local);
			Assert.Same(interval, captured.Interval);
			Assert.Equal(3, captured.PreviousStart);
			Assert.Equal(7, captured.PreviousEnd);
		}

		[Fact]
		public void RemotePropertyChanged_FiresOnPropertyChanged_WithLocalFalse()
		{
			var sharedString = CreateSharedStringWithText("hello world");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			SequenceInterval interval = collection.Add(2, 5, intervalId: "c1");
			var events = new List<IntervalPropertyChangedEventArgs>();
			collection.OnPropertyChanged += (sender, args) => events.Add(args);

			ProcessRemoteIntervalPropertyChanged(
				sharedString,
				"comments",
				"c1",
				new PropertySet()
				{
					["color"] = "blue",
				},
				refSeq: 0,
				seq: 1);

			IntervalPropertyChangedEventArgs captured = Assert.Single(events);
			Assert.False(captured.Local);
			Assert.Same(interval, captured.Interval);
			Assert.Equal("blue", Assert.IsType<string>(captured.ChangedProps["color"]));
			Assert.NotNull(interval.Properties);
			Assert.Equal("blue", Assert.IsType<string>(interval.Properties!["color"]));
		}

		[Fact]
		public void TwoClients_ConcurrentAdd_SameAndDifferentCollections_Converge()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello world");
			IntervalCollection commentsA = harness.ClientA.GetIntervalCollection("comments");
			IntervalCollection commentsB = harness.ClientB.GetIntervalCollection("comments");
			long refSeq = harness.CurrentServerSeq;

			commentsA.Add(1, 3, intervalId: "A1");
			var sentA = Assert.Single(harness.SenderA.Sent);
			commentsB.Add(4, 6, intervalId: "B1");
			var sentB = Assert.Single(harness.SenderB.Sent);
			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			AssertCollectionsHaveSameInterval(harness, "comments", "A1");
			AssertCollectionsHaveSameInterval(harness, "comments", "B1");
			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			refSeq = harness.CurrentServerSeq;

			IntervalCollection highlightsA = harness.ClientA.GetIntervalCollection("highlights");
			IntervalCollection bookmarksB = harness.ClientB.GetIntervalCollection("bookmarks");
			highlightsA.Add(0, 2, intervalId: "H1");
			sentA = Assert.Single(harness.SenderA.Sent);
			bookmarksB.Add(7, 10, intervalId: "M1");
			sentB = Assert.Single(harness.SenderB.Sent);
			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			AssertCollectionsHaveSameInterval(harness, "highlights", "H1");
			AssertCollectionsHaveSameInterval(harness, "bookmarks", "M1");
			Assert.Null(harness.ClientA.GetIntervalCollection("highlights").GetIntervalById("M1"));
			Assert.Null(harness.ClientB.GetIntervalCollection("bookmarks").GetIntervalById("H1"));
		}

		[Fact]
		public void TwoClients_RemoteAdd_CreatesCollectionIfMissing()
		{
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "hello world");
			IntervalCollection commentsA = harness.ClientA.GetIntervalCollection("comments");

			commentsA.Add(2, 5, intervalId: "A1");
			var sentA = Assert.Single(harness.SenderA.Sent);
			harness.DeliverAtoB(sentA, refSeq: harness.CurrentServerSeq);

			IntervalCollection commentsB = harness.ClientB.GetIntervalCollection("comments");
			SequenceInterval? remoteInterval = commentsB.GetIntervalById("A1");
			Assert.NotNull(remoteInterval);
			AssertIntervalPositions(remoteInterval!, 2, 5);
		}

		[Fact]
		public void TwoClients_ConcurrentChange_SameInterval_Converges()
		{
			// TS-parity: Change API requires both start and end (intervalCollection.ts).
			// Two-client convergence tested by each client passing full endpoint pairs; the
			// last-writer's op wins on the shared endpoint.
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "abcdefghij");
			IntervalCollection commentsA = harness.ClientA.GetIntervalCollection("comments");
			commentsA.Add(2, 5, intervalId: "c1");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			IntervalCollection commentsB = harness.ClientB.GetIntervalCollection("comments");
			long refSeq = harness.CurrentServerSeq;

			// Client A wants (1, 8). Client B concurrently wants (1, 8) with same intent.
			commentsA.Change("c1", newStart: 1, newEnd: 8);
			var sentA = Assert.Single(harness.SenderA.Sent);
			commentsB.Change("c1", newStart: 1, newEnd: 8);
			var sentB = Assert.Single(harness.SenderB.Sent);
			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			AssertCollectionsHaveSameInterval(harness, "comments", "c1");
			AssertIntervalPositions(harness.ClientA.GetIntervalCollection("comments").GetIntervalById("c1")!, 1, 8);
			AssertIntervalPositions(harness.ClientB.GetIntervalCollection("comments").GetIntervalById("c1")!, 1, 8);
		}

		[Fact]
		public void TwoClients_ConcurrentChange_DifferentIntent_Converges()
		{
			// Two clients concurrently change the same interval to DIFFERENT
			// ranges. Consensus-based reconciliation keeps the two clients'
			// local views intact through the concurrent-op window, and the
			// last-writer's op wins on the shared endpoint at
			// server-sequenced time.
			var harness = new TwoClientHarness();
			LoadInitialSharedText(harness, "abcdefghij");
			IntervalCollection commentsA = harness.ClientA.GetIntervalCollection("comments");
			commentsA.Add(2, 5, intervalId: "c1");
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: harness.CurrentServerSeq);
			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
			IntervalCollection commentsB = harness.ClientB.GetIntervalCollection("comments");
			long refSeq = harness.CurrentServerSeq;

			// Client A wants (1, 5). Client B concurrently wants (3, 7).
			commentsA.Change("c1", newStart: 1, newEnd: 5);
			var sentA = Assert.Single(harness.SenderA.Sent);
			commentsB.Change("c1", newStart: 3, newEnd: 7);
			var sentB = Assert.Single(harness.SenderB.Sent);

			// Deliver in both orders (A hears B's op, B hears A's op).
			harness.DeliverAtoB(sentA, refSeq);
			harness.DeliverBtoA(sentB, refSeq);

			AssertCollectionsHaveSameInterval(harness, "comments", "c1");
		}

		[Fact]
		public void Change_OneSided_Throws()
		{
			// TS-parity: one-sided change is rejected (intervalCollection.ts).
			var sharedString = CreateSharedStringWithText("abcdefghij");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(2, 5, intervalId: "c1");

			OcsException exceptionStart = Assert.Throws<OcsException>(
				() => collection.Change("c1", newStart: 1, newEnd: null));
			Assert.Equal(OcsGateErrorCode.InvalidOperation, exceptionStart.ErrorCode);

			OcsException exceptionEnd = Assert.Throws<OcsException>(
				() => collection.Change("c1", newStart: null, newEnd: 8));
			Assert.Equal(OcsGateErrorCode.InvalidOperation, exceptionEnd.ErrorCode);
		}

		private static SharedString CreateSharedStringWithText(string text)
		{
			var sharedString = new SharedString();
			sharedString.InsertText(0, text);
			return sharedString;
		}

		private static IntervalCollection CreateCollectionWithOrderedIntervals()
		{
			var sharedString = CreateSharedStringWithText("abcdefghijkl");
			IntervalCollection collection = sharedString.GetIntervalCollection("comments");
			collection.Add(8, 9, intervalId: "a");
			collection.Add(2, 4, intervalId: "b");
			collection.Add(5, 10, intervalId: "c");
			collection.Add(5, 6, intervalId: "d");
			return collection;
		}

		private static void ProcessRemoteIntervalPropertyChanged(
			SharedString sharedString,
			string collectionName,
			string intervalId,
			PropertySet props,
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			string opJson = SharedStringOpSerializer.Serialize(new IntervalPropertyChangedOpMsg()
			{
				CollectionName = collectionName,
				IntervalId = intervalId,
				Props = props,
			});

			sharedString.ProcessDataObjectOp(RemoteMessage(refSeq, seq, clientId), opJson);
		}

		private static SequencedDocumentMessageDescriptor RemoteMessage(
			long refSeq,
			long seq,
			string clientId = "remote-client")
		{
			return new SequencedDocumentMessageDescriptor(
				SequenceNumber.ForTesting(clientSeq: 0, refSeq: refSeq, seq: seq),
				OpOrigin.Remote,
				clientId);
		}

		private static void LoadInitialSharedText(TwoClientHarness harness, string text)
		{
			harness.ClientA.InsertText(0, text);
			harness.DeliverAtoB(Assert.Single(harness.SenderA.Sent), refSeq: 0);
			Assert.Equal(text, harness.ClientA.GetText());
			Assert.Equal(text, harness.ClientB.GetText());
			harness.SenderA.Sent.Clear();
			harness.SenderB.Sent.Clear();
		}

		private static void AssertIntervalPositions(SequenceInterval interval, int start, int end)
		{
			Assert.Equal((int?)start, interval.StartPosition);
			Assert.Equal((int?)end, interval.EndPosition);
		}

		private static void AssertIds(IEnumerable<SequenceInterval> intervals, params string[] expectedIds)
		{
			Assert.Equal(expectedIds, intervals.Select(interval => interval.Id).ToArray());
		}

		private static void AssertCollectionsHaveSameInterval(TwoClientHarness harness, string collectionName, string intervalId)
		{
			SequenceInterval? intervalA = harness.ClientA.GetIntervalCollection(collectionName).GetIntervalById(intervalId);
			SequenceInterval? intervalB = harness.ClientB.GetIntervalCollection(collectionName).GetIntervalById(intervalId);
			Assert.NotNull(intervalA);
			Assert.NotNull(intervalB);
			Assert.Equal(intervalA!.StartPosition, intervalB!.StartPosition);
			Assert.Equal(intervalA.EndPosition, intervalB.EndPosition);
		}

		private sealed class TwoClientHarness
		{
			private const string _clientAId = "client-a";
			private const string _clientBId = "client-b";
			private long _serverSeq;

			public TwoClientHarness()
			{
				SenderA = new FakeFluidDataObjectSender();
				SenderB = new FakeFluidDataObjectSender();
				ClientA = new SharedString("doc", SenderA);
				ClientB = new SharedString("doc", SenderB);
			}

			public SharedString ClientA { get; }

			public SharedString ClientB { get; }

			public FakeFluidDataObjectSender SenderA { get; }

			public FakeFluidDataObjectSender SenderB { get; }

			public long CurrentServerSeq => _serverSeq;

			public void DeliverAtoB(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				long refSeq)
			{
				Deliver(sent, fromClientId: _clientAId, from: ClientA, to: ClientB, refSeq: refSeq);
			}

			public void DeliverBtoA(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				long refSeq)
			{
				Deliver(sent, fromClientId: _clientBId, from: ClientB, to: ClientA, refSeq: refSeq);
			}

			private void Deliver(
				(string Address, string OpTypeName, string OpJson, long ClientSeq) sent,
				string fromClientId,
				SharedString from,
				SharedString to,
				long refSeq)
			{
				long serverSeq = ++_serverSeq;
				var descriptorForTo = new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: serverSeq),
					OpOrigin.Remote,
					fromClientId);
				to.ProcessDataObjectOp(descriptorForTo, sent.OpJson);

				var descriptorForFrom = new SequencedDocumentMessageDescriptor(
					SequenceNumber.ForTesting(clientSeq: sent.ClientSeq, refSeq: refSeq, seq: serverSeq),
					OpOrigin.Local,
					fromClientId);
				from.ProcessDataObjectOp(descriptorForFrom, sent.OpJson);
			}
		}
	}
}

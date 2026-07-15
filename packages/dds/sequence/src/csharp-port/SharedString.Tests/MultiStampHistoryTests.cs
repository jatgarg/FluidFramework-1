// -----------------------------------------------------------------------------
// Multi-stamp removal history tests for SharedString C# port.
// -----------------------------------------------------------------------------

#nullable enable

using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class MultiStampHistoryTests
	{
		[Fact]
		public void OverlappingRemove_TwoClients_KeepsBothStamps()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.MarkRangeRemoved(0, 6, refSeq: 1, seq: 10, clientId: "client-a");
			tree.MarkRangeRemoved(2, 8, refSeq: 1, seq: 11, clientId: "client-b");

			ISegment overlapped = SegmentWithText(tree, "llo ");
			Assert.Equal(new long[] { 10, 11 }, RemoveSeqs(overlapped));
			Assert.All(overlapped.RemoveStamps, stamp => Assert.IsType<SetRemoveOperationStamp>(stamp));
			Assert.Equal("rld", tree.GetText());
			Assert.Equal(5, tree.GetLength(10));
			Assert.Equal(3, tree.GetLength(11));
			// TS-parity: a flattened latest-stamp path would make "llo " visible at refSeq 10 (Finding M1/P4).
			Assert.Equal(9, FlattenedLengthAt(tree, refSeq: 10, useLatestStamp: true));
			Assert.Equal(3, FlattenedLengthAt(tree, refSeq: 11, useLatestStamp: false));
		}

		[Fact]
		public void OverlappingObliterate_SameSegment_KeepsList()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.ObliterateRange(0, 6, refSeq: 1, seq: 10, clientId: "client-a");
			tree.ObliterateRange(2, 8, refSeq: 1, seq: 11, clientId: "client-b");

			ISegment overlapped = SegmentWithText(tree, "llo ");
			Assert.Equal(new long[] { 10, 11 }, RemoveSeqs(overlapped));
			Assert.All(overlapped.RemoveStamps, stamp => Assert.IsType<SliceRemoveOperationStamp>(stamp));
			Assert.Equal("rld", tree.GetText());
		}

		[Fact]
		public void RemoveThenObliterate_SameRange_BothPreserved()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.MarkRangeRemoved(0, 6, refSeq: 1, seq: 10, clientId: "client-a");
			tree.ObliterateRange(0, 6, refSeq: 1, seq: 11, clientId: "client-b");

			ISegment removed = SegmentWithText(tree, "hello ");
			Assert.Equal(new[] { "setRemove", "sliceRemove" }, RemoveTypes(removed));
			Assert.Equal(new long[] { 10, 11 }, RemoveSeqs(removed));
			Assert.Equal("world", tree.GetText());
		}

		[Fact]
		public void ObliterateThenRemove_SameRange_BothPreserved()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.ObliterateRange(0, 6, refSeq: 1, seq: 10, clientId: "client-a");
			tree.MarkRangeRemoved(0, 6, refSeq: 1, seq: 11, clientId: "client-b");

			ISegment removed = SegmentWithText(tree, "hello ");
			Assert.Equal(new[] { "sliceRemove", "setRemove" }, RemoveTypes(removed));
			Assert.Equal(new long[] { 10, 11 }, RemoveSeqs(removed));
			Assert.Equal("world", tree.GetText());
		}

		[Fact]
		public void OverlappingRemove_PartialLength_CorrectAtEachPerspective()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");

			tree.MarkRangeRemoved(0, 6, refSeq: 1, seq: 10, clientId: "client-a");
			tree.MarkRangeRemoved(2, 8, refSeq: 1, seq: 11, clientId: "client-b");

			Assert.Equal(11, tree.GetLength(9));
			Assert.Equal(5, tree.GetLength(10));
			Assert.Equal(3, tree.GetLength(11));
			Assert.Equal(3, tree.GetLength(10, "client-b"));
			// TS-parity: client B's own overlapping remove must be counted even when refSeq is 10 (Finding P4).
			Assert.NotEqual(5, tree.GetLength(10, "client-b"));
		}

		[Fact]
		public void MultiStamp_ZamboniReclaimsOnlyWhenAllStampsAcked()
		{
			MergeTreeModel tree = CreateAckedTree("drop");
			tree.MarkRangeRemoved(0, 4, refSeq: 1, seq: 10, clientId: "client-a");
			tree.MarkRangeRemoved(0, 4, refSeq: 1, seq: 11, clientId: "client-b");
			tree.MarkRangeRemoved(0, 4, refSeq: 1, seq: 12, clientId: "client-c");

			tree.SetMinSeq(12);
			tree.ZamboniSegments();

			ISegment held = Assert.Single(tree.WalkAllSegments());
			Assert.Equal(new long[] { 10, 11, 12 }, RemoveSeqs(held));
			tree.SetMinSeq(13);
			tree.ZamboniSegments();
			Assert.Empty(tree.WalkAllSegments());
		}

		[Fact]
		public void ConcurrentInsertIntoRemovedThenObliterated_Traversal()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");
			tree.MarkRangeRemoved(0, 6, refSeq: 1, seq: 9, clientId: "client-a");
			tree.ObliterateRange(0, 6, refSeq: 1, seq: 10, clientId: "client-b");

			tree.InsertSegments(
				3,
				new ISegment[] { TextSegmentModel.Make("X") },
				refSeq: 5,
				seq: 11,
				clientId: "client-c");

			ISegment inserted = SegmentWithText(tree, "X");
			Assert.Contains(inserted.RemoveStamps, stamp => stamp is SliceRemoveOperationStamp && stamp.Seq == 10);
			Assert.Equal("world", tree.GetText());
		}

		[Fact]
		public void WireRoundTrip_MultiStampSegment_PreservesList()
		{
			string snapshotJson = """
				{
					"version":"1",
					"segmentCount":1,
					"length":5,
					"startIndex":0,
					"segments":[
						{
							"json":{"text":"hello"},
							"seq":1,
							"client":"seed",
							"removedSeq":10,
							"removedClientIds":["client-a","client-b"],
							"movedSeq":12,
							"movedSeqs":[12,13],
							"movedClientIds":["client-c","client-d"]
						}
					],
					"headerMetadata":{
						"minSequenceNumber":0,
						"sequenceNumber":20,
						"orderedChunkMetadata":[{"id":"header"}],
						"totalLength":5,
						"totalSegmentCount":1
					}
				}
				""";
			Client client = new("loader");

			SharedStringSnapshotLoader.PopulateFromSnapshot(client, SharedStringSnapshotLoader.Load(snapshotJson));

			ISegment segment = Assert.Single(client.MergeTree.WalkAllSegments());
			Assert.Equal(new[] { "setRemove", "setRemove", "sliceRemove", "sliceRemove" }, RemoveTypes(segment));
			Assert.Equal(new long[] { 10, 10, 12, 13 }, RemoveSeqs(segment));
			Assert.Equal(string.Empty, client.GetText(0, client.GetLength()));
		}

		[Fact]
		public void Ack_LocalStampInList_ReplacesStamp()
		{
			Client client = CreateAckedClient("client-a", "hello world");
			IMergeTreeRemoveMsg localRemove = client.RemoveText(0, 6);
			ISegment removed = SegmentWithText(client.MergeTree, "hello ");
			SetRemoveOperationStamp localStamp = Assert.Single(FindSetRemoveStamps(removed));
			long? localSeq = localStamp.LocalSeq;

			client.ApplyOp(new MergeTreeRemoveMsg() { Pos1 = 0, Pos2 = 6 }, seq: 10, refSeq: 1, clientId: "client-b");
			client.ApplyOp(localRemove, seq: 11, refSeq: 1, clientId: "client-a");

			RemoveOperationStamp promotedStamp = removed.RemoveStamps.Single(stamp => stamp.Seq == 11);
			Assert.NotSame(localStamp, promotedStamp);
			Assert.DoesNotContain(removed.RemoveStamps, stamp => ReferenceEquals(stamp, localStamp));
			Assert.Equal(MergeTreeModel.UnassignedSequenceNumber, localStamp.Seq);
			Assert.Equal(localSeq, localStamp.LocalSeq);
			Assert.Null(promotedStamp.LocalSeq);
			Assert.Contains(removed.RemoveStamps, stamp => stamp.Seq == 10 && stamp.ClientId != promotedStamp.ClientId);
			Assert.Equal(new long[] { 10, 11 }, RemoveSeqs(removed));
		}

		[Fact]
		public void MultiStamp_LocalReferenceSlide_UsesCorrectStamp()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");
			tree.MarkRangeRemoved(
				0,
				6,
				MergeTreeModel.UnassignedSequenceNumber,
				MergeTreeModel.UnassignedSequenceNumber,
				"client-a");
			ISegment removed = SegmentWithText(tree, "hello ");
			LocalReferencePosition reference = new(
				removed,
				offset: 1,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Forward);
			removed.LocalRefs.Add(reference);

			tree.MarkRangeRemoved(0, 6, refSeq: 1, seq: 10, clientId: "client-b");

			Assert.False(reference.IsDetached);
			Assert.Equal(0, tree.GetPositionOfReference(reference));
			Assert.Equal(new long[] { 10, MergeTreeModel.UnassignedSequenceNumber }, RemoveSeqs(removed));
		}

		private static MergeTreeModel CreateAckedTree(string text)
		{
			MergeTreeModel tree = new();
			tree.InsertSegments(
				0,
				new ISegment[] { TextSegmentModel.Make(text) },
				MergeTreeModel.UnassignedSequenceNumber,
				seq: 1,
				clientId: "seed");
			return tree;
		}

		private static Client CreateAckedClient(string clientId, string text)
		{
			Client client = new(clientId);
			IMergeTreeInsertMsg insert = client.InsertText(0, text);
			client.ApplyOp(insert, seq: 1, refSeq: 0, clientId: clientId);
			return client;
		}

		private static ISegment SegmentWithText(MergeTreeModel tree, string text)
		{
			return Assert.Single(
				tree.WalkAllSegments(),
				segment => segment is TextSegmentModel textSegment && textSegment.Text == text);
		}

		private static long[] RemoveSeqs(ISegment segment)
		{
			return segment.RemoveStamps.Select(stamp => stamp.Seq).ToArray();
		}

		private static string[] RemoveTypes(ISegment segment)
		{
			return segment.RemoveStamps.Select(stamp => stamp.Type).ToArray();
		}

		private static IEnumerable<SetRemoveOperationStamp> FindSetRemoveStamps(ISegment segment)
		{
			return segment.RemoveStamps.OfType<SetRemoveOperationStamp>();
		}

		private static int FlattenedLengthAt(MergeTreeModel tree, long refSeq, bool useLatestStamp)
		{
			int length = 0;
			foreach (ISegment segment in tree.WalkAllSegments())
			{
				if (segment.Seq == MergeTreeModel.UnassignedSequenceNumber || segment.Seq > refSeq)
				{
					continue;
				}

				RemoveOperationStamp? stamp = useLatestStamp
					? segment.RemoveStamps.LastOrDefault()
					: segment.RemoveStamps.FirstOrDefault();
				if (stamp is not null
					&& stamp.Seq != MergeTreeModel.UnassignedSequenceNumber
					&& stamp.Seq <= refSeq)
				{
					continue;
				}

				length += segment.CachedLength;
			}

			return length;
		}
	}
}

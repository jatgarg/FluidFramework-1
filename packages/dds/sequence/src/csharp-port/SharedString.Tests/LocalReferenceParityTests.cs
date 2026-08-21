// -----------------------------------------------------------------------------
// TS-parity regression tests for LocalReferencePosition sliding behavior.
// -----------------------------------------------------------------------------

#nullable enable

using Microsoft.Office.Web.Fluid.Intervals;
using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class LocalReferenceParityTests
	{
		[Fact]
		public void SlideOnRemove_Forward_MiddleRemovalSlidesToNextSegment()
		{
			MergeTreeModel tree = CreateTreeWithSegments("keep", "drop", "tail");
			ISegment droppedSegment = GetRequiredSegment(tree, 5);
			LocalReferencePosition reference = tree.CreateReferencePosition(
				droppedSegment,
				1,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Forward);

			tree.MarkRangeRemoved(4, 8, tree.CurrentSeq, seq: 10, clientId: "remote");

			// TS-parity: forward SlideOnRemove lands before tombstones on the next live segment (Finding L1).
			Assert.False(reference.IsDetached);
			Assert.Same(GetRequiredSegment(tree, 4), reference.Segment);
			Assert.Equal(0, reference.Offset);
			Assert.Equal(4, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void SlideOnRemove_Forward_RemoveToEndFallsBackToPreviousLastCharacter()
		{
			MergeTreeModel tree = CreateTreeWithSegments("keep", "drop");
			ISegment droppedSegment = GetRequiredSegment(tree, 5);
			LocalReferencePosition reference = tree.CreateReferencePosition(
				droppedSegment,
				1,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Forward);

			tree.MarkRangeRemoved(4, 8, tree.CurrentSeq, seq: 10, clientId: "remote");

			// TS-parity: fallback to previous segment uses its last character, not the segment end (Finding L8).
			Assert.False(reference.IsDetached);
			Assert.Equal(3, reference.Offset);
			Assert.Equal(3, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void SlideOnRemove_Forward_RemoveAllWithoutEndpointDetaches()
		{
			MergeTreeModel tree = CreateTreeWithSegments("left", "right");
			LocalReferencePosition reference = tree.CreateReferencePosition(
				GetRequiredSegment(tree, 5),
				1,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Forward);

			tree.MarkRangeRemoved(0, 9, tree.CurrentSeq, seq: 10, clientId: "remote");

			// TS-parity: non-endpoint references slide off an empty tree and detach (Finding L8).
			Assert.True(reference.IsDetached);
			Assert.Equal(MergeTreeModel.DetachedReferencePosition, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void SlideOnRemove_Forward_RemoveAllWithEndpointSlidesToEndSentinel()
		{
			MergeTreeModel tree = CreateTreeWithSegments("left", "right");
			LocalReferencePosition reference = tree.CreateReferencePosition(
				GetRequiredSegment(tree, 5),
				1,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Forward,
				canSlideToEndpoint: true);

			tree.MarkRangeRemoved(0, 9, tree.CurrentSeq, seq: 10, clientId: "remote");

			// TS-parity: canSlideToEndpoint keeps the reference attached to the end sentinel (Finding L8).
			Assert.False(reference.IsDetached);
			Assert.Null(reference.Segment);
			Assert.Equal(0, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void SlideOnRemove_Backward_MiddleRemovalSlidesToPreviousLastCharacter()
		{
			MergeTreeModel tree = CreateTreeWithSegments("abc", "X", "def");
			LocalReferencePosition reference = tree.CreateReferencePosition(
				GetRequiredSegment(tree, 3),
				0,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Backward);

			tree.MarkRangeRemoved(3, 4, tree.CurrentSeq, seq: 10, clientId: "remote");

			// TS-parity: backward SlideOnRemove lands after tombstones on the previous live segment (Finding L4).
			Assert.False(reference.IsDetached);
			Assert.Equal(2, reference.Offset);
			Assert.Equal(2, tree.GetPositionOfReference(reference));
		}

		[Theory]
		[InlineData(0, 1, 0)]
		[InlineData(1, 2, 1)]
		[InlineData(0, 2, 0)]
		public void StayOnRemove_LocalThenRemoteConcurrentDeleteRemainsOnRemovedSegment(
			int start,
			int end,
			int referencePosition)
		{
			MergeTreeModel tree = CreateTreeWithSegments("A", "B");
			ISegment originalSegment = GetRequiredSegment(tree, referencePosition);
			LocalReferencePosition reference = tree.CreateReferencePosition(
				originalSegment,
				0,
				ReferenceType.StayOnRemove,
				SlidingPreference.Forward);

			tree.MarkRangeRemoved(start, end, tree.CurrentSeq, MergeTreeModel.UnassignedSequenceNumber, "local");
			tree.MarkRangeRemoved(start, end, tree.CurrentSeq, seq: 10, clientId: "remote");

			// TS-parity: StayOnRemove is never slid or detached by concurrent delete cleanup (Finding L2).
			Assert.False(reference.IsDetached);
			Assert.Same(originalSegment, reference.Segment);
		}

		[Fact]
		public void TransientReference_CanBeCreatedOnRemovedSegmentWithOffset()
		{
			MergeTreeModel tree = CreateTreeWithSegments("AB", "CD");
			ISegment removedSegment = GetRequiredSegment(tree, 1);
			tree.MarkRangeRemoved(0, 2, tree.CurrentSeq, seq: 10, clientId: "remote");

			LocalReferencePosition reference = tree.CreateReferencePosition(
				removedSegment,
				1,
				ReferenceType.Transient);

			// TS-parity: transient references may point at removed segments and retain their offset (Finding L3).
			Assert.False(reference.IsDetached);
			Assert.Same(removedSegment, reference.Segment);
			Assert.Equal(1, reference.Offset);
		}

		[Fact]
		public void TransientReference_CreatedBeforeRemoveDetachesOnDirectRemove()
		{
			MergeTreeModel tree = CreateTreeWithSegments("A", "B", "C");
			LocalReferencePosition reference = tree.CreateReferencePosition(
				GetRequiredSegment(tree, 1),
				0,
				ReferenceType.Transient);

			tree.MarkRangeRemoved(1, 2, tree.CurrentSeq, seq: 10, clientId: "remote");

			// TS-parity: direct removal of a segment still detaches transient references used by transient intervals (Finding L3).
			Assert.True(reference.IsDetached);
			Assert.Equal(MergeTreeModel.DetachedReferencePosition, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void SplitSegment_MaintainsBoundaryReferencePreferences()
		{
			MergeTreeModel tree = CreateTreeWithSegments("abcdef");
			LocalReferencePosition forward = tree.CreateReferencePosition(
				3,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Forward);
			LocalReferencePosition backward = tree.CreateReferencePosition(
				3,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Backward);

			tree.InsertSegments(
				3,
				new ISegment[] { TextSegment.Make("X") },
				MergeTreeModel.UnassignedSequenceNumber,
				MergeTreeModel.UnassignedSequenceNumber,
				"local");

			// TS-parity: split keeps backward refs on the left and forward refs on the right (Finding L5).
			Assert.Equal(4, tree.GetPositionOfReference(forward));
			Assert.Equal(3, tree.GetPositionOfReference(backward));
		}

		[Fact]
		public void AppendSegment_MaintainsRightSegmentReferenceOffsets()
		{
			MergeTreeModel tree = CreateTreeWithSegments("abc", "def");
			LocalReferencePosition reference = tree.CreateReferencePosition(
				GetRequiredSegment(tree, 4),
				1,
				ReferenceType.SlideOnRemove);

			AdvanceMinSeq(tree, 10);
			tree.ZamboniSegments();

			// TS-parity: append offsets refs from the appended segment by the left segment length (Finding L5).
			Assert.Single(tree.WalkAllSegments());
			Assert.False(reference.IsDetached);
			Assert.Equal(4, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void RebasePendingIntervalAdd_PreservesBackwardSlidingPreference()
		{
			Client client = CreateClientWithAckedText("abcXdef");
			IntervalCollection collection = client.GetOrCreateIntervalCollection("comments", IntervalType.SlideOnRemove);
			collection.Add(3, 3, intervalId: "i1", stickiness: IntervalStickiness.Start);

			client.ApplyOp(
				new MergeTreeRemoveMsg()
				{
					Pos1 = 1,
					Pos2 = 4,
				},
				seq: 2,
				refSeq: 1,
				clientId: "remote");

			IMergeTreeOp rebased = Assert.Single(client.RebasePendingOps());
			IntervalAddOpMsg add = Assert.IsType<IntervalAddOpMsg>(rebased);

			// TS-parity: reconnect/rebase honors endpoint backward sliding preference (Finding L6).
			Assert.Equal(0, add.Start);
			Assert.Equal(0, add.End);
		}

		[Fact]
		public void ObliteratedSimpleReferenceReportsMinusOne()
		{
			MergeTreeModel tree = CreateTreeWithSegments("hello");
			LocalReferencePosition reference = tree.CreateReferencePosition(2, ReferenceType.Simple);

			tree.ObliterateRange(0, 5, tree.CurrentSeq, seq: 10, clientId: "remote");

			// TS-parity: detached/obliterated references report -1, not null (Finding L7).
			Assert.True(reference.IsDetached);
			Assert.Equal(MergeTreeModel.DetachedReferencePosition, tree.GetPositionOfReference(reference));
		}

		[Fact]
		public void SlideOnRemove_LocalUnackedInsert_NotChosenAsSlideTarget()
		{
			// V2-A10 regression. When a remote remove sequences over a
			// segment that holds a SlideOnRemove reference, TS picks the
			// slide target from the "all-acked including this op"
			// perspective — local unacked inserts are excluded so the
			// reference doesn't anchor to a segment that peers cannot see.
			// The port used to use current-view visibility and could pick a
			// local unacked segment as a permanent slide destination.
			Client client = new("local");

			// Two acked segments authored by the remote peer.
			IMergeTreeInsertMsg keep = new MergeTreeInsertMsg() { Pos1 = 0, Seg = "keep" };
			client.ApplyOp(keep, seq: 1, refSeq: 0, clientId: "peer");
			IMergeTreeInsertMsg drop = new MergeTreeInsertMsg() { Pos1 = 4, Seg = "drop" };
			client.ApplyOp(drop, seq: 2, refSeq: 1, clientId: "peer");

			// Reference on "drop" at offset 1.
			ISegment droppedSegment = client.MergeTree.GetContainingSegment(5, MergeTreeModel.UnassignedSequenceNumber).segment;
			LocalReferencePosition reference = client.MergeTree.CreateReferencePosition(
				droppedSegment,
				1,
				ReferenceType.SlideOnRemove,
				SlidingPreference.Forward);

			// Local UNACKED insert AFTER "drop". This is the segment that
			// would be a slide target under current-view visibility but
			// must be excluded under TS's all-acked perspective.
			client.InsertText(8, "LOCAL");

			// Remote remove of "drop" (positions 4..8 in the peer's view,
			// which matches our acked view since local unacked hasn't
			// affected acked positions).
			IMergeTreeRemoveMsg remove = new MergeTreeRemoveMsg() { Pos1 = 4, Pos2 = 8 };
			client.ApplyOp(remove, seq: 3, refSeq: 2, clientId: "peer");

			// After slide: reference must NOT land on the local unacked
			// segment. It should either detach (Forward slide had no
			// visible-in-acked-perspective successor) or land on an acked
			// segment (fallback direction: "keep").
			bool landsOnLocalUnacked = reference.Segment is TextSegment ts && ts.Text == "LOCAL";
			Assert.False(
				landsOnLocalUnacked,
				"V2-A10: SlideOnRemove picked a local unacked segment as target; the all-acked perspective should exclude it.");
		}

		private static MergeTreeModel CreateTreeWithSegments(params string[] texts)
		{
			MergeTreeModel tree = new();
			int position = 0;
			long seq = 1;
			foreach (string text in texts)
			{
				tree.InsertSegments(
					position,
					new ISegment[] { TextSegment.Make(text) },
					refSeq: seq - 1,
					seq: seq,
					clientId: "client");
				position += text.Length;
				seq++;
			}

			return tree;
		}

		private static Client CreateClientWithAckedText(string text)
		{
			Client client = new("local");
			IMergeTreeInsertMsg insert = client.InsertText(0, text);
			client.ApplyOp(insert, seq: 1, refSeq: 0, clientId: "local");
			return client;
		}

		private static ISegment GetRequiredSegment(MergeTreeModel tree, int position)
		{
			return tree.GetContainingSegment(position, MergeTreeModel.UnassignedSequenceNumber, "client").segment;
		}

		private static void AdvanceMinSeq(MergeTreeModel tree, long minSeq)
		{
			if (tree.CurrentSeq < minSeq)
			{
				tree.CurrentSeq = minSeq;
			}

			tree.SetMinSeq(minSeq);
		}
	}
}

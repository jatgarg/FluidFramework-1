#nullable enable

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using TextSegmentModel = Microsoft.Office.Web.Fluid.MergeTree.TextSegment;
using static Microsoft.Office.Web.Fluid.Tests.PortedTestUtilities;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class PartialLengthPerspectiveTestsFromTS
	{
		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "concurrent remote and unsequenced local changes are visible"
		[Fact]
		public void GetLength_LocalSeqIncludesPendingInsertOnlyAtOrBelowLocalSeq()
		{
			var (sharedString, _) = CreateAckedSharedString("localUser", "hello");

			sharedString.InsertText(5, "!");

			Assert.Equal(5, sharedString.GetLength(refSeq: 1));
			Assert.Equal(6, sharedString.GetLength(refSeq: 1, clientId: "localUser"));
			Assert.Equal(5, sharedString.GetLength(refSeq: 1, clientId: "localUser", localSeq: 1));
			Assert.Equal(6, sharedString.GetLength(refSeq: 1, clientId: "localUser", localSeq: 2));
		}

		// Ported from packages/dds/merge-tree/src/test/partialLength.spec.ts — "concurrent remote and unsequenced local changes are visible"
		[Fact]
		public void GetLength_LocalSeqIncludesPendingRemoveOnlyAtOrBelowLocalSeq()
		{
			var (sharedString, _) = CreateAckedSharedString("localUser", "hello");

			sharedString.DeleteText(0, 2);

			Assert.Equal(5, sharedString.GetLength(refSeq: 1));
			Assert.Equal(3, sharedString.GetLength(refSeq: 1, clientId: "localUser"));
			Assert.Equal(5, sharedString.GetLength(refSeq: 1, clientId: "localUser", localSeq: 1));
			Assert.Equal(3, sharedString.GetLength(refSeq: 1, clientId: "localUser", localSeq: 2));
		}

		// Ported from packages/dds/merge-tree/src/mergeTree.ts and perspective.ts — detached refSeq with a local reconnecting perspective
		[Fact]
		public void GetLength_DetachedPerspectiveOnlyIncludesLocalPendingWithinPerspective()
		{
			SharedString sharedString = new("localUser");

			sharedString.InsertText(0, "abc");

			Assert.Equal(0, sharedString.GetLength(refSeq: MergeTreeModel.UnassignedSequenceNumber));
			Assert.Equal(3, sharedString.GetLength(refSeq: MergeTreeModel.UnassignedSequenceNumber, clientId: "localUser"));
			Assert.Equal(0, sharedString.GetLength(refSeq: MergeTreeModel.UnassignedSequenceNumber, clientId: "localUser", localSeq: 0));
			Assert.Equal(3, sharedString.GetLength(refSeq: MergeTreeModel.UnassignedSequenceNumber, clientId: "localUser", localSeq: 1));
		}

		// Ported from packages/dds/merge-tree/src/test/client.rebasePosition.spec.ts — "rebase past remote insert"
		[Fact]
		public void GetPosition_RemoteInsertBeforeSegmentChangesByRefSeq()
		{
			var (sharedString, _) = CreateAckedSharedString("localUser", "hello");
			ISegment originalSegment = GetRequiredContainingSegment(sharedString, 0).Segment;

			ProcessRemoteInsert(sharedString, 0, "XY", refSeq: 1, seq: 2, clientId: "remoteUser");

			Assert.Equal(0, sharedString.GetPosition(originalSegment, refSeq: 1));
			Assert.Equal(2, sharedString.GetPosition(originalSegment, refSeq: 2));
		}

		// Ported from packages/dds/merge-tree/src/test/client.rebasePosition.spec.ts — local edits are ignored above the reconnecting localSeq
		[Fact]
		public void GetPosition_LocalPendingInsertBeforeSegmentRespectsLocalSeq()
		{
			var (sharedString, _) = CreateAckedSharedString("localUser", "ab");

			sharedString.InsertText(1, "X");
			ISegment bSegment = GetRequiredContainingSegment(sharedString, 2).Segment;

			Assert.Equal(1, sharedString.GetPosition(bSegment, refSeq: 1));
			Assert.Equal(1, sharedString.GetPosition(bSegment, refSeq: 1, clientId: "localUser", localSeq: 1));
			Assert.Equal(2, sharedString.GetPosition(bSegment, refSeq: 1, clientId: "localUser", localSeq: 2));
		}

		// Ported from packages/dds/merge-tree/src/test/client.rebasePosition.spec.ts — local removes are ignored above the reconnecting localSeq
		[Fact]
		public void GetPosition_LocalPendingRemoveBeforeSegmentRespectsLocalSeq()
		{
			var (sharedString, _) = CreateAckedSharedString("localUser", "abc");

			sharedString.DeleteText(0, 1);
			ISegment bSegment = GetRequiredContainingSegment(sharedString, 0).Segment;

			Assert.Equal(1, sharedString.GetPosition(bSegment, refSeq: 1));
			Assert.Equal(1, sharedString.GetPosition(bSegment, refSeq: 1, clientId: "localUser", localSeq: 1));
			Assert.Equal(0, sharedString.GetPosition(bSegment, refSeq: 1, clientId: "localUser", localSeq: 2));
		}

		// Ported from packages/dds/merge-tree/src/test/client.rebasePosition.spec.ts — "rebase past remote delete"
		[Fact]
		public void GetPosition_RemoteRemoveBeforeSegmentChangesByRefSeq()
		{
			var (sharedString, _) = CreateAckedSharedString("localUser", "abc");

			ProcessRemoteRemove(sharedString, 0, 1, refSeq: 1, seq: 2, clientId: "remoteUser");
			ISegment bSegment = GetRequiredContainingSegment(sharedString, 0).Segment;

			Assert.Equal(1, sharedString.GetPosition(bSegment, refSeq: 1));
			Assert.Equal(0, sharedString.GetPosition(bSegment, refSeq: 2));
		}

		// Ported from packages/dds/merge-tree/src/test/client.getPosition.spec.ts — "Deleted Segment"
		[Fact]
		public void GetPosition_DeletedSegmentReturnsTombstonePositionAtPerspective()
		{
			MergeTreeModel tree = CreateSingleCharacterTree("abc");
			ISegment bSegment = SegmentWithText(tree, "b");

			tree.MarkRangeRemoved(
				1,
				2,
				refSeq: 3,
				seq: MergeTreeModel.UnassignedSequenceNumber,
				clientId: "localUser");

			// SS-A15: TS returns the collapsed tombstone position (position 1
			// — where "b" used to sit, i.e., between "a" and "c"). Prior
			// port returned null; now matches TS 'Deleted Segment' spec.
			Assert.Equal(1, tree.GetPositionOfSegment(bSegment));
			Assert.Equal(1, tree.GetPositionOfSegmentAt(bSegment, refSeq: 3, clientId: "localUser", localSeq: null));
		}

		// Ported from packages/dds/merge-tree/src/test/client.getPosition.spec.ts — "Detached Segment"
		[Fact]
		public void GetPosition_DetachedSegmentReturnsMinusOne()
		{
			MergeTreeModel tree = CreateSingleCharacterTree("abc");
			ISegment detached = TextSegmentModel.Make("x");

			Assert.Equal(
				MergeTreeModel.DetachedReferencePosition,
				tree.GetPositionOfSegmentAt(detached, refSeq: 3, clientId: "localUser", localSeq: null));
		}
	}
}

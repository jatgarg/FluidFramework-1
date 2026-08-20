#nullable enable

using Microsoft.Office.Web.Fluid.MergeTree;
using Xunit;

using MergeTreeModel = Microsoft.Office.Web.Fluid.MergeTree.MergeTree;
using static Microsoft.Office.Web.Fluid.Tests.PortedTestUtilities;

namespace Microsoft.Office.Web.Fluid.Tests
{
	public sealed class MergeTreeMarkRangeRemovedTestsFromTS
	{
		// Ported from packages/dds/merge-tree/src/test/mergeTree.markRangeRemoved.spec.ts — "local remove followed by local insert"
		[Fact]
		public void LocalRemoveFollowedByLocalInsert_ShowsInsertedText()
		{
			SharedString sharedString = CreateSharedStringWithText("hello world");

			sharedString.DeleteText(0, sharedString.GetLength());
			sharedString.InsertText(0, "text");

			Assert.Equal("text", sharedString.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.markRangeRemoved.spec.ts — "remote remove followed by local insert"
		[Fact]
		public void RemoteRemoveFollowedByLocalInsert_ShowsInsertedText()
		{
			SharedString sharedString = CreateAckedSharedString("client-a", "hello world").SharedString;

			ProcessRemoteRemove(sharedString, 0, sharedString.GetLength(), refSeq: 1, seq: 2, clientId: "remote");
			sharedString.InsertText(0, "text");

			Assert.Equal("text", sharedString.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.markRangeRemoved.spec.ts — "local remove followed by remote insert"
		[Fact]
		public void LocalRemoveFollowedByRemoteInsert_ShowsRemoteInsertedText()
		{
			SharedString sharedString = CreateAckedSharedString("client-a", "hello world").SharedString;
			sharedString.DeleteText(0, sharedString.GetLength());

			ProcessRemoteInsert(sharedString, 0, "text", refSeq: 1, seq: 2, clientId: "remote");

			Assert.Equal("text", sharedString.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.markRangeRemoved.spec.ts — "local remove followed by remote overlapping remove"
		[Fact]
		public void LocalRemoveFollowedByRemoteOverlappingRemove_KeepsBothRemoveStamps()
		{
			MergeTreeModel tree = CreateAckedTree("hello world");
			ISegment segment = tree.GetContainingSegment(0, MergeTreeModel.UnassignedSequenceNumber).segment;

			tree.MarkRangeRemoved(0, 11, refSeq: MergeTreeModel.UnassignedSequenceNumber, seq: MergeTreeModel.UnassignedSequenceNumber, clientId: "local");
			tree.MarkRangeRemoved(0, 11, refSeq: 1, seq: 2, clientId: "remote");

			Assert.Equal(2, segment.RemoveStamps.Count);
			Assert.Contains(segment.RemoveStamps, stamp => stamp.Seq == MergeTreeModel.UnassignedSequenceNumber);
			Assert.Contains(segment.RemoveStamps, stamp => stamp.Seq == 2);
			Assert.Equal(string.Empty, tree.GetText());
		}

		// Ported from packages/dds/merge-tree/src/test/mergeTree.markRangeRemoved.spec.ts — "remote remove followed by remote insert"
		[Fact]
		public void RemoteRemoveFollowedByRemoteInsert_ShowsInsertedText()
		{
			SharedString sharedString = CreateAckedSharedString("client-a", "hello world").SharedString;

			ProcessRemoteRemove(sharedString, 0, sharedString.GetLength(), refSeq: 1, seq: 2, clientId: "remote-1");
			ProcessRemoteInsert(sharedString, 0, "text", refSeq: 1, seq: 3, clientId: "remote-2");

			Assert.Equal("text", sharedString.GetText());
		}
	}
}
